"""Tests for z7_officeletters.core.ai."""

from __future__ import annotations

import json
import logging
from typing import Any
from unittest.mock import MagicMock, patch

import pytest

from tests.conftest import make_ai_response, make_dados_mocao_validos
from z7_officeletters.core.ai import (
    extrair_dados_com_ia,
    limpar_json_da_resposta,
    validar_dados_mocao,
)


# =============================================================================
# limpar_json_da_resposta
# =============================================================================
class TestLimparJsonDaResposta:

    def test_json_puro_sem_marcador(self) -> None:
        assert limpar_json_da_resposta('{"k": 1}') == '{"k": 1}'

    def test_marcador_json(self) -> None:
        texto = '```json\n{"tipo_mocao": "Aplauso"}\n```'
        assert limpar_json_da_resposta(texto) == '{"tipo_mocao": "Aplauso"}'

    def test_marcador_generico(self) -> None:
        texto = '```\n{"tipo_mocao": "Apelo"}\n```'
        assert limpar_json_da_resposta(texto) == '{"tipo_mocao": "Apelo"}'

    def test_espacos_e_quebras_extras(self) -> None:
        texto = '  \n```json\n{"a": 1}\n```\n  '
        assert limpar_json_da_resposta(texto) == '{"a": 1}'

    def test_resultado_e_json_valido_apos_limpeza(self) -> None:
        texto = '```json\n{"tipo_mocao": "Aplauso", "numero_mocao": "42"}\n```'
        assert json.loads(limpar_json_da_resposta(texto))["numero_mocao"] == "42"

    def test_json_array_retornado_intacto(self) -> None:
        texto = '```json\n[{"tipo_mocao": "Apelo"}]\n```'
        parsed = json.loads(limpar_json_da_resposta(texto))
        assert isinstance(parsed, list) and parsed[0]["tipo_mocao"] == "Apelo"


# =============================================================================
# validar_dados_mocao
# =============================================================================
class TestValidarDadosMocao:

    def test_dados_completos_nao_levanta(self) -> None:
        validar_dados_mocao(make_dados_mocao_validos())

    def test_tipo_apelo_valido(self) -> None:
        validar_dados_mocao(make_dados_mocao_validos(tipo_mocao="Apelo"))

    def test_tipo_invalido(self) -> None:
        with pytest.raises(ValueError, match="tipo_mocao"):
            validar_dados_mocao(make_dados_mocao_validos(tipo_mocao="Homenagem"))

    @pytest.mark.parametrize("campo", ["tipo_mocao", "numero_mocao", "autores", "destinatarios"])
    def test_campo_ausente_levanta(self, campo: str) -> None:
        d = make_dados_mocao_validos()
        del d[campo]
        with pytest.raises(ValueError):
            validar_dados_mocao(d)

    def test_autores_lista_vazia(self) -> None:
        with pytest.raises(ValueError):
            validar_dados_mocao(make_dados_mocao_validos(autores=[]))

    def test_destinatarios_lista_vazia(self) -> None:
        with pytest.raises(ValueError):
            validar_dados_mocao(make_dados_mocao_validos(destinatarios=[]))

    def test_autores_nao_e_lista(self) -> None:
        with pytest.raises(ValueError, match="lista"):
            validar_dados_mocao(make_dados_mocao_validos(autores="Alex Dantas"))

    def test_destinatarios_nao_e_lista(self) -> None:
        with pytest.raises(ValueError, match="lista"):
            validar_dados_mocao(make_dados_mocao_validos(destinatarios={"nome": "X"}))

    def test_destinatario_sem_nome(self) -> None:
        with pytest.raises(ValueError, match="nome"):
            validar_dados_mocao(make_dados_mocao_validos(destinatarios=[{"nome": ""}]))

    def test_destinatario_sem_chave_nome(self) -> None:
        with pytest.raises(ValueError, match="nome"):
            validar_dados_mocao(make_dados_mocao_validos(destinatarios=[{"cargo": "X"}]))

    def test_multiplos_destinatarios_validos(self) -> None:
        validar_dados_mocao(make_dados_mocao_validos(
            destinatarios=[{"nome": "Fulano"}, {"nome": "Ciclano"}]
        ))

    def test_segundo_destinatario_sem_nome_levanta(self) -> None:
        with pytest.raises(ValueError):
            validar_dados_mocao(make_dados_mocao_validos(
                destinatarios=[{"nome": "Fulano"}, {"nome": ""}]
            ))


# =============================================================================
# extrair_dados_com_ia
# =============================================================================
class TestExtrairDadosComIA:
    """All Gemini calls are mocked."""

    def _client(self, *responses: Any) -> MagicMock:
        client = MagicMock()
        client.models.generate_content.side_effect = list(responses)
        return client

    # --- Happy path ---
    def test_retorna_dados_validos_na_primeira_tentativa(self) -> None:
        client = self._client(make_ai_response(make_dados_mocao_validos()))
        r = extrair_dados_com_ia("MOÇÃO Nº 1 texto", client)
        assert r["tipo_mocao"] == "Aplauso"
        assert r["numero_mocao"] == "123"

    def test_aceita_resposta_com_markdown_json_fence(self) -> None:
        payload = make_dados_mocao_validos()
        mock_resp = MagicMock()
        mock_resp.text = f"```json\n{json.dumps(payload)}\n```"
        client = self._client(mock_resp)
        assert extrair_dados_com_ia("MOÇÃO", client)["tipo_mocao"] == "Aplauso"

    def test_aceita_resposta_em_lista(self) -> None:
        mock_resp = MagicMock()
        mock_resp.text = json.dumps([make_dados_mocao_validos()])
        client = self._client(mock_resp)
        assert extrair_dados_com_ia("MOÇÃO", client)["tipo_mocao"] == "Aplauso"

    # --- Retry on invalid response ---
    def test_retenta_quando_json_invalido(self) -> None:
        bad = MagicMock()
        bad.text = "não é json"
        good = make_ai_response(make_dados_mocao_validos())
        client = self._client(bad, good)
        r = extrair_dados_com_ia("MOÇÃO", client)
        assert r["tipo_mocao"] == "Aplauso"
        assert client.models.generate_content.call_count == 2

    def test_retenta_quando_tipo_mocao_invalido(self) -> None:
        bad = make_ai_response(make_dados_mocao_validos(tipo_mocao="Homenagem"))
        good = make_ai_response(make_dados_mocao_validos())
        client = self._client(bad, good)
        r = extrair_dados_com_ia("MOÇÃO", client)
        assert r["tipo_mocao"] == "Aplauso"
        assert client.models.generate_content.call_count == 2

    def test_levanta_apos_todas_as_tentativas_falharem(self) -> None:
        bad = MagicMock()
        bad.text = "não é json"
        client = self._client(*([bad] * 5))
        with pytest.raises((ValueError, json.JSONDecodeError)):
            extrair_dados_com_ia("MOÇÃO", client)
        assert client.models.generate_content.call_count == 5

    # --- Rate limit (429) ---
    def test_aguarda_e_retenta_em_rate_limit(self) -> None:
        good = make_ai_response(make_dados_mocao_validos())
        client = self._client(
            Exception("Error 429: retry_delay { seconds: 1 }"),
            good,
        )
        with patch("time.sleep") as mock_sleep:
            r = extrair_dados_com_ia("MOÇÃO", client)
        mock_sleep.assert_called_once()
        assert r["tipo_mocao"] == "Aplauso"

    def test_extrai_espera_do_retry_delay(self) -> None:
        good = make_ai_response(make_dados_mocao_validos())
        client = self._client(
            Exception("Error 429: retry_delay { seconds: 10 }"),
            good,
        )
        with patch("time.sleep") as mock_sleep:
            extrair_dados_com_ia("MOÇÃO", client)
        mock_sleep.assert_called_once_with(12)

    def test_erro_nao_429_relancado_imediatamente(self) -> None:
        client = self._client(ConnectionError("falha de rede"))
        with pytest.raises(ConnectionError):
            extrair_dados_com_ia("MOÇÃO", client)
        assert client.models.generate_content.call_count == 1

    # --- Logging ---
    def test_loga_resposta_bruta_em_debug(self, caplog: pytest.LogCaptureFixture) -> None:
        client = self._client(make_ai_response(make_dados_mocao_validos()))
        with caplog.at_level(logging.DEBUG, logger="z7_officeletters"):
            extrair_dados_com_ia("MOÇÃO", client)
        assert any("Resposta bruta" in r.message for r in caplog.records)

    def test_loga_warning_em_resposta_invalida(self, caplog: pytest.LogCaptureFixture) -> None:
        bad = MagicMock()
        bad.text = "não é json"
        good = make_ai_response(make_dados_mocao_validos())
        client = self._client(bad, good)
        with caplog.at_level(logging.WARNING, logger="z7_officeletters"):
            extrair_dados_com_ia("MOÇÃO", client)
        assert any(r.levelno == logging.WARNING for r in caplog.records)

    def test_resposta_longa_nao_aparece_inteira_no_log(
        self, caplog: pytest.LogCaptureFixture
    ) -> None:
        long_text = "x" * 2000
        bad = MagicMock()
        bad.text = long_text
        good = make_ai_response(make_dados_mocao_validos())
        client = self._client(bad, good)
        with caplog.at_level(logging.DEBUG, logger="z7_officeletters"):
            extrair_dados_com_ia("MOÇÃO", client)
        bruta_msgs = [r.message for r in caplog.records if "Resposta bruta" in r.message]
        assert all(len(m) < len(long_text) for m in bruta_msgs)
