"""Tests for z7_officeletters.core.documents."""

from __future__ import annotations

import pytest

from z7_officeletters.core.documents import construir_nome_arquivo, normalizar_numero_mocao


# =============================================================================
# normalizar_numero_mocao
# =============================================================================
class TestNormalizarNumeroMocao:

    def test_numero_puro_nao_alterado(self) -> None:
        assert normalizar_numero_mocao("124") == "124"

    def test_remove_espacos(self) -> None:
        assert normalizar_numero_mocao("  124  ") == "124"

    def test_sufixo_nao_numerico_preservado(self) -> None:
        assert normalizar_numero_mocao("124-A") == "124-A"

    @pytest.mark.parametrize("entrada,esperado", [
        ("124/2026", "124"),
        ("124-2026", "124"),
        ("124/26",   "124"),
        ("124-26",   "124"),
        ("001/2026", "001"),
        ("999-2025", "999"),
        ("7",        "7"),
    ])
    def test_variantes(self, entrada: str, esperado: str) -> None:
        assert normalizar_numero_mocao(entrada) == esperado


# =============================================================================
# construir_nome_arquivo
# =============================================================================
class TestConstruirNomeArquivo:

    def _nome(self, **overrides: object) -> str:
        params: dict[str, object] = dict(
            num_oficio_str="001",
            sigla_servidor="js",
            tipo_mocao="Aplauso",
            num_mocao="124",
            envio="E-mail",
            nome_dest="Fulano de Tal",
            sigla_autores="AD",
            ano=2026,
        )
        params.update(overrides)
        return construir_nome_arquivo(**params)  # type: ignore[arg-type]

    def test_extensao_docx(self) -> None:
        assert self._nome().endswith(".docx")

    def test_contem_numero_oficio(self) -> None:
        assert "001" in self._nome()

    def test_contem_tipo_mocao(self) -> None:
        assert "Aplauso" in self._nome()

    def test_contem_numero_mocao_com_sufixo_26(self) -> None:
        assert "124-26" in self._nome()

    def test_sufixo_26_aparece_uma_vez(self) -> None:
        assert self._nome().count("-26") == 1

    def test_envio_convertido_para_minusculo(self) -> None:
        assert "e-mail" in self._nome(envio="E-mail")

    def test_sigla_servidor_refletida(self) -> None:
        assert "redator" in self._nome(sigla_servidor="redator")

    def test_sigla_autores_refletida(self) -> None:
        assert "ad e outros" in self._nome(sigla_autores="ad e outros")

    def test_remove_caracteres_invalidos_windows(self) -> None:
        nome = construir_nome_arquivo(
            num_oficio_str="001",
            sigla_servidor="js",
            tipo_mocao="Aplauso",
            num_mocao="124",
            envio="Em Mãos",
            nome_dest='Nome "Ilegal" <teste>',
            sigla_autores="AD",
            ano=2026,
        )
        for ch in r'\/*?:"<>|':
            assert ch not in nome
