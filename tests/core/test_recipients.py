"""Tests for z7_officeletters.core.recipients."""

from __future__ import annotations

import pytest

from tests.conftest import make_dest_simples
from z7_officeletters.core.recipients import processar_destinatario


# =============================================================================
# processar_destinatario
# =============================================================================
class TestProcessarDestinatario:

    # --- Mayor rule ---
    def test_prefeito_por_flag(self) -> None:
        r = processar_destinatario(make_dest_simples(nome="Rafael", is_prefeito=True))
        assert r["vocativo"] == "Excelentíssimo Senhor Prefeito"
        assert r["pronome_corpo"] == "Vossa Excelência"
        assert r["envio"] == "Protocolo"
        assert r["destinatario_nome"] == "RAFAEL PIOVEZAN"

    def test_prefeito_por_nome(self) -> None:
        r = processar_destinatario(make_dest_simples(nome="o Prefeito Municipal"))
        assert r["pronome_corpo"] == "Vossa Excelência"

    def test_prefeito_endereco_fixo(self) -> None:
        r = processar_destinatario(make_dest_simples(is_prefeito=True))
        assert "Oeste/SP" in r["destinatario_endereco"]
        assert "Prefeito Municipal" in r["destinatario_endereco"]

    # --- Delivery method ---
    def test_envio_email_tem_prioridade_sobre_endereco(self) -> None:
        r = processar_destinatario(make_dest_simples(endereco="Rua X", email="x@y.com"))
        assert r["envio"] == "E-mail"

    def test_envio_carta_sem_email(self) -> None:
        r = processar_destinatario(make_dest_simples(endereco="Rua X"))
        assert r["envio"] == "Carta"

    def test_envio_em_maos_sem_contato(self) -> None:
        r = processar_destinatario(make_dest_simples())
        assert r["envio"] == "Em Mãos"

    # --- Formatting ---
    def test_nome_em_maiusculas(self) -> None:
        r = processar_destinatario(make_dest_simples(nome="João Silva"))
        assert r["destinatario_nome"] == "JOÃO SILVA"

    def test_pessoa_fisica_masculino_tratamento(self) -> None:
        r = processar_destinatario(make_dest_simples(genero="M"))
        assert r["tratamento_rodape"] == "Ao Ilustríssimo Senhor"

    def test_pessoa_fisica_feminino_tratamento(self) -> None:
        r = processar_destinatario(make_dest_simples(genero="F"))
        assert r["tratamento_rodape"] == "À Ilustríssima Senhora"

    def test_pronome_pessoa_fisica(self) -> None:
        r = processar_destinatario(make_dest_simples())
        assert r["pronome_corpo"] == "Vossa Senhoria"

    def test_instituicao_tratamento_ao(self) -> None:
        r = processar_destinatario(
            make_dest_simples(nome="Câmara Municipal", is_instituicao=True)
        )
        assert r["tratamento_rodape"] == "Ao"

    def test_instituicao_comeca_com_a_usa_crase(self) -> None:
        r = processar_destinatario(
            make_dest_simples(nome="ABCD Fundação", is_instituicao=True)
        )
        assert r["tratamento_rodape"] == "À"

    def test_instituicao_pronome_plural(self) -> None:
        r = processar_destinatario(
            make_dest_simples(nome="Câmara Municipal", is_instituicao=True)
        )
        assert r["pronome_corpo"] == "Vossas Senhorias"

    def test_instituicao_masculina_vocativo(self) -> None:
        r = processar_destinatario(
            make_dest_simples(nome="Câmara Municipal", is_instituicao=True, genero="M")
        )
        assert r["vocativo"] == "Ilustríssimos Senhores"

    def test_instituicao_feminina_vocativo(self) -> None:
        r = processar_destinatario(
            make_dest_simples(nome="Associação das Mães", is_instituicao=True, genero="F")
        )
        assert r["vocativo"] == "Ilustríssimas Senhoras"

    def test_pessoa_fisica_masculino_vocativo(self) -> None:
        r = processar_destinatario(make_dest_simples(genero="M"))
        assert r["vocativo"] == "Ilustríssimo Senhor"

    def test_pessoa_fisica_feminino_vocativo(self) -> None:
        r = processar_destinatario(make_dest_simples(nome="Maria Silva", genero="F"))
        assert r["vocativo"] == "Ilustríssima Senhora"

    def test_endereco_concatena_cargo_e_logradouro(self) -> None:
        r = processar_destinatario(make_dest_simples(
            cargo_ou_tratamento="Secretário de Saúde",
            endereco="Av. das Flores, 100",
        ))
        assert "Secretário de Saúde" in r["destinatario_endereco"]
        assert "Av. das Flores, 100" in r["destinatario_endereco"]

    def test_endereco_inclui_email(self) -> None:
        r = processar_destinatario(
            make_dest_simples(cargo_ou_tratamento="Diretor", email="d@e.com")
        )
        assert "d@e.com" in r["destinatario_endereco"]

    def test_honorifico_barra_cargo_remove_honorifico(self) -> None:
        r = processar_destinatario(
            make_dest_simples(cargo_ou_tratamento="Sr. / Ex-servidor")
        )
        assert r["destinatario_endereco"] == "Ex-servidor"

    def test_honorifico_barra_cargo_variante_sem_ponto(self) -> None:
        r = processar_destinatario(
            make_dest_simples(cargo_ou_tratamento="Sr / Diretor Geral")
        )
        assert r["destinatario_endereco"] == "Diretor Geral"
