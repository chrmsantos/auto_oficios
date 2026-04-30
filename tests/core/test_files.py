"""Tests for z7_officeletters.core.files."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

import z7_officeletters.core.files as _files_mod
from z7_officeletters.core.files import (
    ler_arquivo_mocoes,
    listar_proposituras,
    resolver_arquivo_preferencial,
)


# =============================================================================
# resolver_arquivo_preferencial
# =============================================================================
class TestResolverArquivoPreferencial:

    def test_retorna_proprio_quando_unico(self, tmp_path: Path) -> None:
        f = tmp_path / "mocoes.txt"
        f.write_text("c")
        assert resolver_arquivo_preferencial(str(f)) == str(f)

    def test_prefere_txt_sobre_docx(self, tmp_path: Path) -> None:
        (tmp_path / "mocoes.txt").write_text("c")
        docx = tmp_path / "mocoes.docx"
        docx.write_bytes(b"c")
        assert resolver_arquivo_preferencial(str(docx)) == str(tmp_path / "mocoes.txt")

    def test_prefere_docx_sobre_odt(self, tmp_path: Path) -> None:
        docx = tmp_path / "mocoes.docx"
        odt = tmp_path / "mocoes.odt"
        docx.write_bytes(b"c")
        odt.write_bytes(b"c")
        assert resolver_arquivo_preferencial(str(odt)) == str(docx)

    def test_prefere_odt_sobre_pdf(self, tmp_path: Path) -> None:
        odt = tmp_path / "mocoes.odt"
        pdf = tmp_path / "mocoes.pdf"
        odt.write_bytes(b"c")
        pdf.write_bytes(b"c")
        assert resolver_arquivo_preferencial(str(pdf)) == str(odt)

    def test_retorna_melhor_variante_sem_superior(self, tmp_path: Path) -> None:
        odt = tmp_path / "mocoes.odt"
        pdf = tmp_path / "mocoes.pdf"
        odt.write_bytes(b"c")
        pdf.write_bytes(b"c")
        assert resolver_arquivo_preferencial(str(pdf)) == str(odt)

    def test_retorna_original_sem_variante(self, tmp_path: Path) -> None:
        caminho = str(tmp_path / "naoexiste.pdf")
        assert resolver_arquivo_preferencial(caminho) == caminho

    def test_nao_cruza_diretorios(self, tmp_path: Path) -> None:
        sub = tmp_path / "sub"
        sub.mkdir()
        odt = sub / "arq.odt"
        odt.write_bytes(b"c")
        (tmp_path / "arq.txt").write_text("x")
        assert resolver_arquivo_preferencial(str(odt)) == str(odt)


# =============================================================================
# listar_proposituras
# =============================================================================
class TestListarProposituras:

    def _set_pasta(self, monkeypatch: pytest.MonkeyPatch, pasta: str) -> None:
        monkeypatch.setattr(_files_mod, "PASTA_PROPOSITURAS", pasta)

    def test_pasta_inexistente_retorna_vazio(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        self._set_pasta(monkeypatch, str(tmp_path / "nao_existe"))
        assert listar_proposituras() == []

    def test_pasta_vazia_retorna_vazio(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        self._set_pasta(monkeypatch, str(tmp_path))
        assert listar_proposituras() == []

    def test_arquivo_txt_retornado(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        self._set_pasta(monkeypatch, str(tmp_path))
        (tmp_path / "mocoes.txt").write_text("x")
        resultado = listar_proposituras()
        assert len(resultado) == 1 and resultado[0].suffix == ".txt"

    def test_formatos_nao_suportados_ignorados(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        self._set_pasta(monkeypatch, str(tmp_path))
        (tmp_path / "mocoes.txt").write_text("x")
        (tmp_path / "imagem.png").write_bytes(b"x")
        (tmp_path / "dados.csv").write_text("x")
        assert len(listar_proposituras()) == 1

    def test_multiplos_arquivos_distintos(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        self._set_pasta(monkeypatch, str(tmp_path))
        (tmp_path / "mocoes_marco.txt").write_text("x")
        (tmp_path / "mocoes_abril.docx").write_bytes(b"x")
        assert len(listar_proposituras()) == 2

    def test_duplicata_retorna_apenas_preferencial(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        self._set_pasta(monkeypatch, str(tmp_path))
        (tmp_path / "mocoes.txt").write_text("x")
        (tmp_path / "mocoes.docx").write_bytes(b"x")
        (tmp_path / "mocoes.pdf").write_bytes(b"x")
        resultado = listar_proposituras()
        assert len(resultado) == 1 and resultado[0].suffix == ".txt"

    def test_duplicata_prefere_docx_sem_txt(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        self._set_pasta(monkeypatch, str(tmp_path))
        (tmp_path / "mocoes.docx").write_bytes(b"x")
        (tmp_path / "mocoes.odt").write_bytes(b"x")
        resultado = listar_proposituras()
        assert len(resultado) == 1 and resultado[0].suffix == ".docx"

    def test_lista_ordenada_alfabeticamente(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        self._set_pasta(monkeypatch, str(tmp_path))
        (tmp_path / "z_ultimo.txt").write_text("x")
        (tmp_path / "a_primeiro.txt").write_text("x")
        (tmp_path / "m_meio.txt").write_text("x")
        nomes = [p.name for p in listar_proposituras()]
        assert nomes == sorted(nomes)

    def test_gitkeep_ignorado(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        self._set_pasta(monkeypatch, str(tmp_path))
        (tmp_path / ".gitkeep").write_bytes(b"")
        (tmp_path / "mocoes.txt").write_text("x")
        assert len(listar_proposituras()) == 1


# =============================================================================
# ler_arquivo_mocoes
# =============================================================================
class TestLerArquivoMocoes:

    def test_le_txt_utf8(self, tmp_path: Path) -> None:
        f = tmp_path / "mocoes.txt"
        f.write_text("MOÇÃO Nº 1\nTexto.", encoding="utf-8")
        assert ler_arquivo_mocoes(str(f)) == "MOÇÃO Nº 1\nTexto."

    def test_txt_preserva_conteudo_completo(self, tmp_path: Path) -> None:
        conteudo = "MOÇÃO Nº 1\n\nMOÇÃO Nº 2\nSegundo texto."
        f = tmp_path / "mocoes.txt"
        f.write_text(conteudo, encoding="utf-8")
        assert ler_arquivo_mocoes(str(f)) == conteudo

    def test_formato_invalido_levanta_value_error(self, tmp_path: Path) -> None:
        f = tmp_path / "mocoes.xyz"
        f.write_text("x")
        with pytest.raises(ValueError, match="suportado"):
            ler_arquivo_mocoes(str(f))

    def test_le_docx_via_mock(self, tmp_path: Path) -> None:
        mock_doc = MagicMock()
        mock_doc.paragraphs = [
            MagicMock(text="MOÇÃO Nº 1"),
            MagicMock(text="Texto."),
        ]
        with patch("docx.Document", return_value=mock_doc):
            resultado = ler_arquivo_mocoes(str(tmp_path / "mocoes.docx"))
        assert resultado == "MOÇÃO Nº 1\nTexto."

    def test_pdf_sem_pypdf_levanta_import_error(self, tmp_path: Path) -> None:
        f = tmp_path / "mocoes.pdf"
        f.write_bytes(b"%PDF-1.4")
        with patch.dict("sys.modules", {"pypdf": None}):
            with pytest.raises(ImportError, match="pypdf"):
                ler_arquivo_mocoes(str(f))
