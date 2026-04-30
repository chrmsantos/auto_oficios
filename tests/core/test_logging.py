"""Tests for z7_officeletters.core.logging_setup."""

from __future__ import annotations

import logging
import sys
from logging.handlers import RotatingFileHandler
from pathlib import Path
from unittest.mock import patch

import pytest

import z7_officeletters.core.logging_setup as _ls_mod
from z7_officeletters.core.logging_setup import SESSAO_ID, configurar_logging, logger


# =============================================================================
# configurar_logging
# =============================================================================
class TestConfigurarLogging:

    def setup_method(self) -> None:
        logger.handlers.clear()

    def teardown_method(self) -> None:
        logger.handlers.clear()
        sys.excepthook = sys.__excepthook__

    def test_cria_arquivo_de_log(self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setattr(_ls_mod, "PASTA_LOGS", str(tmp_path))
        assert Path(configurar_logging()).exists()

    def test_nome_arquivo_contem_sessao_id(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setattr(_ls_mod, "PASTA_LOGS", str(tmp_path))
        assert SESSAO_ID in configurar_logging()

    def test_usa_rotating_file_handler(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setattr(_ls_mod, "PASTA_LOGS", str(tmp_path))
        configurar_logging()
        rfhs = [h for h in logger.handlers if isinstance(h, RotatingFileHandler)]
        assert len(rfhs) == 1

    def test_rotating_handler_max_bytes(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setattr(_ls_mod, "PASTA_LOGS", str(tmp_path))
        configurar_logging()
        fh = next(h for h in logger.handlers if isinstance(h, RotatingFileHandler))
        assert fh.maxBytes == 2 * 1024 * 1024

    def test_rotating_handler_backup_count(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setattr(_ls_mod, "PASTA_LOGS", str(tmp_path))
        configurar_logging()
        fh = next(h for h in logger.handlers if isinstance(h, RotatingFileHandler))
        assert fh.backupCount == 5

    def test_console_level_warning_por_padrao(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setattr(_ls_mod, "PASTA_LOGS", str(tmp_path))
        configurar_logging(verbose=False)
        stream_hs = [
            h for h in logger.handlers
            if isinstance(h, logging.StreamHandler) and not isinstance(h, RotatingFileHandler)
        ]
        assert stream_hs[0].level == logging.WARNING

    def test_console_level_info_quando_verbose(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setattr(_ls_mod, "PASTA_LOGS", str(tmp_path))
        configurar_logging(verbose=True)
        stream_hs = [
            h for h in logger.handlers
            if isinstance(h, logging.StreamHandler) and not isinstance(h, RotatingFileHandler)
        ]
        assert stream_hs[0].level == logging.INFO

    def test_chamadas_repetidas_nao_duplicam_handlers(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setattr(_ls_mod, "PASTA_LOGS", str(tmp_path))
        configurar_logging()
        configurar_logging()
        assert len(logger.handlers) == 2  # 1 file + 1 console

    def test_excepthook_instalado(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setattr(_ls_mod, "PASTA_LOGS", str(tmp_path))
        configurar_logging()
        assert sys.excepthook is not sys.__excepthook__

    def test_excepthook_delega_keyboard_interrupt(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setattr(_ls_mod, "PASTA_LOGS", str(tmp_path))
        configurar_logging()
        with patch("sys.__excepthook__") as mock_orig:
            sys.excepthook(KeyboardInterrupt, KeyboardInterrupt(), None)
            mock_orig.assert_called_once()

    def test_excepthook_loga_excecao_nao_tratada(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setattr(_ls_mod, "PASTA_LOGS", str(tmp_path))
        configurar_logging()
        with patch.object(logger, "critical") as mock_crit:
            try:
                raise RuntimeError("erro de teste")
            except RuntimeError:
                sys.excepthook(*sys.exc_info())
            mock_crit.assert_called_once()

    def test_mensagem_debug_gravada_no_arquivo(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setattr(_ls_mod, "PASTA_LOGS", str(tmp_path))
        log_path = configurar_logging()
        logger.debug("mensagem-debug-xyz")
        assert "mensagem-debug-xyz" in Path(log_path).read_text(encoding="utf-8")

    def test_sessao_id_aparece_nas_linhas_de_log(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setattr(_ls_mod, "PASTA_LOGS", str(tmp_path))
        log_path = configurar_logging()
        logger.info("linha qualquer")
        assert SESSAO_ID in Path(log_path).read_text(encoding="utf-8")
