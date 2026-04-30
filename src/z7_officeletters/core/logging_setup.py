"""Logging configuration for Z7 OfficeLetters.

Provides a single ``configurar_logging()`` call that sets up rotating file
handlers, a console handler, and a session-scoped unique identifier included
in every log record.

Public exports:
    SESSAO_ID: Short random hex string that identifies the current process run.
    logger: Module-level logger (name ``z7_officeletters``).
    configurar_logging: Configures all handlers and returns the log file path.
"""

from __future__ import annotations

import logging
import sys
import types
import uuid
from logging.handlers import RotatingFileHandler
from pathlib import Path

from z7_officeletters.constants import PASTA_LOGS

__all__ = ["SESSAO_ID", "logger", "configurar_logging"]

# Unique identifier for the current process run, embedded in every log line.
SESSAO_ID: str = uuid.uuid4().hex[:8]

logger: logging.Logger = logging.getLogger("z7_officeletters")


def configurar_logging(verbose: bool = False) -> str:
    """Configure rotating file and console log handlers.

    Sets up:
    - A ``RotatingFileHandler`` (DEBUG level, 2 MB, 5 backups) in
      ``PASTA_LOGS``.
    - A ``StreamHandler`` (WARNING by default; INFO when *verbose* is True).
    - A ``sys.excepthook`` that captures unhandled exceptions into the log.

    Calling this function multiple times is safe: existing handlers are cleared
    before new ones are added.

    Args:
        verbose: When True the console handler is raised to INFO level.

    Returns:
        Absolute path of the log file created in this session.
    """
    # Prevent handler accumulation on repeated calls (e.g., during testing).
    logger.handlers.clear()

    Path(PASTA_LOGS).mkdir(parents=True, exist_ok=True)

    from datetime import datetime

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = str(Path(PASTA_LOGS) / f"z7_officeletters_{timestamp}_{SESSAO_ID}.log")

    fmt = logging.Formatter(
        f"%(asctime)s [{SESSAO_ID}] [%(levelname)-8s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    file_handler = RotatingFileHandler(
        log_path,
        encoding="utf-8",
        maxBytes=2 * 1024 * 1024,  # 2 MB per file
        backupCount=5,
    )
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(fmt)

    console_level = logging.INFO if verbose else logging.WARNING
    console_handler = logging.StreamHandler()
    console_handler.setLevel(console_level)
    console_handler.setFormatter(logging.Formatter("[%(levelname)s] %(message)s"))

    logger.setLevel(logging.DEBUG)
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    def _excepthook(
        exc_type: type[BaseException],
        exc_value: BaseException,
        exc_tb: types.TracebackType | None,
    ) -> None:
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_tb)
            return
        logger.critical(
            "Exceção não tratada — o processo será encerrado.",
            exc_info=(exc_type, exc_value, exc_tb),
        )

    sys.excepthook = _excepthook  # type: ignore[assignment]
    logger.debug("Sessão de log iniciada. ID=%s", SESSAO_ID)
    return log_path
