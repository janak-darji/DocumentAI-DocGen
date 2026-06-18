"""Application logging configuration."""

from __future__ import annotations

import logging
import os
from logging.handlers import RotatingFileHandler
from pathlib import Path

SERVICE_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_ERROR_LOG_FILE = SERVICE_ROOT / "logs" / "errors.log"

logger = logging.getLogger("docgen-service")

_LOG_FORMAT = "%(asctime)s %(levelname)s [%(name)s] %(message)s"
_CONFIGURED = False


def _resolve_error_log_path() -> Path:
    """Resolves the error log file path from env or the default location."""
    configured = os.getenv("ERROR_LOG_FILE", "").strip()
    if not configured:
        return DEFAULT_ERROR_LOG_FILE

    path = Path(configured)
    if not path.is_absolute():
        path = SERVICE_ROOT / path
    return path


def configure_logging() -> Path:
    """
    Configures console and file logging for docgen-service.

    ERROR and above (including exception tracebacks) are written to a rotating
    log file. INFO and above continue to go to the console.

    Returns:
        Path to the error log file.
    """
    global _CONFIGURED
    if _CONFIGURED:
        return _resolve_error_log_path()

    level_name = os.getenv("LOG_LEVEL", "INFO").upper()
    level = getattr(logging, level_name, logging.INFO)

    logger.setLevel(level)
    logger.propagate = False
    logger.handlers.clear()

    console_handler = logging.StreamHandler()
    console_handler.setLevel(level)
    console_handler.setFormatter(logging.Formatter(_LOG_FORMAT))
    logger.addHandler(console_handler)

    error_log_path = _resolve_error_log_path()
    error_log_path.parent.mkdir(parents=True, exist_ok=True)

    max_bytes = int(os.getenv("ERROR_LOG_MAX_BYTES", str(10 * 1024 * 1024)))
    backup_count = int(os.getenv("ERROR_LOG_BACKUP_COUNT", "5"))

    file_handler = RotatingFileHandler(
        error_log_path,
        maxBytes=max_bytes,
        backupCount=backup_count,
        encoding="utf-8",
    )
    file_handler.setLevel(logging.ERROR)
    file_handler.setFormatter(logging.Formatter(_LOG_FORMAT))
    logger.addHandler(file_handler)

    logger.info("Detailed exception logs will be written to %s", error_log_path)

    _CONFIGURED = True
    return error_log_path
