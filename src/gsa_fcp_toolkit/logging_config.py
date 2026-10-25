"""Centralized logging configuration for gsa_fcp_toolkit.

Contract:
- Library modules must only use ``logging.getLogger(__name__)`` and never
  add handlers or call ``basicConfig`` themselves (import-time logging setup
  is a side effect and is forbidden).
- The CLI entry point calls :func:`configure_logging` exactly once at startup.

Design: the console shows the requested level (default INFO) while the
rotating file handler always captures DEBUG, giving a complete audit trail
on disk — timestamps and levels are required for auditing.
"""

from __future__ import annotations

import logging
import sys
from logging.handlers import RotatingFileHandler
from pathlib import Path

from gsa_fcp_toolkit import config

LOG_FORMAT = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
DATE_FORMAT = "%Y-%m-%d %H:%M:%S"
LOG_FILENAME = "gsa_fcp_toolkit.log"
MAX_LOG_BYTES = 1_000_000
LOG_BACKUP_COUNT = 3

_configured = False


def _log_file_path(log_dir: Path | str | None) -> Path:
    """Resolve the log file path, defaulting to the configured output dir."""
    base = Path(log_dir) if log_dir else config.OUTPUT_DIR
    return base / LOG_FILENAME


def configure_logging(
    level: int = logging.INFO,
    log_dir: Path | str | None = None,
    *,
    force: bool = False,
) -> Path:
    """Configure root logging with console and rotating-file handlers.

    Safe to call multiple times: repeat calls are no-ops unless ``force`` is
    true. Any handlers left by an earlier ``basicConfig`` call are removed so
    this module owns the root logger configuration.

    Args:
        level: Threshold for console output (default INFO). The file handler
            always records DEBUG and above.
        log_dir: Directory for the log file; defaults to ``config.OUTPUT_DIR``.
        force: Rebuild the handler set even if already configured.

    Returns:
        The path of the log file in use.
    """
    global _configured
    if _configured and not force:
        return _log_file_path(log_dir)

    root = logging.getLogger()
    root.setLevel(logging.DEBUG)  # let handlers filter; file keeps everything
    for handler in list(root.handlers):
        root.removeHandler(handler)

    formatter = logging.Formatter(LOG_FORMAT, datefmt=DATE_FORMAT)

    console = logging.StreamHandler(sys.stdout)
    console.setLevel(level)
    console.setFormatter(formatter)
    root.addHandler(console)

    log_file = _log_file_path(log_dir)
    log_file.parent.mkdir(parents=True, exist_ok=True)
    file_handler = RotatingFileHandler(
        log_file, maxBytes=MAX_LOG_BYTES, backupCount=LOG_BACKUP_COUNT, encoding="utf-8"
    )
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)
    root.addHandler(file_handler)

    _configured = True
    return log_file
