"""Shared data-preparation pipeline for the CLI commands.

One load -> clean -> sort flow used by both the Excel update and the Word
CPL generation, so the two outputs can never disagree on ordering or
sanitization.
"""

from __future__ import annotations

import logging
from pathlib import Path

from gsa_fcp_toolkit import config
from gsa_fcp_toolkit.data import DataHandler
from gsa_fcp_toolkit.models import LaborItem

logger = logging.getLogger(__name__)


def prepare_labor_data(json_path: str | Path | None = None) -> list[LaborItem]:
    """Load labor data from JSON and return it cleaned and UCID-sorted.

    Args:
        json_path: Source JSON file; defaults to ``config.LABOR_DATA_JSON``.

    Raises:
        FileNotFoundError: If the file does not exist.
        json.JSONDecodeError: If the file is not valid JSON.
        TypeError: If the JSON root is not a list.
    """
    path = Path(json_path) if json_path else config.LABOR_DATA_JSON
    handler = DataHandler()
    logger.info("Loading labor data from %s...", path)
    labor_data = handler.load_labor_data(path)
    logger.info("Applying sanitization and sorting rules to %d items...", len(labor_data))
    labor_data = handler.clean_labor_data(labor_data)
    return handler.sort_labor_data(labor_data)
