"""Ingestion, sanitization, and ordering of GSA labor category data.

All operations are pure: they return new lists and never mutate the input.
"""

from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import cast

from gsa_fcp_toolkit.config import TITLE_KEY, UCID_KEY, YEARS_EXP_KEY
from gsa_fcp_toolkit.models import LaborItem

logger = logging.getLogger(__name__)


def _sanitize_years_experience(value: int | float | str | None, title: str | None) -> int:
    """Coerce a years-of-experience value to an int for Excel compatibility.

    Handles numeric strings with noise (e.g. ``'5+'``, ``'10 years'``) by
    extracting digits, and falls back to ``0`` for unparseable values.
    """
    try:
        if isinstance(value, str):
            digits = "".join(filter(str.isdigit, value))
            return int(digits) if digits else 0
        return int(value) if value is not None else 0
    except (ValueError, TypeError):
        logger.warning("Could not sanitize %s for item: %s", YEARS_EXP_KEY, title)
        return 0


class DataHandler:
    """Lifecycle of labor data from raw JSON to cleaned, sorted structures."""

    @staticmethod
    def load_labor_data(json_path: str | Path) -> list[LaborItem]:
        """Load labor data from a JSON file.

        Args:
            json_path: Path to the source JSON file.

        Returns:
            A list of labor category records.

        Raises:
            FileNotFoundError: If the file does not exist.
            json.JSONDecodeError: If the file is not valid JSON.
            TypeError: If the JSON root is not a list.
        """
        path = Path(json_path)
        try:
            with path.open("r", encoding="utf-8") as f:
                data = json.load(f)
        except FileNotFoundError:
            logger.error("JSON data not found at: %s", path)
            raise
        except json.JSONDecodeError as e:
            logger.error("Failed to decode JSON from %s: %s", path, e)
            raise

        if not isinstance(data, list):
            raise TypeError(f"Expected list of items in {path.name}, got {type(data).__name__}")

        logger.info("Successfully loaded %d items from %s.", len(data), path.name)
        return cast("list[LaborItem]", data)

    @staticmethod
    def sort_labor_data(labor_data: list[LaborItem]) -> list[LaborItem]:
        """Return a new list sorted by UCID for consistent GSA display order.

        Sorts on the string representation to handle mixed-type IDs safely,
        defaulting to ``""`` to avoid None comparison errors. The input list
        is not mutated.
        """
        logger.info("Sorting %d labor items by UCID...", len(labor_data))
        return sorted(labor_data, key=lambda item: str(item.get(UCID_KEY, "")))

    @staticmethod
    def clean_labor_data(labor_data: list[LaborItem]) -> list[LaborItem]:
        """Return a new list with Excel-compatible types.

        Standardizes years-of-experience to ints and strips whitespace from
        titles. The input records are not mutated.
        """
        cleaned: list[LaborItem] = []
        for item in labor_data:
            new_item = cast(LaborItem, dict(item))
            new_item[YEARS_EXP_KEY] = _sanitize_years_experience(
                item.get(YEARS_EXP_KEY), item.get(TITLE_KEY)
            )
            title = item.get(TITLE_KEY)
            if isinstance(title, str):
                new_item[TITLE_KEY] = title.strip()
            cleaned.append(new_item)
        return cleaned
