"""Excel processing engine for GSA FCP file generation.

Uses xlwings to interact with Excel templates, preserving complex
formatting, macros, and data validation while injecting raw data.

This is the only module coupled to Excel (Windows/macOS); everything else
in gsa_fcp_toolkit is platform-independent.
"""

from __future__ import annotations

import logging
from pathlib import Path
from types import TracebackType
from typing import Any

import xlwings as xw

from gsa_fcp_toolkit.config import (
    DATA_START_ROW,
    EXPECTED_TEMPLATE_VERSION,
    HEADER_ROW_TECHNICAL,
    MAX_HEADER_SCAN_COLUMNS,
    PRICING_SHEET_NAME,
    README_SHEET_NAME,
    VERSION_CELL,
    XL_PASTE_FORMATS,
)
from gsa_fcp_toolkit.models import LaborItem

logger = logging.getLogger(__name__)


class ExcelProcessor:
    """Manage the Excel lifecycle, from template loading to data injection.

    A context manager: the Excel process is terminated on exit even if an
    error occurs.
    """

    def __init__(self, template_path: str | Path, visible: bool = False) -> None:
        """Capture metadata; Excel is not launched until first use.

        Args:
            template_path: Path to the .xlsx template.
            visible: If True, bring the Excel window to the foreground.
        """
        self.template_path = Path(template_path).resolve()
        self.visible = visible
        self.app: xw.App | None = None
        self.wb: xw.Book | None = None
        self.ws: xw.Sheet | None = None

    def __enter__(self) -> ExcelProcessor:
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc_val: BaseException | None,
        exc_tb: TracebackType | None,
    ) -> None:
        self.close()

    def _ensure_active(self) -> tuple[xw.App, xw.Book, xw.Sheet]:
        """Lazily launch Excel and open the template; return the live handles.

        Safe to call repeatedly — work already done is not repeated.

        Raises:
            FileNotFoundError: If the template file does not exist.
            ValueError: If the pricing sheet is missing from the template.
        """
        app = self.app
        if app is None:
            # add_book=False prevents Excel from opening a default 'Book1'
            app = xw.App(visible=self.visible, add_book=False)
            self.app = app

        wb = self.wb
        if wb is None:
            if not self.template_path.exists():
                raise FileNotFoundError(f"Excel template not found: {self.template_path}")
            logger.info("Opening workbook: %s", self.template_path.name)
            wb = app.books.open(self.template_path)
            self.wb = wb

        ws = self.ws
        if ws is None:
            if PRICING_SHEET_NAME not in [s.name for s in wb.sheets]:
                logger.error("Sheet '%s' missing from template.", PRICING_SHEET_NAME)
                raise ValueError(f"Workbook must contain a sheet named '{PRICING_SHEET_NAME}'")
            ws = wb.sheets[PRICING_SHEET_NAME]
            self.ws = ws

        return app, wb, ws

    def validate_template_version(self) -> str:
        """Verify the template matches the expected GSA schema version.

        Returns:
            The version string when it matches.

        Raises:
            ValueError: If version tracking is missing or the version mismatches.
        """
        _, wb, _ = self._ensure_active()
        if README_SHEET_NAME not in [s.name for s in wb.sheets]:
            logger.error("Could not find version info on '%s' sheet.", README_SHEET_NAME)
            raise ValueError(f"Template missing version tracking on {README_SHEET_NAME} sheet.")

        current_version = wb.sheets[README_SHEET_NAME].range(VERSION_CELL).value
        if current_version != EXPECTED_TEMPLATE_VERSION:
            logger.error(
                "Template version mismatch! Expected: %s, Found: %s",
                EXPECTED_TEMPLATE_VERSION,
                current_version,
            )
            raise ValueError(f"Incompatible template version: {current_version}")

        logger.info("Verified template version: %s", current_version)
        return str(current_version)

    def _get_headers_mapping(self) -> tuple[dict[str, int], int]:
        """Map technical header names to 1-based column indices.

        Reads the header row in a single call, then returns
        ``(mapping, column_count)``.
        """
        _, _, ws = self._ensure_active()
        header_values = ws.range(
            (HEADER_ROW_TECHNICAL, 1), (HEADER_ROW_TECHNICAL, MAX_HEADER_SCAN_COLUMNS)
        ).value

        headers: dict[str, int] = {}
        last_index = 0
        for i, val in enumerate(header_values):
            if val:
                headers[str(val).strip()] = i + 1  # 1-based indexing for Excel
                last_index = i + 1

        logger.debug("Mapped %d columns from template.", len(headers))
        return headers, last_index

    def append_data(self, labor_data: list[LaborItem], output_path: str | Path) -> None:
        """Inject data, clone formatting, and save the resulting workbook.

        Args:
            labor_data: Cleaned labor category records.
            output_path: Destination path for the generated file.

        Raises:
            ValueError: If ``labor_data`` is empty.
        """
        if not labor_data:
            raise ValueError("No labor data to write; refusing to generate an empty pricelist.")

        _, wb, ws = self._ensure_active()
        abs_output = Path(output_path).resolve()
        headers, col_count = self._get_headers_mapping()

        # Transform JSON -> 2D matrix for bulk writing (far faster than cell-by-cell)
        rows: list[list[Any]] = []
        for record in labor_data:
            row: list[Any] = [None] * col_count
            for key, value in record.items():
                if key in headers:
                    row[headers[key] - 1] = value
            rows.append(row)

        start_row = DATA_START_ROW
        end_row = start_row + len(rows) - 1
        target_range = ws.range((start_row, 1), (end_row, col_count))

        logger.info("Writing %d rows to Excel...", len(rows))
        target_range.value = rows

        self._apply_style_template(start_row, end_row)

        ws.autofit("columns")
        logger.info("Saving output to: %s", abs_output)
        wb.save(abs_output)

    def _apply_style_template(self, start_row: int, end_row: int) -> None:
        """Clone formatting from the first data row across the new data range."""
        app, _, ws = self._ensure_active()
        source_row = ws.range(f"{DATA_START_ROW}:{DATA_START_ROW}")
        target_rows = ws.range(f"{start_row}:{end_row}")

        source_row.api.Copy()
        target_rows.api.PasteSpecial(Paste=XL_PASTE_FORMATS)
        app.api.CutCopyMode = False

    def close(self) -> None:
        """Terminate the Excel process and release handles (never raises)."""
        try:
            if self.wb:
                self.wb.close()
            if self.app:
                self.app.quit()
        except Exception as e:  # noqa: BLE001 — teardown must not raise
            logger.debug("Error during cleanup: %s", e)
        finally:
            self.app = None
            self.wb = None
            self.ws = None
