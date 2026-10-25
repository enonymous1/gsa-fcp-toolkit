"""Central configuration for the gsa_fcp_toolkit package.

Consolidates file paths, Excel schema definitions, business constants, and
document styling in one place. Directory locations support environment
variable overrides so deployments can relocate them without code changes.
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Literal

# =============================================================================
# PATH RESOLUTION
# =============================================================================
# Project root. This module lives at <root>/src/gsa_fcp_toolkit/config.py,
# so the root is three levels up (previously two, when config.py sat directly
# under src/).
BASE_DIR = Path(__file__).resolve().parent.parent.parent


def _get_env_path(var_name: str, default_path: Path) -> Path:
    """Return ``Path(os.environ[var_name])`` when set, else ``default_path``."""
    env_val = os.getenv(var_name)
    return Path(env_val) if env_val else default_path


# Core directories (overridable via environment)
DATA_DIR = _get_env_path("GSA_DATA_DIR", BASE_DIR / "data")
TEMPLATE_DIR = _get_env_path("GSA_TEMPLATE_DIR", BASE_DIR / "templates")
OUTPUT_DIR = _get_env_path("GSA_OUTPUT_DIR", BASE_DIR / "output")

# File references
LABOR_DATA_JSON: Path = DATA_DIR / "labor_data.json"
FCP_TEMPLATE_XLSX: Path = TEMPLATE_DIR / "FCP_Services_Plus_File.xlsx"
OUTPUT_BASE_NAME = "FCP_Services_Plus_File_PERFECT"

# =============================================================================
# EXCEL SCHEMA DEFINITIONS
# =============================================================================
PRICING_SHEET_NAME = "Pricing"
README_SHEET_NAME = "Read Me"

# Row offsets (1-based, matching the Excel coordinate system)
HEADER_ROW_TECHNICAL = 2  # system/technical identifiers
DATA_START_ROW = 4  # first record row (row 3 holds instructions)

# Template versioning
VERSION_CELL = "B2"
EXPECTED_TEMPLATE_VERSION = "1.1-20260602"

# Upper bound on columns scanned when mapping the technical header row
MAX_HEADER_SCAN_COLUMNS = 100

# =============================================================================
# LABOR DATA KEYS (map directly to the labor_data.json schema)
# =============================================================================
# Typed as Literals so mypy accepts them as TypedDict keys in LaborItem.
UCID_KEY: Literal["unique_catalog_item_id"] = "unique_catalog_item_id"
YEARS_EXP_KEY: Literal["minimum_years_of_experience"] = "minimum_years_of_experience"
TITLE_KEY: Literal["title"] = "title"

# =============================================================================
# BUSINESS CONSTANTS (moved out of generate_cpl.py so pricing rules live here)
# =============================================================================
IFF_RATE = 0.0075  # GSA Industrial Funding Fee, used when deriving offered rates

COMPANY_NAME = "Millennium Partners"
COMPANY_LOCATION = "Spotsylvania, Virginia"
COMPANY_POC = "Justin, IT Manager"
PAYMENT_TERMS = "Net 30"

PROPOSAL_CPL_FILENAME = "Millennium_Partners_Proposal_CPL.docx"
FSS_PRICELIST_FILENAME = "Millennium_Partners_Authorized_FSS_Pricelist.docx"

# =============================================================================
# DOCUMENT STYLING (RGB tuples; the word subpackage converts them to RGBColor)
# =============================================================================
DOC_MARGIN_INCHES = 1.0
DOC_FONT_NAME = "Arial"
COLOR_CORPORATE_BLUE: tuple[int, int, int] = (31, 73, 125)
COLOR_SLATE_GRAY: tuple[int, int, int] = (79, 98, 114)
COLOR_WHITE: tuple[int, int, int] = (255, 255, 255)
COLOR_PLACEHOLDER_RED: tuple[int, int, int] = (255, 0, 0)
CPL_TABLE_COLUMN_WIDTHS: tuple[float, float, float] = (1.5, 4.0, 1.0)

# =============================================================================
# EXCEL AUTOMATION CONSTANTS
# =============================================================================
XL_PASTE_FORMATS = -4122  # Excel XlPasteType.xlPasteFormats
