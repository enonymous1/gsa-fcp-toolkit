"""Command-line interface for gsa_fcp_toolkit.

Subcommands:
    update  Regenerate the FCP Excel pricelist from labor data.
    cpl     Generate the Word Commercial Price List documents.

This module owns all process control (exit codes); library modules only
raise.
"""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from gsa_fcp_toolkit import config
from gsa_fcp_toolkit.excel import ExcelProcessor
from gsa_fcp_toolkit.logging_config import configure_logging
from gsa_fcp_toolkit.pipeline import prepare_labor_data
from gsa_fcp_toolkit.word import generate_word_cpl

logger = logging.getLogger(__name__)

EXIT_VALIDATION = 1
EXIT_MISSING_INPUT = 2


def initialize_fs() -> None:
    """Create the working directories if they do not exist."""
    for folder in (config.DATA_DIR, config.TEMPLATE_DIR, config.OUTPUT_DIR):
        if not folder.exists():
            logger.info("Creating directory: %s", folder)
        folder.mkdir(parents=True, exist_ok=True)


def _require_exists(path: Path, what: str) -> None:
    """Exit with code 2 when a required input file is missing."""
    if not path.exists():
        logger.error("Pre-flight check failed. Missing %s: %s", what, path)
        sys.exit(EXIT_MISSING_INPUT)


def _resolve_input(input_arg: str | None) -> Path:
    """Resolve the labor JSON path from the CLI arg or the configuration."""
    return Path(input_arg) if input_arg else config.LABOR_DATA_JSON


def cmd_update(args: argparse.Namespace) -> int:
    """Run the FCP Excel pricing update. Returns the process exit code."""
    initialize_fs()
    input_path = _resolve_input(args.input)
    _require_exists(input_path, "labor data file")
    _require_exists(config.FCP_TEMPLATE_XLSX, "Excel template")

    try:
        # Single Excel session: validate the template, then inject the data.
        with ExcelProcessor(config.FCP_TEMPLATE_XLSX) as processor:
            logger.info("Validating template version...")
            template_version = processor.validate_template_version()

            output_xlsx = config.OUTPUT_DIR / f"{config.OUTPUT_BASE_NAME}_v{template_version}.xlsx"

            labor_data = prepare_labor_data(input_path)
            processor.append_data(labor_data, output_xlsx)
    except ValueError as e:
        logger.error("VALIDATION ERROR: %s", e)
        return EXIT_VALIDATION
    except TypeError as e:
        logger.error("DATA TYPE ERROR: %s", e)
        return EXIT_VALIDATION
    except Exception:
        logger.critical("FATAL SYSTEM ERROR", exc_info=True)
        return EXIT_VALIDATION

    logger.info("=" * 50)
    logger.info("SUCCESS: GSA FCP UPDATE COMPLETE")
    logger.info("Items Processed: %d", len(labor_data))
    logger.info("Output Location: %s", output_xlsx)
    logger.info("=" * 50)
    return 0


def cmd_cpl(args: argparse.Namespace) -> int:
    """Generate the Word CPL documents. Returns the process exit code."""
    initialize_fs()
    input_path = _resolve_input(args.input)
    _require_exists(input_path, "labor data file")

    try:
        labor_data = prepare_labor_data(input_path)
        if args.type in ("proposal", "both"):
            generate_word_cpl("proposal", labor_data)
        if args.type in ("fss", "both"):
            generate_word_cpl("fss", labor_data)
    except (ValueError, TypeError) as e:
        logger.error("CPL generation failed: %s", e)
        return EXIT_VALIDATION
    except Exception:
        logger.critical("FATAL SYSTEM ERROR", exc_info=True)
        return EXIT_VALIDATION

    logger.info("SUCCESS: CPL documents generated.")
    return 0


def build_parser() -> argparse.ArgumentParser:
    """Build the argument parser for the gsa-fcp CLI."""
    parser = argparse.ArgumentParser(
        prog="gsa-fcp",
        description="GSA FCP pricing update toolkit.",
    )
    parser.add_argument(
        "--log-level",
        default="INFO",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        help="Console log verbosity (default: INFO).",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    p_update = sub.add_parser("update", help="Regenerate the FCP Excel pricelist.")
    p_update.add_argument(
        "--input",
        default=None,
        help="Labor JSON path (default: configured LABOR_DATA_JSON).",
    )
    p_update.set_defaults(func=cmd_update)

    p_cpl = sub.add_parser("cpl", help="Generate the Word Commercial Price List documents.")
    p_cpl.add_argument(
        "--input",
        default=None,
        help="Labor JSON path (default: configured LABOR_DATA_JSON).",
    )
    p_cpl.add_argument(
        "--type",
        choices=["proposal", "fss", "both"],
        default="both",
        help="Which documents to generate (default: both).",
    )
    p_cpl.set_defaults(func=cmd_cpl)
    return parser


def main(argv: list[str] | None = None) -> int:
    """CLI entry point: configure logging once, dispatch, return the exit code."""
    args = build_parser().parse_args(argv)
    configure_logging(level=getattr(logging, args.log_level, logging.INFO))
    exit_code: int = args.func(args)
    return exit_code


if __name__ == "__main__":
    sys.exit(main())
