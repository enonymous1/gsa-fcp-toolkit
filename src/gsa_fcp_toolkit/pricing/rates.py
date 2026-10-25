"""GSA and commercial hourly rate calculations.

Pure functions over :class:`LaborItem` records. Unparseable values produce a
warning and fall back to the next derivation rule (or ``0.0``); nothing here
raises for bad data.
"""

from __future__ import annotations

import logging

from gsa_fcp_toolkit.config import IFF_RATE, TITLE_KEY
from gsa_fcp_toolkit.models import LaborItem

logger = logging.getLogger(__name__)


def _safe_float(value: object, *, title: str | None, field: str) -> float | None:
    """Parse ``value`` to float; warn and return None when unparseable.

    ``None`` input returns ``None`` silently (missing data is normal);
    any other unparseable value logs a warning identifying the record.
    """
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        try:
            return float(value)
        except ValueError:
            pass
    logger.warning("Could not parse %s=%r for item: %s", field, value, title)
    return None


def _parse_discount(value: str | None, title: str | None) -> float | None:
    """Parse a discount like ``'15%'`` into a fraction (``0.15``).

    Returns None when the value is missing or unparseable.
    """
    if value is None:
        return 0.0
    try:
        return float(value.replace("%", "").strip()) / 100
    except (ValueError, TypeError):
        logger.warning("Could not parse discount %r for item: %s", value, title)
        return None


def calculate_gsa_rate(item: LaborItem) -> float:
    """Final GSA hourly rate, inclusive of the Industrial Funding Fee.

    Uses ``price_offered_to_gsa_including_iff`` directly when available;
    otherwise derives it as net / (1 - IFF_RATE), rounded to cents.
    Returns 0.0 when no rate can be determined.
    """
    title = item.get(TITLE_KEY)

    including_iff = _safe_float(
        item.get("price_offered_to_gsa_including_iff"),
        title=title,
        field="price_offered_to_gsa_including_iff",
    )
    if including_iff is not None:
        return including_iff

    gsa_net = _safe_float(
        item.get("price_offered_to_gsa_excluding_iff"),
        title=title,
        field="price_offered_to_gsa_excluding_iff",
    )
    if gsa_net is not None:
        try:
            return round(gsa_net / (1 - IFF_RATE), 2)
        except ZeroDivisionError:
            logger.warning("Cannot derive GSA rate for %s: IFF rate is 100%%.", title)
    return 0.0


def calculate_commercial_rate(item: LaborItem) -> float:
    """Commercial hourly rate for proposal documents.

    Uses ``commercial_price`` directly when available (an unparseable value
    yields 0.0 immediately, without fallback); otherwise reverse-derives it
    as net / (1 - discount), rounded to cents. Returns 0.0 when no rate can
    be determined.
    """
    title = item.get(TITLE_KEY)

    commercial = item.get("commercial_price")
    if commercial is not None:
        rate = _safe_float(commercial, title=title, field="commercial_price")
        return rate if rate is not None else 0.0

    gsa_net = _safe_float(
        item.get("price_offered_to_gsa_excluding_iff"),
        title=title,
        field="price_offered_to_gsa_excluding_iff",
    )
    discount = _parse_discount(
        item.get("discount_offered_to_gsa_off_commercial_price", "0%"), title
    )
    if gsa_net is not None and discount is not None and (1 - discount) > 0:
        return round(gsa_net / (1 - discount), 2)

    logger.error("No valid rate found or calculable for: %s", title)
    return 0.0
