"""Typed data models for gsa_fcp_toolkit."""

from __future__ import annotations

from typing import TypedDict


class LaborItem(TypedDict, total=False):
    """One labor category record from ``labor_data.json``.

    All keys are optional (``total=False``) because the source JSON is
    heterogeneous — records may omit keys or carry mixed types. Consumers
    must use ``.get()`` with sensible defaults.
    """

    unique_catalog_item_id: str
    title: str
    description: str
    minimum_years_of_experience: int | float | str | None
    minimum_education: str
    price_offered_to_gsa_including_iff: float | str | None
    price_offered_to_gsa_excluding_iff: float | str | None
    commercial_price: float | str | None
    discount_offered_to_gsa_off_commercial_price: str
