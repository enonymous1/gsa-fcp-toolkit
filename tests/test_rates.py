"""Tests for gsa_fcp_toolkit.pricing.rates."""

import pytest

from gsa_fcp_toolkit.pricing import calculate_commercial_rate, calculate_gsa_rate

# (item, expected_gsa, expected_commercial)
CASES = [
    ({"title": "A", "price_offered_to_gsa_including_iff": "100.456"}, 100.456, 0.0),
    ({"title": "B", "price_offered_to_gsa_including_iff": 88.5}, 88.5, 0.0),
    (
        {
            "title": "C",
            "price_offered_to_gsa_including_iff": "abc",
            "price_offered_to_gsa_excluding_iff": "100",
        },
        100.76,
        100.0,
    ),
    ({"title": "D", "price_offered_to_gsa_excluding_iff": 100}, 100.76, 100.0),
    ({"title": "E"}, 0.0, 0.0),
    ({"title": "F", "price_offered_to_gsa_excluding_iff": "n/a"}, 0.0, 0.0),
    ({"title": "G", "commercial_price": "75.50"}, 0.0, 75.5),
    (
        {
            "title": "H",
            "commercial_price": "bad",
            "price_offered_to_gsa_excluding_iff": 100,
            "discount_offered_to_gsa_off_commercial_price": "10%",
        },
        100.76,
        0.0,
    ),
    (
        {
            "title": "I",
            "price_offered_to_gsa_excluding_iff": 90,
            "discount_offered_to_gsa_off_commercial_price": "10%",
        },
        90.68,
        100.0,
    ),
    ({"title": "J", "price_offered_to_gsa_excluding_iff": 100}, 100.76, 100.0),
    (
        {
            "title": "K",
            "price_offered_to_gsa_excluding_iff": 100,
            "discount_offered_to_gsa_off_commercial_price": "100%",
        },
        100.76,
        0.0,
    ),
    (
        {
            "title": "L",
            "price_offered_to_gsa_excluding_iff": 100,
            "discount_offered_to_gsa_off_commercial_price": "lots",
        },
        100.76,
        0.0,
    ),
    ({"title": "M", "price_offered_to_gsa_excluding_iff": "???"}, 0.0, 0.0),
    (
        {
            "title": "N",
            "commercial_price": None,
            "price_offered_to_gsa_excluding_iff": 90,
            "discount_offered_to_gsa_off_commercial_price": "10%",
        },
        90.68,
        100.0,
    ),
    (
        {
            "title": "O",
            "price_offered_to_gsa_including_iff": None,
            "price_offered_to_gsa_excluding_iff": 100,
        },
        100.76,
        100.0,
    ),
    ({"title": "P", "commercial_price": 60}, 0.0, 60.0),
]


@pytest.mark.parametrize(("item", "expected_gsa", "expected_commercial"), CASES)
def test_rate_cases(item, expected_gsa, expected_commercial):
    assert calculate_gsa_rate(item) == expected_gsa
    assert calculate_commercial_rate(item) == expected_commercial


def test_iff_derivation_arithmetic():
    # 100 / (1 - 0.0075), rounded to cents
    assert calculate_gsa_rate({"price_offered_to_gsa_excluding_iff": 100}) == 100.76


def test_exotic_types_fall_back_instead_of_raising():
    assert (
        calculate_gsa_rate(
            {
                "title": "Y",
                "price_offered_to_gsa_including_iff": [1],
                "price_offered_to_gsa_excluding_iff": 100,
            }
        )
        == 100.76
    )
    assert calculate_commercial_rate({"title": "Z", "commercial_price": [1]}) == 0.0
