"""Tests for gsa_fcp_toolkit.data.handler."""

import copy
import json

import pytest

from gsa_fcp_toolkit.data import DataHandler

ITEMS = [
    {"unique_catalog_item_id": "B2", "title": "  Analyst ", "minimum_years_of_experience": "5+"},
    {"unique_catalog_item_id": "A10", "title": "Engineer", "minimum_years_of_experience": None},
    {"unique_catalog_item_id": "A2", "title": "Manager", "minimum_years_of_experience": "10 years"},
    {"unique_catalog_item_id": "C3", "title": "Clerk", "minimum_years_of_experience": "abc"},
    {"title": "NoUCID", "minimum_years_of_experience": 7},
    {"unique_catalog_item_id": "D4", "minimum_years_of_experience": 2.5},
]


def test_clean_sanitizes_types_and_titles():
    cleaned = DataHandler.clean_labor_data(copy.deepcopy(ITEMS))
    by_ucid = {i.get("unique_catalog_item_id"): i for i in cleaned}
    assert by_ucid["B2"]["minimum_years_of_experience"] == 5
    assert by_ucid["B2"]["title"] == "Analyst"
    assert by_ucid["A10"]["minimum_years_of_experience"] == 0
    assert by_ucid["A2"]["minimum_years_of_experience"] == 10
    assert by_ucid["C3"]["minimum_years_of_experience"] == 0  # no digits -> 0
    assert by_ucid[None]["minimum_years_of_experience"] == 7
    assert by_ucid["D4"]["minimum_years_of_experience"] == 2


def test_clean_does_not_mutate_input():
    original = copy.deepcopy(ITEMS)
    DataHandler.clean_labor_data(ITEMS)
    assert original == ITEMS


def test_sort_orders_by_ucid_string():
    sorted_items = DataHandler.sort_labor_data(copy.deepcopy(ITEMS))
    assert [i.get("unique_catalog_item_id") for i in sorted_items] == [
        None,
        "A10",
        "A2",
        "B2",
        "C3",
        "D4",
    ]


def test_sort_does_not_mutate_input():
    original = copy.deepcopy(ITEMS)
    DataHandler.sort_labor_data(ITEMS)
    assert original == ITEMS


def test_load_roundtrip(tmp_path):
    path = tmp_path / "labor.json"
    path.write_text(json.dumps(ITEMS))
    assert DataHandler.load_labor_data(path) == ITEMS
    assert DataHandler.load_labor_data(str(path)) == ITEMS  # str accepted too


def test_load_missing_file(tmp_path):
    with pytest.raises(FileNotFoundError):
        DataHandler.load_labor_data(tmp_path / "nope.json")


def test_load_invalid_json(tmp_path):
    path = tmp_path / "bad.json"
    path.write_text("{not json")
    with pytest.raises(json.JSONDecodeError):
        DataHandler.load_labor_data(path)


def test_load_non_list_root(tmp_path):
    path = tmp_path / "obj.json"
    path.write_text(json.dumps({"a": 1}))
    with pytest.raises(TypeError):
        DataHandler.load_labor_data(path)
