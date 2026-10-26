"""Template version policy, tested without Excel.

``validate_template_version`` only needs the workbook's sheet names and the
Read Me B2 value, so the Excel session is stubbed out with fakes that mimic
xlwings' Sheets collection (iterable, plus lookup by name).
"""

import logging

import pytest

from gsa_fcp_toolkit.config import README_SHEET_NAME
from gsa_fcp_toolkit.excel import ExcelProcessor


class _FakeRange:
    def __init__(self, value):
        self.value = value


class _FakeSheet:
    def __init__(self, name, b2_value=None):
        self.name = name
        self._b2_value = b2_value

    def range(self, _cell):
        return _FakeRange(self._b2_value)


class _FakeSheets:
    def __init__(self, sheets):
        self._by_name = {s.name: s for s in sheets}

    def __iter__(self):
        return iter(self._by_name.values())

    def __getitem__(self, name):
        return self._by_name[name]


class _FakeBook:
    def __init__(self, sheets):
        self.sheets = _FakeSheets(sheets)


def _processor_with_version(tmp_path, monkeypatch, version=None, include_readme=True):
    sheets = [_FakeSheet("Pricing")]
    if include_readme:
        sheets.append(_FakeSheet(README_SHEET_NAME, version))
    book = _FakeBook(sheets)
    processor = ExcelProcessor(tmp_path / "template.xlsx")
    monkeypatch.setattr(processor, "_ensure_active", lambda: (None, book, None))
    return processor


def test_latest_version_accepted(tmp_path, monkeypatch):
    processor = _processor_with_version(tmp_path, monkeypatch, "1.1-20260602")
    assert processor.validate_template_version() == "1.1-20260602"


def test_older_known_version_accepted_with_warning(tmp_path, monkeypatch, caplog):
    processor = _processor_with_version(tmp_path, monkeypatch, "1.1-20260128")
    with caplog.at_level(logging.WARNING):
        assert processor.validate_template_version() == "1.1-20260128"
    assert "superseded by 1.1-20260602" in caplog.text


def test_unknown_version_rejected_with_known_list(tmp_path, monkeypatch):
    processor = _processor_with_version(tmp_path, monkeypatch, "1.1-20270101")
    with pytest.raises(ValueError) as exc_info:
        processor.validate_template_version()
    message = str(exc_info.value)
    assert "1.1-20270101" in message
    assert "1.1-20260128" in message and "1.1-20260602" in message


def test_missing_readme_sheet_rejected(tmp_path, monkeypatch):
    processor = _processor_with_version(tmp_path, monkeypatch, include_readme=False)
    with pytest.raises(ValueError, match="missing version tracking"):
        processor.validate_template_version()


def test_missing_version_value_rejected(tmp_path, monkeypatch):
    processor = _processor_with_version(tmp_path, monkeypatch, None)
    with pytest.raises(ValueError) as exc_info:
        processor.validate_template_version()
    assert "Known versions" in str(exc_info.value)
