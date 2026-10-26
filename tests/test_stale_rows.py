"""Stale-row clearing in ExcelProcessor.append_data, tested without Excel.

The fakes mimic only the xlwings surface append_data touches: range reads
and writes, used_range, clear_contents, Copy/PasteSpecial, autofit, save.
"""

import logging
from types import SimpleNamespace

from gsa_fcp_toolkit.excel import ExcelProcessor

HEADERS = ["unique_catalog_item_id", "title", "commercial_price"]


class _FakeApi:
    def __init__(self):
        self.calls = []

    def Copy(self):
        self.calls.append("Copy")

    def PasteSpecial(self, Paste=None):
        self.calls.append(("PasteSpecial", Paste))


class _FakeRange:
    def __init__(self, sheet, coords=None):
        self._sheet = sheet
        self._coords = coords  # (r1, c1, r2, c2), or None for whole-row strings
        self.api = sheet.api

    @property
    def value(self):
        r1, c1, r2, c2 = self._coords
        if r1 == r2 and c1 == c2:
            return self._sheet.grid.get((r1, c1))
        rows = [
            [self._sheet.grid.get((r, c)) for c in range(c1, c2 + 1)] for r in range(r1, r2 + 1)
        ]
        if r1 == r2:
            return rows[0]
        return rows

    @value.setter
    def value(self, new):
        r1, c1, r2, c2 = self._coords
        if r1 == r2 and c1 == c2:
            self._sheet.grid[(r1, c1)] = new
            return
        if r1 == r2:
            for offset, v in enumerate(new):
                self._sheet.grid[(r1, c1 + offset)] = v
            return
        for r_offset, row in enumerate(new):
            for c_offset, v in enumerate(row):
                self._sheet.grid[(r1 + r_offset, c1 + c_offset)] = v

    def clear_contents(self):
        r1, c1, r2, c2 = self._coords
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                self._sheet.grid[(r, c)] = None
        self._sheet.clears.append((r1, c1, r2, c2))


class _FakeSheet:
    def __init__(self, grid, used_last_row, fail_scan=False):
        self.grid = grid
        self._used_last_row = used_last_row
        self._fail_scan = fail_scan
        self.api = _FakeApi()
        self.clears = []

    @property
    def used_range(self):
        if self._fail_scan:
            raise RuntimeError("used_range unavailable")
        return SimpleNamespace(last_cell=SimpleNamespace(row=self._used_last_row))

    def range(self, *args):
        if len(args) == 1:
            return _FakeRange(self, None)
        (r1, c1), (r2, c2) = args
        return _FakeRange(self, (r1, c1, r2, c2))

    def autofit(self, _what):
        pass


class _FakeBook:
    def __init__(self):
        self.saved = None

    def save(self, path):
        self.saved = path


def _record(ucid):
    return {"unique_catalog_item_id": ucid, "title": "Role " + ucid, "commercial_price": "100.00"}


def _make_processor(tmp_path, monkeypatch, dummy_rows, used_last_row, fail_scan=False):
    grid = {(2, i): h for i, h in enumerate(HEADERS, start=1)}
    for row, ucid in dummy_rows:
        grid[(row, 1)] = ucid
    sheet = _FakeSheet(grid, used_last_row, fail_scan)
    book = _FakeBook()
    app = SimpleNamespace(api=SimpleNamespace())
    processor = ExcelProcessor(tmp_path / "template.xlsx")
    monkeypatch.setattr(processor, "_ensure_active", lambda: (app, book, sheet))
    return processor, sheet, book


def test_clears_stale_rows_beyond_written_data(tmp_path, monkeypatch):
    dummies = [(r, f"DUMMY-EXAMPLE-00{r - 3}") for r in range(4, 8)]
    processor, sheet, book = _make_processor(tmp_path, monkeypatch, dummies, 7)
    out = tmp_path / "out.xlsx"

    processor.append_data([_record("A1"), _record("A2")], out)

    assert sheet.clears == [(6, 1, 7, 3)]
    assert sheet.grid[(4, 1)] == "A1" and sheet.grid[(5, 1)] == "A2"
    assert sheet.grid[(6, 1)] is None and sheet.grid[(7, 1)] is None
    assert book.saved == out


def test_no_clear_when_data_covers_dummy_rows(tmp_path, monkeypatch):
    dummies = [(r, f"DUMMY-EXAMPLE-00{r - 3}") for r in range(4, 8)]
    processor, sheet, book = _make_processor(tmp_path, monkeypatch, dummies, 7)

    processor.append_data([_record(f"A{i}") for i in range(1, 6)], tmp_path / "out.xlsx")

    assert sheet.clears == []
    assert [sheet.grid[(r, 1)] for r in range(4, 9)] == ["A1", "A2", "A3", "A4", "A5"]


def test_no_clear_when_no_preexisting_rows(tmp_path, monkeypatch):
    processor, sheet, book = _make_processor(tmp_path, monkeypatch, [], 3)

    processor.append_data([_record("A1"), _record("A2")], tmp_path / "out.xlsx")

    assert sheet.clears == []
    assert book.saved is not None


def test_single_row_scan_flat_value(tmp_path, monkeypatch):
    # used_range ends on row 4, so the scan read comes back as one flat row.
    processor, sheet, book = _make_processor(tmp_path, monkeypatch, [(4, "DUMMY-EXAMPLE-001")], 4)

    processor.append_data([_record("A1"), _record("A2")], tmp_path / "out.xlsx")

    assert sheet.clears == []
    assert sheet.grid[(4, 1)] == "A1"


def test_scan_failure_warns_and_proceeds(tmp_path, monkeypatch, caplog):
    dummies = [(r, f"DUMMY-EXAMPLE-00{r - 3}") for r in range(4, 8)]
    processor, sheet, book = _make_processor(tmp_path, monkeypatch, dummies, 7, fail_scan=True)

    with caplog.at_level(logging.WARNING):
        processor.append_data([_record("A1"), _record("A2")], tmp_path / "out.xlsx")

    assert "stale rows" in caplog.text
    assert sheet.clears == []
    assert sheet.grid[(4, 1)] == "A1" and book.saved is not None
