"""The shipped example input stays valid: it loads through the real pipeline
and generates both documents with the rates the README's worked example quotes."""

from pathlib import Path

from docx import Document

from gsa_fcp_toolkit import config
from gsa_fcp_toolkit.pipeline import prepare_labor_data
from gsa_fcp_toolkit.word import generate_word_cpl

SAMPLE = Path(__file__).parent.parent / "examples" / "sample_labor_data.json"


def _rates(docx_path: Path) -> list[str]:
    doc = Document(str(docx_path))
    table = doc.tables[0]
    return [row.cells[2].text for row in table.rows[1:]]


def test_sample_loads_sorted(tmp_path, monkeypatch):
    items = prepare_labor_data(SAMPLE)
    assert [item["unique_catalog_item_id"] for item in items] == ["PM-001", "SE-002"]


def test_sample_generates_both_documents(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "OUTPUT_DIR", tmp_path)
    items = prepare_labor_data(SAMPLE)

    proposal = generate_word_cpl("proposal", items)
    fss = generate_word_cpl("fss", items)

    assert proposal == tmp_path / "Proposal_CPL.docx"
    assert fss == tmp_path / "Authorized_FSS_Pricelist.docx"
    # Program Manager: direct prices. Systems Engineer: derived —
    # $135 net / (1 - 10%) = $150.00 commercial, $135 / (1 - 0.0075 IFF) = $136.02.
    assert _rates(proposal) == ["$185.00", "$150.00"]
    assert _rates(fss) == ["$172.50", "$136.02"]
