"""Tests for gsa_fcp_toolkit.config."""

from pathlib import Path

from gsa_fcp_toolkit import config


def test_base_dir_resolves_to_project_root():
    assert (config.BASE_DIR / "pyproject.toml").is_file()
    assert config.LABOR_DATA_JSON == config.DATA_DIR / "labor_data.json"
    assert config.FCP_TEMPLATE_XLSX == config.TEMPLATE_DIR / "FCP_Services_Plus_File.xlsx"


def test_env_override(monkeypatch, tmp_path):
    monkeypatch.setenv("GSA_DATA_DIR", str(tmp_path))
    assert config._get_env_path("GSA_DATA_DIR", Path("/default")) == tmp_path


def test_env_fallback(monkeypatch):
    monkeypatch.delenv("GSA_DATA_DIR", raising=False)
    assert config._get_env_path("GSA_DATA_DIR", Path("/default")) == Path("/default")


def test_env_str_override(monkeypatch):
    monkeypatch.setenv("GSA_COMPANY_NAME", "Acme Corp")
    assert config._get_env_str("GSA_COMPANY_NAME", "default") == "Acme Corp"


def test_env_str_fallback(monkeypatch):
    monkeypatch.delenv("GSA_COMPANY_NAME", raising=False)
    assert config._get_env_str("GSA_COMPANY_NAME", "default") == "default"
    monkeypatch.setenv("GSA_COMPANY_NAME", "")
    assert config._get_env_str("GSA_COMPANY_NAME", "default") == "default"


def test_company_defaults_are_placeholders():
    assert config.COMPANY_NAME == "[ENTER COMPANY NAME]"
    assert config.COMPANY_LOCATION == "[ENTER CITY, STATE]"
    assert config.COMPANY_POC == "[ENTER POINT OF CONTACT]"
    assert config.PAYMENT_TERMS == "Net 30"


def test_business_constants():
    assert config.IFF_RATE == 0.0075
    assert config.EXPECTED_TEMPLATE_VERSION == "1.1-20260602"
    assert config.XL_PASTE_FORMATS == -4122
    assert config.PROPOSAL_CPL_FILENAME.endswith(".docx")
    assert config.FSS_PRICELIST_FILENAME.endswith(".docx")
