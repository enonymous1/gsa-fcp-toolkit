"""Tests for gsa_fcp_toolkit.logging_config."""

import logging

from gsa_fcp_toolkit import logging_config


def test_configure_logging_is_idempotent(tmp_path):
    log_file = logging_config.configure_logging(log_dir=tmp_path, force=True)
    before = list(logging.getLogger().handlers)
    assert logging_config.configure_logging(log_dir=tmp_path) == log_file
    assert list(logging.getLogger().handlers) == before
    assert log_file == tmp_path / "gsa_fcp_toolkit.log"
    assert log_file.is_file()


def test_console_level_and_file_audit_trail(tmp_path, capfd):
    logging_config.configure_logging(level=logging.WARNING, log_dir=tmp_path, force=True)
    log = logging.getLogger("gsa_fcp_toolkit.test_logging")
    log.info("info-message")
    log.warning("warning-message")

    console_out = capfd.readouterr().out
    assert "info-message" not in console_out
    assert "warning-message" in console_out

    content = (tmp_path / "gsa_fcp_toolkit.log").read_text(encoding="utf-8")
    assert "info-message" in content  # file always captures DEBUG+
    assert "warning-message" in content
