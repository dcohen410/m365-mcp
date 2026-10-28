"""Read-only mode, tool groups, the signature image setting and log redaction."""

import logging

import pytest

from m365_mcp import cli, logs, server
from m365_mcp import config as config_mod
from m365_mcp.config import Config


def test_every_tool_is_in_exactly_one_group():
    grouped = [name for names in server.TOOL_GROUPS.values() for name in names]
    assert sorted(grouped) == sorted(server.registered_tools())
    assert set(server.TOOL_GROUPS) == set(config_mod.TOOL_GROUPS)


def test_default_settings_hide_nothing():
    assert server.hidden_tools(Config(), server.registered_tools()) == set()


def test_read_only_leaves_only_reading_tools():
    tools = server.registered_tools()
    left = set(tools) - server.hidden_tools(Config(read_only=True), tools)
    assert left == {"search_mail", "read_email", "list_mail_folders", "find_sharepoint_sites", "list_libraries",
                    "list_folder", "search_files", "read_file", "excel_list_sheets", "excel_read", "list_events"}


def test_tool_groups_switch_off_whole_areas():
    tools = server.registered_tools()
    left = set(tools) - server.hidden_tools(Config(tool_groups=["mail", "drafts"]), tools)
    assert left == set(server.TOOL_GROUPS["mail"]) | set(server.TOOL_GROUPS["drafts"])


def test_parse_tool_groups():
    assert config_mod.parse_tool_groups("Files, mail") == ["mail", "files"]
    assert config_mod.parse_tool_groups("all") == list(config_mod.TOOL_GROUPS)
    with pytest.raises(ValueError):
        config_mod.parse_tool_groups("mail,contacts")


def test_settings_from_environment(monkeypatch):
    monkeypatch.setenv("M365_MCP_READ_ONLY", "true")
    monkeypatch.setenv("M365_MCP_TOOLS", "excel")
    cfg = config_mod.load()
    assert cfg.read_only and cfg.tool_groups == ["excel"]
    monkeypatch.setenv("M365_MCP_TOOLS", "nonsense")
    assert config_mod.load().tool_groups == list(config_mod.TOOL_GROUPS)


def test_redact():
    jwt = "eyJhbGciOiJSUzI1NiJ9.eyJzdWIiOiJ4In0.c2lnbmF0dXJl"
    text = logs.redact(f"Authorization: Bearer abc.def-123 token={jwt} "
                       f'{{"refresh_token": "0.AXyz", "x": 1}} from jane.doe+bills@example.com.au')
    assert "abc.def" not in text and jwt not in text and "0.AXyz" not in text and "jane" not in text
    assert "[email]" in text and "Bearer [redacted]" in text


def test_log_records_are_redacted():
    record = logging.LogRecord("x", logging.INFO, __file__, 1, "sent to %s", ("sam@example.org",), None)
    logs.RedactingFilter().filter(record)
    assert record.getMessage() == "sent to [email]"


def test_settings_saved_before_calendar_still_mean_everything(monkeypatch):
    monkeypatch.delenv("M365_MCP_TOOLS", raising=False)
    path = config_mod.config_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text('{"tool_groups": ["mail", "drafts", "files", "excel"]}')
    try:
        assert config_mod.load().tool_groups == list(config_mod.TOOL_GROUPS)
        path.write_text('{"tool_groups": ["mail", "files"]}')
        assert config_mod.load().tool_groups == ["mail", "files"]
    finally:
        path.unlink()


def test_signature_image_is_copied_into_settings_and_removable(monkeypatch, tmp_path):
    settings = tmp_path / "settings"
    monkeypatch.setenv("M365_MCP_CONFIG_DIR", str(settings))
    original = tmp_path / "My Signature.PNG"
    original.write_bytes(b"\x89PNG\r\n\x1a\n")
    cli.main(["configure", "--signature-image", str(original)])
    stored = settings / "signature-image.png"
    assert config_mod.load().signature_image == str(stored)
    original.unlink()  # the copy keeps working
    assert config_mod.load().signature_image_file() == stored

    cli.main(["configure", "--signature-image", ""])
    assert config_mod.load().signature_image == "" and not stored.exists()


def test_signature_image_rejects_other_files_and_keeps_the_old_one(monkeypatch, tmp_path, capsys):
    monkeypatch.setenv("M365_MCP_CONFIG_DIR", str(tmp_path / "settings"))
    good = tmp_path / "sig.jpg"
    good.write_bytes(b"\xff\xd8")
    cli.main(["configure", "--signature-image", str(good)])
    bad = tmp_path / "sig.pdf"
    bad.write_bytes(b"%PDF")
    assert cli.main(["configure", "--signature-image", str(bad)]) == 1
    assert "PNG, JPEG or GIF" in capsys.readouterr().err
    assert config_mod.load().signature_image_file() is not None
