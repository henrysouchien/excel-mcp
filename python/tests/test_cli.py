from __future__ import annotations

import json
import re

import pytest

from excel_mcp import cli


def _run_generate_mcp_key(capsys, *args: str) -> list[str]:
  exit_code = cli.main(["generate-mcp-key", *args])

  assert exit_code == 0
  return capsys.readouterr().out.splitlines()


def test_generate_mcp_key_default_label_drops_label_segment(capsys) -> None:
  lines = _run_generate_mcp_key(capsys, "--user", "henry", "--email", "h@e.com", "--risk-user-id", "1")

  assert re.fullmatch(r"sk_henry_mcp_[A-Za-z0-9_-]{32,}", lines[0])


def test_generate_mcp_key_with_label_includes_label(capsys) -> None:
  lines = _run_generate_mcp_key(
    capsys,
    "--user",
    "henry",
    "--email",
    "h@e.com",
    "--risk-user-id",
    "1",
    "--label",
    "excelmcp",
  )

  assert re.fullmatch(r"sk_henry_mcp_excelmcp_[A-Za-z0-9_-]{32,}", lines[0])


def test_generate_mcp_key_emits_valid_json_entry(capsys) -> None:
  lines = _run_generate_mcp_key(capsys, "--user", "henry", "--email", "h@e.com", "--risk-user-id", "1")

  entry = json.loads(lines[1])

  assert set(entry) == {"key", "channel", "slug", "email", "risk_user_id", "role"}
  assert entry["key"] == lines[0]
  assert entry["channel"] == "mcp"
  assert entry["slug"] == "henry"
  assert entry["email"] == "h@e.com"
  assert entry["risk_user_id"] == 1
  assert entry["role"] == "owner"


def test_generate_mcp_key_rejects_invalid_role(capsys) -> None:
  with pytest.raises(SystemExit) as exc:
    cli.main(
      [
        "generate-mcp-key",
        "--user",
        "henry",
        "--email",
        "h@e.com",
        "--risk-user-id",
        "1",
        "--role",
        "engineer",
      ]
    )

  assert exc.value.code != 0
  stderr = capsys.readouterr().err
  assert "invalid choice" in stderr
  assert "owner" in stderr
  assert "invite" in stderr


def test_generate_mcp_key_rejects_missing_required(capsys) -> None:
  with pytest.raises(SystemExit) as exc:
    cli.main(["generate-mcp-key", "--email", "h@e.com", "--risk-user-id", "1"])

  assert exc.value.code != 0
  assert "--user" in capsys.readouterr().err
