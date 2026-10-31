"""The CLI is the product's front door; these tests are its contract."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from docviews import cli

from conftest import make_docx

REVISION_FILE = "docx-revisions-mai_tests_OXML_TrackChanges_Test.docx"


def _revision_file(corpus_path: Path) -> Path:
    f = corpus_path / REVISION_FILE
    if not f.exists():
        pytest.skip("corpus not present")
    return f


def test_text_command_prints_the_requested_view(corpus_path, capsys):
    f = _revision_file(corpus_path)
    assert cli.main(["text", str(f), "--view", "accepted"]) == 0
    out = capsys.readouterr().out
    assert "We prefer  color spelling." in out
    assert cli.main(["text", str(f), "--view", "original"]) == 0
    assert "We prefer colour  spelling." in capsys.readouterr().out


def test_revisions_command_lists_authors(corpus_path, capsys):
    f = _revision_file(corpus_path)
    assert cli.main(["revisions", str(f)]) == 0
    out = capsys.readouterr().out
    assert "tracked changes" in out
    assert "insert" in out and "delete" in out


def test_parts_command_shows_every_part(corpus_path, capsys):
    f = _revision_file(corpus_path)
    assert cli.main(["parts", str(f)]) == 0
    assert "document.xml" in capsys.readouterr().out


def test_audit_json_round_trips(corpus_path, capsys):
    f = _revision_file(corpus_path)
    assert cli.main(["audit", str(f), "--json"]) == 0
    payload = json.loads(capsys.readouterr().out)
    assert len(payload) == 1
    row = payload[0]
    assert row["openable"] is True
    assert row["verdict"] in ("clean", "lossy")
    assert row["revision_paragraphs"] > 0


def test_ci_mode_gates_on_revisions(corpus_path, capsys):
    f = _revision_file(corpus_path)
    code = cli.main(["audit", str(f), "--ci", "--fail-on", "revisions"])
    assert code == 2
    assert "--ci: 1 file(s) failed" in capsys.readouterr().err


def test_ci_mode_passes_on_a_clean_file(tmp_path, capsys):
    clean = tmp_path / "clean.docx"
    clean.write_bytes(make_docx("<w:p><w:r><w:t>nothing to see</w:t></w:r></w:p>"))
    assert cli.main(["audit", str(clean), "--ci", "--fail-on", "revisions"]) == 0
    assert "1 files: 0 lossy" in capsys.readouterr().out


def test_missing_file_exits_nonzero(tmp_path, capsys):
    assert cli.main(["text", str(tmp_path / "nope.docx")]) == 1
    assert "error:" in capsys.readouterr().err
