"""Regression against the census (CENSUS.md).

These are the numbers the project is built on. If a change here is not
accompanied by a change in the census, the change is wrong — that is the whole
point of the census existing.
"""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

import pytest

from docviews import open_docx
from docviews.audit import audit_file

pytestmark = pytest.mark.corpus


def corpus_files(corpus_path: Path):
    return sorted(corpus_path.glob("*.docx"))


def test_docviews_reads_the_files_python_docx_cannot_open(corpus_path):
    """Census §3.1: python-docx raises on 75/832; 61 of them are ISO-Strict."""
    files = corpus_files(corpus_path)
    assert len(files) >= 800, f"corpus looks wrong: {len(files)} files"
    strict = unreadable = 0
    for f in files:
        try:
            doc = open_docx(f)
        except Exception:
            unreadable += 1
            continue
        strict += bool(doc.strict)
    assert strict == 61, f"ISO-Strict packages read: {strict}"
    assert unreadable == 1, "only the password-protected file should be unreadable"


def test_unreachable_text_matches_the_census(corpus_path):
    """Census §3.3: ~147k characters live in parts python-docx has no API for."""
    totals = {"footnotes": 0, "endnotes": 0, "textboxes": 0}
    files_with = 0
    for f in corpus_files(corpus_path):
        try:
            row = audit_file(f)
        except Exception:
            continue
        if row.unreachable_chars:
            files_with += 1
        for k, v in row.unreachable_chars.items():
            totals[k] += v
    assert 130 <= files_with <= 190, f"documents with unreachable text: {files_with}"
    assert abs(totals["footnotes"] - 93_676) / 93_676 < 0.05, totals
    assert abs(totals["endnotes"] - 39_053) / 39_053 < 0.05, totals
    assert abs(totals["textboxes"] - 14_453) / 14_453 < 0.10, totals


def test_the_impossible_text_is_gone(corpus_path):
    """Census §3.2: python-docx returns 'We prefer   spelling.' — no Word view
    produces that string. docviews must return exactly the two that exist."""
    f = corpus_path / "docx-revisions-mai_tests_OXML_TrackChanges_Test.docx"
    doc = open_docx(f)
    accepted, original = doc.text("accepted"), doc.text("original")
    assert "We prefer  color spelling." in accepted
    assert "We prefer colour  spelling." in original
    for view_text in (accepted, original):
        assert "We prefer   spelling." not in view_text


def test_tree_and_stream_agree(corpus_path):
    """The bounded-memory path is not a different reader."""
    import random
    files = corpus_files(corpus_path)
    random.Random(0).shuffle(files)
    checked = 0
    for f in files[:40]:
        try:
            doc = open_docx(f)
        except Exception:
            continue
        tree = [p.text for p in doc.paragraphs("accepted")
                if p.part_name == doc.main_part]
        stream = [p.text for p in doc.stream("accepted")]
        assert tree == stream, f"stream and tree disagree on {f.name}"
        checked += 1
    assert checked >= 20


def test_revision_counts_and_authors(corpus_path):
    """Census §6: 95 inserts, 150 deletes, 6 moveTo, 4 moveFrom in 28 documents.

    Counted independently from the raw XML (one revision container = one
    revision), so a double-count or a skip here is a real bug.
    """
    from collections import Counter
    kinds: Counter = Counter()
    scopes: Counter = Counter()
    files_with = 0
    paragraphs = 0
    for f in corpus_files(corpus_path):
        try:
            doc = open_docx(f)
            revs = doc.revisions()
        except Exception:
            continue
        if revs:
            files_with += 1
            kinds.update(r.kind for r in revs)
            scopes.update(r.scope for r in revs)
        assert all(r.author is not None for r in revs)
        paragraphs += sum(1 for p in doc.iter_paragraphs("accepted") if p.is_revised)
    assert kinds == {"insert": 95, "delete": 150, "move_to": 6, "move_from": 4}, kinds
    assert scopes == {"text": 134, "paragraph-mark": 89, "row": 32}, scopes
    assert files_with == 28, files_with
    assert paragraphs == 142, paragraphs


def test_the_expansion_bomb_streams_in_constant_memory(corpus_path):
    """Census §3.6: 0.49 MB of zip hides 28.5 MB of XML, which costs python-docx
    416 MB. docviews' stream path must read it with flat memory."""
    f = corpus_path / "layout-probes_stops2_word-stops-font-widths.docx"
    info = zipfile.ZipFile(f).getinfo("word/document.xml")
    assert info.file_size > 20e6
    doc = open_docx(f)
    paragraphs = list(doc.stream("accepted"))
    assert len(paragraphs) > 4000
    assert sum(len(p.text) for p in paragraphs) > 1_000_000


def test_a_tight_limit_refuses_the_bomb(corpus_path):
    from docviews import Limits
    from docviews.ooxml import PartTooLarge
    f = corpus_path / "layout-probes_stops2_word-stops-font-widths.docx"
    with pytest.raises(PartTooLarge):
        open_docx(f, limits=Limits(max_part_bytes=8 * 1024 * 1024)).text()


# ------------------------------------------------- v0.2: saving is not a loss

@pytest.mark.corpus
def test_saving_an_opened_document_changes_nothing(corpus_path):
    """Open every sampled file, save it unchanged, compare the packages.

    Entries must match exactly, and every part must come back byte-identical —
    which is the claim that makes editing an existing document safe.
    """
    files = sorted(corpus_path.glob("*.docx"))
    sample = files[::12]                      # ~70 files, deterministic
    for extra in ("docx-revisions-mai_tests_OXML_TrackChanges_Test.docx",
                  "steps_test_files_comments-rich-para.docx",
                  "wordprocessing_MiscNode_WithMiscNodes.docx",
                  "assets_TestFiles_Strict01.docx"):
        candidate = corpus_path / extra
        if candidate.exists() and candidate not in sample:
            sample.append(candidate)
    checked = rewrote = 0
    for path in sample:
        try:
            original = path.read_bytes()
            doc = open_docx(original)
            saved = doc.to_bytes()
        except Exception:                                     # noqa: BLE001
            continue
        checked += 1
        with zipfile.ZipFile(io.BytesIO(original)) as a, \
                zipfile.ZipFile(io.BytesIO(saved)) as b:
            assert sorted(a.namelist()) == sorted(b.namelist()), path.name
            for name in a.namelist():
                if name.endswith("/"):
                    continue
                if a.read(name) != b.read(name):
                    rewrote += 1
                    assert name.endswith(".xml"), \
                        f"{path.name}: {name} is not byte-identical"
    assert checked >= 50
    assert rewrote == 0, "an unedited part was rewritten"


@pytest.mark.corpus
def test_editing_a_document_keeps_its_other_parts(corpus_path):
    """Add a paragraph to a tracked-changes file and to a comments file."""
    for name in ("docx-revisions-mai_tests_OXML_TrackChanges_Test.docx",
                 "steps_test_files_comments-rich-para.docx"):
        path = corpus_path / name
        if not path.exists():
            continue
        original = path.read_bytes()
        doc = open_docx(original)
        before = (len(doc.revisions()), len(doc.comments()),
                  sum(len(p.text.strip()) for p in doc.iter_paragraphs(
                      "accepted", kinds=["footnotes", "endnotes"])))
        doc.add_paragraph("added by the test")
        saved = doc.to_bytes()
        with zipfile.ZipFile(io.BytesIO(original)) as a, \
                zipfile.ZipFile(io.BytesIO(saved)) as b:
            assert sorted(a.namelist()) == sorted(b.namelist())
            for part in a.namelist():
                if part in ("word/document.xml",) or part.endswith("/"):
                    continue
                assert a.read(part) == b.read(part), f"{name}: {part} changed"
        after_doc = open_docx(saved)
        after = (len(after_doc.revisions()), len(after_doc.comments()),
                 sum(len(p.text.strip()) for p in after_doc.iter_paragraphs(
                     "accepted", kinds=["footnotes", "endnotes"])))
        assert after == before, f"{name}: content changed {before} -> {after}"
        assert "added by the test" in after_doc.text()
