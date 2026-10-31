"""Authoring: sections, styles, properties, settings, comments (v0.3.0).

Every check is a check on a *file*: whatever is written here is opened again,
with python-docx where python-docx has an opinion — a document only this
library can read is not a .docx.
"""

from __future__ import annotations

import datetime
import io
import sys
import zipfile
from pathlib import Path

import pytest

from docviews import DocxError, open_docx
from docviews.compat.enum import (WD_ORIENT, WD_SECTION_START, WD_STYLE_TYPE)
from docviews.document import Document
from docviews.units import Inches, Pt

sys.path.insert(0, str(Path(__file__).resolve().parent))
from conftest import CORPUS                                   # noqa: E402

docx = pytest.importorskip("docx")                            # python-docx 1.2.0


# --------------------------------------------------------------------- sections
def test_first_and_even_page_headers_are_written_and_read(tmp_path):
    doc = Document.new()
    section = doc.sections[0]
    section.first_page_header.add_paragraph("first page only")
    section.first_page_footer.add_paragraph("footer, page one")
    section.even_page_header.add_paragraph("even pages")
    doc.settings.odd_and_even_pages_header_footer = True
    path = tmp_path / "furniture.docx"
    doc.save(path)

    theirs = docx.Document(str(path))
    assert theirs.sections[0].first_page_header.paragraphs[0].text == "first page only"
    assert theirs.sections[0].first_page_footer.paragraphs[0].text == "footer, page one"
    assert theirs.sections[0].even_page_header.paragraphs[0].text == "even pages"
    assert theirs.settings.odd_and_even_pages_header_footer is True

    ours = open_docx(path.read_bytes()).sections[0]
    assert ours.first_page_header.text == "first page only"
    assert ours.even_page_header.text == "even pages"
    assert ours.first_page_header.is_linked_to_previous is False


def test_a_first_page_header_sets_titlepg():
    doc = Document.new()
    doc.sections[0].first_page_header.add_paragraph("x")
    body = doc.to_bytes()
    with zipfile.ZipFile(io.BytesIO(body)) as zf:
        assert b"titlePg" in zf.read("word/document.xml")


def test_orientation_start_type_and_gutter(tmp_path):
    doc = Document.new()
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.start_type = WD_SECTION_START.ODD_PAGE
    section.gutter = Inches(0.25)
    path = tmp_path / "layout.docx"
    doc.save(path)

    theirs = docx.Document(str(path))
    assert theirs.sections[0].orientation == WD_ORIENT.LANDSCAPE
    assert theirs.sections[0].start_type == WD_SECTION_START.ODD_PAGE
    assert round(theirs.sections[0].gutter.inches, 2) == 0.25
    ours = open_docx(path.read_bytes()).sections[0]
    assert ours.orientation == WD_ORIENT.LANDSCAPE
    assert ours.start_type == WD_SECTION_START.ODD_PAGE


def test_a_header_can_hold_a_table(tmp_path):
    doc = Document.new()
    header = doc.sections[0].header
    header.add_paragraph("Report")
    header.add_table(rows=1, cols=2, text=[["left", "right"]])
    path = tmp_path / "header-table.docx"
    doc.save(path)

    theirs = docx.Document(str(path)).sections[0].header
    assert [type(item).__name__ for item in theirs.iter_inner_content()] == \
        ["Paragraph", "Table"]
    assert theirs.tables[0].cell(0, 1).text == "right"
    ours = open_docx(path.read_bytes()).sections[0].header
    assert [type(item).__name__ for item in ours.iter_inner_content()] == \
        ["Paragraph", "Table"]
    assert len(ours.tables) == 1


# ----------------------------------------------------------------------- styles
def test_styles_can_be_looked_up_three_ways():
    doc = Document.new()
    assert doc.styles.default().name == "Normal"
    assert doc.styles.default(WD_STYLE_TYPE.PARAGRAPH).name == "Normal"
    assert doc.styles.get_by_id("Heading1").name == "Heading 1"
    assert doc.styles.get_style_id("Heading 1") == "Heading1"
    assert doc.styles.get_style_id() == "Normal"
    with pytest.raises(KeyError):
        doc.styles.get_by_id("NoSuchStyle")


def test_latent_styles_read_and_write():
    doc = Document.new()
    latent = doc.styles.latent_styles
    assert latent.default_font_name                    # the template names a font
    latent.add_latent_style("Callout", ui_priority=10, semi_hidden=True)
    assert "Callout" in latent
    assert latent["Callout"].ui_priority == 10
    assert len(latent) == 1
    raw = doc.to_bytes()
    again = open_docx(raw).styles.latent_styles
    assert again["Callout"].semi_hidden is True


def test_style_properties_are_read_and_written(tmp_path):
    doc = Document.new()
    style = doc.styles["Heading 1"]
    assert style.builtin is True
    assert style.base_style.name == "Normal"
    style.quick_style = True
    style.priority = 5
    style.hidden = False
    style.font.size = Pt(18)
    style.paragraph_format.space_before = Pt(6)
    path = tmp_path / "styles.docx"
    doc.save(path)

    theirs = docx.Document(str(path)).styles["Heading 1"]
    assert theirs.quick_style is True
    assert theirs.priority == 5
    assert theirs.font.size.pt == 18
    assert theirs.paragraph_format.space_before.pt == 6

    doc.styles.add_style("Temporary", WD_STYLE_TYPE.PARAGRAPH)
    assert "Temporary" in doc.styles
    doc.styles["Temporary"].delete()
    assert "Temporary" not in doc.styles


# ------------------------------------------------------------- core properties
def test_core_properties_round_trip(tmp_path):
    doc = Document.new()
    props = doc.core_properties
    props.title = "Q3 report"
    props.created = datetime.datetime(2026, 10, 16, 12, 0, 0)
    props.last_printed = datetime.datetime(2026, 10, 15, 9, 30, 0)
    props.language = "en-IN"
    props.identifier = "urn:docviews:1"
    props.version = "3"
    props.content_status = "Draft"
    path = tmp_path / "props.docx"
    doc.save(path)

    theirs = docx.Document(str(path)).core_properties
    assert theirs.created == datetime.datetime(2026, 10, 16, 12, 0, 0,
                                              tzinfo=datetime.timezone.utc)
    assert theirs.last_printed.year == 2026
    assert theirs.language == "en-IN"
    assert theirs.identifier == "urn:docviews:1"
    assert theirs.version == "3"
    assert theirs.content_status == "Draft"

    ours = open_docx(path.read_bytes()).core_properties
    assert ours.created.year == 2026
    assert ours.language == "en-IN"


# --------------------------------------------------------------------- comments
def test_a_new_comment_is_a_comment_python_docx_can_read(tmp_path):
    doc = Document.new()
    paragraph = doc.add_paragraph()
    paragraph.add_run("The total is ")
    number = paragraph.add_run("1,204,332")
    paragraph.add_run(" units.")
    comment = doc.add_comment([number], text="Confirm against the ledger.",
                              author="AM", initials="AM")
    assert comment.comment_id == "0"
    path = tmp_path / "comment.docx"
    doc.save(path)

    with zipfile.ZipFile(path) as zf:
        assert "word/comments.xml" in zf.namelist()
        assert b"comments+xml" in zf.read("[Content_Types].xml")
        assert b"relationships/comments" in zf.read("word/_rels/document.xml.rels")
        document_xml = zf.read("word/document.xml").decode()
        for marker in ("commentRangeStart", "commentRangeEnd", "commentReference"):
            assert marker in document_xml

    theirs = docx.Document(str(path))
    assert [(c.comment_id, c.author, c.text) for c in theirs.comments] == \
        [(0, "AM", "Confirm against the ledger.")]

    ours = open_docx(path.read_bytes()).comments()
    assert [c.anchor_text for c in ours] == ["1,204,332"]      # the range python-docx cannot give


def test_a_comment_can_span_several_runs():
    doc = Document.new()
    paragraph = doc.add_paragraph()
    runs = [paragraph.add_run(part) for part in ("one ", "two ", "three")]
    doc.add_comment(runs, text="whole clause", author="RS")
    comments = open_docx(doc.to_bytes()).comments()
    assert comments[0].anchor_text == "one two three"


def test_adding_a_comment_leaves_the_rest_of_the_file_alone():
    path = CORPUS / "steps_test_files_comments-rich-para.docx"
    if not path.exists():
        pytest.skip("corpus is not present")
    original = path.read_bytes()
    live = Document(original)
    before = len(live.comments())
    target = live.paragraphs()[1].add_run(" [needs a cite]")
    live.add_comment([target], text="Add the source.", author="Editor")

    with zipfile.ZipFile(io.BytesIO(original)) as a, \
            zipfile.ZipFile(io.BytesIO(live.to_bytes())) as b:
        changed = {name for name in a.namelist() if a.read(name) != b.read(name)}
    assert changed == {"word/comments.xml", "word/document.xml"}

    reopened = open_docx(live.to_bytes())
    assert len(reopened.comments()) == before + 1
    added = [c for c in reopened.comments() if c.text == "Add the source."]
    assert added and added[0].anchor_text == " [needs a cite]"


def test_mark_comment_range_is_python_docx_s_own_call(tmp_path):
    """python-docx's primitive: same signature, same markers, same meaning.

    A comment can own more than one range, so marking a second range with an
    existing id is not a mistake — the reader reports both ranges.
    """
    doc = Document.new()
    paragraph = doc.add_paragraph()
    first = paragraph.add_run("alpha ")
    last = paragraph.add_run("beta")
    comment = doc.add_comment([first], text="start", author="A")
    last.mark_comment_range(last, int(comment.comment_id))
    raw = doc.to_bytes()
    body = open_docx(raw)
    anchors = [anchor.text for anchor in body.comments()[0].anchors]
    assert anchors == ["alpha ", "beta"]
    with zipfile.ZipFile(io.BytesIO(raw)) as zf:
        document_xml = zf.read("word/document.xml").decode()
    assert document_xml.count("commentRangeStart") == 2


def test_a_run_from_nowhere_is_refused():
    from docviews.write import Run
    from xml.etree import ElementTree as ET
    orphan = Run(ET.Element("{http://x}r"), document=None)
    with pytest.raises(DocxError):
        orphan.mark_comment_range(orphan, 0)


# ------------------------------------------------------------- through the shim
def test_the_shim_offers_the_same_authoring_surface(tmp_path):
    import docviews.compat as compat
    from docviews.compat.enum import WD_ORIENT, WD_SECTION_START, WD_STYLE_TYPE
    from docviews.compat.shared import Inches

    doc = compat.Document()
    doc.add_heading("Q3", 0)
    para = doc.add_paragraph("Revenue grew ")
    run = para.add_run("18%")
    comment = doc.add_comment([run], text="Check the rounding.", author="AM")
    assert comment.anchor_text == "18%"

    section = doc.sections[0]
    section.header.add_paragraph("Internal")
    section.first_page_header.add_paragraph("Cover")
    section.orientation = WD_ORIENT.LANDSCAPE
    section.start_type = WD_SECTION_START.NEW_PAGE
    section.gutter = Inches(0.2)
    doc.settings.odd_and_even_pages_header_footer = True

    style = doc.styles["Heading 1"]
    style.quick_style = True
    assert doc.styles.default(WD_STYLE_TYPE.PARAGRAPH).name == "Normal"
    assert doc.styles.get_by_id(
        "Heading1", WD_STYLE_TYPE.PARAGRAPH).name == "Heading 1"
    assert doc.styles.get_style_id(
        "Heading 1", WD_STYLE_TYPE.PARAGRAPH) == "Heading1"
    assert len(doc.styles.latent_styles) >= 0
    assert [c.comment_id for c in doc.comments] == [0]

    path = tmp_path / "shim.docx"
    doc.save(path)
    theirs = docx.Document(str(path))
    assert [p.text for p in theirs.sections[0].first_page_header.paragraphs] \
        == ["", "Cover"]                 # the starter paragraph, then the text
    assert theirs.sections[0].orientation == WD_ORIENT.LANDSCAPE
    assert theirs.settings.odd_and_even_pages_header_footer is True
    assert theirs.styles["Heading 1"].quick_style is True
    assert [(c.comment_id, c.text) for c in theirs.comments] == \
        [(0, "Check the rounding.")]
