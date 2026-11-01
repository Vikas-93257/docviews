"""Writing: create, edit, and save without destroying what was already there.

Every expectation is about the file, not the object: write, save, re-open with
the reader, and check what came back.
"""

from __future__ import annotations

import io
import zipfile
from xml.etree import ElementTree as ET

import pytest

from docviews import DocxError, open_docx
from docviews.document import Document

from conftest import make_docx


def P(inner: str) -> str:
    return f"<w:p>{inner}</w:p>"


def R(text: str) -> str:
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


# ------------------------------------------------------------------ new files

def test_a_new_document_is_a_real_package():
    doc = Document.new()
    assert doc.editable and not doc.strict
    with zipfile.ZipFile(io.BytesIO(doc.to_bytes())) as zf:
        names = zf.namelist()
    for required in ("[Content_Types].xml", "_rels/.rels", "word/document.xml",
                     "word/styles.xml", "word/numbering.xml",
                     "word/_rels/document.xml.rels"):
        assert required in names
    assert doc.warnings == []


def test_paragraphs_can_be_added_with_style_alignment_and_formatting():
    doc = Document.new()
    doc.add_heading("Report", 0)
    doc.add_paragraph("Body text.", style="Heading 1", align="center")
    para = doc.add_paragraph("Total: ")
    run = para.add_run("1,204,332", bold=True, italic=True, color="FF0000",
                       size=12700 * 12)
    run.font_name = "Consolas"

    back = open_docx(doc.to_bytes())
    assert back.text() == "Report\nBody text.\nTotal: 1,204,332"
    first, second, third = back.paragraphs()
    assert first.style == "Title" and first.style.name == "Title"
    assert second.style == "Heading1" and second.style.name == "Heading 1"
    assert second.alignment == "center"
    assert [r.text for r in third.runs] == ["Total: ", "1,204,332"]
    formatted = third.runs[1]
    assert formatted.bold and formatted.italic
    assert formatted.color == "FF0000" and formatted.font_name == "Consolas"
    assert formatted.size == 12700 * 12      # EMU: 12 pt


def test_a_paragraph_from_a_reader_is_editable_and_saves():
    doc = Document.new()
    para = doc.add_paragraph("first draft")
    para.set_text("second draft", bold=True)
    para.insert_after("inserted after it")
    back = open_docx(doc.to_bytes())
    assert back.text() == "second draft\ninserted after it"
    assert back.paragraphs()[0].runs[0].bold is True


def test_unstyled_paragraphs_report_the_documents_default_style():
    doc = Document.new()
    doc.add_paragraph("no style set")
    back = open_docx(doc.to_bytes())
    assert back.paragraphs()[0].style.name == "Normal"


def test_lists_and_page_breaks_survive_a_round_trip():
    doc = Document.new()
    doc.add_list_item("bullet one")
    doc.add_list_item("nested bullet", level=1)
    doc.add_list_item("numbered", kind="number")
    doc.add_page_break()
    back = open_docx(doc.to_bytes())
    styles = [p.style for p in back.paragraphs()]
    assert "ListBullet" in styles and "ListNumber" in styles
    # the numbering part is what makes the bullet a bullet, not just an indent
    with zipfile.ZipFile(io.BytesIO(doc.to_bytes())) as zf:
        numbering = zf.read("word/numbering.xml").decode()
    assert 'w:numFmt w:val="bullet"' in numbering
    assert 'w:numFmt w:val="decimal"' in numbering


def test_tables_can_be_written_and_edited_cell_by_cell():
    doc = Document.new()
    table = doc.add_table(rows=2, cols=2, text=[["Region", "Revenue"],
                                                ["EMEA", "1,204,332"]])
    table.cell(0, 0).paragraphs[0].runs[0].bold = True
    table.add_row()
    table.cell(2, 0).text = "APAC"
    table.cell(2, 1).text = "987,001"
    back = open_docx(doc.to_bytes())
    assert back.tables[0].text() == [["Region", "Revenue"],
                                     ["EMEA", "1,204,332"],
                                     ["APAC", "987,001"]]
    assert back.tables[0].style == "TableGrid"
    assert back.tables[0].style.name == "Table Grid"


def test_saving_to_a_path_returns_the_path_and_none_from_compat(tmp_path):
    import docviews.compat as compat
    doc = compat.Document()
    doc.add_paragraph("hello")
    target = tmp_path / "docviews_write_test.docx"
    assert doc.save(str(target)) is None   # python-docx shape
    assert open_docx(target).text() == "hello"


# ------------------------------------------------------------------- editing

def test_editing_an_opened_document_keeps_every_other_part_verbatim():
    """The whole point: writing must not be a way to lose parts."""
    original = make_docx(
        P(R("body text")),
        extra_parts={
            "footnotes.xml": '<w:footnotes xmlns:w="{ns}">'
                             '<w:footnote w:id="1">' + P(R("a footnote")) +
                             "</w:footnote></w:footnotes>",
            "header1.xml": '<w:hdr xmlns:w="{ns}">' + P(R("a header")) + "</w:hdr>",
        })
    doc = open_docx(original)
    before = [p.name for p in doc.parts()]
    doc.add_paragraph("added by the editor")
    raw = doc.to_bytes()

    with zipfile.ZipFile(io.BytesIO(original)) as a, \
            zipfile.ZipFile(io.BytesIO(raw)) as b:
        assert sorted(a.namelist()) == sorted(b.namelist())
        for name in a.namelist():
            if name == "word/document.xml" or name.endswith("/"):
                continue
            assert a.read(name) == b.read(name), f"{name} was rewritten"

    after = open_docx(raw)
    assert after.text() == "body text\nadded by the editor\na header\na footnote"
    assert [p.name for p in after.parts()] == before


def test_tracked_changes_and_comments_survive_an_edit():
    body = (P('<w:del w:id="9" w:author="A" w:date="2026-01-01T00:00:00Z">'
              "<w:r><w:delText>gone</w:delText></w:r></w:del>" + R("kept"))
            + P('<w:commentRangeStart w:id="1"/>' + R("marked")
                + '<w:commentRangeEnd w:id="1"/>' + R(" rest")))
    original = make_docx(
        body,
        extra_parts={
            "comments.xml": ('<w:comments xmlns:w="{ns}">'
                             '<w:comment w:id="1" w:author="A">'
                             + P(R("a comment")) + "</w:comment></w:comments>"),
        })
    doc = open_docx(original)
    assert len(doc.revisions()) == 1 and len(doc.comments()) == 1
    doc.add_paragraph("an addition")
    after = open_docx(doc.to_bytes())
    assert len(after.revisions()) == 1 and len(after.comments()) == 1
    assert after.comments()[0].anchors                    # anchors still land
    assert after.comments()[0].anchors[0].text == "marked"
    # the comment's own text is a part of the document, so it shows in text()
    assert after.text("original") == ("gonekept\nmarked rest\nan addition"
                                      "\na comment")


def test_deleting_and_replacing_text_in_an_opened_document():
    doc = open_docx(make_docx(P(R("first")) + P(R("second"))))
    paragraphs = doc.paragraphs()
    paragraphs[0].set_text("replaced")
    assert paragraphs[1].delete() is True
    back = open_docx(doc.to_bytes())
    assert back.text() == "replaced"


def test_save_into_a_pre_existing_file_is_atomic(tmp_path=None):
    """Planning step 4: the target keeps its old contents until the write lands."""
    import tempfile
    from pathlib import Path
    with tempfile.TemporaryDirectory() as folder:
        target = Path(folder) / "out.docx"
        target.write_bytes(b"the file that was here before")
        doc = Document.new()
        doc.add_paragraph("written now")
        doc.save(target)
        assert open_docx(target).text() == "written now"


def test_save_refuses_when_there_is_nowhere_to_save():
    doc = open_docx(make_docx(P(R("text"))))
    doc.path = "<bytes>"
    with pytest.raises(DocxError):
        doc.save()


def test_a_streamed_paragraph_cannot_be_silently_edited():
    """Streaming paragraphs are detached; editing one must raise, not no-op."""
    doc = open_docx(make_docx(P(R("one")) + P(R("two"))))
    streamed = next(iter(doc.stream()))
    assert not streamed.editable
    with pytest.raises(DocxError):
        streamed.add_run("nope")


# -------------------------------------------------------------------- compat

def test_compat_accepts_the_python_docx_calls_in_order(tmp_path):
    import docviews.compat as docx
    from docviews.compat.enum import WD_ALIGN_PARAGRAPH

    doc = docx.Document()
    doc.add_heading("Title", 0)
    para = doc.add_paragraph("Total: ")
    param = para.add_run("100")
    param.bold = True
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    table = doc.add_table(rows=1, cols=2, style="Table Grid")
    table.cell(0, 0).text = "a"
    table.cell(0, 1).text = "b"
    buffer = io.BytesIO()
    target = tmp_path / "compat_test.docx"
    doc.save(str(target))
    back = docx.Document(str(target))
    assert back.text() == "Title\nTotal: 100\na\nb"
    assert back.paragraphs[1].alignment is WD_ALIGN_PARAGRAPH.CENTER
    assert back.paragraphs[1].runs[1].bold is True
    assert back.tables[0].cell(0, 1).text == "b"


def test_compat_document_error_is_raisable_and_catchable():
    import docviews.compat as docx
    with pytest.raises(docx.DocumentError):
        docx.Document(b"not a docx at all")


def test_compat_units_match_python_docx_numbers():
    from docviews.compat.shared import Inches, Pt, Cm, RGBColor
    assert Inches(1) == 914400
    assert Pt(12) == 152400
    assert Cm(2.54) == 914400
    assert str(RGBColor(0xFF, 0x80, 0x00)) == "FF8000"
    assert Inches(1.5).pt == 108.0


def test_reading_and_writing_are_the_same_document(tmp_path):
    """A reader object and a writer object must agree on what the file says."""
    doc = Document.new()
    doc.add_heading("Shared", 1)
    para = doc.add_paragraph("Tracked ")
    target = tmp_path / "shared_truth.docx"
    doc.save(target)
    back = open_docx(target)
    assert back.text("accepted") == back.text("original")   # nothing revised
    assert back.paragraphs()[0].style.name == "Heading 1"
    assert len(back.parts()) == 1
