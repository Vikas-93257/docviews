"""Writing: creation, editing, the compat shim, and what a save may not destroy.

Everything is asserted on the *file*: build or edit, save, reopen, and check
what came back — with python-docx where python-docx has an opinion, because a
document only this library can read is not a .docx.
"""

from __future__ import annotations

import io
import os
import sys
import zipfile
from pathlib import Path

import pytest

from docviews import DocxError, open_docx
from docviews.document import Document

sys.path.insert(0, str(Path(__file__).resolve().parent))
import blobs                                              # noqa: E402
from conftest import CORPUS, make_docx                    # noqa: E402

docx = pytest.importorskip("docx")                        # python-docx 1.2.0
EMU_PER_INCH = 914400


# ----------------------------------------------------------------- new files
def test_a_new_document_opens_with_python_docx(tmp_path):
    doc = Document.new()
    doc.add_heading("Report", 1)
    doc.add_paragraph("Body text.")
    path = tmp_path / "new.docx"
    doc.save(path)
    theirs = docx.Document(str(path))
    assert [p.text for p in theirs.paragraphs] == ["Report", "Body text."]
    assert theirs.paragraphs[0].style.name == "Heading 1"


def test_editing_a_document_you_did_not_create_keeps_the_rest_byte_for_byte():
    """The claim the whole project rests on, on a file that has everything."""
    path = CORPUS / "steps_test_files_comments-rich-para.docx"
    if not path.exists():
        pytest.skip("corpus is not present")
    original = path.read_bytes()
    doc = open_docx(original)
    doc.paragraphs()[0].set_text("edited")
    doc.add_paragraph("added by docviews")
    saved = doc.to_bytes()
    with zipfile.ZipFile(io.BytesIO(original)) as a, \
            zipfile.ZipFile(io.BytesIO(saved)) as b:
        changed = {name for name in a.namelist()
                   if a.read(name) != b.read(name)}
    assert changed == {"word/document.xml"}          # comments, media, theme: untouched


def test_saving_an_unedited_document_changes_nothing_at_all():
    path = CORPUS / "wordprocessing_MiscNode_WithMiscNodes.docx"
    if not path.exists():
        pytest.skip("corpus is not present")
    original = path.read_bytes()
    saved = Document(original).to_bytes()
    with zipfile.ZipFile(io.BytesIO(original)) as a, \
            zipfile.ZipFile(io.BytesIO(saved)) as b:
        assert a.namelist() == b.namelist()
        for name in a.namelist():
            if not name.endswith("/"):
                assert a.read(name) == b.read(name)


# ------------------------------------------------------------- the compat door
def test_compat_document_builds_the_same_file_as_python_docx(tmp_path):
    import docviews.compat as compat
    from docviews.compat.enum import WD_ALIGN_PARAGRAPH
    from docviews.compat.shared import Inches, Pt

    doc = compat.Document()
    doc.add_heading("Q3", 0)
    para = doc.add_paragraph("Total: ")
    run = para.add_run("1,204,332")
    run.bold = True
    run.font.size = Pt(14)
    run.font.all_caps = True
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("first", style="List Bullet")
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "Region"
    table.cell(0, 1).text = "Revenue"
    table.style = "Table Grid"
    doc.sections[0].left_margin = Inches(1.5)
    doc.core_properties.author = "compat"
    path = tmp_path / "compat.docx"
    doc.save(path)

    theirs = docx.Document(str(path))
    assert theirs.paragraphs[0].style.name == "Title"
    assert theirs.paragraphs[0].text == "Q3"
    assert theirs.paragraphs[1].runs[1].bold is True
    assert theirs.paragraphs[1].runs[1].font.size.pt == 14
    assert theirs.paragraphs[1].runs[1].font.all_caps is True
    assert theirs.paragraphs[1].alignment.name == "CENTER"
    assert theirs.paragraphs[2].style.name == "List Bullet"
    assert theirs.tables[0].cell(0, 0).text == "Region"
    assert theirs.sections[0].left_margin.inches == 1.5
    assert theirs.core_properties.author == "compat"


def test_the_shim_keeps_the_reader_a_python_docx_program_cannot_see():
    """A shim that lost the views would be a downgrade, not a replacement."""
    import docviews.compat as compat
    path = CORPUS / "docx-revisions-mai_tests_OXML_TrackChanges_Test.docx"
    if not path.exists():
        pytest.skip("corpus is not present")
    doc = compat.Document(str(path))
    assert len(doc.revisions()) == 8                 # python-docx sees none
    assert doc.text("markup") != doc.text("accepted")


# -------------------------------------------------------------------- pictures
def test_add_picture_writes_everything_word_needs(tmp_path):
    doc = Document.new()
    doc.add_picture(blobs.PNG_96, width=2 * EMU_PER_INCH)
    path = tmp_path / "picture.docx"
    doc.save(path)
    with zipfile.ZipFile(path) as zf:
        assert zf.read("word/media/image1.png") == blobs.PNG_96
        rels = zf.read("word/_rels/document.xml.rels").decode()
        assert "relationships/image" in rels and 'Target="media/image1.png"' in rels
        assert b'Extension="png"' in zf.read("[Content_Types].xml")
    theirs = docx.Document(str(path))
    assert len(theirs.inline_shapes) == 1
    assert round(theirs.inline_shapes[0].width.inches, 2) == 2.0
    assert round(theirs.inline_shapes[0].height.inches, 2) == 1.0


def test_an_unreadable_image_is_refused_with_a_reason():
    from docviews.images import ImageError
    doc = Document.new()
    with pytest.raises(ImageError):
        doc.add_picture(b"<svg/>")
    with pytest.raises(ImageError):
        doc.add_picture(b"")


# ---------------------------------------------------------------------- tables
def test_table_cells_and_rows_round_trip():
    doc = Document.new()
    table = doc.add_table(rows=2, cols=2, text=[["a", "b"], ["c", "d"]])
    table.add_row()
    table.cell(0, 0).text = "A"
    reopened = open_docx(doc.to_bytes())
    assert reopened.tables[0].text() == [["A", "b"], ["c", "d"], ["", ""]]


def test_a_merged_table_reads_back_the_way_python_docx_reads_it(tmp_path):
    doc = Document.new()
    table = doc.add_table(rows=2, cols=3, text=[["a", "b", "c"], ["d", "e", "f"]])
    table.cell(0, 0).merge(table.cell(0, 2))
    path = tmp_path / "merge.docx"
    doc.save(path)
    ours = [[c.text for c in row.cells]
            for row in open_docx(path.read_bytes()).tables[0].rows]
    theirs = [[c.text for c in row.cells]
              for row in docx.Document(str(path)).tables[0].rows]
    assert ours == theirs
    assert ours[0] == ["a\nb\nc"] * 3


# ------------------------------------------------------------- pages and units
def test_page_geometry_and_sections(tmp_path):
    from docviews.units import Inches, Pt
    assert Inches(1) == 914400 and Pt(12) == 152400
    doc = Document.new()
    doc.add_paragraph("one")
    doc.add_section()
    doc.add_paragraph("two")
    doc.save(tmp_path / "sections.docx")
    theirs = docx.Document(str(tmp_path / "sections.docx"))
    assert len(theirs.sections) == 2


def test_margins_and_orientation_notes():
    from docviews.units import Inches
    doc = Document.new()
    doc.sections[0].left_margin = Inches(0.75)
    assert open_docx(doc.to_bytes()).sections[0].left_margin == int(0.75 * EMU_PER_INCH)
