"""Authoring: sections, headers and footers, styles, merged cells, metadata.

Like the rest of the write suite, expectations are checked on the saved file —
and, where python-docx has an equivalent, on what python-docx makes of it.
"""

from __future__ import annotations

import io
import zipfile

import pytest

from docviews import DocxError, open_docx
from docviews.document import Document

from conftest import CORPUS, make_docx

docx = pytest.importorskip("docx")

EMU_PER_INCH = 914400


def P(inner: str) -> str:
    return f"<w:p>{inner}</w:p>"


def R(text: str) -> str:
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


def _corpus_file_with_a_header():
    for candidate in sorted(CORPUS.glob("*.docx")):
        try:
            with zipfile.ZipFile(candidate) as zf:
                if any("/header" in name for name in zf.namelist()):
                    return candidate
        except zipfile.BadZipFile:
            continue
    return None


# ------------------------------------------------------------------- sections
def test_a_new_document_has_one_section_with_word_page_size():
    doc = Document.new()
    assert len(doc.sections) == 1
    section = doc.sections[0]
    # Letter, 1.25in sides — what an empty python-docx document says, so a
    # program that never touches page setup sees the page it saw before
    assert round(section.page_width / EMU_PER_INCH, 2) == 8.5
    assert round(section.page_height / EMU_PER_INCH, 2) == 11.0
    assert round(section.left_margin / EMU_PER_INCH, 2) == 1.25


def test_adding_many_paragraphs_stays_linear():
    """The index of a new paragraph must not cost a walk of the whole document.

    800 paragraphs: O(n) per append is 20x slower than O(1) and was the shape of
    the first implementation. This is a coarse guard — it fails at ~5x the
    measured time, not at a millisecond.
    """
    import time
    from docviews import Document

    doc = Document.new()
    started = time.perf_counter()
    for i in range(800):
        doc.add_paragraph(f"paragraph {i}")
    elapsed = time.perf_counter() - started
    assert [p.index for p in doc.paragraphs()][:3] == [0, 1, 2]
    assert doc.paragraphs()[-1].index == 799
    assert elapsed < 2.0, f"{elapsed:.2f}s to add 800 paragraphs"


def test_the_document_class_defines_each_method_once():
    """A duplicated method silently shadows the first — read the source, not vars().

    ``vars(Document)`` shows the winner only; the duplicate is invisible there.
    """
    import ast
    import collections
    import inspect
    import docviews.document as module

    tree = ast.parse(inspect.getsource(module))
    document = next(node for node in tree.body
                    if isinstance(node, ast.ClassDef) and node.name == "Document")
    names = [node.name for node in document.body
             if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))]
    duplicates = {k: v for k, v in collections.Counter(names).items() if v > 1}
    assert duplicates == {}, duplicates


def test_page_geometry_is_settable():
    doc = Document.new()
    doc.sections[0].page_width = 5 * EMU_PER_INCH
    doc.sections[0].bottom_margin = 0.5 * EMU_PER_INCH
    reopened = open_docx(doc.to_bytes())
    assert reopened.sections[0].page_width == 5 * EMU_PER_INCH
    assert reopened.sections[0].bottom_margin == int(0.5 * EMU_PER_INCH)


def test_add_section_starts_a_new_one_and_python_docx_agrees(tmp_path):
    doc = Document.new()
    doc.add_paragraph("first section")
    doc.add_section()
    doc.add_paragraph("second section")
    assert len(doc.sections) == 2
    path = tmp_path / "sections.docx"
    doc.save(path)
    assert len(docx.Document(str(path)).sections) == 2
    assert len(open_docx(path.read_bytes()).sections) == 2


def test_every_section_is_reachable_in_a_real_file():
    path = _corpus_file_with_a_header()
    if path is None:
        pytest.skip("no corpus file with a header in this checkout")
    ours = len(open_docx(path.read_bytes()).sections)
    theirs = len(docx.Document(str(path)).sections)
    assert ours >= 1 and ours == theirs


# ------------------------------------------------------------ headers, footers
def test_a_new_document_header_is_linked_until_written_to():
    doc = Document.new()
    assert doc.sections[0].header.is_linked_to_previous
    doc.sections[0].header.add_paragraph("Confidential")
    assert not doc.sections[0].header.is_linked_to_previous
    assert doc.sections[0].header.text == "Confidential"


def test_a_header_is_a_real_header_part(tmp_path):
    doc = Document.new()
    doc.sections[0].header.add_paragraph("Confidential")
    doc.sections[0].footer.add_paragraph("page 1")
    path = tmp_path / "hdr.docx"
    doc.save(path)
    with zipfile.ZipFile(path) as zf:
        assert zf.read("word/header1.xml").decode().startswith("<?xml")
        assert b"<w:hdr" in zf.read("word/header1.xml")
        assert b"<w:ftr" in zf.read("word/footer1.xml")
        types = zf.read("[Content_Types].xml").decode()
        assert "wordprocessingml.header+xml" in types
        assert "wordprocessingml.footer+xml" in types
        rels = zf.read("word/_rels/document.xml.rels").decode()
        assert 'Target="header1.xml"' in rels
        assert 'Target="footer1.xml"' in rels
    theirs = docx.Document(str(path))
    assert [p.text for p in theirs.sections[0].header.paragraphs] == ["Confidential"]
    assert [p.text for p in theirs.sections[0].footer.paragraphs] == ["page 1"]


def test_a_header_created_and_left_empty_is_still_valid(tmp_path):
    doc = Document.new()
    doc.sections[0].header                       # touch it, write nothing
    doc.sections[0].footer
    path = tmp_path / "empty.docx"
    doc.save(path)
    theirs = docx.Document(str(path))            # must not raise
    assert theirs.sections[0].header.paragraphs is not None


def test_editing_an_existing_header_changes_only_that_part():
    path = _corpus_file_with_a_header()
    if path is None:
        pytest.skip("no corpus file with a header in this checkout")
    original = path.read_bytes()
    doc = Document(original)
    before = doc.sections[0].header.text
    doc.sections[0].header.add_paragraph("docviews was here")
    saved = doc.to_bytes()
    assert "docviews was here" in open_docx(saved).sections[0].header.text
    assert "docviews was here" in docx.Document(io.BytesIO(saved)) \
        .sections[0].header.paragraphs[-1].text
    with zipfile.ZipFile(io.BytesIO(original)) as a, \
            zipfile.ZipFile(io.BytesIO(saved)) as b:
        changed = [name for name in a.namelist()
                   if not name.endswith("/") and a.read(name) != b.read(name)]
    assert changed and all(name.startswith("word/header") for name in changed), changed
    assert before in doc.sections[0].header.text       # nothing was thrown away


def test_the_python_docx_shape_works_too():
    """``section.header.add_paragraph`` is python-docx's call; keep it working."""
    from docviews import compat
    doc = compat.Document()
    doc.sections[0].header.add_paragraph("compat header")
    doc.sections[0].footer.add_paragraph("compat footer")
    reopened = compat.Document(doc.to_bytes())
    # python-docx's new header part starts with one empty, Header-styled
    # paragraph, so the added text is the *second* paragraph — same as there
    header = reopened.sections[0].header
    assert [p.text for p in header.paragraphs] == ["", "compat header"]
    assert header.paragraphs[0].style.name == "Header"
    assert reopened.sections[0].header.text == "\ncompat header"
    assert reopened.sections[0].footer.paragraphs[-1].text == "compat footer"


# --------------------------------------------------------------------- styles
def test_a_new_style_is_declared_and_linked(tmp_path):
    doc = Document.new()
    # sizes are EMU, like everywhere else in this library (Pt(12) == 152400)
    doc.styles.add_style("Callout", "paragraph", bold=True, size=152400)
    assert "Callout" in doc.styles
    assert doc.styles["Callout"].name == "Callout"
    doc.add_paragraph("styled", style="Callout")
    path = tmp_path / "styles.docx"
    doc.save(path)
    with zipfile.ZipFile(path) as zf:
        styles = zf.read("word/styles.xml").decode()
    assert 'w:styleId="Callout"' in styles and "Callout" in styles
    theirs = docx.Document(str(path))
    assert theirs.paragraphs[-1].style.name == "Callout"
    assert theirs.styles["Callout"].font.bold is True
    assert theirs.styles["Callout"].font.size.pt == 12


def test_adding_the_same_style_twice_is_refused():
    doc = Document.new()
    doc.styles.add_style("Callout", "paragraph")
    with pytest.raises(ValueError):
        doc.styles.add_style("Callout", "paragraph")


def test_styles_of_an_existing_document_are_readable_and_extendable():
    doc = Document.new()
    doc.add_heading("Title", 0)
    doc.styles.add_style("Aside", "character", italic=True)
    reopened = open_docx(doc.to_bytes())
    assert "Aside" in reopened.styles
    assert "Normal" in reopened.styles        # the ones the file came with
    assert reopened.styles["Aside"].type == "character"
    with pytest.raises(KeyError):
        reopened.styles["NoSuchStyle"]


def test_a_style_reference_is_explainable():
    """``paragraph.style.style_id`` is what Word writes; ``.name`` is the label."""
    doc = Document.new()
    paragraph = doc.add_paragraph("x", style="Heading1")
    assert paragraph.style == "Heading1" or paragraph.style.style_id == "Heading1"


# ---------------------------------------------------------------- cell merging
def test_a_horizontal_merge_writes_a_grid_span():
    doc = Document.new()
    table = doc.add_table(rows=2, cols=3,
                          text=[["a", "b", "c"], ["d", "e", "f"]])
    merged = table.cell(0, 0).merge(table.cell(0, 2))
    assert merged.text == "a\nb\nc"
    assert merged.grid_span == 3
    assert len(table.rows[0].cells) == 3          # three columns, one cell wide
    assert [c.text for c in table.rows[0].cells] == ["a\nb\nc"] * 3


def test_a_vertical_merge_says_restart_and_continue():
    doc = Document.new()
    table = doc.add_table(rows=3, cols=2,
                          text=[["a", "b"], ["c", "d"], ["e", "f"]])
    table.cell(0, 0).merge(table.cell(2, 0))
    assert table.rows[0].cells[0].vertical_merge == "restart"
    assert table.rows[1].cells[0].vertical_merge == "continue"
    # the text lives in the top cell; readers that do not follow w:vMerge
    # show blanks below it, so this library follows it for them
    assert table.rows[1].cells[0].text == "a\nc\ne"


def test_merges_read_back_the_same_way_python_docx_reads_them(tmp_path):
    doc = Document.new()
    table = doc.add_table(rows=3, cols=3,
                          text=[["a", "b", "c"], ["d", "e", "f"], ["g", "h", "i"]])
    table.cell(0, 0).merge(table.cell(0, 2))       # horizontal
    table.cell(1, 0).merge(table.cell(2, 0))       # vertical
    path = tmp_path / "merged.docx"
    doc.save(path)
    ours = [[cell.text for cell in row.cells]
            for row in open_docx(path.read_bytes()).tables[0].rows]
    theirs = [[cell.text for cell in row.cells]
              for row in docx.Document(str(path)).tables[0].rows]
    assert ours == theirs
    assert ours[0] == ["a\nb\nc"] * 3
    assert ours[1][0] == ours[2][0] == "d\ng"


def test_merging_cells_from_different_tables_is_refused():
    doc = Document.new()
    first = doc.add_table(rows=1, cols=1)
    second = doc.add_table(rows=1, cols=1)
    with pytest.raises(ValueError):
        first.cell(0, 0).merge(second.cell(0, 0))


# ---------------------------------------------------------------- core metadata
def test_core_properties_are_written_and_read_by_python_docx(tmp_path):
    doc = Document.new()
    doc.core_properties.title = "Quarterly report"
    doc.core_properties.author = "docviews"
    doc.core_properties.subject = "Q3"
    path = tmp_path / "props.docx"
    doc.save(path)
    with zipfile.ZipFile(path) as zf:
        rels = zf.read("_rels/.rels").decode()
        assert "http://schemas.openxmlformats.org/package/2006/relationships" in rels
        assert "metadata/core-properties" in rels
        assert "docProps/core.xml" in zf.namelist()
    theirs = docx.Document(str(path))
    assert theirs.core_properties.title == "Quarterly report"
    assert theirs.core_properties.author == "docviews"
    assert theirs.core_properties.subject == "Q3"
    assert open_docx(path.read_bytes()).core_properties.title == "Quarterly report"


def test_a_package_without_core_properties_reads_as_empty():
    data = make_docx(P(R("plain")))
    assert open_docx(data).core_properties.title == ""


def test_editing_core_properties_of_an_existing_file_touches_only_that_part():
    if not list(CORPUS.glob("*.docx")):
        pytest.skip("no corpus in this checkout")
    path = None
    for candidate in sorted(CORPUS.glob("*.docx")):
        try:
            open_docx(candidate.read_bytes())
            path = candidate
            break
        except Exception:                                     # noqa: BLE001
            continue
    doc = Document(path.read_bytes())
    doc.core_properties.title = "Rewritten"
    saved = doc.to_bytes()
    with zipfile.ZipFile(path) as a, zipfile.ZipFile(io.BytesIO(saved)) as b:
        changed = [name for name in a.namelist()
                   if not name.endswith("/") and a.read(name) != b.read(name)]
    assert changed and all(name.startswith("docProps/") for name in changed), changed
    assert docx.Document(io.BytesIO(saved)).core_properties.title == "Rewritten"


# ------------------------------------------------------------------- round trip
def test_an_untouched_document_still_comes_back_byte_for_byte():
    """The regression that matters most: v0.3 must not rewrite a file it only read."""
    if not list(CORPUS.glob("*.docx")):
        pytest.skip("no corpus in this checkout")
    checked = 0
    for path in sorted(CORPUS.glob("*.docx"))[::24]:
        try:
            original = path.read_bytes()
            saved = Document(original).to_bytes()
        except Exception:                                     # noqa: BLE001
            continue
        checked += 1
        with zipfile.ZipFile(io.BytesIO(original)) as a, \
                zipfile.ZipFile(io.BytesIO(saved)) as b:
            assert a.namelist() == b.namelist()
            for name in a.namelist():
                if not name.endswith("/"):
                    assert a.read(name) == b.read(name), (path, name)
    assert checked >= 5
