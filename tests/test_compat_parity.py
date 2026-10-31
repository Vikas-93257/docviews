"""Regression tests for python-docx-shaped values and mutation behavior."""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

import docx
import pytest
from docx.oxml import OxmlElement
import docviews.compat as compat
from docviews.compat.enum import (
    MSO_COLOR_TYPE,
    MSO_THEME_COLOR,
    WD_ALIGN_PARAGRAPH,
    WD_BREAK,
    WD_CELL_VERTICAL_ALIGNMENT,
    WD_INLINE_SHAPE,
    WD_LINE_SPACING,
    WD_ROW_HEIGHT_RULE,
    WD_SECTION_START,
    WD_STYLE_TYPE,
    WD_TAB_ALIGNMENT,
    WD_TAB_LEADER,
    WD_TABLE_ALIGNMENT,
)
from docviews.compat.shared import Inches, Pt, RGBColor

from conftest import make_docx
import blobs


def _P(text: str) -> str:
    return f"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>"


def test_sampled_public_inventory_and_callable_shapes_have_no_gaps():
    import compat_coverage as audit

    rows = audit.measure()
    assert all(not row["missing"] for row in rows)
    checked, compared, matched, mismatches, unavailable = audit.signature_audit()
    # Some inherited built-ins (for example int.conjugate on Length) are not
    # introspectable by inspect.signature on older Python runtimes. They remain
    # present on both implementations; compare every available signature and
    # make sure no callable was silently omitted from the audit.
    assert checked == compared + unavailable
    assert compared == matched
    assert mismatches == []


def test_paragraph_text_style_and_alignment_are_read_write_compatible(tmp_path):
    document = compat.Document()
    paragraph = document.add_paragraph("before")
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    assert paragraph.alignment is WD_ALIGN_PARAGRAPH.RIGHT
    assert paragraph.paragraph_format.alignment is WD_ALIGN_PARAGRAPH.RIGHT

    paragraph.style = document.styles["Heading 1"]
    paragraph.text = "after"
    assert paragraph.text == "after"
    assert paragraph.style.name == "Heading 1"

    path = tmp_path / "paragraph.docx"
    document.save(path)
    reopened = docx.Document(str(path))
    assert reopened.paragraphs[0].text == "after"
    assert reopened.paragraphs[0].alignment is docx.enum.text.WD_ALIGN_PARAGRAPH.RIGHT
    assert reopened.paragraphs[0].style.name == "Heading 1"


def test_document_paragraphs_and_inner_content_use_python_docx_body_boundary():
    body = (_P("body") +
            '<w:sdt><w:sdtContent>' + _P("inside content control") +
            "</w:sdtContent></w:sdt>")
    source = make_docx(body)
    ours = compat.Document(source)
    theirs = docx.Document(io.BytesIO(source))

    assert [p.text for p in ours.paragraphs] == [p.text for p in theirs.paragraphs]
    assert [p.text for p in ours.paragraphs] == ["body"]
    # The enhanced reader still exposes nested content through its full walk.
    assert "inside content control" in ours.text()
    assert [type(item).__name__ for item in ours.iter_inner_content()] == [
        "_CompatParagraph"
    ]


def test_document_inner_content_yields_paragraphs_and_tables_in_order():
    document = compat.Document()
    document.add_paragraph("before")
    document.add_table(rows=1, cols=1).cell(0, 0).text = "cell"
    document.add_paragraph("after")

    content = list(document.iter_inner_content())
    assert [type(item).__name__ for item in content] == [
        "_CompatParagraph", "_CompatTable", "_CompatParagraph"
    ]
    assert [p.text for p in document.paragraphs] == ["before", "after"]


def test_default_table_style_matches_python_docx():
    ours = compat.Document().add_table(rows=1, cols=1)
    theirs = docx.Document().add_table(rows=1, cols=1)
    assert ours.style.name == theirs.style.name == "Normal Table"


def test_latent_styles_match_python_docx_public_properties(tmp_path):
    document = compat.Document()
    latent = document.styles.latent_styles
    latent.default_priority = 43
    latent.default_to_hidden = False
    latent.default_to_locked = True
    latent.default_to_quick_style = True
    latent.default_to_unhide_when_used = False
    latent.load_count = 276

    entry = latent.add_latent_style("Coverage Latent")
    entry.hidden = True
    entry.locked = False
    entry.priority = 17
    entry.quick_style = True
    entry.unhide_when_used = None
    alias = latent.add_latent_style("Heading 1")
    assert alias.name == "Heading 1"
    assert latent["Coverage Latent"].name == "Coverage Latent"
    assert len(latent) >= 2

    path = tmp_path / "latent-styles.docx"
    document.save(path)
    reference = docx.Document(str(path)).styles.latent_styles
    assert reference.default_priority == latent.default_priority == 43
    assert reference.default_to_hidden == latent.default_to_hidden is False
    assert reference.default_to_locked == latent.default_to_locked is True
    assert reference.default_to_quick_style == latent.default_to_quick_style is True
    assert reference.default_to_unhide_when_used == \
        latent.default_to_unhide_when_used is False
    assert reference.load_count == latent.load_count == 276
    reference_entry = reference["Coverage Latent"]
    assert reference_entry.name == entry.name
    assert reference_entry.hidden == entry.hidden is True
    assert reference_entry.locked == entry.locked is False
    assert reference_entry.priority == entry.priority == 17
    assert reference_entry.quick_style == entry.quick_style is True
    assert reference_entry.unhide_when_used == entry.unhide_when_used is None

    entry.delete()
    with pytest.raises(KeyError):
        latent["Coverage Latent"]
    with pytest.raises(AttributeError):
        _ = entry.name


def test_table_cell_paragraphs_and_enum_properties_round_trip(tmp_path):
    document = compat.Document()
    table = document.add_table(rows=2, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = document.styles["Table Grid"]
    cell = table.cell(row_idx=0, col_idx=0)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    paragraph = cell.paragraphs[0]
    paragraph.text = "cell value"
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    table.rows[0].height = Inches(0.4)
    table.rows[0].height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

    assert table.alignment is WD_ALIGN_PARAGRAPH.CENTER
    assert cell.vertical_alignment is WD_CELL_VERTICAL_ALIGNMENT.CENTER
    assert paragraph.alignment is WD_ALIGN_PARAGRAPH.CENTER
    assert table.rows[0].height_rule is WD_ROW_HEIGHT_RULE.EXACTLY
    assert table.row_cells(row_idx=0)[0].text == "cell value"
    assert table.column_cells(column_idx=0)[0].text == "cell value"
    assert table.rows.table is table
    assert table.columns.table is table
    assert table.rows[0].table is table
    assert table.columns[0].table is table

    path = tmp_path / "table.docx"
    document.save(path)
    reopened = docx.Document(str(path))
    assert reopened.tables[0].alignment is docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    assert reopened.tables[0].cell(0, 0).vertical_alignment is \
        docx.enum.table.WD_CELL_VERTICAL_ALIGNMENT.CENTER
    assert reopened.tables[0].cell(0, 0).paragraphs[0].alignment is \
        docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    assert reopened.tables[0].rows[0].height_rule is \
        docx.enum.table.WD_ROW_HEIGHT_RULE.EXACTLY
    assert reopened.tables[0].cell(0, 0).text == "cell value"


def test_shared_unit_constructors_match_python_docx():
    cases = (
        (docx.shared.Inches, compat.Inches, 1.25),
        (docx.shared.Cm, compat.Cm, 3.5),
        (docx.shared.Mm, compat.Mm, 12.75),
        (docx.shared.Pt, compat.Pt, 13.5),
        (docx.shared.Emu, compat.Emu, 123456),
        (docx.shared.Twips, compat.Twips, 1.5),
    )
    for reference_type, compat_type, value in cases:
        reference = reference_type(value)
        actual = compat_type(value)
        assert int(actual) == int(reference)
        assert isinstance(actual, compat.Length)
        assert isinstance(actual, compat_type)
        for unit in ("emu", "inches", "cm", "mm", "pt", "twips"):
            assert getattr(actual, unit) == getattr(reference, unit)


def test_font_color_object_matches_python_docx_defaults_and_writes():
    document = compat.Document()
    run = document.add_paragraph("colored").runs[0]
    color = run.font.color
    assert color is not None
    assert color.rgb is None
    assert color.theme_color is None
    assert color.type is None

    color.rgb = RGBColor(12, 34, 56)
    assert color.rgb == RGBColor(12, 34, 56)
    assert isinstance(color.rgb, compat.RGBColor)
    assert color.rgb == docx.shared.RGBColor(12, 34, 56)
    assert RGBColor.from_string("0C2238") == docx.shared.RGBColor.from_string(
        "0C2238")
    assert str(RGBColor.from_string("0C2238")) == "0C2238"
    with pytest.raises(TypeError, match=r"RGBColor\(\) takes three integer values"):
        RGBColor(12.5, 34, 56)
    assert color.type is MSO_COLOR_TYPE.RGB

    style_font = document.styles["Heading 1"].font
    style_font.color.rgb = RGBColor(1, 2, 3)
    assert isinstance(style_font.color.rgb, compat.RGBColor)
    color.theme_color = MSO_THEME_COLOR.ACCENT_1
    assert color.theme_color is MSO_THEME_COLOR.ACCENT_1
    assert color.type is MSO_COLOR_TYPE.THEME

    # ``rgb=None`` removes the color completely, including a prior theme.
    color.rgb = None
    assert color.rgb is None
    assert color.theme_color is None
    assert color.type is None


def test_pictures_get_unique_parts_mutable_dimensions_and_story_rels(tmp_path):
    document = compat.Document()
    first = document.add_picture(blobs.PNG_96, width=Inches(1))
    second = document.add_picture(blobs.PNG_96, width=Inches(1))
    first.width = Inches(2)
    body_run = document.add_paragraph().add_run()
    third = body_run.add_picture(blobs.PNG_96, width=Inches(0.75))

    section = document.sections[0]
    header_run = section.header.add_paragraph("header image").add_run()
    header_shape = header_run.add_picture(blobs.PNG_96, width=Inches(0.5))
    header_shape_2 = header_run.add_picture(blobs.PNG_96, width=Inches(0.4))
    footer_run = section.footer.add_paragraph("footer image").add_run()
    footer_shape = footer_run.add_picture(blobs.PNG_96, width=Inches(0.25))
    for run in (body_run, header_run, footer_run):
        drawing = next(item for item in run.iter_inner_content()
                       if type(item).__name__ == "Drawing")
        assert drawing.has_picture is True
        assert drawing.image.blob == blobs.PNG_96

    assert [first.name, second.name, third.name, header_shape.name,
            header_shape_2.name, footer_shape.name] == [
        "Picture 1", "Picture 2", "Picture 3", "Picture 4", "Picture 5",
        "Picture 6"
    ]
    assert all(shape.type is WD_INLINE_SHAPE.PICTURE
               for shape in (first, second, third, header_shape,
                             header_shape_2, footer_shape))
    assert first.width == Inches(2)
    # Like python-docx, Document.inline_shapes covers the main body, not headers.
    assert len(document.inline_shapes) == 3
    # The core all-parts view resolves relationships in unsaved story parts too.
    all_shapes = document._inline_shapes()
    assert len(all_shapes) == 6
    assert {shape.media for shape in all_shapes} == {
        f"word/media/image{i}.png" for i in range(1, 7)
    }

    path = tmp_path / "pictures.docx"
    document.save(path)
    reopened = docx.Document(str(path))
    assert len(reopened.inline_shapes) == 3
    assert [shape.width for shape in reopened.inline_shapes] == [
        Inches(2), Inches(1), Inches(0.75)
    ]
    with zipfile.ZipFile(path) as package:
        names = set(package.namelist())
        assert {f"word/media/image{i}.png" for i in range(1, 7)} <= names
        header_rels = package.read("word/_rels/header1.xml.rels")
        footer_rels = package.read("word/_rels/footer1.xml.rels")
        assert b'Id="rId1"' in header_rels
        assert b'Target="media/image4.png"' in header_rels
        assert b'Id="rId2"' in header_rels
        assert b'Target="media/image5.png"' in header_rels
        assert b'Id="rId1"' in footer_rels
        assert b'Target="media/image6.png"' in footer_rels
        header_xml = package.read("word/header1.xml")
        footer_xml = package.read("word/footer1.xml")
        assert b"<wp:inline" in header_xml and b"rId2" in header_xml
        assert b"<wp:inline" in footer_xml and b"rId1" in footer_xml

    assert len(reopened.sections[0].header.paragraphs[-1].runs[-1]._r.xpath(
        "w:drawing/wp:inline")) == 2
    assert len(reopened.sections[0].footer.paragraphs[-1].runs[-1]._r.xpath(
        "w:drawing/wp:inline")) == 1
    assert len(compat.Document(path.read_bytes()).inline_shapes) == 3


def test_run_picture_adds_relationship_to_existing_header_part(tmp_path):
    source = (Path(__file__).resolve().parents[1] / "corpus" /
              "resources_AlteredParts_header-simple-plus-image.docx")
    if not source.exists():
        pytest.skip("the optional repository corpus is not present")
    document = compat.Document(source.read_bytes())
    header = document.sections[0].header
    paragraph = header.paragraphs[0]
    shape = paragraph.add_run().add_picture(blobs.PNG_96, width=Inches(0.5))
    assert shape.name == "Picture 2"
    all_shapes = document._inline_shapes()
    assert len(all_shapes) == 2
    assert {item.media for item in all_shapes} == {
        "word/media/image1.png", "word/media/image2.png"
    }

    path = tmp_path / "existing-header-picture.docx"
    document.save(path)
    reopened = docx.Document(str(path))
    assert sum(len(run._r.xpath("w:drawing/wp:inline"))
               for run in reopened.sections[0].header.paragraphs[0].runs) == 2
    with zipfile.ZipFile(path) as package:
        rels = package.read("word/_rels/header2.xml.rels")
        assert b'Id="rId1"' in rels and b'Target="media/image1.png"' in rels
        assert b'Id="rId2"' in rels and b'Target="media/image2.png"' in rels
        assert package.read("word/header2.xml").count(b"<wp:inline") == 2


def test_python_docx_keywords_character_styles_and_stream_save():
    document = compat.Document()
    character_style = document.styles.add_style(
        "Inline Label", WD_STYLE_TYPE.CHARACTER)
    builtin_style = document.styles.add_style(
        "Synthetic Builtin", WD_STYLE_TYPE.PARAGRAPH, True)
    assert builtin_style.builtin is True
    paragraph = document.add_paragraph()
    run = paragraph.add_run(text=None, style=character_style)
    run.add_break(break_type=WD_BREAK.PAGE)
    run.add_picture(image_path_or_stream=io.BytesIO(blobs.PNG_96),
                    width=Inches(0.5))
    document.add_picture(image_path_or_stream=io.BytesIO(blobs.PNG_96),
                         width=Inches(1))

    assert run.style.name == "Inline Label"
    assert document.styles.get_by_id(
        character_style.style_id,
        style_type=WD_STYLE_TYPE.CHARACTER).name == "Inline Label"
    assert document.styles.get_style_id(
        style_or_name=character_style,
        style_type=WD_STYLE_TYPE.CHARACTER) == character_style.style_id

    stream = io.BytesIO()
    assert document.save(path_or_stream=stream) is None
    stream.seek(0)
    reopened = docx.Document(stream)
    assert len(reopened.inline_shapes) == 2
    assert [shape.width for shape in reopened.inline_shapes] == [
        Inches(0.5), Inches(1)
    ]
    saved_run = reopened.paragraphs[0].runs[0]
    assert saved_run.style.name == "Inline Label"
    assert saved_run._r.br_lst[0].type == "page"


def test_add_section_uses_python_docx_start_type_default_and_argument(tmp_path):
    document = compat.Document()
    assert document.add_section().start_type is WD_SECTION_START.NEW_PAGE
    assert document.add_section(WD_SECTION_START.ODD_PAGE).start_type is \
        WD_SECTION_START.ODD_PAGE
    sections = document.sections
    reference = docx.Document()
    reference.add_section()
    reference.add_section(docx.enum.section.WD_SECTION_START.ODD_PAGE)
    reference_sections = reference.sections
    # python-docx creates a fresh Section proxy on each access and does not
    # define equality, so its inherited Sequence methods return zero/raise.
    assert sections.count(sections[0]) == reference_sections.count(
        reference_sections[0]
    ) == 0
    with pytest.raises(ValueError):
        sections.index(sections[0])
    with pytest.raises(ValueError):
        reference_sections.index(reference_sections[0])
    assert len(sections[0:2]) == len(reference_sections[0:2]) == 2

    path = tmp_path / "sections.docx"
    document.save(path)
    sections = docx.Document(str(path)).sections
    assert sections[1].start_type is docx.enum.section.WD_SECTION_START.NEW_PAGE
    assert sections[2].start_type is docx.enum.section.WD_SECTION_START.ODD_PAGE


def test_hyperlink_objects_match_python_docx_public_properties():
    body = ("<w:p><w:hyperlink w:anchor='section1'>"
            "<w:r><w:t>linked</w:t><w:lastRenderedPageBreak/></w:r>"
            "<w:r><w:rPr><w:b/></w:rPr><w:t> text</w:t></w:r>"
            "</w:hyperlink></w:p>")
    source = make_docx(body)
    ours = compat.Document(source).paragraphs[0]
    theirs = docx.Document(io.BytesIO(source)).paragraphs[0]
    our_link, = ours.hyperlinks
    their_link, = theirs.hyperlinks

    assert our_link.address == their_link.address == ""
    assert our_link.fragment == their_link.fragment == "section1"
    assert our_link.url == their_link.url == ""
    assert our_link.text == their_link.text == "linked text"
    assert our_link.contains_page_break == their_link.contains_page_break is True
    assert [run.text for run in our_link.runs] == [
        run.text for run in their_link.runs
    ] == ["linked", " text"]
    assert our_link.paragraph is ours
    assert list(ours.iter_inner_content())[0].text == their_link.text


def test_run_inner_content_and_rendered_break_fragments_match_python_docx():
    source = docx.Document()
    paragraph = source.add_paragraph()
    paragraph.add_run("before")
    rendered = paragraph.add_run()
    rendered.add_text("left")
    rendered._r.append(OxmlElement("w:lastRenderedPageBreak"))
    rendered.add_text("right")
    rendered.add_break(docx.enum.text.WD_BREAK.PAGE)
    rendered.add_text("tail")
    hard_only = paragraph.add_run()
    hard_only.add_break(docx.enum.text.WD_BREAK.PAGE)

    payload = io.BytesIO()
    source.save(payload)
    reference = docx.Document(io.BytesIO(payload.getvalue()))
    document = compat.Document(io.BytesIO(payload.getvalue()))

    expected_run = reference.paragraphs[0].runs[1]
    actual_run = document.paragraphs[0].runs[1]
    expected_items = list(expected_run.iter_inner_content())
    actual_items = list(actual_run.iter_inner_content())
    signature = lambda items: [
        item if isinstance(item, str) else type(item).__name__ for item in items
    ]
    assert signature(actual_items) == signature(expected_items)
    assert actual_run.contains_page_break is expected_run.contains_page_break is True
    assert (document.paragraphs[0].contains_page_break is
            reference.paragraphs[0].contains_page_break is True)

    rendered_break = next(item for item in actual_items
                          if type(item).__name__ == "RenderedPageBreak")
    expected_break = next(item for item in expected_items
                          if type(item).__name__ == "RenderedPageBreak")
    assert (rendered_break.preceding_paragraph_fragment.text ==
            expected_break.preceding_paragraph_fragment.text == "beforeleft")
    assert (rendered_break.following_paragraph_fragment.text ==
            expected_break.following_paragraph_fragment.text == "righttail")
    actual_hard_only = document.paragraphs[0].runs[2]
    expected_hard_only = reference.paragraphs[0].runs[2]
    assert (actual_hard_only.contains_page_break is
            expected_hard_only.contains_page_break is False)

    picture_source = docx.Document()
    picture_source.add_paragraph().add_run().add_picture(
        io.BytesIO(blobs.PNG_96)
    )
    picture_payload = io.BytesIO()
    picture_source.save(picture_payload)
    picture_reference = docx.Document(io.BytesIO(picture_payload.getvalue()))
    picture_document = compat.Document(io.BytesIO(picture_payload.getvalue()))
    expected_drawing = list(
        picture_reference.paragraphs[0].runs[0].iter_inner_content()
    )
    actual_drawing = list(
        picture_document.paragraphs[0].runs[0].iter_inner_content()
    )
    assert [type(item).__name__ for item in actual_drawing] == \
        [type(item).__name__ for item in expected_drawing] == ["Drawing"]
    assert actual_drawing[0]._drawing.tag == str(expected_drawing[0]._drawing.tag)
    assert actual_drawing[0].has_picture is expected_drawing[0].has_picture is True
    actual_image = actual_drawing[0].image
    expected_image = expected_drawing[0].image
    for attribute in ("blob", "content_type", "ext", "filename", "px_width",
                      "px_height", "horz_dpi", "vert_dpi", "width", "height",
                      "sha1"):
        assert getattr(actual_image, attribute) == getattr(expected_image, attribute)
    assert actual_image.scaled_dimensions(width=914400) == \
        expected_image.scaled_dimensions(width=914400)

    unsaved = compat.Document()
    unsaved.add_picture(io.BytesIO(blobs.PNG_96), width=914400)
    unsaved_drawing = next(
        item for item in unsaved.paragraphs[0].runs[0].iter_inner_content()
        if type(item).__name__ == "Drawing"
    )
    assert unsaved_drawing.image.blob == blobs.PNG_96


def test_hyperlink_rendered_break_fragments_match_python_docx():
    source = make_docx(
        "<w:p><w:r><w:t>before</w:t></w:r>"
        "<w:hyperlink w:anchor='fragment'>"
        "<w:r><w:t>left</w:t><w:lastRenderedPageBreak/>"
        "<w:t>right</w:t></w:r></w:hyperlink>"
        "<w:r><w:t>after</w:t></w:r></w:p>"
    )
    reference = docx.Document(io.BytesIO(source))
    document = compat.Document(source)
    expected_run = reference.paragraphs[0].hyperlinks[0].runs[0]
    actual_run = document.paragraphs[0].hyperlinks[0].runs[0]
    expected_break = next(
        item for item in expected_run.iter_inner_content()
        if type(item).__name__ == "RenderedPageBreak"
    )
    actual_break = next(
        item for item in actual_run.iter_inner_content()
        if type(item).__name__ == "RenderedPageBreak"
    )
    assert (actual_break.preceding_paragraph_fragment.text ==
            expected_break.preceding_paragraph_fragment.text == "beforeleftright")
    assert (actual_break.following_paragraph_fragment.text ==
            expected_break.following_paragraph_fragment.text == "after")


def test_paragraph_spacing_and_tab_stops_match_python_docx(tmp_path):
    document = compat.Document()
    paragraph = document.add_paragraph("tabbed")
    paragraph_format = paragraph.paragraph_format

    paragraph_format.line_spacing_rule = WD_LINE_SPACING.DOUBLE
    assert paragraph_format.line_spacing_rule is WD_LINE_SPACING.DOUBLE
    assert paragraph_format.line_spacing == 2.0
    paragraph_format.line_spacing = 1
    assert paragraph_format.line_spacing_rule is WD_LINE_SPACING.SINGLE
    assert paragraph_format.line_spacing == 1.0
    paragraph_format.line_spacing = Pt(14)
    assert paragraph_format.line_spacing_rule is WD_LINE_SPACING.EXACTLY
    assert paragraph_format.line_spacing == Pt(14)
    for rule in (WD_LINE_SPACING.AT_LEAST, WD_LINE_SPACING.MULTIPLE,
                 WD_LINE_SPACING.EXACTLY):
        paragraph_format.line_spacing_rule = rule
        assert paragraph_format.line_spacing_rule is rule
    paragraph_format.line_spacing = Pt(14)

    stops = paragraph_format.tab_stops
    right = stops.add_tab_stop(Inches(3), WD_TAB_ALIGNMENT.RIGHT,
                               WD_TAB_LEADER.DOTS)
    stops.add_tab_stop(Inches(1))
    assert len(stops) == 2
    assert [stop.position for stop in stops] == [Inches(1), Inches(3)]
    assert stops[0].alignment is WD_TAB_ALIGNMENT.LEFT
    assert stops[0].leader is WD_TAB_LEADER.SPACES
    assert right.alignment is WD_TAB_ALIGNMENT.RIGHT
    assert right.leader is WD_TAB_LEADER.DOTS
    right.position = Inches(0.5)
    right.leader = None
    assert [stop.position for stop in stops] == [Inches(0.5), Inches(1)]
    assert stops[0].alignment is WD_TAB_ALIGNMENT.RIGHT
    assert right.leader is WD_TAB_LEADER.SPACES

    path = tmp_path / "paragraph-format.docx"
    document.save(path)
    reopened = docx.Document(str(path))
    saved_format = reopened.paragraphs[0].paragraph_format
    assert saved_format.line_spacing_rule is docx.enum.text.WD_LINE_SPACING.EXACTLY
    assert saved_format.line_spacing == Pt(14)
    saved_stops = saved_format.tab_stops
    assert [stop.position for stop in saved_stops] == [Inches(0.5), Inches(1)]
    assert saved_stops[0].alignment is docx.enum.text.WD_TAB_ALIGNMENT.RIGHT
    assert saved_stops[0].leader is docx.enum.text.WD_TAB_LEADER.SPACES

    stops.clear_all()
    assert len(stops) == 0
    with pytest.raises(IndexError, match="TabStops object is empty"):
        stops[0]
    stops.add_tab_stop(Inches(1))
    del stops[0]
    assert len(stops) == 0


def test_comments_collection_and_rich_comment_api_match_python_docx(tmp_path):
    document = compat.Document()
    paragraph = document.add_paragraph("review this")
    run = paragraph.runs[0]

    comments = document.comments
    comment = comments.add_comment("First line\nSecond line",
                                   author="Reviewer", initials="RV")
    assert len(comments) == 1
    assert comments.get(comment.comment_id).text == "First line\nSecond line"
    assert comments.get(999) is None
    assert comment.comment_id == 0
    assert comment.author == "Reviewer"
    assert comment.initials == "RV"
    assert comment.timestamp.tzinfo is not None
    assert [p.text for p in comment.paragraphs] == ["First line", "Second line"]
    assert comment.anchor_text == ""

    run.mark_comment_range(run, comment.comment_id)
    assert comments.get(0).anchor_text == "review this"
    comment.initials = None
    assert comment.initials is None

    # Comment is a block-item container: rich paragraphs and tables remain
    # editable through the compatibility proxies and survive save/reopen.
    comment.paragraphs[0].add_run(" emphasized").bold = True
    comment.add_paragraph("A third paragraph")
    table = comment.add_table(rows=1, cols=1, width=Inches(2))
    table.cell(0, 0).text = "comment table"
    assert [type(item).__name__ for item in comment.iter_inner_content()] == [
        "_CompatParagraph", "_CompatParagraph", "_CompatParagraph", "_CompatTable"
    ]

    path = tmp_path / "comments-collection.docx"
    document.save(path)
    reopened = docx.Document(str(path))
    assert len(reopened.comments) == 1
    saved = reopened.comments.get(0)
    assert saved.text == "First line emphasized\nSecond line\nA third paragraph"
    assert saved.author == "Reviewer" and saved.initials is None
    assert len(saved.tables) == 1
    assert saved.tables[0].cell(0, 0).text == "comment table"
    assert reopened.paragraphs[0].runs[0].text == "review this"
