"""Focused edge-case comparisons against the pinned python-docx release."""

from __future__ import annotations

from io import BytesIO

import docx
import pytest

import docviews.compat as compat
from conftest import make_docx
from docviews.compat.enum import MSO_THEME_COLOR, WD_ALIGN_PARAGRAPH, WD_UNDERLINE


_XML_TEXT_ERROR = (
    "All strings must be XML compatible: Unicode or ASCII, "
    "no NULL bytes or control characters"
)


def _exception_result(call):
    try:
        call()
    except Exception as exc:  # noqa: BLE001 - comparing public error behavior
        return type(exc), str(exc)
    return None


def _table_snapshot(document):
    return [
        [
            [(cell.grid_span, cell.text) for cell in row.cells]
            for row in table.rows
        ]
        for table in document.tables
    ]


@pytest.mark.parametrize("value", [1.5, "12", True, b"12", None])
def test_length_invalid_constructor_inputs_match_python_docx(value):
    reference = _exception_result(lambda: docx.shared.Length(value))
    candidate = _exception_result(lambda: compat.Length(value))
    assert candidate == reference


def test_hyperlink_without_rendered_break_matches_python_docx():
    source = make_docx(
        "<w:p><w:hyperlink w:anchor='fragment'><w:r><w:t>linked</w:t>"
        "</w:r></w:hyperlink></w:p>"
    )
    reference = docx.Document(BytesIO(source)).paragraphs[0]
    candidate = compat.Document(source).paragraphs[0]

    assert candidate.contains_page_break is reference.contains_page_break is False
    assert candidate.hyperlinks[0].contains_page_break is reference.hyperlinks[0].contains_page_break is False
    assert candidate.rendered_page_breaks == reference.rendered_page_breaks == []


def test_empty_heading_defaults_match_python_docx(tmp_path):
    reference = docx.Document()
    candidate = compat.Document()

    reference.add_heading()
    candidate.add_heading()
    reference.add_heading("")
    candidate.add_heading("")

    def snapshot(document):
        return [(paragraph.text, paragraph.style.name) for paragraph in document.paragraphs]

    assert snapshot(candidate) == snapshot(reference)
    reference_path = tmp_path / "reference-empty-headings.docx"
    candidate_path = tmp_path / "candidate-empty-headings.docx"
    reference.save(reference_path)
    candidate.save(candidate_path)
    assert snapshot(docx.Document(reference_path)) == snapshot(
        compat.Document(candidate_path.read_bytes())
    )


@pytest.mark.parametrize(
    "value",
    [" plain ", b" bytes ", bytearray(b" bytearray ")],
    ids=["str", "bytes", "bytearray"],
)
def test_run_add_text_supported_input_types_match_python_docx(value):
    reference = docx.Document()
    candidate = compat.Document()
    reference_run = reference.add_paragraph().add_run()
    candidate_run = candidate.add_paragraph().add_run()

    reference_result = reference_run.add_text(value)
    candidate_result = candidate_run.add_text(value)

    assert reference_result is not None
    assert candidate_result is not None
    expected = bytes(value).decode("utf-8") if isinstance(value, (bytes, bytearray)) else value
    assert candidate_run.text == reference_run.text == expected


@pytest.mark.parametrize(
    "value",
    [None, 1, 1.5, True, ["text"], memoryview(b"text"), b"\xff", bytearray(b"\xff"), "bad\x00text"],
    ids=["none", "int", "float", "bool", "list", "memoryview", "invalid-bytes", "invalid-bytearray", "xml-control"],
)
def test_run_add_text_error_behavior_matches_python_docx(value):
    reference = docx.Document().add_paragraph().add_run()
    candidate = compat.Document().add_paragraph().add_run()

    reference_error = _exception_result(lambda: reference.add_text(value))
    candidate_error = _exception_result(lambda: candidate.add_text(value))

    assert candidate_error == reference_error
    if value in (b"\xff", bytearray(b"\xff"), "bad\x00text"):
        assert candidate_error == (ValueError, _XML_TEXT_ERROR)


@pytest.mark.parametrize("index", [0, -1])
def test_empty_inline_shape_index_errors_match_python_docx(index):
    reference = docx.Document()
    candidate = compat.Document()

    reference_error = _exception_result(lambda: reference.inline_shapes[index])
    candidate_error = _exception_result(lambda: candidate.inline_shapes[index])

    assert candidate_error == reference_error


@pytest.mark.parametrize(
    "first_corner,second_corner",
    [((0, 0), (1, 1)), ((1, 1), (0, 0)), ((0, 1), (1, 0)), ((1, 0), (0, 1))],
)
def test_cross_table_cell_merge_coordinates_match_python_docx(
    first_corner, second_corner, tmp_path
):
    reference = docx.Document()
    candidate = compat.Document()
    reference_tables = [reference.add_table(rows=2, cols=2) for _ in range(2)]
    candidate_tables = [candidate.add_table(rows=2, cols=2) for _ in range(2)]

    for table_index, (reference_table, candidate_table) in enumerate(
        zip(reference_tables, candidate_tables)
    ):
        for row_index in range(2):
            for column_index in range(2):
                text = f"{table_index}:{row_index}:{column_index}"
                reference_table.cell(row_index, column_index).text = text
                candidate_table.cell(row_index, column_index).text = text

    reference_result = reference_tables[0].cell(*first_corner).merge(
        reference_tables[1].cell(*second_corner)
    )
    candidate_result = candidate_tables[0].cell(*first_corner).merge(
        candidate_tables[1].cell(*second_corner)
    )

    assert candidate_result.text == reference_result.text
    assert _table_snapshot(candidate) == _table_snapshot(reference)

    reference_path = tmp_path / "reference-cross-table-merge.docx"
    candidate_path = tmp_path / "candidate-cross-table-merge.docx"
    reference.save(reference_path)
    candidate.save(candidate_path)
    reopened_reference = docx.Document(reference_path)
    reopened_candidate = compat.Document(candidate_path.read_bytes())
    assert _table_snapshot(reopened_candidate) == _table_snapshot(reopened_reference)


def test_paragraph_add_run_style_arguments_match_python_docx(tmp_path):
    reference = docx.Document()
    candidate = compat.Document()
    ref_paragraph = reference.add_paragraph()
    our_paragraph = candidate.add_paragraph()

    ref_string_style_run = ref_paragraph.add_run("string style", style="Emphasis")
    our_string_style_run = our_paragraph.add_run("string style", style="Emphasis")
    ref_object_style_run = ref_paragraph.add_run(
        "object style", style=reference.styles["Strong"]
    )
    our_object_style_run = our_paragraph.add_run(
        "object style", style=candidate.styles["Strong"]
    )
    ref_empty_run = ref_paragraph.add_run()
    our_empty_run = our_paragraph.add_run()

    def snapshot(paragraph):
        return [
            (run.text, run.style.style_id, run.style.name)
            for run in paragraph.runs
        ]

    assert snapshot(our_paragraph) == snapshot(ref_paragraph)
    assert (ref_string_style_run.style.style_id, our_string_style_run.style.style_id) == (
        "Emphasis",
        "Emphasis",
    )
    assert ref_empty_run.text == our_empty_run.text == ""

    reference_path = tmp_path / "reference-add-run-styles.docx"
    candidate_path = tmp_path / "candidate-add-run-styles.docx"
    reference.save(reference_path)
    candidate.save(candidate_path)
    assert snapshot(compat.Document(candidate_path.read_bytes()).paragraphs[0]) == snapshot(
        docx.Document(reference_path).paragraphs[0]
    )


def test_larger_reversed_rectangle_and_invalid_table_merge_match_python_docx(tmp_path):
    reference = docx.Document()
    candidate = compat.Document()
    ref_table = reference.add_table(rows=3, cols=3)
    our_table = candidate.add_table(rows=3, cols=3)
    for row_index in range(3):
        for column_index in range(3):
            text = f"{row_index}:{column_index}"
            ref_table.cell(row_index, column_index).text = text
            our_table.cell(row_index, column_index).text = text

    ref_merged = ref_table.cell(2, 2).merge(ref_table.cell(0, 0))
    our_merged = our_table.cell(2, 2).merge(our_table.cell(0, 0))
    assert our_merged.text == ref_merged.text
    assert _table_snapshot(candidate) == _table_snapshot(reference)

    reference_path = tmp_path / "reference-reversed-merge.docx"
    candidate_path = tmp_path / "candidate-reversed-merge.docx"
    reference.save(reference_path)
    candidate.save(candidate_path)
    assert _table_snapshot(compat.Document(candidate_path.read_bytes())) == _table_snapshot(
        docx.Document(reference_path)
    )

    reference = docx.Document()
    candidate = compat.Document()
    ref_table = reference.add_table(rows=3, cols=2)
    our_table = candidate.add_table(rows=3, cols=2)
    for row_index in range(3):
        for column_index in range(2):
            text = f"{row_index}:{column_index}"
            ref_table.cell(row_index, column_index).text = text
            our_table.cell(row_index, column_index).text = text
    ref_table.cell(0, 0).merge(ref_table.cell(2, 0))
    our_table.cell(0, 0).merge(our_table.cell(2, 0))

    ref_error = _exception_result(lambda: ref_table.cell(0, 0).merge(ref_table.cell(1, 1)))
    our_error = _exception_result(lambda: our_table.cell(0, 0).merge(our_table.cell(1, 1)))
    assert (our_error[0].__name__, our_error[1]) == (
        ref_error[0].__name__,
        ref_error[1],
    ) == ("InvalidSpanError", "requested span not rectangular")
    assert _table_snapshot(candidate) == _table_snapshot(reference)


@pytest.mark.parametrize("attribute", ["bold", "italic", "underline"])
def test_run_direct_formatting_tristates_match_python_docx(attribute):
    reference = docx.Document()
    candidate = compat.Document()
    reference_run = reference.add_paragraph().add_run("formatted")
    candidate_run = candidate.add_paragraph().add_run("formatted")

    for value in (None, True, False):
        setattr(reference_run, attribute, value)
        setattr(candidate_run, attribute, value)
        assert getattr(candidate_run, attribute) is getattr(reference_run, attribute)


def test_run_font_none_values_preserve_style_inheritance_match_python_docx():
    reference = docx.Document()
    candidate = compat.Document()
    ref_style = reference.styles["Heading 2"]
    our_style = candidate.styles["Heading 2"]
    ref_style.font.bold = our_style.font.bold = True
    ref_style.font.italic = our_style.font.italic = False
    ref_style.font.underline = docx.enum.text.WD_UNDERLINE.DOUBLE
    our_style.font.underline = WD_UNDERLINE.DOUBLE

    ref_run = reference.add_paragraph(style=ref_style).add_run("inherited")
    our_run = candidate.add_paragraph(style=our_style).add_run("inherited")
    assert (our_run.bold, our_run.italic, our_run.underline) == (
        ref_run.bold,
        ref_run.italic,
        ref_run.underline,
    ) == (None, None, None)
    assert our_run.font.bold is ref_run.font.bold is None
    assert our_run.font.italic is ref_run.font.italic is None
    assert our_run.font.underline is ref_run.font.underline is None


def test_tab_stop_alignment_and_leader_enums_match_python_docx(tmp_path):
    reference = docx.Document()
    candidate = compat.Document()
    ref_format = reference.add_paragraph().paragraph_format
    our_format = candidate.add_paragraph().paragraph_format

    position = 1
    for alignment_name in docx.enum.text.WD_TAB_ALIGNMENT.__members__:
        ref_alignment = getattr(docx.enum.text.WD_TAB_ALIGNMENT, alignment_name)
        our_alignment = getattr(compat.WD_TAB_ALIGNMENT, alignment_name)
        for leader_name in docx.enum.text.WD_TAB_LEADER.__members__:
            ref_leader = getattr(docx.enum.text.WD_TAB_LEADER, leader_name)
            our_leader = getattr(compat.WD_TAB_LEADER, leader_name)
            ref_format.tab_stops.add_tab_stop(docx.shared.Inches(position), ref_alignment, ref_leader)
            our_format.tab_stops.add_tab_stop(compat.Inches(position), our_alignment, our_leader)
            position += 1

    def snapshot(paragraph_format):
        return [
            (int(stop.position), stop.alignment.name, stop.leader.name)
            for stop in paragraph_format.tab_stops
        ]

    assert snapshot(our_format) == snapshot(ref_format)
    reference_path = tmp_path / "reference-tab-stop-enums.docx"
    candidate_path = tmp_path / "candidate-tab-stop-enums.docx"
    reference.save(reference_path)
    candidate.save(candidate_path)
    assert snapshot(docx.Document(reference_path).paragraphs[0].paragraph_format) == snapshot(
        compat.Document(candidate_path.read_bytes()).paragraphs[0].paragraph_format
    )


def test_paragraph_text_replacement_and_clear_preserve_formatting_match(tmp_path):
    reference = docx.Document()
    candidate = compat.Document()
    reference_paragraph = reference.add_paragraph(style="Heading 2")
    candidate_paragraph = candidate.add_paragraph(style="Heading 2")
    reference_paragraph.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.RIGHT
    candidate_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    for paragraph in (reference_paragraph, candidate_paragraph):
        paragraph.add_run("old ").bold = True
        paragraph.add_run("inline content").italic = True

    reference_paragraph.text = "replacement\ttext\nnext line"
    candidate_paragraph.text = "replacement\ttext\nnext line"

    def snapshot(paragraph):
        return (
            paragraph.text,
            paragraph.style.style_id,
            paragraph.alignment.name if paragraph.alignment is not None else None,
            [(run.text, run.bold, run.italic) for run in paragraph.runs],
        )

    assert snapshot(candidate_paragraph) == snapshot(reference_paragraph)
    assert len(candidate_paragraph.runs) == len(reference_paragraph.runs) == 1

    assert candidate_paragraph.clear() is candidate_paragraph
    assert reference_paragraph.clear() is reference_paragraph
    assert snapshot(candidate_paragraph) == snapshot(reference_paragraph)
    assert candidate_paragraph.text == ""
    assert candidate_paragraph.style.style_id == "Heading2"
    assert candidate_paragraph.alignment is WD_ALIGN_PARAGRAPH.RIGHT

    reference_path = tmp_path / "reference-paragraph-clear.docx"
    candidate_path = tmp_path / "candidate-paragraph-clear.docx"
    reference.save(reference_path)
    candidate.save(candidate_path)
    assert snapshot(compat.Document(candidate_path.read_bytes()).paragraphs[0]) == snapshot(
        docx.Document(reference_path).paragraphs[0]
    )


def test_font_color_rgb_and_theme_values_round_trip_match_python_docx(tmp_path):
    reference = docx.Document()
    candidate = compat.Document()
    ref_rgb_run = reference.add_paragraph().add_run("rgb").font
    our_rgb_run = candidate.add_paragraph().add_run("rgb").font
    ref_theme_run = reference.add_paragraph().add_run("theme").font
    our_theme_run = candidate.add_paragraph().add_run("theme").font

    def color_snapshot(font):
        color = font.color
        return (
            str(color.rgb) if color.rgb is not None else None,
            color.theme_color.name if color.theme_color is not None else None,
            color.type.name if color.type is not None else None,
        )

    ref_rgb_run.color.rgb = docx.shared.RGBColor(12, 34, 56)
    our_rgb_run.color.rgb = compat.RGBColor(12, 34, 56)
    for theme_name in docx.enum.dml.MSO_THEME_COLOR_INDEX.__members__:
        ref_theme_run.color.theme_color = getattr(
            docx.enum.dml.MSO_THEME_COLOR_INDEX, theme_name
        )
        our_theme_run.color.theme_color = getattr(MSO_THEME_COLOR, theme_name)
        assert color_snapshot(ref_theme_run) == color_snapshot(our_theme_run)
    reference.styles["Heading 1"].font.color.rgb = docx.shared.RGBColor(1, 2, 3)
    candidate.styles["Heading 1"].font.color.rgb = compat.RGBColor(1, 2, 3)

    def snapshot(document):
        return (
            color_snapshot(document.paragraphs[0].runs[0].font),
            color_snapshot(document.paragraphs[1].runs[0].font),
            color_snapshot(document.styles["Heading 1"].font),
        )

    assert snapshot(candidate) == snapshot(reference)
    reference_path = tmp_path / "reference-font-colors.docx"
    candidate_path = tmp_path / "candidate-font-colors.docx"
    reference.save(reference_path)
    candidate.save(candidate_path)
    assert snapshot(compat.Document(candidate_path.read_bytes())) == snapshot(
        docx.Document(reference_path)
    )


def test_comment_defaults_and_none_text_errors_match_python_docx():
    reference = docx.Document()
    candidate = compat.Document()

    reference_comment = reference.comments.add_comment()
    candidate_comment = candidate.comments.add_comment()
    assert (
        candidate_comment.comment_id,
        candidate_comment.author,
        candidate_comment.initials,
        candidate_comment.text,
    ) == (
        reference_comment.comment_id,
        reference_comment.author,
        reference_comment.initials,
        reference_comment.text,
    )
    assert candidate_comment.text == ""
    assert candidate_comment.author == candidate_comment.initials == ""
    assert candidate_comment.timestamp.tzinfo is not None
    assert reference_comment.timestamp.tzinfo is not None

    reference_run = reference.add_paragraph("anchor").runs[0]
    candidate_run = candidate.add_paragraph("anchor").runs[0]
    reference_error = _exception_result(
        lambda: reference.add_comment(reference_run, text=None)
    )
    candidate_error = _exception_result(
        lambda: candidate.add_comment(candidate_run, text=None)
    )
    assert candidate_error == reference_error
    assert candidate_error == (AttributeError, "'NoneType' object has no attribute 'split'")
    assert len(candidate.comments) == len(reference.comments) == 2

    reference_error = _exception_result(lambda: reference.comments.add_comment(None))
    candidate_error = _exception_result(lambda: candidate.comments.add_comment(None))
    assert candidate_error == reference_error
    assert candidate_error == (AttributeError, "'NoneType' object has no attribute 'split'")
    assert len(candidate.comments) == len(reference.comments) == 3


def _comment_marker_snapshot(document):
    def events(paragraph):
        element = getattr(paragraph, "_p", None)
        if element is None:
            element = paragraph._element
        result = []
        for child in element:
            local = child.tag.rsplit("}", 1)[-1]
            if local in {"commentRangeStart", "commentRangeEnd"}:
                result.append((local, child.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id")))
            elif local == "r":
                text = "".join(
                    text_element.text or ""
                    for text_element in child.iter()
                    if isinstance(text_element.tag, str)
                    and text_element.tag.endswith("}t")
                )
                references = tuple(
                    reference.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id")
                    for reference in child.iter()
                    if isinstance(reference.tag, str)
                    and reference.tag.endswith("}commentReference")
                )
                result.append(("r", text, references))
        return result

    body = [events(paragraph) for paragraph in document.paragraphs]
    table_cells = [
        events(paragraph)
        for table in document.tables
        for row in table.rows
        for cell in row.cells
        for paragraph in cell.paragraphs
    ]
    header = [events(paragraph) for paragraph in document.sections[0].header.paragraphs]
    footer = [events(paragraph) for paragraph in document.sections[0].footer.paragraphs]
    return body, table_cells, header, footer


@pytest.mark.parametrize("story_name", ["header", "footer"])
def test_reversed_and_cross_story_comment_ranges_match_python_docx(tmp_path, story_name):
    reference = docx.Document()
    candidate = compat.Document()
    ref_paragraph = reference.add_paragraph()
    our_paragraph = candidate.add_paragraph()
    ref_runs = [ref_paragraph.add_run(text) for text in "ABC"]
    our_runs = [our_paragraph.add_run(text) for text in "ABC"]
    ref_runs[2].mark_comment_range(ref_runs[0], 17)
    our_runs[2].mark_comment_range(our_runs[0], 17)
    assert _comment_marker_snapshot(candidate) == _comment_marker_snapshot(reference)

    reference = docx.Document()
    candidate = compat.Document()
    ref_body_run = reference.add_paragraph("body").runs[0]
    our_body_run = candidate.add_paragraph("body").runs[0]
    ref_story_run = getattr(reference.sections[0], story_name).add_paragraph(story_name).runs[0]
    our_story_run = getattr(candidate.sections[0], story_name).add_paragraph(story_name).runs[0]
    ref_comment = reference.add_comment([ref_body_run, ref_story_run], text="cross-story")
    our_comment = candidate.add_comment([our_body_run, our_story_run], text="cross-story")
    assert (our_comment.comment_id, our_comment.author, our_comment.text) == (
        ref_comment.comment_id,
        ref_comment.author,
        ref_comment.text,
    )
    assert _comment_marker_snapshot(candidate) == _comment_marker_snapshot(reference)

    reference_path = tmp_path / "reference-cross-story-comment.docx"
    candidate_path = tmp_path / "candidate-cross-story-comment.docx"
    reference.save(reference_path)
    candidate.save(candidate_path)
    reopened_reference = docx.Document(reference_path)
    reopened_candidate = compat.Document(candidate_path.read_bytes())
    assert _comment_marker_snapshot(reopened_candidate) == _comment_marker_snapshot(
        reopened_reference
    )
    assert reopened_candidate.comments.get(0).text == reopened_reference.comments.get(0).text


def test_document_comment_anchor_in_table_cell_matches_python_docx(tmp_path):
    reference = docx.Document()
    candidate = compat.Document()
    ref_cell = reference.add_table(rows=1, cols=1).cell(0, 0)
    our_cell = candidate.add_table(rows=1, cols=1).cell(0, 0)
    ref_run = ref_cell.paragraphs[0].add_run("table-cell anchor")
    our_run = our_cell.paragraphs[0].add_run("table-cell anchor")
    ref_comment = reference.add_comment(ref_run, text="review table cell")
    our_comment = candidate.add_comment(our_run, text="review table cell")

    assert (our_comment.comment_id, our_comment.text) == (
        ref_comment.comment_id,
        ref_comment.text,
    )
    assert _comment_marker_snapshot(candidate) == _comment_marker_snapshot(reference)

    reference_path = tmp_path / "reference-table-cell-comment.docx"
    candidate_path = tmp_path / "candidate-table-cell-comment.docx"
    reference.save(reference_path)
    candidate.save(candidate_path)
    reopened_reference = docx.Document(reference_path)
    reopened_candidate = compat.Document(candidate_path.read_bytes())
    assert _comment_marker_snapshot(reopened_candidate) == _comment_marker_snapshot(
        reopened_reference
    )
