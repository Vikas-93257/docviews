"""The compat enums: same members, same values, and a usable token.

python-docx code says ``WD_ALIGN_PARAGRAPH.CENTER``, compares it to ``1``, and
expects the file to end up with ``w:jc="center"``. All three have to hold here,
so all three are checked — including against python-docx itself, member by
member, because a member list that drifts silently is a migration bug.
"""

from __future__ import annotations

import io
import re
import zipfile

import pytest

from docviews.compat import enum as ours
from docviews.compat.enum import _generated

docx = pytest.importorskip("docx")                        # python-docx 1.2.0


def test_every_generated_member_has_python_docx_s_value():
    checked = 0
    for name, members in _generated.ENUMS.items():
        enum_class = getattr(ours, name, None)
        assert enum_class is not None, f"{name} is missing from the compat enums"
        theirs = None
        for module_name in ("docx.enum.text", "docx.enum.style",
                            "docx.enum.section", "docx.enum.table",
                            "docx.enum.shape", "docx.enum.dml"):
            module = __import__(module_name, fromlist=["*"])
            if hasattr(module, name):
                theirs = getattr(module, name)
                break
        for member, value in members.items():
            mine = getattr(enum_class, member)
            assert mine.value == value, f"{name}.{member}"
            if theirs is not None and hasattr(theirs, member):
                assert getattr(theirs, member).value == value, f"{name}.{member}"
            checked += 1
    assert checked >= 400, f"only {checked} members checked"


def test_aliases_point_at_the_same_class():
    assert ours.WD_ALIGN_PARAGRAPH is ours.WD_PARAGRAPH_ALIGNMENT
    assert ours.WD_INLINE_SHAPE is ours.WD_INLINE_SHAPE_TYPE
    assert ours.WD_COLOR is ours.WD_COLOR_INDEX


def test_members_compare_like_python_docx_members():
    assert ours.WD_ALIGN_PARAGRAPH.CENTER == 1
    assert int(ours.WD_LINE_SPACING.DOUBLE) == 2
    assert ours.WD_ORIENT.LANDSCAPE.name == "LANDSCAPE"
    # str() is python-docx's, not the token: a program that printed a member
    # before the switch prints the same thing after it
    assert str(ours.WD_ALIGN_PARAGRAPH.CENTER) == "CENTER (1)"
    assert repr(ours.WD_ALIGN_PARAGRAPH.CENTER) == \
        "<WD_PARAGRAPH_ALIGNMENT.CENTER: 1>"
    # ...and the token the file wants is a property away
    assert ours.WD_ALIGN_PARAGRAPH.CENTER.xml_value == "center"
    assert ours.WD_COLOR_INDEX.YELLOW.xml_value == "yellow"
    assert ours.WD_SECTION_START.ODD_PAGE.xml_value == "oddPage"
    assert ours.WD_ALIGN_PARAGRAPH.CENTER.ooxml == "center"      # our name


def test_members_are_enum_members_not_lookalikes():
    """The documented protocol, not just the names: iteration, lookup, isinstance."""
    from docviews.compat.enum.base import BaseEnum, BaseXmlEnum

    member = ours.WD_ALIGN_PARAGRAPH.CENTER
    assert isinstance(member, int) and isinstance(member, ours.WD_ALIGN_PARAGRAPH)
    assert isinstance(member, BaseXmlEnum) and not isinstance(member, BaseEnum)
    assert issubclass(ours.WD_ALIGN_PARAGRAPH, BaseXmlEnum)
    assert list(ours.WD_ALIGN_PARAGRAPH)[:2] == [ours.WD_ALIGN_PARAGRAPH.LEFT,
                                                 ours.WD_ALIGN_PARAGRAPH.CENTER]
    assert ours.WD_ALIGN_PARAGRAPH(1) is member
    assert ours.WD_ALIGN_PARAGRAPH["CENTER"] is member
    assert sorted([0, 1]) == sorted(int(m) for m in
                                     [ours.WD_ORIENT.LANDSCAPE,
                                      ours.WD_ORIENT.PORTRAIT])
    assert ours.WD_ALIGN_PARAGRAPH.from_xml("center") is member
    assert ours.WD_ALIGN_PARAGRAPH.to_xml(member) == "center"
    with pytest.raises(ValueError):
        ours.WD_ALIGN_PARAGRAPH.from_xml("nowhere")


def test_the_per_module_imports_work():
    """``from docx.enum.text import …`` — the import line, not just the name."""
    from docviews.compat.enum import dml, section, shape, style, table, text
    from docviews.compat.enum.base import BaseXmlEnum as _  # noqa: F401

    assert text.WD_ALIGN_PARAGRAPH is ours.WD_PARAGRAPH_ALIGNMENT
    assert text.WD_BREAK is ours.WD_BREAK_TYPE
    assert style.WD_STYLE_TYPE is ours.WD_STYLE_TYPE
    assert section.WD_ORIENT is ours.WD_ORIENTATION
    assert table.WD_ROW_HEIGHT_RULE is ours.WD_ROW_HEIGHT
    assert shape.WD_INLINE_SHAPE is ours.WD_INLINE_SHAPE_TYPE
    assert dml.MSO_THEME_COLOR is ours.MSO_THEME_COLOR_INDEX
    assert ours.WD_BREAK.PAGE.xml_value == "page"        # ours, python-docx has none
    assert str(ours.WD_BREAK.PAGE) == "WD_BREAK_TYPE.PAGE"


def test_every_member_matches_python_docx_member_by_member():
    """Value, str(), repr() and xml token — 471 members, against python-docx."""
    # The only deliberate differences; both are about what Word's XML accepts,
    # and both are documented in COMPAT.md.
    known = {
        ("WD_COLOR_INDEX", "AUTO", "xml_value"),        # "default" vs "none"
        ("WD_COLOR", "AUTO", "xml_value"),
        ("WD_LINE_SPACING", "SINGLE", "xml_value"),     # python-docx: "UNMAPPED"
        ("WD_LINE_SPACING", "ONE_POINT_FIVE", "xml_value"),
        ("WD_LINE_SPACING", "DOUBLE", "xml_value"),
        # python-docx's "UNMAPPED" marker; None here, so the writer never
        # writes a placeholder token into a file
        ("MSO_THEME_COLOR", "NOT_THEME_COLOR", "xml_value"),
        ("MSO_THEME_COLOR_INDEX", "NOT_THEME_COLOR", "xml_value"),
    }
    checked, additions, problems = 0, 0, []
    for module_name in ("docx.enum.text", "docx.enum.style", "docx.enum.section",
                        "docx.enum.table", "docx.enum.shape", "docx.enum.dml"):
        module = __import__(module_name, fromlist=["*"])
        for name in dir(module):
            if not (name.startswith("WD_") or name.startswith("MSO_")):
                continue
            theirs = getattr(module, name)
            if not isinstance(theirs, type) or not hasattr(theirs, "__members__"):
                continue
            mine = getattr(ours, name, None)
            assert mine is not None, f"{name} is missing"
            assert list(mine.__members__) == list(theirs.__members__), name
            for member in theirs.__members__:
                theirs_member, our_member = getattr(theirs, member), getattr(mine, member)
                checked += 1
                for field, a, b in (("value", theirs_member.value, our_member.value),
                                    ("str", str(theirs_member), str(our_member)),
                                    ("repr", repr(theirs_member), repr(our_member))):
                    if a != b and (name, member, field) not in known:
                        problems.append(f"{name}.{member}.{field}: {a!r} != {b!r}")
                their_token = getattr(theirs_member, "xml_value", None)
                our_token = getattr(our_member, "xml_value", None)
                if their_token != our_token:
                    if their_token is None:
                        additions += 1         # ours knows a token python-docx does not
                    elif (name, member, "xml_value") not in known:
                        problems.append(f"{name}.{member}.xml_value: "
                                        f"{their_token!r} != {our_token!r}")
    assert checked == 471, f"compared {checked} members"
    assert problems == [], problems
    assert additions >= 20, additions


def test_every_enum_name_python_docx_defines_is_importable():
    missing = []
    for module_name in ("docx.enum.text", "docx.enum.style", "docx.enum.section",
                        "docx.enum.table", "docx.enum.shape", "docx.enum.dml"):
        module = __import__(module_name, fromlist=["*"])
        for name in dir(module):
            if name.startswith("_") or not name.startswith("WD_"):
                continue
            if not isinstance(getattr(module, name), type):
                continue
            if getattr(getattr(module, name), "__module__", "").startswith("docx.enum") \
                    and not hasattr(ours, name):
                missing.append(f"{module_name}.{name}")
    assert missing == []


def test_enum_members_reach_the_file_as_tokens():
    from docviews import Document
    from docviews.compat.enum import (WD_ALIGN_PARAGRAPH,
                                      WD_CELL_VERTICAL_ALIGNMENT,
                                      WD_COLOR_INDEX, WD_ORIENT,
                                      WD_ROW_HEIGHT_RULE, WD_SECTION_START,
                                      WD_TABLE_ALIGNMENT, WD_UNDERLINE)

    doc = Document.new()
    paragraph = doc.add_paragraph("centered", align=WD_ALIGN_PARAGRAPH.CENTER)
    run = paragraph.add_run("highlighted")
    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    run.underline = WD_UNDERLINE.DOUBLE
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.cell(0, 0).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    table.rows[0].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.start_type = WD_SECTION_START.ODD_PAGE

    xml = zipfile.ZipFile(io.BytesIO(doc.to_bytes())).read(
        "word/document.xml").decode()
    assert re.findall(r'<w:jc w:val="([^"]+)"', xml) == ["center", "center"]
    assert re.findall(r'<w:highlight w:val="([^"]+)"', xml) == ["yellow"]
    assert re.findall(r'<w:u w:val="([^"]+)"', xml) == ["double"]
    assert re.findall(r'<w:vAlign w:val="([^"]+)"', xml) == ["center"]
    assert re.findall(r'w:hRule="([^"]+)"', xml) == ["atLeast"]
    assert re.findall(r'<w:type w:val="([^"]+)"', xml) == ["oddPage"]
    assert re.findall(r'w:orient="([^"]+)"', xml) == ["landscape"]
