"""The compat enums: same members, same values, and a usable token.

python-docx code says ``WD_ALIGN_PARAGRAPH.CENTER``, compares it to ``1``, and
expects the file to end up with ``w:jc="center"``. All three have to hold here,
so all three are checked — including against python-docx itself, member by
member, because a member list that drifts silently is a migration bug.
"""

from __future__ import annotations

import enum
import importlib
import io
import json
import re
import zipfile
from pathlib import Path

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
    for alias, canonical in _generated.ALIASES.items():
        assert getattr(ours, alias) is getattr(ours, canonical)
    assert ours.WD_ALIGN_PARAGRAPH is ours.WD_PARAGRAPH_ALIGNMENT
    assert ours.WD_INLINE_SHAPE is ours.WD_INLINE_SHAPE_TYPE
    assert ours.WD_COLOR is ours.WD_COLOR_INDEX


def test_documented_enum_inventory_matches_runtime_and_proxy_surface():
    """Check the source-backed enum inventory against both runtimes."""
    inventory_path = (Path(__file__).resolve().parents[1] / "docs" / "api-parity"
                      / "python_docx_1_2_0_inventory.json")
    records = json.loads(inventory_path.read_text(encoding="utf-8"))["records"]
    enum_records = [record for record in records
                    if record["kind"].startswith("enum_")]

    def resolve_enum(name, root):
        parts = name.split(".")
        assert parts[:2] == ["docx", "enum"]
        module = importlib.import_module(f"{root}.{parts[2]}")
        return module, parts[3:]

    class_count = alias_count = member_count = documented_only_count = 0
    for record in enum_records:
        path = record["qualified_name"]
        reference_module, reference_parts = resolve_enum(path, "docx.enum")
        candidate_module, candidate_parts = resolve_enum(
            path, "docviews.compat.enum")
        if record["kind"] == "enum_class":
            reference_class = getattr(reference_module, reference_parts[-1])
            candidate_class = getattr(candidate_module, candidate_parts[-1])
            reference_names = list(reference_class.__members__)
            candidate_names = list(candidate_class.__members__)
            if path == "docx.enum.dml.MSO_THEME_COLOR_INDEX":
                # Follow the published enum page while recording the one
                # member missing from python-docx 1.2.0's runtime.
                assert candidate_names == reference_names + ["MIXED"]
                assert candidate_class.MIXED.value == -2
            else:
                assert candidate_names == reference_names
            class_count += 1
        elif record["kind"] == "enum_alias":
            alias = getattr(reference_module, reference_parts[-1])
            candidate_alias = getattr(candidate_module, candidate_parts[-1])
            _, target_parts = resolve_enum(record["alias_of"], "docx.enum")
            target = getattr(reference_module, target_parts[-1])
            _, candidate_target_parts = resolve_enum(
                record["alias_of"], "docviews.compat.enum")
            candidate_target = getattr(candidate_module,
                                       candidate_target_parts[-1])
            assert alias is target
            assert candidate_alias is candidate_target
            alias_count += 1
        else:
            reference_class = getattr(reference_module,
                                      reference_parts[-2])
            candidate_class = getattr(candidate_module,
                                      candidate_parts[-2])
            member_name = reference_parts[-1]
            if not record.get("present_in_python_docx_1_2_0", True):
                # The docs page lists MIXED, the runtime omits it, and the
                # Office API supplies its value. The candidate follows docs.
                assert path == "docx.enum.dml.MSO_THEME_COLOR_INDEX.MIXED"
                assert record.get("python_docx_value") is None
                assert record.get("documented_value") == -2
                assert not hasattr(reference_class, member_name)
                candidate_member = getattr(candidate_class, member_name)
                assert candidate_member.value == record["documented_value"]
                assert candidate_member.xml_value is None
                assert candidate_member.__doc__ == record["behavior"]
                documented_only_count += 1
                continue
            reference_member = getattr(reference_class, member_name)
            candidate_member = getattr(candidate_class, member_name)
            assert reference_member.value == record["python_docx_value"]
            assert candidate_member.value == reference_member.value
            assert str(candidate_member) == str(reference_member)
            assert repr(candidate_member) == repr(reference_member)
            assert getattr(candidate_member, "xml_value", None) == \
                getattr(reference_member, "xml_value", None)
            member_count += 1

    assert (class_count, alias_count, member_count, documented_only_count) == \
        (18, 10, 245, 1)


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
    assert ours.WD_LINE_SPACING.from_xml("UNMAPPED") is ours.WD_LINE_SPACING.SINGLE
    assert ours.WD_LINE_SPACING.to_xml(ours.WD_LINE_SPACING.SINGLE) == "UNMAPPED"
    assert ours.MSO_THEME_COLOR_INDEX.to_xml(
        ours.MSO_THEME_COLOR_INDEX.NOT_THEME_COLOR) == "UNMAPPED"
    with pytest.raises(ValueError):
        ours.WD_ALIGN_PARAGRAPH.from_xml("nowhere")

    # These legacy enums derive directly from Enum in python-docx, not from
    # BaseEnum/BaseXmlEnum; keep their standard str/repr and non-int members.
    assert ours.WD_BREAK_TYPE.__bases__ == (enum.Enum,)
    assert ours.WD_BREAK.PAGE.value == docx.enum.text.WD_BREAK.PAGE.value
    assert str(ours.WD_BREAK.PAGE) == str(docx.enum.text.WD_BREAK.PAGE)
    assert repr(ours.WD_BREAK.PAGE) == repr(docx.enum.text.WD_BREAK.PAGE)
    assert not isinstance(ours.WD_BREAK.PAGE, int)
    assert not hasattr(ours.WD_BREAK.PAGE, "xml_value")
    assert ours.WD_INLINE_SHAPE_TYPE.__bases__ == (enum.Enum,)
    assert not isinstance(ours.WD_INLINE_SHAPE_TYPE.PICTURE, int)


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
    assert not hasattr(ours.WD_BREAK.PAGE, "xml_value")
    assert str(ours.WD_BREAK.PAGE) == "WD_BREAK_TYPE.PAGE"


def test_every_member_matches_python_docx_member_by_member():
    """Value, str(), repr(), and xml_value match for all 471 runtime members."""
    checked, problems = 0, []
    doc_only_aliases = {"MSO_THEME_COLOR", "MSO_THEME_COLOR_INDEX"}
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
            reference_names = list(theirs.__members__)
            candidate_names = list(mine.__members__)
            if name in doc_only_aliases:
                assert candidate_names == reference_names + ["MIXED"]
                assert mine.MIXED.value == -2
                assert mine.MIXED.xml_value is None
            else:
                assert candidate_names == reference_names, name
            for member in theirs.__members__:
                theirs_member, our_member = getattr(theirs, member), getattr(mine, member)
                checked += 1
                for field, a, b in (("value", theirs_member.value, our_member.value),
                                    ("str", str(theirs_member), str(our_member)),
                                    ("repr", repr(theirs_member), repr(our_member))):
                    if a != b:
                        problems.append(f"{name}.{member}.{field}: {a!r} != {b!r}")
                their_token = getattr(theirs_member, "xml_value", None)
                our_token = getattr(our_member, "xml_value", None)
                if their_token != our_token:
                    problems.append(f"{name}.{member}.xml_value: "
                                    f"{their_token!r} != {our_token!r}")
    assert checked == 471, f"compared {checked} members"
    assert problems == [], problems


def test_python_docx_xml_sentinels_reach_the_writer_unchanged():
    import docviews.compat as compat
    from docviews.write import token

    assert token(ours.WD_COLOR_INDEX.AUTO) == "default"
    assert token(ours.WD_LINE_SPACING.SINGLE) == "UNMAPPED"
    assert token(ours.MSO_THEME_COLOR_INDEX.NOT_THEME_COLOR) == "UNMAPPED"
    assert token(ours.WD_UNDERLINE.INHERITED) is None

    reference = docx.Document()
    candidate = compat.Document()
    reference_run = reference.add_paragraph().add_run("highlight")
    candidate_run = candidate.add_paragraph().add_run("highlight")
    reference_run.font.highlight_color = docx.enum.text.WD_COLOR_INDEX.AUTO
    candidate_run.font.highlight_color = ours.WD_COLOR_INDEX.AUTO

    payload = io.BytesIO()
    reference.save(payload)
    reference_xml = zipfile.ZipFile(io.BytesIO(payload.getvalue())).read(
        "word/document.xml").decode()
    candidate_xml = zipfile.ZipFile(io.BytesIO(candidate.to_bytes())).read(
        "word/document.xml").decode()
    assert re.findall(r'<w:highlight w:val="([^"]+)"', reference_xml) == ["default"]
    assert re.findall(r'<w:highlight w:val="([^"]+)"', candidate_xml) == ["default"]

    with pytest.raises(ValueError, match="WD_UNDERLINE.INHERITED has no XML representation"):
        reference_run.underline = docx.enum.text.WD_UNDERLINE.INHERITED
    with pytest.raises(ValueError, match="WD_UNDERLINE.INHERITED has no XML representation"):
        candidate_run.underline = ours.WD_UNDERLINE.INHERITED
    with pytest.raises(ValueError, match="MSO_THEME_COLOR_INDEX.MIXED has no XML representation"):
        candidate_run.font.color.theme_color = ours.MSO_THEME_COLOR_INDEX.MIXED


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
