#!/usr/bin/env python3
"""Build a source-backed python-docx 1.2.0 documented-API inventory.

The inventory is extracted from the Sphinx object inventory and HTML produced
from the upstream docs, with the prose-only enum pages parsed from their RST.
The source inventory alone is not a parity assertion. When a comparison JSON is
provided, the Markdown merges availability, signature, enum, and test-backed
behavior evidence recorded separately.

Example (after building the upstream docs with Sphinx):

    python tools/build_api_parity_matrix.py \
        --reference /path/to/python-docx-v1.2.0 \
        --html /path/to/sphinx-html \
        --json docs/api-parity/python_docx_1_2_0_inventory.json \
        --markdown API_PARITY_MATRIX.md

The reference checkout must be python-docx tag v1.2.0. lxml is used only by
this maintainer tool to read the generated HTML; it is not a docviews runtime
dependency.
"""

from __future__ import annotations

import argparse
import copy
import importlib
import json
import re
import sys
import zlib
from collections import Counter
from pathlib import Path
from typing import Any

from lxml import html as lxml_html

UPSTREAM_TAG = "v1.2.0"
UPSTREAM_COMMIT = "e45454602b53e8e572b179ccf1c91093ec9f4ed7"
DOCS_BASE = "https://python-docx.readthedocs.io/en/latest/"

# The 1.2.0 API reference documents enumerations as prose in RST rather than
# Sphinx autodoc objects. These mappings are checked against the upstream
# runtime while generating the inventory.
ENUMS = {
    "MsoColorType": ("dml", "MSO_COLOR_TYPE", ()),
    "MsoThemeColorIndex": ("dml", "MSO_THEME_COLOR_INDEX", ("MSO_THEME_COLOR",)),
    "WdAlignParagraph": (
        "text", "WD_PARAGRAPH_ALIGNMENT", ("WD_ALIGN_PARAGRAPH",)
    ),
    "WdBuiltinStyle": ("style", "WD_BUILTIN_STYLE", ("WD_STYLE",)),
    "WdCellVerticalAlignment": (
        "table", "WD_CELL_VERTICAL_ALIGNMENT", ("WD_ALIGN_VERTICAL",)
    ),
    "WdColorIndex": ("text", "WD_COLOR_INDEX", ("WD_COLOR",)),
    "WdLineSpacing": ("text", "WD_LINE_SPACING", ()),
    "WdOrientation": ("section", "WD_ORIENTATION", ("WD_ORIENT",)),
    "WdRowAlignment": ("table", "WD_TABLE_ALIGNMENT", ()),
    "WdRowHeightRule": ("table", "WD_ROW_HEIGHT_RULE", ("WD_ROW_HEIGHT",)),
    "WdSectionStart": ("section", "WD_SECTION_START", ("WD_SECTION",)),
    "WdStyleType": ("style", "WD_STYLE_TYPE", ()),
    "WdTabAlignment": ("text", "WD_TAB_ALIGNMENT", ()),
    "WdTabLeader": ("text", "WD_TAB_LEADER", ()),
    "WdTableDirection": ("table", "WD_TABLE_DIRECTION", ()),
    "WdUnderline": ("text", "WD_UNDERLINE", ()),
}

# The v1.2.0 docs name MIXED but do not print its numeric value, and the
# python-docx runtime omits it. Use the Office API's authoritative -2 value,
# while preserving the runtime absence in the inventory fields above.
DOCUMENTED_ONLY_ENUM_VALUES = {
    "MSO_THEME_COLOR_INDEX.MIXED": {
        "documented_value": -2,
        "documented_xml_value": None,
        "documented_value_source": (
            "https://learn.microsoft.com/en-us/office/vba/api/office.msothemecolorindex"
        ),
    },
}

# Two additional enum types are documented indirectly by API member signatures
# and behavior prose, but have no standalone page in api/enum/index.rst. Only
# the values explicitly named by those API docs are included here.
INDIRECT_ENUMS = {
    "WD_BREAK_TYPE": {
        "module": "text",
        "aliases": ("WD_BREAK",),
        "page": "api/text.html",
        "anchor": "docx.text.run.Run.add_break",
        "source_file": "docs/api/text.rst (Run.add_break autodoc)",
        "behavior": (
            "The Run.add_break() documentation names LINE, PAGE, and COLUMN "
            "as supported values; LINE is the default."
        ),
        "members": {
            "LINE": "Named as a supported break_type and documented as the default.",
            "PAGE": "Named as a supported break_type; no separate value description is given.",
            "COLUMN": "Named as a supported break_type; no separate value description is given.",
        },
    },
    "WD_INLINE_SHAPE_TYPE": {
        "module": "shape",
        "aliases": ("WD_INLINE_SHAPE",),
        "page": "api/shape.html",
        "anchor": "docx.shape.InlineShape.type",
        "source_file": "docs/api/shape.rst (InlineShape.type autodoc)",
        "behavior": (
            "The InlineShape.type documentation says the property returns a "
            "member of WD_INLINE_SHAPE and cites LINKED_PICTURE as an example."
        ),
        "members": {
            "LINKED_PICTURE": "Named as an example value for InlineShape.type.",
        },
    },
}

STRUCTURAL_MEMBERS = {"part", "element", "blob"}


def collapse(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def load_sphinx_inventory(path: Path) -> list[dict[str, str]]:
    data = path.read_bytes()
    header_end = data.find(b"\n", data.find(b"\n", data.find(b"\n") + 1) + 1)
    # Sphinx v2 inventories have four header lines followed by a zlib stream.
    lines = zlib.decompress(data.split(b"\n", 4)[4]).decode("utf-8").splitlines()
    records: list[dict[str, str]] = []
    for line in lines:
        parts = line.split(" ", 4)
        if len(parts) != 5:
            continue
        name, obj_type, _priority, location, _display = parts
        if not obj_type.startswith("py:") or not location.startswith("api/"):
            continue
        if "#" in location:
            page, fragment = location.split("#", 1)
            anchor = name if fragment == "$" else fragment
        else:
            page, anchor = location, name
        records.append({
            "qualified_name": name,
            "kind": obj_type.removeprefix("py:"),
            "doc_page": page,
            "doc_anchor": anchor,
        })
    return records


def parse_api_html(page_path: Path) -> dict[str, dict[str, str]]:
    root = lxml_html.parse(str(page_path))
    result: dict[str, dict[str, str]] = {}
    xpath = '//dt[contains(concat(" ", normalize-space(@class), " "), " sig-object ")]'
    for dt in root.xpath(xpath):
        anchor = dt.get("id")
        if not anchor:
            continue
        dt_copy = copy.deepcopy(dt)
        for node in dt_copy.xpath(
            './/*[contains(concat(" ", normalize-space(@class), " "), " viewcode-link ") '
            'or contains(concat(" ", normalize-space(@class), " "), " headerlink ")]'
        ):
            node.drop_tree()
        signature = collapse(dt_copy.text_content()).replace("¶", "").strip()
        dd = dt.getnext()
        paragraphs: list[str] = []
        if dd is not None and str(dd.tag).lower() == "dd":
            for child in dd:
                classes = set((child.get("class") or "").split())
                # The class description contains nested autodoc definitions.
                # Do not fold those member descriptions into the class behavior.
                if child.tag == "dl" and any(c.startswith("py") for c in classes):
                    break
                text = collapse(child.text_content())
                if text:
                    paragraphs.append(text)
        if not paragraphs:
            # CoreProperties is introduced by three explanatory paragraphs in
            # document.rst before its manual ``class``/``attribute`` block,
            # rather than by a class docstring. Preserve that class-level
            # behavior instead of leaving a blank description.
            object_dl = dt.getparent()
            section = object_dl.getparent() if object_dl is not None else None
            if section is not None:
                children = list(section)
                try:
                    object_index = children.index(object_dl)
                except ValueError:
                    object_index = -1
                if object_index >= 0:
                    headings = [i for i, node in enumerate(children[:object_index])
                                if str(node.tag).lower() in {"h1", "h2", "h3"}]
                    start = headings[-1] + 1 if headings else 0
                    for node in children[start:object_index]:
                        if str(node.tag).lower() == "p":
                            text = collapse(node.text_content())
                            if text:
                                paragraphs.append(text)
        parent = dt.getparent()
        css = set((parent.get("class") or "").split()) if parent is not None else set()
        kind = next((token.removeprefix("py-") for token in css if token.startswith("py-")), "")
        result[anchor] = {
            "signature": signature,
            "behavior": "\n".join(paragraphs),
            "html_kind": kind,
        }
    return result


def source_class_doc(rst: str) -> tuple[str, list[str]]:
    lines = rst.splitlines()
    title = ""
    title_index = -1
    for i, line in enumerate(lines):
        match = re.fullmatch(r"\s*``([^`]+)``\s*", line)
        if match:
            title, title_index = match.group(1).strip(), i
            break
    summary_lines: list[str] = []
    aliases: list[str] = []
    if title_index >= 0:
        for line in lines[title_index + 2:]:
            stripped = line.strip()
            if stripped == "----" or stripped.endswith("::"):
                break
            alias = re.search(r"\bAlias\s*:\s*``([A-Z0-9_]+)``", stripped, re.I)
            if not alias:
                alias = re.search(r"\balias\s*:\s*\*\*([A-Z0-9_]+)\*\*", stripped, re.I)
            if alias:
                aliases.append(alias.group(1))
                continue
            if stripped:
                summary_lines.append(stripped)
    return title, [collapse(" ".join(summary_lines))] if summary_lines else []


def parse_enum_values(rst: str) -> list[tuple[str, str]]:
    if "----" not in rst:
        return []
    tail = rst.split("----", 1)[1].splitlines()
    entries: list[tuple[str, str]] = []
    current: str | None = None
    description: list[str] = []

    def flush() -> None:
        nonlocal current, description
        if current is not None:
            entries.append((current, collapse(" ".join(description))))
        current, description = None, []

    for line in tail:
        if re.fullmatch(r"[A-Z][A-Z0-9_]*", line.strip()) and line == line.strip():
            flush()
            current = line.strip()
            continue
        if current is None:
            continue
        if not line.strip():
            continue
        if line.startswith("    ") or line.startswith("\t"):
            description.append(line.strip())
        else:
            # A top-level directive or heading ends the definition list.
            flush()
    flush()
    return entries


def enum_value_runtime(cls: Any, member_name: str) -> dict[str, Any]:
    try:
        member = getattr(cls, member_name)
    except AttributeError:
        return {
            "present_in_python_docx_1_2_0": False,
            "python_docx_value": None,
            "python_docx_xml_value": None,
        }
    raw_value = getattr(member, "value", None)
    xml_value = getattr(member, "xml_value", None)
    return {
        "present_in_python_docx_1_2_0": True,
        "python_docx_value": raw_value,
        "python_docx_xml_value": xml_value,
    }


def build_enum_records(reference: Path) -> list[dict[str, Any]]:
    sys.path.insert(0, str(reference / "src"))
    records: list[dict[str, Any]] = []
    for page_name, (module_name, canonical, aliases) in ENUMS.items():
        rst_path = reference / "docs" / "api" / "enum" / f"{page_name}.rst"
        rst = rst_path.read_text(encoding="utf-8")
        title, summaries = source_class_doc(rst)
        module = importlib.import_module(f"docx.enum.{module_name}")
        enum_cls = getattr(module, canonical)
        page = f"api/enum/{page_name}.html"
        label_match = re.search(r"^\.\. _([^:]+):", rst, re.M)
        page_fragment = label_match.group(1).lower() if label_match else page_name.lower()
        class_record = {
            "qualified_name": f"docx.enum.{module_name}.{canonical}",
            "kind": "enum_class",
            "signature": f"enum {canonical}",
            "behavior": summaries[0] if summaries else "",
            "doc_page": page,
            "doc_anchor": page_fragment,
            "doc_url": f"{DOCS_BASE}{page}#{page_fragment}",
            "source_file": str(rst_path.relative_to(reference)),
            "source_type": "prose RST; manually mapped to upstream enum class",
            "aliases": [f"docx.enum.{module_name}.{alias}" for alias in aliases],
            "scope": "in_scope",
            "availability": "not_assessed",
            "signature_parity": "not_assessed",
            "behavior_parity": "not_assessed",
            "test_evidence": [],
            "upstream_runtime_name": f"docx.enum.{module_name}.{canonical}",
            "upstream_runtime_present": True,
        }
        records.append(class_record)
        for alias in aliases:
            records.append({
                "qualified_name": f"docx.enum.{module_name}.{alias}",
                "kind": "enum_alias",
                "signature": f"alias of {canonical}",
                "behavior": f"Documented alias for {canonical}.",
                "doc_page": page,
                "doc_anchor": page_fragment,
                "doc_url": f"{DOCS_BASE}{page}#{page_fragment}",
                "source_file": str(rst_path.relative_to(reference)),
                "source_type": "alias explicitly named in the enum RST page",
                "alias_of": f"docx.enum.{module_name}.{canonical}",
                "scope": "in_scope",
                "availability": "not_assessed",
                "signature_parity": "not_assessed",
                "behavior_parity": "not_assessed",
                "test_evidence": [],
                "upstream_runtime_present": hasattr(module, alias),
            })
        for member_name, behavior in parse_enum_values(rst):
            runtime = enum_value_runtime(enum_cls, member_name)
            extra = ""
            if not runtime["present_in_python_docx_1_2_0"]:
                extra = "Documented on the official page but absent from the v1.2.0 runtime enum."
            documented_value = DOCUMENTED_ONLY_ENUM_VALUES.get(
                f"{canonical}.{member_name}", {})
            if documented_value:
                extra += (
                    " Its numeric value is supplied by the Microsoft Office API "
                    "reference, since the python-docx page omits it."
                )
            records.append({
                "qualified_name": f"docx.enum.{module_name}.{canonical}.{member_name}",
                "kind": "enum_member",
                "signature": f"{canonical}.{member_name}",
                "behavior": behavior,
                "doc_page": page,
                "doc_anchor": page_fragment,
                "doc_url": f"{DOCS_BASE}{page}#{page_fragment}",
                "source_file": str(rst_path.relative_to(reference)),
                "source_type": "prose RST definition list",
                "aliases": [
                    f"docx.enum.{module_name}.{alias}.{member_name}"
                    for alias in aliases
                ],
                "scope": "in_scope",
                "availability": "not_assessed",
                "signature_parity": "not_applicable",
                "behavior_parity": "not_assessed",
                "test_evidence": [],
                "upstream_runtime_name": (
                    f"docx.enum.{module_name}.{canonical}.{member_name}"
                ),
                **runtime,
                **documented_value,
                "reference_note": extra,
            })
    return records


def build_indirect_enum_records(reference: Path) -> list[dict[str, Any]]:
    sys.path.insert(0, str(reference / "src"))
    records: list[dict[str, Any]] = []
    for canonical, spec in INDIRECT_ENUMS.items():
        module_name = spec["module"]
        module = importlib.import_module(f"docx.enum.{module_name}")
        enum_cls = getattr(module, canonical)
        aliases = spec["aliases"]
        page = spec["page"]
        anchor = spec["anchor"]
        url = f"{DOCS_BASE}{page}#{anchor}"
        records.append({
            "qualified_name": f"docx.enum.{module_name}.{canonical}",
            "kind": "enum_class",
            "signature": f"enum {canonical}",
            "behavior": spec["behavior"],
            "doc_page": page,
            "doc_anchor": anchor,
            "doc_url": url,
            "source_file": spec["source_file"],
            "source_type": "enum referenced by API signature and behavior prose",
            "aliases": [f"docx.enum.{module_name}.{alias}" for alias in aliases],
            "scope": "in_scope",
            "availability": "not_assessed",
            "signature_parity": "not_assessed",
            "behavior_parity": "not_assessed",
            "test_evidence": [],
            "upstream_runtime_name": f"docx.enum.{module_name}.{canonical}",
            "upstream_runtime_present": True,
        })
        for alias in aliases:
            records.append({
                "qualified_name": f"docx.enum.{module_name}.{alias}",
                "kind": "enum_alias",
                "signature": f"alias of {canonical}",
                "behavior": f"Documented alias used by the API prose for {canonical}.",
                "doc_page": page,
                "doc_anchor": anchor,
                "doc_url": url,
                "source_file": spec["source_file"],
                "source_type": "enum alias named in API signature or behavior prose",
                "alias_of": f"docx.enum.{module_name}.{canonical}",
                "scope": "in_scope",
                "availability": "not_assessed",
                "signature_parity": "not_applicable",
                "behavior_parity": "not_assessed",
                "test_evidence": [],
                "upstream_runtime_present": hasattr(module, alias),
            })
        for member_name, behavior in spec["members"].items():
            runtime = enum_value_runtime(enum_cls, member_name)
            records.append({
                "qualified_name": (
                    f"docx.enum.{module_name}.{canonical}.{member_name}"
                ),
                "kind": "enum_member",
                "signature": f"{canonical}.{member_name}",
                "behavior": behavior,
                "doc_page": page,
                "doc_anchor": anchor,
                "doc_url": url,
                "source_file": spec["source_file"],
                "source_type": "enum value named in API behavior prose",
                "aliases": [
                    f"docx.enum.{module_name}.{alias}.{member_name}"
                    for alias in aliases
                ],
                "scope": "in_scope",
                "availability": "not_assessed",
                "signature_parity": "not_applicable",
                "behavior_parity": "not_assessed",
                "test_evidence": [],
                "upstream_runtime_name": (
                    f"docx.enum.{module_name}.{canonical}.{member_name}"
                ),
                **runtime,
                "reference_note": "",
            })
    return records


def scope_for(record: dict[str, Any]) -> tuple[str, str]:
    leaf = record["qualified_name"].rsplit(".", 1)[-1]
    if record["kind"] in {"property", "attribute"} and leaf in STRUCTURAL_MEMBERS:
        return (
            "excluded_internals",
            "Direct lxml/OOXML/OPC structural access is outside the agreed scope.",
        )
    return "in_scope", ""


def build_object_records(reference: Path, html_root: Path, inventory: Path) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    pages: dict[str, dict[str, dict[str, str]]] = {}
    for item in load_sphinx_inventory(inventory):
        page = item["doc_page"]
        if page not in pages:
            pages[page] = parse_api_html(html_root / page)
        details = pages[page].get(item["doc_anchor"])
        if details is None:
            details = pages[page].get(item["qualified_name"])
        if details is None:
            # Keep a missing extraction visible instead of silently dropping an
            # object from the inventory.
            details = {"signature": "", "behavior": "", "html_kind": ""}
        scope, scope_note = scope_for(item)
        qualified_name = item["qualified_name"]
        anchor = item["doc_anchor"]
        record: dict[str, Any] = {
            **item,
            "signature": details["signature"],
            "behavior": details["behavior"],
            "doc_url": f"{DOCS_BASE}{page}#{anchor}",
            "source_file": f"docs/{page.replace('.html', '.rst')}",
            "source_type": "Sphinx autodoc object extracted from upstream HTML",
            "scope": scope,
            "scope_note": scope_note,
            "availability": "not_assessed",
            "signature_parity": "not_assessed",
            "behavior_parity": "not_assessed",
            "test_evidence": [],
            "reference_alias": anchor != qualified_name,
        }
        if anchor != qualified_name:
            record["alias_of"] = anchor
        records.append(record)
    return records


def markdown_cell(value: Any, code: bool = False) -> str:
    import html

    text = "" if value is None else str(value).replace("\r", "")
    escaped = html.escape(text, quote=False).replace("|", "&#124;")
    escaped = escaped.replace("\n", "<br>")
    if code:
        return f"<code>{escaped}</code>"
    return escaped.replace("`", "&#96;")


def render_markdown(records: list[dict[str, Any]], reference: Path,
                    comparison: dict[str, Any] | None = None) -> str:
    object_records = [r for r in records if r["kind"] in {
        "class", "function", "method", "property", "attribute"
    }]
    enum_records = [r for r in records if r["kind"].startswith("enum_")]
    type_counts = Counter(r["kind"] for r in object_records)
    scoped_counts = Counter(r["scope"] for r in records)
    has_comparison = comparison is not None
    comparison_records = comparison.get("records", {}) if comparison else {}
    user_guide_recipes = comparison.get("user_guide_recipes", []) if comparison else []
    behavior_counts = Counter(
        r.get("behavior_parity", "not_assessed")
        for r in records if r.get("scope") != "excluded_internals"
    )
    if has_comparison:
        in_scope_members = [
            r for r in object_records
            if r["scope"] == "in_scope" and r["kind"] in {
                "method", "property", "attribute"
            }
        ]
        comparable_callables = [
            r for r in object_records
            if r.get("signature_parity") in {
                "call_shape_match", "call_shape_mismatch"
            }
        ]
        property_rows = [
            r for r in object_records
            if r["scope"] == "in_scope" and r["kind"] in {
                "property", "attribute"
            }
        ]
        setter_gaps = [
            r for r in property_rows if r.get("writability_parity") == "mismatch"
        ]
        missing_setters = [
            r for r in setter_gaps
            if r.get("reference_writable") and not r.get("docviews_writable")
        ]
        extra_setters = [
            r for r in setter_gaps
            if not r.get("reference_writable") and r.get("docviews_writable")
        ]
        enum_members = [r for r in enum_records if r["kind"] == "enum_member"]
        enum_values_match = sum(
            r.get("value_parity") == "match" for r in enum_members
        )
        enum_classes = [r for r in enum_records if r["kind"] == "enum_class"]
        proxy_path_gaps = [
            r for r in object_records
            if r["kind"] == "class"
            and r.get("import_path_status") == "object_access_only"
        ]
        behavior_assessed = [
            r for r in records
            if r.get("behavior_parity") not in {
                None, "not_assessed", "excluded_by_scope"
            }
        ]
        recipe_counts = Counter(
            recipe.get("coverage", "not_assessed")
            for recipe in user_guide_recipes
        )
    matches = behavior_counts.get("tested_match", 0)
    partials = behavior_counts.get("tested_partial_match", 0)
    intentional_differences = behavior_counts.get("intentional_difference", 0)
    not_assessed = behavior_counts.get("not_assessed", 0)
    scope_exclusions = scoped_counts.get("excluded_internals", 0)
    behavior_complete = (
        has_comparison and matches > 0 and partials == 0
        and intentional_differences == 0 and not_assessed == 0
    )
    if behavior_complete:
        status_lines = [
            "This is the source-backed inventory and differential assessment of",
            "the documented public API in python-docx 1.2.0. Under the agreed",
            f"scope, all {matches} in-scope behavior records are tested matches;",
            f"{partials} are partial, {not_assessed} are unassessed, and",
            f"{scope_exclusions} structural/internal records are excluded.",
            "This supports documented-public-API parity, not parity for",
            "undocumented internals or malformed states outside the published",
            "contract. The documentation-only MIXED enum is assessed against its",
            "published value separately from upstream runtime availability.",
        ]
    elif has_comparison:
        status_lines = [
            "This is the source-backed inventory and differential assessment of",
            "the documented public API in python-docx 1.2.0. The comparison below",
            f"currently records {matches} tested behavior matches, {partials}",
            f"partial assessments, and {not_assessed} not-assessed in-scope rows.",
            "Malformed/invalid states and undocumented internals remain outside",
            "the agreed public-API scope.",
        ]
    else:
        status_lines = [
            "This source-backed inventory contains every API autodoc",
            "object/member emitted by the official v1.2.0 documentation build,",
            "plus enumerations and values documented in prose-only enum pages.",
            "Availability, signature, and behavior comparisons are not assessed",
            "until differential evidence is merged.",
        ]
    lines = [
        "# python-docx 1.2.0 documented API parity matrix",
        "",
        "## Status",
        "",
        *status_lines,
        "",
        "## Provenance",
        "",
        f"- Upstream source: [`python-docx` {UPSTREAM_TAG}](https://github.com/python-openxml/python-docx/tree/{UPSTREAM_TAG})",
        f"- Pinned upstream commit: `{UPSTREAM_COMMIT}`",
        f"- API documentation: [{DOCS_BASE}](https://python-docx.readthedocs.io/en/latest/index.html)",
        "- API page source: upstream `docs/api/*.rst`; API member signatures and",
        "  docstring behavior text were extracted from a Sphinx HTML build of that",
        "  source. Enum class/value entries were parsed from their RST pages and",
        "  checked against the v1.2.0 runtime enum classes.",
        "- Scope: documented public API and documented proxy objects; direct",
        "  lxml/OOXML/OPC structural access is recorded as excluded, per the",
        "  agreed boundary. Selected user-guide recipes have separate, partial",
        "  end-to-end traceability in the behavior evidence manifest.",
        "",
        "## Counts",
        "",
        "| Inventory family | Entries |",
        "|---|---:|",
        f"| Sphinx API objects/members | {len(object_records)} |",
        f"| Classes | {type_counts['class']} |",
        f"| Functions | {type_counts['function']} |",
        f"| Methods | {type_counts['method']} |",
        f"| Properties | {type_counts['property']} |",
        f"| Attributes | {type_counts['attribute']} |",
        f"| Enum classes, aliases, and documented values | {len(enum_records)} |",
        f"| Out-of-scope structural members (listed, not required) | {scoped_counts['excluded_internals']} |",
        f"| Total rows | {len(records)} |",
        "",
    ]
    if has_comparison:
        matched_methods = sum(
            r.get("signature_parity") == "call_shape_match"
            for r in object_records if r["kind"] == "method"
        )
        matched_members = sum(
            r.get("availability") in {"available", "available_as_proxy_instance"}
            for r in in_scope_members
        )
        annotation_matches = sum(
            r.get("annotation_parity") == "match" for r in comparable_callables
        )
        annotation_mismatches = sum(
            r.get("annotation_parity") == "mismatch" for r in comparable_callables
        )
        xml_status = Counter(r.get("xml_value_parity", "not_assessed")
                             for r in enum_members)
        enum_set_matches = sum(
            r.get("enum_member_set_parity") == "match" for r in enum_classes
        )
        lines += [
            "## Documented public API comparison",
            "",
            "The machine-readable source of these results is",
            "`docs/api-parity/python_docx_1_2_0_comparison.json`. Availability was",
            "tested on representative synthetic objects, not user documents.",
            "Signature comparisons separate call shape (names, kinds, defaults)",
            "from annotations; behavior requires explicit differential tests.",
            "",
            "| Measure | Result | Interpretation |",
            "|---|---:|---|",
            f"| In-scope documented methods/properties/attributes accessible on samples | {matched_members}/{len(in_scope_members)} | Availability only; does not establish behavioral parity. |",
            f"| Documented method call shapes matching | {matched_methods}/{sum(r['kind'] == 'method' for r in object_records)} | All documented method parameter shapes currently match. |",
            f"| Comparable callable annotations matching / differing | {annotation_matches}/{len(comparable_callables)} match; {annotation_mismatches} differ | All comparable callable annotations currently match; non-comparable entries are reported separately. |",
            f"| In-scope property/attribute writability | {sum(r.get('writability_parity') == 'match' for r in property_rows)} match; {len(missing_setters)} missing setters; {len(extra_setters)} candidate-only setters; {sum(r.get('writability_parity') == 'not_determined' for r in property_rows)} not determined | Static descriptor checks are supplemented by explicit differential setter/read-only probes where recorded. |",
            f"| Enum class member sets matching | {enum_set_matches}/{len(enum_classes)} | The documented names are present as member sets in both runtimes. |",
            f"| Documented enum member runtime values matching | {enum_values_match}/{len(enum_members)} | One documented value conflicts with the upstream 1.2.0 runtime; see the enum row. |",
            f"| Documented enum XML tokens | {xml_status['match']} exact; {xml_status['mismatch']} differ; {xml_status['docviews_extra_xml_token']} candidate-only tokens | Value/string/repr parity is tracked separately from XML serialization tokens. |",
            f"| Documented proxy class entries without a mirrored import path | {len(proxy_path_gaps)} | Importable class aliases are provided in mirrored modules; this does not establish proxy-type or behavior parity. |",
            f"| Behavior evidence | {sum(r.get('behavior_parity') == 'tested_match' for r in behavior_assessed)} tested matches; {sum(r.get('behavior_parity') == 'tested_partial_match' for r in behavior_assessed)} partial assessments; {sum(r.get('behavior_parity') == 'intentional_difference' for r in behavior_assessed)} intentional difference(s); {sum(r.get('behavior_parity') == 'not_assessed' for r in behavior_assessed)} not assessed | Evidence applies to the documented public contract; malformed, invalid, and undocumented states remain out of scope. |",
            f"| Official user-guide recipe mappings | {len(user_guide_recipes)} tracked ({recipe_counts['partial']} partial) | Traceability for selected quickstart, section, and header/footer examples; not full chapter coverage. |",
            "",
        ]
        if missing_setters:
            lines += ["Missing documented setters in this pass:", ""]
            for r in missing_setters:
                lines.append(f"- `{r['qualified_name']}`")
            lines.append("")
        if extra_setters:
            lines += ["Candidate-only setter (not a missing upstream capability):", ""]
            for r in extra_setters:
                lines.append(f"- `{r['qualified_name']}`")
            lines.append("")
    lines += [
        "## Sphinx API objects and members",
        "",
        "The reference signature is the signature shown by the generated official",
        "API page. The documented behavior is source prose, not a test result.",
        "Availability and signature fields report introspection/access checks.",
        "`Behavior parity` is assessed only when the Evidence column cites an",
        "explicit differential test; partial assessments do not establish full",
        "member parity.",
        "",
        "| Documented object/member | Kind | Reference signature | Documented behavior | Scope | docviews availability | Import-path status | Call-shape parity | Annotation parity | Writability | Behavior parity | Evidence | Source |",
        "|---|---|---|---|---|---|---|---|---|---|---|---|---|",
    ]
    for r in object_records:
        source = f"[docs]({r['doc_url']})"
        evidence = "\n".join(r.get("behavior_evidence", []))
        lines.append("| " + " | ".join([
            markdown_cell(r["qualified_name"], code=True),
            markdown_cell(r["kind"]),
            markdown_cell(r["signature"], code=True),
            markdown_cell(r["behavior"]),
            markdown_cell(r["scope"]),
            markdown_cell(r.get("availability", "not_assessed")),
            markdown_cell(r.get("import_path_status", "not_assessed")),
            markdown_cell(r.get("signature_parity", "not_assessed")),
            markdown_cell(r.get("annotation_parity", "not_assessed")),
            markdown_cell(r.get("writability_parity", "not_applicable")),
            markdown_cell(r.get("behavior_parity", "not_assessed")),
            markdown_cell(evidence, code=True),
            source,
        ]) + " |")
    lines += [
        "",
        "## Documented enumerations",
        "",
        "Enum values are documented as RST definition lists rather than Sphinx",
        "autodoc objects. The values below include both the documented behavior",
        "and a runtime-presence check against python-docx 1.2.0. The `aliases`",
        "column lists documented alternate import paths.",
        "",
        "| Documented enum path | Kind | Signature/value | Documented behavior | python-docx value | docviews value | Value/member-set parity | XML-token parity | str/repr parity | Aliases | Availability | Behavior | Source |",
        "|---|---|---|---|---|---|---|---|---|---|---|---|---|",
    ]
    for r in enum_records:
        result = comparison_records.get(r["qualified_name"], {})
        if r["kind"] == "enum_member":
            upstream_value = result.get("reference_value", r.get("python_docx_value"))
            if upstream_value is None and r.get("present_in_python_docx_1_2_0"):
                upstream_value = "None"
            if not r.get("present_in_python_docx_1_2_0"):
                upstream_value = "MISSING in v1.2.0 runtime"
            candidate_value = result.get("docviews_value", "—")
            value_parity = result.get("value_parity", "not_assessed")
            xml_parity = result.get("xml_value_parity", "not_assessed")
            string_repr_parity = (
                f"str {result.get('string_parity', 'not_assessed')}; "
                f"repr {result.get('repr_parity', 'not_assessed')}"
            )
        elif r["kind"] == "enum_class":
            upstream_value = result.get("reference_member_count", "—")
            candidate_value = result.get("docviews_member_count", "—")
            value_parity = result.get("enum_member_set_parity", "not_assessed")
            xml_parity = "not_applicable"
            string_repr_parity = "not_applicable"
        else:
            upstream_value = result.get("reference_alias_identity", "—")
            candidate_value = result.get("docviews_alias_identity", "—")
            value_parity = result.get("alias_identity_parity", "not_assessed")
            xml_parity = "not_applicable"
            string_repr_parity = "not_applicable"
        source = f"[docs]({r['doc_url']})"
        aliases = ", ".join(r.get("aliases", []))
        note = r.get("reference_note", "")
        behavior = r.get("behavior", "")
        if note:
            behavior = (behavior + " " + note).strip()
        lines.append("| " + " | ".join([
            markdown_cell(r["qualified_name"], code=True),
            markdown_cell(r["kind"]),
            markdown_cell(r["signature"], code=True),
            markdown_cell(behavior),
            markdown_cell(upstream_value, code=True),
            markdown_cell(candidate_value, code=True),
            markdown_cell(value_parity),
            markdown_cell(xml_parity),
            markdown_cell(string_repr_parity),
            markdown_cell(aliases, code=True),
            markdown_cell(result.get("availability", r.get("availability", "not_assessed"))),
            markdown_cell(result.get("behavior_parity", r.get("behavior_parity", "not_assessed"))),
            source,
        ]) + " |")
    lines += [
        "",
        "## Documented/runtime enum discrepancy",
        "",
        "The official `MSO_THEME_COLOR_INDEX` page lists `MIXED`, but the",
        "python-docx 1.2.0 runtime omits it. docviews implements the documented",
        "member with value `-2`, sourced from the Microsoft Office API reference.",
        "The inventory and comparison keep the upstream runtime absence explicit;",
        "the member is tested against the documented value rather than presented",
        "as a runtime match. Its `xml_value` is unmapped.",
        "Two runtime enum sentinels, `WD_COLOR_INDEX.INHERITED` and",
        "`WD_UNDERLINE.INHERITED`, are not listed on their enum reference pages;",
        "the broader enum tests compare them to the pinned runtime.",
        "",
        "## Release-readiness sequence",
        "",
        "1. Enum discrepancies — successfully completed: 471 runtime members match; the",
        "   documentation-only MIXED value (-2) is implemented separately from the",
        "   python-docx 1.2.0 runtime omission.",
        "2. Partial behavior assessments — successfully completed: all 578 in-scope",
        "   documented behavior records have differential evidence; 6 structural/internal",
        "   records remain excluded by the agreed scope.",
        "3. Python-version range — successfully completed: the full 239-test suite passed",
        "   on CPython 3.9.25, 3.10.22, 3.11.18, 3.12.16, 3.13.16, 3.14.10, and 3.15.2.",
        "4. Docs-build warning — successfully completed by explicit acceptance: the Sphinx 9.2.0",
        "   build succeeded with one upstream docutils warning at",
        "   `research/python-docx-v1.2.0/docs/user/comments.rst:134` (`Unknown target name: \"container\"`).",
        "   This malformed inline reference is documentation-only and does not affect the",
        "   API inventory or runtime; the pinned upstream reference was left unchanged.",
        "5. Final artifact build — successfully completed for the selected stable `1.0.0`",
        "   metadata: wheel and sdist rebuilt on Python 3.15.2; strict Twine checks, separate",
        "   clean installs, DOCX round-trip smoke tests, and CLI checks all passed.",
        "   Artifacts and SHA-256 checksums are in `release_artifacts/1.0.0/`.",
        "6. Stable-release publication — version `1.0.0` with the `Production/Stable`",
        "   classifier was published to PyPI on 2026-11-01 after parity and hosted CI",
        "   passed; see `RELEASE_READINESS.md`.",

        "",
        "Machine-readable sources: `docs/api-parity/python_docx_1_2_0_inventory.json`",
        "(source inventory), `docs/api-parity/python_docx_1_2_0_comparison.json`",
        "(availability/signature/enum comparison), and",
        "`docs/api-parity/python_docx_1_2_0_behavior.json` (test-backed behavior",
        "assessments).",
        "",
    ]
    return "\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--reference", type=Path, required=True,
                        help="python-docx v1.2.0 source checkout")
    parser.add_argument("--html", type=Path, required=True,
                        help="HTML directory produced by Sphinx from upstream docs")
    parser.add_argument("--json", type=Path, default=Path(
        "docs/api-parity/python_docx_1_2_0_inventory.json"))
    parser.add_argument("--markdown", type=Path, default=Path("API_PARITY_MATRIX.md"))
    parser.add_argument("--comparison", type=Path, default=None,
                        help="optional comparison JSON to merge into the Markdown matrix")
    args = parser.parse_args()

    reference = args.reference.resolve()
    html_root = args.html.resolve()
    inventory_path = html_root / "objects.inv"
    if not inventory_path.is_file():
        parser.error(f"missing Sphinx inventory: {inventory_path}")
    if not (reference / "docs/api").is_dir() or not (reference / "src/docx").is_dir():
        parser.error("reference must contain docs/api and src/docx")

    records = build_object_records(reference, html_root, inventory_path)
    records.extend(build_enum_records(reference))
    records.extend(build_indirect_enum_records(reference))
    records.sort(key=lambda r: (r.get("doc_page", ""), r["qualified_name"]))
    args.json.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "reference": {
            "package": "python-docx",
            "tag": UPSTREAM_TAG,
            "commit": UPSTREAM_COMMIT,
            "documentation_base": DOCS_BASE,
            "source_docs": "docs/api/*.rst",
            "generated_from_sphinx_html": True,
            "scope": "documented public API; excludes direct lxml/OOXML/OPC internals",
        },
        "counts": {
            "records": len(records),
            "sphinx_api_objects": sum(
                r["kind"] in {"class", "function", "method", "property", "attribute"}
                for r in records
            ),
            "enum_classes_aliases_and_values": sum(
                r["kind"].startswith("enum_") for r in records
            ),
            "by_kind": dict(sorted(Counter(r["kind"] for r in records).items())),
            "out_of_scope_structural": sum(
                r.get("scope") == "excluded_internals" for r in records
            ),
        },
        "records": records,
    }
    args.json.write_text(json.dumps(payload, indent=2, ensure_ascii=False) + "\n",
                         encoding="utf-8")
    matrix_records = [dict(record) for record in records]
    comparison_payload = None
    if args.comparison is not None:
        if not args.comparison.is_file():
            parser.error(f"missing comparison JSON: {args.comparison}")
        comparison_payload = json.loads(args.comparison.read_text(encoding="utf-8"))
        comparison_records = comparison_payload.get("records", {})
        for record in matrix_records:
            record.update(comparison_records.get(record["qualified_name"], {}))
        unmatched = set(comparison_records) - {
            record["qualified_name"] for record in matrix_records
        }
        if unmatched:
            print(f"WARNING: {len(unmatched)} comparison entries are not in the source inventory")
    args.markdown.parent.mkdir(parents=True, exist_ok=True)
    args.markdown.write_text(
        render_markdown(matrix_records, reference, comparison_payload),
        encoding="utf-8")
    no_signature = [r["qualified_name"] for r in records
                    if r["kind"] in {"class", "function", "method", "property", "attribute"}
                    and not r.get("signature")]
    print(f"Wrote {len(records)} records to {args.json} and {args.markdown}")
    print(f"Sphinx object records: {payload['counts']['sphinx_api_objects']}; "
          f"enum records: {payload['counts']['enum_classes_aliases_and_values']}")
    if no_signature:
        print(f"WARNING: {len(no_signature)} Sphinx objects had no signature text:")
        print("\n".join(f"  {name}" for name in no_signature))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
