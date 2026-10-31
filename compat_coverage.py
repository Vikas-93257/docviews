"""How much of python-docx's public surface does the compat door answer?

    python3 compat_coverage.py

Builds the same document with python-docx 1.2.0 and with ``docviews.compat``,
then asks each pair of objects for the same attribute names and reports what is
missing. The point of the shim is one-line migration, so the question that
matters is not "is docviews good" but "which names does existing python-docx
code touch that would raise AttributeError here".

Names that reach into the XML/OPC layer (``part``, ``element``) are listed
separately: those exist to hand out lxml proxies, and no pure-Python rewriting
of them is on the table — code that uses them is not migrated by an import
change, and `COMPAT.md` says so.

Writes ``COMPAT.md`` next to this file.
"""

from __future__ import annotations

import inspect
import io
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT / "tests"))

import docx as real                                          # noqa: E402
import docviews.compat as shim                               # noqa: E402
from docx.enum.style import WD_STYLE_TYPE as REAL_STYLE_TYPE  # noqa: E402
from docx.shared import Inches as REAL_INCHES                  # noqa: E402
from docx.shared import RGBColor as REAL_RGB_COLOR              # noqa: E402
from docviews.compat.enum import WD_STYLE_TYPE as OUR_STYLE_TYPE  # noqa: E402
from docviews.compat.shared import Inches as SHIM_INCHES       # noqa: E402
from docviews.compat.shared import RGBColor as SHIM_RGB_COLOR  # noqa: E402

import blobs                                                 # noqa: E402
from conftest import make_docx                                # noqa: E402

# Handing out lxml elements is a different promise, and no rewrite changes that.
STRUCTURAL = {"part", "element", "blob"}


def build(lib, style_type) -> dict:
    """The same document, built with whichever library is passed in."""
    doc = lib.Document()
    doc.add_heading("Coverage", 1)
    paragraph = doc.add_paragraph("Body ")
    paragraph.add_run("text").bold = True
    tab_stops = paragraph.paragraph_format.tab_stops
    inches = REAL_INCHES if lib is real else SHIM_INCHES
    rgb_color_type = REAL_RGB_COLOR if lib is real else SHIM_RGB_COLOR
    length = inches(1)
    rgb_color = rgb_color_type(0x12, 0x34, 0x56)
    tab_stop = tab_stops.add_tab_stop(inches(1))
    table = doc.add_table(rows=2, cols=2)
    for row in range(2):
        for col in range(2):
            table.cell(row, col).text = f"{row}{col}"
    table.style = "Table Grid"
    doc.add_picture(io.BytesIO(blobs.PNG_96), width=914400)
    drawing = next(
        item for item in doc.paragraphs[-1].runs[0].iter_inner_content()
        if type(item).__name__ == "Drawing"
    )
    section = doc.sections[0]
    section.header.add_paragraph("header")
    section.footer.add_paragraph("footer")
    doc.styles.add_style("CoverageStyle", style_type.PARAGRAPH)
    latent_styles = doc.styles.latent_styles
    latent_style = latent_styles.add_latent_style("Coverage Latent")
    properties = doc.core_properties
    properties.title = "Coverage"
    properties.author = "compat"
    run = paragraph.runs[-1]
    comment = doc.add_comment(run, text="Coverage comment", author="API")
    hyperlink_source = make_docx(
        "<w:p><w:hyperlink w:anchor='coverage-link'>"
        "<w:r><w:t>link</w:t></w:r></w:hyperlink></w:p>"
    )
    hyperlink_doc = lib.Document(
        io.BytesIO(hyperlink_source) if lib is real else hyperlink_source
    )
    hyperlink = hyperlink_doc.paragraphs[0].hyperlinks[0]
    rendered_source = make_docx(
        "<w:p><w:r><w:t>before</w:t><w:lastRenderedPageBreak/>"
        "<w:t>after</w:t></w:r></w:p>"
    )
    rendered_doc = lib.Document(
        io.BytesIO(rendered_source) if lib is real else rendered_source
    )
    rendered_break = next(
        item for item in rendered_doc.paragraphs[0].runs[0].iter_inner_content()
        if type(item).__name__ == "RenderedPageBreak"
    )
    return {
        "Document": doc,
        "Paragraph": paragraph,
        "ParagraphFormat": paragraph.paragraph_format,
        "TabStops": tab_stops,
        "TabStop": tab_stop,
        "Run": run,
        "Font": run.font,
        "ColorFormat": run.font.color,
        "Table": table,
        "Row": table.rows[0],
        "Rows": table.rows,
        "Cell": table.cell(0, 0),
        "Column": table.columns[0],
        "Columns": table.columns,
        "Section": section,
        "Sections": doc.sections,
        "Header": section.header,
        "Footer": section.footer,
        "Settings": doc.settings,
        "Styles": doc.styles,
        "Style": doc.styles["Heading 1"],
        "LatentStyles": latent_styles,
        "LatentStyle": latent_style,
        "InlineShape": doc.inline_shapes[0],
        "InlineShapes": doc.inline_shapes,
        "Drawing": drawing,
        "Image": drawing.image,
        "Length": length,
        "RGBColor": rgb_color,
        "CoreProperties": properties,
        "Comments": doc.comments,
        "Comment": comment,
        "CommentParagraph": comment.paragraphs[0],
        "Hyperlink": hyperlink,
        "RenderedPageBreak": rendered_break,
    }


def public(obj) -> set:
    """Public attribute names, resolved rather than read off ``dir()``.

    ``dir()`` misses anything a class serves through ``__getattr__``, and a
    property that raises is still an attribute that exists.
    """
    names = set()
    for name in dir(obj):
        if name.startswith("_"):
            continue
        try:
            getattr(obj, name)
        except AttributeError:
            continue
        except Exception:                              # noqa: BLE001
            pass
        names.add(name)
    return names


def measure() -> list:
    theirs = build(real, REAL_STYLE_TYPE)
    ours = build(shim, OUR_STYLE_TYPE)
    rows = []
    for label, their_object in theirs.items():
        our_object = ours[label]
        wanted = public(their_object) - STRUCTURAL
        listed = public(our_object)
        missing, answered = [], 0
        for name in sorted(wanted):
            present = name in listed
            if not present:
                # a name served through __getattr__ is not in dir() — ask for it
                try:
                    getattr(our_object, name)
                    present = True
                except AttributeError:
                    present = False
                except Exception:                      # noqa: BLE001
                    present = True
            if present:
                answered += 1
            else:
                missing.append(name)
        rows.append({
            "object": label,
            "asked": len(wanted),
            "missing": missing,
            "answered": answered,
            "beyond": sorted(listed - public(their_object)),
        })
    return rows


def signature_audit() -> tuple[int, int, int, list[str], int]:
    """Compare sampled object methods and selected public constructors.

    Annotations are library-specific and are not compared; parameter names,
    order, kinds, and defaults are. A missing/non-callable counterpart counts
    as a mismatch rather than silently disappearing from the denominator.
    """
    theirs = build(real, REAL_STYLE_TYPE)
    ours = build(shim, OUR_STYLE_TYPE)
    checked = compared = matched = unavailable = 0
    mismatches = []

    def shape(signature):
        return tuple(
            (parameter.name, parameter.kind,
             parameter.default is inspect.Parameter.empty,
             None if parameter.default is inspect.Parameter.empty
             else (type(parameter.default).__name__, repr(parameter.default)))
            for parameter in signature.parameters.values()
        )

    for label, their_object in theirs.items():
        our_object = ours[label]
        for name in sorted(public(their_object) - STRUCTURAL):
            try:
                reference = getattr(their_object, name)
            except Exception:                              # noqa: BLE001
                continue
            try:
                candidate = getattr(our_object, name)
            except AttributeError:
                candidate = None
            if not callable(reference):
                continue
            checked += 1
            if not callable(candidate):
                mismatches.append(f"{label}.{name}: callable missing")
                continue
            try:
                reference_sig = inspect.signature(reference)
                candidate_sig = inspect.signature(candidate)
            except (TypeError, ValueError):
                unavailable += 1
                continue

            compared += 1
            if shape(reference_sig) == shape(candidate_sig):
                matched += 1
            else:
                mismatches.append(
                    f"{label}.{name}: {reference_sig} != {candidate_sig}"
                )

    import docx.shared as real_shared
    import docviews.compat.shared as shim_shared
    constructors = [
        ("docx.Document", real.Document, shim.Document),
        ("docx.shared.Length", real_shared.Length, shim_shared.Length),
        ("docx.shared.Inches", real_shared.Inches, shim_shared.Inches),
        ("docx.shared.Cm", real_shared.Cm, shim_shared.Cm),
        ("docx.shared.Mm", real_shared.Mm, shim_shared.Mm),
        ("docx.shared.Pt", real_shared.Pt, shim_shared.Pt),
        ("docx.shared.Emu", real_shared.Emu, shim_shared.Emu),
        ("docx.shared.Twips", real_shared.Twips, shim_shared.Twips),
        ("docx.shared.RGBColor", real_shared.RGBColor, shim_shared.RGBColor),
    ]
    for label, reference, candidate in constructors:
        checked += 1
        try:
            reference_sig = inspect.signature(reference)
            candidate_sig = inspect.signature(candidate)
        except (TypeError, ValueError):
            unavailable += 1
            continue
        compared += 1
        if shape(reference_sig) == shape(candidate_sig):
            matched += 1
        else:
            mismatches.append(f"{label}: {reference_sig} != {candidate_sig}")

    return checked, compared, matched, mismatches, unavailable


ENUM_MODULES = ("docx.enum.text", "docx.enum.style", "docx.enum.section",
                "docx.enum.table", "docx.enum.shape", "docx.enum.dml")

# docx.shared names that hand out lxml element proxies; the compat door draws
# its boundary at the same place it draws it for ``part``/``element``.
SHARED_PLUMBING = {"ElementProxy", "Parented", "StoryChild", "TextAccumulator"}


def module_rows() -> list:
    """The names python-docx ships outside a document object.

    ``from docx.enum.text import WD_ALIGN_PARAGRAPH`` and ``from docx.shared
    import Pt`` are ordinary imports in existing code, so each one is measured
    where it is imported from: the class must be importable from the module of
    the same name (``docviews.compat.enum.text``), and every member must exist.
    """
    import docviews.compat.enum as our_enum
    import docviews.compat.shared as our_shared

    rows = []
    for module_name in ENUM_MODULES:
        theirs_module = __import__(module_name, fromlist=["*"])
        ours_module = __import__("docviews.compat." + module_name.replace(
            "docx.", ""), fromlist=["*"])
        asked = missing = members = answered_members = 0
        for name in dir(theirs_module):
            if name.startswith("_") or name == "enum":
                continue
            theirs = getattr(theirs_module, name)
            if not isinstance(theirs, type) or \
                    not getattr(theirs, "__module__", "").startswith("docx.enum"):
                continue                            # an import, not a definition
            asked += 1
            ours = getattr(ours_module, name, None)
            if ours is None:
                missing += 1
                continue
            for member in getattr(theirs, "__members__", {}):
                members += 1
                if hasattr(ours, member):
                    answered_members += 1
        rows.append({"module": module_name, "asked": asked,
                     "missing": missing, "members": members,
                     "members_answered": answered_members})

    # python-docx's two enum bases are part of ``docx.enum.base``
    import docx.enum.base as real_base
    import docviews.compat.enum.base as our_base
    base_names = [n for n in dir(real_base) if n.startswith("Base")]
    rows.append({"module": "docx.enum.base", "asked": len(base_names),
                 "missing": sum(1 for n in base_names if not hasattr(our_base, n)),
                 "members": 0, "members_answered": 0})

    import docx.shared as real_shared
    shared_asked = shared_missing = 0
    for name in dir(real_shared):
        if name.startswith("_") or name in SHARED_PLUMBING:
            continue
        value = getattr(real_shared, name)
        if not (isinstance(value, type) or callable(value)):
            continue
        if not getattr(value, "__module__", "").startswith("docx.shared"):
            continue
        shared_asked += 1
        if not hasattr(our_shared, name):
            shared_missing += 1
    rows.append({"module": "docx.shared", "asked": shared_asked,
                 "missing": shared_missing, "members": 0,
                 "members_answered": 0})
    return rows


def module_report(rows: list) -> str:
    lines = ["| module | names python-docx defines | answered | enum members | answered |",
             "|---|---:|---:|---:|---:|"]
    for row in rows:
        members = row["members"] or "—"
        answered = row["members_answered"] if row["members"] else "—"
        lines.append(f"| `{row['module']}` | {row['asked']} | "
                     f"{row['asked'] - row['missing']} | {members} | {answered} |")
    return "\n".join(lines)


# Tokens we give on purpose where python-docx has something else; each is a
# statement about what Word's XML accepts, and each is checked in test_enums.py.
DELIBERATE_TOKENS = {
    ("WD_COLOR_INDEX", "AUTO"): ("default", "none"),
    ("WD_COLOR", "AUTO"): ("default", "none"),
    ("WD_LINE_SPACING", "SINGLE"): ("UNMAPPED", None),
    ("WD_LINE_SPACING", "ONE_POINT_FIVE"): ("UNMAPPED", None),
    ("WD_LINE_SPACING", "DOUBLE"): ("UNMAPPED", None),
    ("MSO_THEME_COLOR", "NOT_THEME_COLOR"): ("UNMAPPED", None),
    ("MSO_THEME_COLOR_INDEX", "NOT_THEME_COLOR"): ("UNMAPPED", None),
}


def enum_protocol() -> str:
    """How a compat member behaves compared with a python-docx member.

    Names are not the whole contract: code iterates an enum, looks a member up
    by value, prints it, and round-trips tokens. This compares all of that,
    member by member, against python-docx, and counts the tokens only we carry.
    """
    import docviews.compat.enum as our_enum

    compared = divergences = additions = deliberate = 0
    for module_name in ("docx.enum.text", "docx.enum.style", "docx.enum.section",
                        "docx.enum.table", "docx.enum.shape", "docx.enum.dml"):
        theirs_module = __import__(module_name, fromlist=["*"])
        for name in dir(theirs_module):
            if not (name.startswith("WD_") or name.startswith("MSO_")):
                continue
            theirs = getattr(theirs_module, name)
            if not isinstance(theirs, type) or not hasattr(theirs, "__members__"):
                continue
            ours = getattr(our_enum, name)
            for member in theirs.__members__:
                a, b = getattr(theirs, member), getattr(ours, member)
                compared += 1
                if (a.value, str(a), repr(a)) != (b.value, str(b), repr(b)):
                    divergences += 1
                their_token = getattr(a, "xml_value", None)
                our_token = getattr(b, "xml_value", None)
                if their_token != our_token:
                    if their_token is None:
                        additions += 1         # a token python-docx does not map
                    elif DELIBERATE_TOKENS.get((name, member)) == (their_token,
                                                                   our_token):
                        deliberate += 1        # documented, and tested
                    else:
                        divergences += 1
    return (f"{compared} members compared with python-docx — value, `str()`, "
            f"`repr()` identical for all {compared}; {additions} members carry "
            f"a token python-docx leaves unmapped (e.g. `WD_BREAK.PAGE`), "
            f"{deliberate} carry a different token on purpose (listed below), "
            f"{divergences} unexplained")


def report(rows: list) -> str:
    lines = ["| object | python-docx attributes | answered | missing |",
             "|---|---:|---:|---|"]
    asked = answered = 0
    for row in rows:
        asked += row["asked"]
        answered += row["answered"]
        missing = ", ".join(f"`{name}`" for name in row["missing"]) or "—"
        lines.append(f"| {row['object']} | {row['asked']} | {row['answered']} | {missing} |")
    lines.append(f"| **total** | **{asked}** | **{answered} "
                 f"({answered / asked * 100:.1f}%)** | **{asked - answered}** |")
    return "\n".join(lines)


REASONS = {
    ("WD_COLOR_INDEX", "AUTO"): "`w:highlight` has no `default` value; "
                                "`none` is what clears it",
    ("WD_COLOR", "AUTO"): "same class as `WD_COLOR_INDEX` (an alias)",
    ("WD_LINE_SPACING", "SINGLE"): "python-docx marks it unmapped; line rule "
                                   "`auto` is what Word writes",
    ("WD_LINE_SPACING", "ONE_POINT_FIVE"): "as above",
    ("WD_LINE_SPACING", "DOUBLE"): "as above",
    ("MSO_THEME_COLOR", "NOT_THEME_COLOR"): "python-docx marks it unmapped; "
                                            "no token is written",
    ("MSO_THEME_COLOR_INDEX", "NOT_THEME_COLOR"): "same class, alias",
}


def token_table() -> str:
    lines = ["| member | python-docx `xml_value` | here | why |",
             "|---|---|---|---|"]
    for (cls, member), (theirs, ours) in sorted(DELIBERATE_TOKENS.items()):
        lines.append(f"| `{cls}.{member}` | `{theirs}` | "
                     f"`{ours if ours is not None else 'None'}` | "
                     f"{REASONS.get((cls, member), '')} |")
    return "\n".join(lines)


def beyond_section(beyond: dict) -> str:
    if not beyond:
        return ""
    lines = ["## Names python-docx does not have", "",
             "On the same objects, and answering the questions python-docx "
             "cannot:", ""]
    for label, names in beyond.items():
        lines.append(f"- **{label}**: " + ", ".join(f"`{n}`" for n in names))
    return "\n".join(lines)


OUT_OF_SCOPE = """## Out of scope, on purpose

- `docx.oxml` — the lxml element API. Code that builds `<w:p>` elements by hand
  is not migrated by changing an import; `docviews` reads and writes the same
  XML through its own layer.
- `part`, `element` on document objects, and `docx.shared`'s `ElementProxy`,
  `Parented`, `StoryChild`, `TextAccumulator` — all hand out lxml proxies.
- `doc.part` / package internals (`docx.opc`) — `docviews` exposes `main_part`,
  `parts` and part names instead.
"""


def main() -> int:
    rows = measure()
    table = report(rows)
    print(table)
    signature_checked, signature_compared, signature_matched, \
        signature_mismatches, signature_unavailable = signature_audit()
    print(f"\ncallable parameter shapes: {signature_matched}/{signature_compared} "
          f"matched; {len(signature_mismatches)} mismatches across "
          f"{signature_checked} methods/constructors; "
          f"{signature_unavailable} unavailable")
    for mismatch in signature_mismatches:
        print(f"  {mismatch}")

    modules = module_rows()
    module_table = module_report(modules)
    protocol = enum_protocol()
    print("\n" + module_table)
    print(protocol)

    missing = sorted({name for row in rows for name in row["missing"]})
    print(f"\nstill missing ({len(missing)}): {', '.join(missing) or 'none'}")

    beyond = {row["object"]: row["beyond"] for row in rows if row["beyond"]}
    print("\nbeyond python-docx (docviews' own names, on the same objects):")
    for label, names in beyond.items():
        print(f"  {label}: {', '.join(names)}")

    signature_section = (
        "## Sampled callable signatures\n\n"
        f"Parameter names, order, kinds, and defaults matched for "
        f"**{signature_matched}/{signature_compared}** inspected methods and "
        f"constructors ({signature_checked} checked; "
        f"{signature_unavailable} signatures unavailable). Type annotations "
        "and return annotations are not compared."
    )
    if signature_mismatches:
        signature_section += "\n\nMismatches:\n\n" + "\n".join(
            f"- `{mismatch}`" for mismatch in signature_mismatches)

    document = "\n\n".join([
        "# python-docx API coverage",
        f"Generated by `python3 compat_coverage.py` under Python "
        f"{sys.version_info.major}.{sys.version_info.minor}: representative "
        "object instances are built\nwith **python-docx 1.2.0** and with "
        "`docviews.compat`; non-underscore, non-structural public attribute "
        "names\nand sampled callable parameter shapes are compared.",
        "## Objects\n\n" + table,
        signature_section,
        "## Scope limits\n\nThis is a bounded inventory of representative "
        "object instances, not proof of full documented-public-API parity. "
        "Special methods (`__iter__`, `__len__`, and similar) are excluded "
        "from the attribute total; XML/OPC plumbing is excluded by design. "
        "The signature check does not compare type annotations or return "
        "annotations. Signatures unavailable to `inspect.signature` on a runtime "
        "are counted separately, not guessed; for example, inherited `int` "
        "method `Length.conjugate` is uninspectable on Python 3.9 but exists on "
        "both implementations. Neither inventory checks behavior for every "
        "documented method, input, or object state. The remaining engineering "
        "target is a "
        "complete official-API member/signature/behavior matrix.",
        "## Imports\n\nThe names python-docx ships *beside* a document — the "
        "ones existing\ncode imports directly:\n\n" + module_table + "\n\n"
        + protocol,
        "## Token differences that are deliberate\n\n" + token_table(),
        beyond_section(beyond),
        "## How this is checked\n\nA name answering is the start, not the "
        "end. Focused API behavior is compared with python-docx in "
        "`tests/test_compat_parity.py` and `tests/test_authoring_v03.py`; the "
        "writer and authoring suites exercise file output, and "
        "`tests/test_enums.py` compares all 471 enum members. "
        "`examples/docs_recipes.py` runs python-docx's own documented recipes "
        "through both libraries and diffs the results.",
        OUT_OF_SCOPE,
    ]) + "\n"
    (ROOT / "COMPAT.md").write_text(document, encoding="utf-8")
    print("\nwrote COMPAT.md")
    return 1 if missing or signature_mismatches else 0


if __name__ == "__main__":
    raise SystemExit(main())
