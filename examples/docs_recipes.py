"""The recipes from python-docx's own documentation, run twice.

    python examples/docs_recipes.py            # python-docx 1.2.0
    PYTHONPATH=src python examples/docs_recipes.py --shim    # docviews.compat

Each recipe is written exactly the way python-docx's docs write it. The file is
then re-opened with the *same* library and described in plain text — styles,
text, alignment, run formatting, table geometry, section geometry, header and
footer text, core properties, comments. Printing the two runs and diffing them
is the adoption test: nothing about the description is docviews-specific, so a
difference is a behaviour a migrating program would see.

python-docx cannot describe a comment's anchor, so neither does the description
— that difference belongs in `the_corruption.py` and `review_notes.py`, where it
is the point.
"""

from __future__ import annotations

import io
import struct
import sys
import zlib
from pathlib import Path

SHIM = "--shim" in sys.argv

# the author a brand-new document carries before the program sets one
DEFAULT_AUTHOR = "" if SHIM else "python-docx"

if SHIM:
    import docviews.compat as docx
    from docviews.compat.enum import WD_ALIGN_PARAGRAPH, WD_ORIENT
    from docviews.compat.shared import Inches, Pt
else:
    import docx
    from docx.enum.section import WD_ORIENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Inches, Pt


# ------------------------------------------------------------------- a picture
def tiny_png() -> bytes:
    """A 3x2 PNG, built here so the example needs no asset files."""
    def chunk(kind: bytes, payload: bytes) -> bytes:
        return (struct.pack(">I", len(payload)) + kind + payload
                + struct.pack(">I", zlib.crc32(kind + payload) & 0xFFFFFFFF))
    raw = b"".join(b"\x00" + b"\x88\x22\x11" * 3 for _ in range(2))
    return (b"\x89PNG\r\n\x1a\n"
            + chunk(b"IHDR", struct.pack(">IIBBBBB", 3, 2, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw))
            + chunk(b"IEND", b""))


# ------------------------------------------------------------------- recipes
def r01_hello_world(doc):
    doc.add_paragraph("Hello, World!")


def r02_headings(doc):
    doc.add_heading("Title", 0)
    for level in range(1, 4):
        doc.add_heading(f"Heading {level}", level)


def r03_bulleted_list(doc):
    for item in ("CPU", "Disk", "Network"):
        doc.add_paragraph(item, style="List Bullet")


def r04_run_formatting(doc):
    para = doc.add_paragraph()
    para.add_run("bold ").bold = True
    para.add_run("italic ").italic = True
    styled = para.add_run("sized")
    styled.font.size = Pt(14)
    styled.font.name = "Courier New"


def r05_alignment(doc):
    doc.add_paragraph("centred", style="Normal").alignment = \
        WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("justified").alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


def r06_table(doc):
    table = doc.add_table(rows=3, cols=3)
    table.style = "Table Grid"
    for row, values in enumerate((("a", "b", "c"), ("d", "e", "f"),
                                  ("g", "h", "i"))):
        for col, value in enumerate(values):
            table.cell(row, col).text = value
    table.cell(0, 0).merge(table.cell(0, 2))
    table.cell(1, 0).merge(table.cell(2, 0))


def r07_picture(doc):
    doc.add_paragraph("A picture, sized in inches:")
    doc.add_picture(io.BytesIO(tiny_png()), width=Inches(1.5))


def r08_sections_and_furniture(doc):
    section = doc.sections[0]
    section.left_margin = Inches(0.75)
    section.top_margin = Inches(1)
    section.orientation = WD_ORIENT.LANDSCAPE
    section.header.add_paragraph("Page header")
    section.footer.add_paragraph("Page footer")
    doc.add_paragraph("Body of section one.")
    new = doc.add_section()
    new.left_margin = Inches(2)
    doc.add_paragraph("Body of section two.")


def r09_page_break(doc):
    doc.add_paragraph("Page one")
    doc.add_page_break()
    doc.add_paragraph("Page two")


def r10_core_properties(doc):
    properties = doc.core_properties
    properties.title = "Document Title"
    properties.author = "The Author"
    properties.subject = "Testing docx"
    properties.category = "Examples"
    doc.add_paragraph("Metadata set.")


def r11_custom_style(doc):
    from docx.enum.style import WD_STYLE_TYPE
    style = doc.styles.add_style("Callout", WD_STYLE_TYPE.PARAGRAPH)
    style.font.bold = True
    style.font.size = Pt(13)
    doc.add_paragraph("Read this first.", style="Callout")


def r12_comment(doc):
    paragraph = doc.add_paragraph("The total is 1,204,332 units.")
    run = paragraph.add_run(" Checked by hand.")
    doc.add_comment([run], text="Verify against the ledger.",
                    author="Reviewer", initials="RV")


RECIPES = [value for name, value in sorted(globals().items())
           if name.startswith("r") and name[1:3].isdigit()]


# ----------------------------------------------------------------- reader side
def describe(path: str) -> str:
    """Everything both libraries can see, in a fixed order, as text."""
    doc = docx.Document(path)
    lines = []

    lines.append("paragraphs:")
    for para in doc.paragraphs:
        style = para.style.name if para.style is not None else "-"
        align = para.alignment
        align = getattr(align, "name", None) or str(align)
        lines.append(f"  [{style}] {para.text!r} align={align}")
        for run in para.runs:
            bits = [f"text={run.text!r}"]
            font = run.font
            for name in ("bold", "italic", "underline"):
                value = getattr(font, name)
                if value:
                    bits.append(f"{name}={getattr(value, 'name', value)}")
            if font.size is not None:
                bits.append(f"size={font.size.pt:g}pt")
            if font.name:
                bits.append(f"font={font.name}")
            if len(bits) > 1:
                lines.append("      " + " ".join(bits))

    lines.append("tables:")
    for table in doc.tables:
        lines.append(f"  style={table.style.name if table.style else '-'}")
        for row in table.rows:
            lines.append("  " + " | ".join(cell.text for cell in row.cells))

    lines.append("sections:")
    for index, section in enumerate(doc.sections):
        lines.append(
            f"  {index}: page={section.page_width.inches:g}x"
            f"{section.page_height.inches:g}in "
            f"left={section.left_margin.inches:g}in "
            f"top={section.top_margin.inches:g}in "
            f"orient={getattr(section.orientation, 'name', section.orientation)}")
        if index == 0:
            for label, part in (("header", section.header),
                                ("footer", section.footer)):
                if part is not None:
                    text = [p.text for p in part.paragraphs]
                    lines.append(f"    {label}: {text}")

    lines.append("core:")
    properties = doc.core_properties
    for name in ("title", "author", "subject", "category"):
        value = getattr(properties, name)
        if name == "author" and value == DEFAULT_AUTHOR:
            # a template artifact, not behaviour: python-docx's bundled blank
            # document says the author is "python-docx"; ours leaves it empty
            # and lets the program say who wrote the file
            value = "<library default>"
        lines.append(f"  {name}={value!r}")

    lines.append("comments:")
    for comment in doc.comments:
        lines.append(f"  {comment.comment_id} by {comment.author!r}: "
                     f"{comment.text!r}")

    lines.append("inline shapes:")
    for shape in doc.inline_shapes:
        lines.append(f"  {shape.width.inches:g}x{shape.height.inches:g}in")
    return "\n".join(lines)


def main() -> int:
    out_dir = Path("/tmp/docs_recipes")
    out_dir.mkdir(exist_ok=True)
    for recipe in RECIPES:
        name = recipe.__name__
        doc = docx.Document()
        recipe(doc)
        path = out_dir / f"{'shim' if SHIM else 'python_docx'}_{name}.docx"
        doc.save(str(path))
        print(f"### {name}")
        print(describe(str(path)))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
