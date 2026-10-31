"""Build a report from scratch: pictures, a header, merged cells, a style.

    python3 examples/build_report.py out/report.docx

Everything here is one call per idea, and every measurement is EMU, so
``Inches(2.5)`` and ``Pt(9)`` read the way python-docx code reads. Run it with
``--shim`` to build the same file through ``docviews.compat`` instead, which is
the door python-docx code walks through:

    python3 examples/build_report.py out/report_shim.docx --shim
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

SAMPLE_PNG = Path(__file__).resolve().parent.parent / "tests" / "blobs.py"


def load_logo_bytes() -> bytes:
    """A logo to insert: the tiny PNG the test suite carries."""
    sys.path.insert(0, str(SAMPLE_PNG.parent))
    import blobs
    return blobs.PNG_96


def build(document, Inches, Pt) -> None:
    """The report. ``document`` is either Document or the compat Document."""
    document.add_heading("Revenue review", 0)
    document.add_paragraph("Prepared by Finance", style="Subtitle")

    section = document.sections[0]
    section.header.add_paragraph("Confidential — internal distribution only")
    section.footer.add_paragraph("Q3 FY26")
    section.left_margin = Inches(1.0)

    document.add_paragraph("Totals are in thousands of dollars.")
    document.add_picture(load_logo_bytes(), width=Inches(2.5))
    document.add_paragraph()

    document.styles.add_style("Callout", "paragraph", bold=True, size=Pt(11))
    document.add_paragraph("Excludes inter-company revenue.", style="Callout")

    table = document.add_table(rows=4, cols=3)
    rows = [["Region", "Q2", "Q3"],
            ["EMEA", "1,204", "1,332"],
            ["Americas", "2,910", "3,004"],
            ["Asia Pacific", "980", "1,150"]]
    for row_index, values in enumerate(rows):
        for column, value in enumerate(values):
            table.cell(row_index, column).text = value
            if row_index == 0:
                table.cell(row_index, column).paragraphs[0].runs[0].bold = True

    table.cell(0, 1).merge(table.cell(0, 2))          # "Q2"/"Q3" stay a header
    table.cell(3, 0).merge(table.cell(3, 2))          # a totals row, spanning
    table.cell(3, 0).text = "Total                    5,094    5,486"

    document.core_properties.title = "Revenue review — Q3 FY26"
    document.core_properties.author = "Finance"


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    target = Path(sys.argv[1])
    target.parent.mkdir(parents=True, exist_ok=True)

    if "--shim" in sys.argv:
        import docviews.compat as docx
        from docviews.compat.shared import Inches, Pt
        document = docx.Document()
    else:
        from docviews import Document
        from docviews.units import Inches, Pt
        document = Document.new()

    build(document, Inches, Pt)
    document.save(target)

    # read it back with the same library and report what the file contains
    from docviews import open_docx
    reopened = open_docx(target)
    print(f"wrote {target} ({target.stat().st_size:,} bytes)")
    print(f"  paragraphs in the body: {len([p for p in reopened.paragraphs()])}")
    print(f"  pictures: {len(reopened.inline_shapes)} "
          f"({reopened.inline_shapes[0].media} at "
          f"{reopened.inline_shapes[0].width / 914400:.2f} in wide)")
    # the paragraphs, not their join: a header written through python-docx's
    # shape starts with Word's empty Header-styled paragraph, and the shim run
    # of this example reproduces that too
    print(f"  header paragraphs: "
          f"{[p.text for p in reopened.sections[0].header.paragraphs]!r}")
    print(f"  footer paragraphs: "
          f"{[p.text for p in reopened.sections[0].footer.paragraphs]!r}")
    print(f"  table row 0: {[cell.text for cell in reopened.tables[0].rows[0].cells]}")
    print(f"  styles: Callout in file: {'Callout' in reopened.styles}")
    print(f"  title: {reopened.core_properties.title!r}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
