"""The adoption test: one script, two libraries, same output.

Run it as written (python-docx) and then with the import swapped to the shim:

    python examples/switch_check.py            # python-docx
    python examples/switch_check.py --shim     # docviews.compat

Everything below is ordinary python-docx code: no docviews-specific call.
"""

from __future__ import annotations

import io
import sys
import tempfile
from pathlib import Path

SHIM = "--shim" in sys.argv

if SHIM:
    import docviews.compat as docx
    from docviews.compat.enum import WD_ALIGN_PARAGRAPH
else:
    import docx
    from docx.enum.text import WD_ALIGN_PARAGRAPH


def build():
    doc = docx.Document()
    doc.add_heading("Quarterly Report", 0)
    doc.add_paragraph("Prepared for the board.", style="Subtitle")

    doc.add_heading("Highlights", level=1)
    for item in ("Revenue grew 18% year over year.",
                 "Churn fell to 2.1%.",
                 "Pipeline coverage is 3.4x."):
        doc.add_paragraph(item, style="List Bullet")

    doc.add_heading("Breakdown", level=2)
    para = doc.add_paragraph("Total revenue: ")
    para.add_run("1,204,332").bold = True
    para.add_run(" INR, up from ")
    para.add_run("1,020,900").italic = True
    para.add_run(" last year.")
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT

    table = doc.add_table(rows=3, cols=2)
    table.style = "Table Grid"
    for row, (region, revenue) in enumerate((("Region", "Revenue"),
                                             ("EMEA", "1,204,332"),
                                             ("APAC", "987,001"))):
        table.cell(row, 0).text = region
        table.cell(row, 1).text = revenue

    doc.add_paragraph("Approved by Finance.", style="Quote")
    return doc


def report(path: str) -> str:
    """Read the file back and describe it in plain text."""
    doc = docx.Document(path)
    lines = ["paragraphs:"]
    for para in doc.paragraphs:
        bold = [run.text for run in para.runs if run.bold]
        lines.append(f"  [{para.style.name if para.style else '-':<12}] {para.text!r}"
                     + (f"  bold={bold}" if bold else ""))
    lines.append("tables:")
    for table in doc.tables:
        for row in table.rows:
            lines.append("  " + " | ".join(cell.text for cell in row.cells))
    lines.append(f"sections: {len(doc.sections)}")
    return "\n".join(lines)


def main() -> int:
    out = Path(tempfile.gettempdir()) / (
        f"switch_check_{'shim' if SHIM else 'python_docx'}.docx")
    doc = build()
    doc.save(str(out))
    print(report(str(out)))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
