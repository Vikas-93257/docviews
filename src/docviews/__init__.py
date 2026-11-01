"""docviews — read and write .docx the way Word does: all three views, every part.

Read:

    from docviews import open_docx
    doc = open_docx("contract.docx")
    doc.text("accepted")        # Word with changes accepted
    doc.text("original")        # Word with changes rejected
    doc.text("markup")          # everything the file contains
    doc.revisions()             # kind, author, date, position, scope
    doc.comments()              # text *and* the characters it is about
    doc.parts()                 # footnotes, endnotes, headers, footer, comments

Write:

    from docviews import Document
    doc = Document.new()                      # or open_docx("existing.docx")
    doc.add_heading("Report", 1)
    doc.add_paragraph("Body text.", style="Heading 2")
    table = doc.add_table(rows=2, cols=2, text=[["a", "b"], ["c", "d"]])
    table.cell(0, 0).merge(table.cell(0, 1))
    doc.add_picture("chart.png", width=Inches(5))   # Pt/Cm/Mm/Inches/Emu
    doc.sections[0].header.add_paragraph("Confidential")
    doc.styles.add_style("Callout", "paragraph", bold=True, size=Pt(12))
    doc.core_properties.title = "Quarterly report"
    doc.save("out.docx")

    # already written against python-docx?
    import docviews.compat as docx            # the same calls, same shapes

Writing never loses what it does not understand: every part the caller did not
edit is copied out of the package byte-for-byte, so footnotes, comments,
tracked changes, embedded media and extension parts survive an edit (measured
on 831 public documents — see CENSUS.md §6.3). A picture you insert, a header
you write and a merge you make are read back by python-docx unchanged.

Nothing is dropped silently: anything unreadable is reported in ``doc.warnings``.
"""

from .authoring import (CoreProperties, InlineShape, Section, Sections, Style,
                        Styles, find_inline_shapes)
from .comments import Anchor, Comment
from .document import (Document, NumberToken, PartInfo, RevisionInfo,
                       open_docx)
from .exceptions import InvalidSpanError, InvalidXmlError, PythonDocxError
from .images import ImageError, ImageInfo, detect
from .ooxml import DocxError, Limits, Package, Part, PartTooLarge
from .views import (VIEWS, AlignRef, Paragraph, StyleRef, Views, views_of)
from .units import Cm, Emu, Inches, Length, Mm, Pt, RGBColor, Twips
from .write import Builder, Cell, Row, Table

__version__ = "1.0.0"
__all__ = [
    "Document", "open_docx", "PartInfo", "RevisionInfo", "NumberToken",
    "Comment", "Anchor",
    "Package", "Part", "Limits", "DocxError", "PartTooLarge",
    "PythonDocxError", "InvalidSpanError", "InvalidXmlError",
    "VIEWS", "Paragraph", "Views", "views_of", "StyleRef", "AlignRef",
    "Builder", "Table", "Row", "Cell",
    "Section", "Sections", "Style", "Styles", "CoreProperties",
    "InlineShape", "find_inline_shapes",
    "detect", "ImageInfo", "ImageError",
    "Pt", "Inches", "Cm", "Mm", "Emu", "Twips", "Length", "RGBColor",
    "__version__",
]
