"""Authoring: sections, headers and footers, styles, pictures.

Split out of ``document.py`` because these are the parts that *grow* a package
rather than read it. Every one of them follows the same rule as the rest of the
writer: state the relationship, the content-type entry and the XML that refers
to each other consistently, and touch nothing else in the file.
"""

from __future__ import annotations

import copy
from dataclasses import dataclass, field
from typing import Dict, Iterable, List, Optional, Sequence, Tuple, Union

from xml.etree import ElementTree as ET

from .images import ImageInfo, detect
from .units import Length
from .ooxml import TRANSITIONAL, DocxError, Package, parse_xml, root_ns
from .views import Paragraph, StyleRef, _attr, _tag, views_of
from .write import (CONTENT_TYPES, HEADER_FOOTER_TYPES, IMAGE_CONTENT_TYPES,
                    Builder, serialize)

# Python-docx's own default page: US Letter, 1in top/bottom, 1.25in sides.
# Matching it means a program that never sets page geometry keeps the page it
# had; a document that *has* geometry is still read exactly as it is.
DEFAULT_SECTION = ("<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/>"
                   "<w:pgMar w:top=\"1440\" w:right=\"1800\" w:bottom=\"1440\" "
                   "w:left=\"1800\" w:header=\"720\" w:footer=\"720\" "
                   "w:gutter=\"0\"/></w:sectPr>")


# ------------------------------------------------------------------ inline image
@dataclass
class InlineShape:
    """A picture in the document, as the file describes it."""

    part_name: str
    name: str
    width: int                      # EMU
    height: int
    embed: str                      # relationship id inside that part
    media: str = ""                 # the part the relationship points at
    filename: str = ""
    graphic_uri: str = ""            # graphicData URI: picture, chart, diagram, …
    _element: Optional[ET.Element] = field(default=None, repr=False, compare=False)

    @property
    def size(self):
        """``(width, height)`` as :class:`~docviews.units.Length` values."""
        return (Length(self.width), Length(self.height))

    @property
    def type(self) -> str:
        """python-docx reports an enum; the format string is more useful."""
        return self.filename.rsplit(".", 1)[-1].lower() if self.filename else ""

    @property
    def sha1(self) -> str:
        return self.embed

    def __repr__(self) -> str:
        return (f"InlineShape({self.name!r}, {self.width / 914400:.2f}x"
                f"{self.height / 914400:.2f} in, {self.media or 'unresolved'})")


def _inline_shape(part_name: str, inline: ET.Element,
                  rels: Dict[str, str]) -> InlineShape:
    """Build a shape record from one ``wp:inline`` element."""
    width = height = 0
    name = filename = embed = graphic_uri = ""
    for child in inline.iter():
        if not isinstance(child.tag, str):
            continue
        local = _tag(child)
        if local == "extent":
            width = int(child.get("cx") or 0)
            height = int(child.get("cy") or 0)
        elif local == "docPr":
            name = child.get("name") or ""
        elif local == "cNvPr":
            filename = filename or (child.get("name") or "")
        elif local == "graphicData":
            graphic_uri = child.get("uri") or ""
        elif local == "blip":
            for attr, value in child.attrib.items():
                if attr.endswith("}embed") or attr.endswith("}link"):
                    embed = value
    return InlineShape(part_name=part_name, name=name, width=width,
                       height=height, embed=embed, media=rels.get(embed, ""),
                       filename=filename, graphic_uri=graphic_uri,
                       _element=inline)


def find_inline_shapes(package: Package,
                       roots: Optional[Dict[str, ET.Element]] = None
                       ) -> List[InlineShape]:
    """Every ``wp:inline`` in every part, with its relationship resolved.

    ``roots`` overrides the tree to read for the parts it names — pass the live
    tree of a part being edited, so a picture inserted a moment ago is visible
    before anything is saved.
    """
    shapes: List[InlineShape] = []
    part_roots = {part.name: part.root for part in package.text_parts()}
    part_roots.update(roots or {})
    for part_name, root in part_roots.items():
        rels = find_part_rels(package, part_name)
        for inline in root.iter():
            if _tag(inline) != "inline":
                continue
            shapes.append(_inline_shape(part_name, inline, rels))
    return shapes


def find_body_inline_shapes(package: Package, part_name: str,
                            body: ET.Element) -> List[InlineShape]:
    """The inline drawings python-docx exposes on ``Document.inline_shapes``.

    python-docx's public collection uses the body XPath
    ``.//w:p/w:r/w:drawing/wp:inline``: paragraphs may be nested in tables or
    content controls, but runs and drawings must be direct children at each
    step. Drawings in headers, footers and comments are outside this collection.
    The core reader has a broader package-wide view; the compat door follows
    python-docx's body selector so a migration gets the same boundary.
    """
    rels = find_part_rels(package, part_name)
    shapes: List[InlineShape] = []
    for paragraph in body.iter():
        if _tag(paragraph) != "p":
            continue
        for run in paragraph:
            if _tag(run) != "r":
                continue
            for drawing in run:
                if _tag(drawing) != "drawing":
                    continue
                for inline in drawing:
                    if _tag(inline) == "inline":
                        shapes.append(_inline_shape(part_name, inline, rels))
    return shapes


def _enum_name(value) -> str:
    """The name of an enum member, whichever library's member it is."""
    for attribute in ("name", "_name"):
        found = getattr(value, attribute, None)
        if isinstance(found, str) and found:
            return found
    return str(getattr(value, "value", value))


def _as_datetime(value: str):
    """Core-properties dates are ISO-8601 strings; callers want datetimes."""
    if not value:
        return None
    from datetime import datetime
    for pattern in ("%Y-%m-%dT%H:%M:%SZ", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(value.strip(), pattern)
        except ValueError:
            continue
    return None


def _as_text(value) -> str:
    """A datetime back to the ISO-8601 string the file wants."""
    if hasattr(value, "strftime"):
        return value.strftime("%Y-%m-%dT%H:%M:%SZ")
    return "" if value is None else str(value)

def find_part_rels(package: Package, part_name: str) -> Dict[str, str]:
    """relationship id -> part name (or URL) for one part."""
    """relationship id -> part name, for one part."""
    if "/" in part_name:
        folder, tail = part_name.rsplit("/", 1)
        rels_name = f"{folder}/_rels/{tail}.rels"
    else:
        rels_name = f"_rels/{part_name}.rels"
    if not package.exists(rels_name):
        return {}
    table: Dict[str, str] = {}
    try:
        root = parse_xml(package._zip.read(rels_name))
    except Exception:                                        # noqa: BLE001
        return {}
    base = part_name.rsplit("/", 1)[0] if "/" in part_name else ""
    for rel in root:
        target = rel.get("Target", "")
        if not target:
            continue
        if rel.get("TargetMode") == "External":
            table[rel.get("Id", "")] = target        # a URL, not a part
            continue
        name = target.lstrip("/")
        if not name.startswith("word/") and base and not name.startswith("../"):
            name = f"{base}/{name}"
        elif name.startswith("../"):
            name = name.replace("../", "", 1)
        table[rel.get("Id", "")] = name
    return table


# ------------------------------------------------------------------ sections
class Section:
    """One ``w:sectPr``: page size, margins, header, footer."""

    def __init__(self, document, element: Optional[ET.Element] = None):
        self._document = document
        self._element = element
        self._header: Optional["_HeaderFooter"] = None
        self._footer: Optional["_HeaderFooter"] = None

    # -- the elements this section edits
    def _sect(self) -> ET.Element:
        if self._element is None:
            self._element = self._document._sectPr()
        return self._element

    def _child(self, name: str) -> Optional[ET.Element]:
        for child in self._sect():
            if _tag(child) == name:
                return child
        return None

    def _replace(self, name: str, **attrs) -> ET.Element:
        section = self._sect()
        existing = self._child(name)
        if existing is not None:
            for key, value in attrs.items():
                existing.set(self._document._w(f"{key}"), str(int(value)))
            return existing
        element = ET.SubElement(section, self._document._w(name))
        for key, value in attrs.items():
            element.set(self._document._w(f"{key}"), str(int(value)))
        return element

    # -- geometry, in EMU like python-docx
    @property
    def page_width(self) -> int:
        element = self._child("pgSz")
        value = _attr(element, "w") if element is not None else ""
        return Length(int(value) * 635) if value else Length(12240 * 635)

    @page_width.setter
    def page_width(self, value) -> None:
        height = self._child("pgSz")
        self._replace("pgSz", w=int(round(value / 635)),
                      h=int(_attr(height, "h") or 16838))

    @property
    def page_height(self) -> int:
        element = self._child("pgSz")
        value = _attr(element, "h") if element is not None else ""
        return Length(int(value) * 635) if value else Length(15840 * 635)

    @page_height.setter
    def page_height(self, value) -> None:
        self._replace("pgSz", w=int(_attr(self._child("pgSz"), "w") or 11906),
                      h=int(round(value / 635)))

    # what a document that states no geometry gets: the same numbers the blank
    # document carries (Letter, 1in top/bottom, 1.25in sides, 0.5in header)
    MARGIN_DEFAULTS = {"top": 1440, "bottom": 1440, "left": 1800, "right": 1800,
                       "header": 720, "footer": 720, "gutter": 0}

    def _margin(self, name: str) -> int:
        element = self._child("pgMar")
        value = _attr(element, name) if element is not None else ""
        if not value:
            value = str(self.MARGIN_DEFAULTS.get(name, 1440))
        return Length(int(value) * 635)

    def _set_margin(self, name: str, value) -> None:
        margin = self._child("pgMar")
        if margin is None:
            margin = self._replace("pgMar", top=1440, right=1440, bottom=1440,
                                   left=1440, header=708, footer=708, gutter=0)
        margin.set(self._document._w(name), str(int(round(value / 635))))

    header_distance = property(lambda self: self._margin("header"),
                               lambda self, v: self._set_margin("header", v))
    footer_distance = property(lambda self: self._margin("footer"),
                               lambda self, v: self._set_margin("footer", v))

    @property
    def different_first_page_header_footer(self) -> bool:
        """True when the first page gets its own header/footer (``w:titlePg``)."""
        return self._child("titlePg") is not None

    @different_first_page_header_footer.setter
    def different_first_page_header_footer(self, value) -> None:
        section = self._sect()
        for child in list(section):
            if _tag(child) == "titlePg":
                section.remove(child)
        if value:
            section.append(ET.Element(self._document._w("titlePg")))
        self._touch()

    left_margin = property(lambda self: self._margin("left"),
                           lambda self, v: self._set_margin("left", v))
    right_margin = property(lambda self: self._margin("right"),
                            lambda self, v: self._set_margin("right", v))
    top_margin = property(lambda self: self._margin("top"),
                          lambda self, v: self._set_margin("top", v))
    bottom_margin = property(lambda self: self._margin("bottom"),
                             lambda self, v: self._set_margin("bottom", v))

    # -- header and footer
    def _typed(self, kind: str, which: str) -> "_HeaderFooter":
        cache_name = f"_{which}_{kind}"
        cached = getattr(self, cache_name, None)
        if cached is None:
            cached = _HeaderFooter(self._document, self, kind, which=which)
            setattr(self, cache_name, cached)
        return cached

    @property
    def first_page_header(self) -> "_HeaderFooter":
        """The header on this section's first page (``w:titlePg``)."""
        return self._typed("header", "first")

    @property
    def first_page_footer(self) -> "_HeaderFooter":
        return self._typed("footer", "first")

    @property
    def even_page_header(self) -> "_HeaderFooter":
        """The header on even pages — sets ``w:evenAndOddHeaders`` when written."""
        return self._typed("header", "even")

    @property
    def even_page_footer(self) -> "_HeaderFooter":
        return self._typed("footer", "even")

    @property
    def orientation(self):
        """``WD_ORIENT.PORTRAIT`` or ``WD_ORIENT.LANDSCAPE`` — ``pgSz/@w:orient``.

        Read exactly the way python-docx reads it: the attribute on ``w:pgSz``,
        defaulting to portrait, *not* inferred from which measurement is larger.
        A program that compares this value keeps comparing the same value after
        the switch.
        """
        from .compat.enum import WD_ORIENT
        element = self._child("pgSz")
        declared = (_attr(element, "orient") or "").lower() if element is not None \
            else ""
        return WD_ORIENT.LANDSCAPE if declared.startswith("land") \
            else WD_ORIENT.PORTRAIT

    @orientation.setter
    def orientation(self, value) -> None:
        """Set ``w:orient``; the page measurements are left alone.

        python-docx 1.2.0 does the same thing (``Section.orientation`` writes the
        attribute and nothing else), so a program that sets landscape and reads
        ``page_width`` back sees the same number it saw before the migration.
        Use ``page_width``/``page_height`` to change the paper.
        """
        landscape = "landscape" in _enum_name(value).lower()
        if self._child("pgSz") is None:
            self._replace("pgSz", w=12240, h=15840)
        element = self._child("pgSz")
        element.set(self._document._w("orient"),
                    "landscape" if landscape else "portrait")
        self._touch()

    @property
    def start_type(self):
        """How the section begins: ``CONTINUOUS``, ``NEW_PAGE``, ``ODD_PAGE`` …"""
        from .compat.enum import WD_SECTION_START
        names = {"continuous": "CONTINUOUS", "nextColumn": "NEW_COLUMN",
                 "nextPage": "NEW_PAGE", "evenPage": "EVEN_PAGE",
                 "oddPage": "ODD_PAGE"}
        element = self._child("type")
        value = (_attr(element, "val") if element is not None else "") or "nextPage"
        return getattr(WD_SECTION_START, names.get(value, "NEW_PAGE"))

    @start_type.setter
    def start_type(self, value) -> None:
        names = {"CONTINUOUS": "continuous", "NEW_COLUMN": "nextColumn",
                 "NEW_PAGE": "nextPage", "EVEN_PAGE": "evenPage",
                 "ODD_PAGE": "oddPage"}
        name = _enum_name(value).upper()
        element = self._child("type")
        if element is None:
            element = ET.Element(self._document._w("type"))
            section = self._sect()
            index = 0                       # w:type follows the w:*Reference
            for position, child in enumerate(section):
                if _tag(child) in ("headerReference", "footerReference"):
                    index = position + 1
            section.insert(index, element)
        element.set(self._document._w("val"), names.get(name, "nextPage"))
        self._touch()

    @property
    def gutter(self) -> int:
        """The extra binding space Word leaves at the edge of the page."""
        return self._margin("gutter")

    @gutter.setter
    def gutter(self, value) -> None:
        self._set_margin("gutter", value)

    def iter_inner_content(self):
        """The header and footer, in python-docx's order."""
        yield self.header
        yield self.footer

    def _touch(self) -> None:
        """Say the body changed, so an edited ``w:sectPr`` is written back."""
        self._document._mark_changed(self._document._package.main_part)

    @property
    def header(self) -> "_HeaderFooter":
        """The section's header. Writing to it creates the part on the spot."""
        if self._header is None:
            self._header = _HeaderFooter(self._document, self, "header")
        return self._header

    @property
    def footer(self) -> "_HeaderFooter":
        if self._footer is None:
            self._footer = _HeaderFooter(self._document, self, "footer")
        return self._footer

    def __repr__(self) -> str:
        return (f"Section({self.page_width / 914400:.2f}x"
                f"{self.page_height / 914400:.2f} in)")


class Sections:
    """``doc.sections`` — every section, in document order."""

    def __init__(self, document):
        self._document = document

    def _elements(self) -> List[ET.Element]:
        found: List[ET.Element] = []

        def walk(element) -> None:
            for child in element:
                if not isinstance(child.tag, str):
                    continue
                if _tag(child) == "sectPr":
                    if _tag(element) == "pPr":
                        found.append(child)          # a section break
                    else:
                        continue
                walk(child)

        body = self._document._body()
        walk(body)
        # the last section's properties live directly in the body, or after the
        # last paragraph when the document has one
        if len(body) and _tag(body[-1]) == "sectPr":
            found.append(body[-1])
        elif not found:
            found.append(self._document._sectPr())
        return found

    def __len__(self) -> int:
        return len(self._elements())

    def __getitem__(self, index):
        elements = self._elements()
        try:
            element = elements[index]
        except IndexError:
            raise IndexError(f"no section {index}: the document has "
                             f"{len(elements)}") from None
        return Section(self._document, element)

    def __iter__(self):
        for element in self._elements():
            yield Section(self._document, element)

    def __repr__(self) -> str:
        return f"Sections({len(self)})"


class _HeaderFooter:
    """``section.header`` / ``section.footer`` — read, and created on demand."""

    def __init__(self, document, section: Section, kind: str,
                 which: str = "default"):
        self._document = document
        self._section = section
        self._kind = kind                 # "header" | "footer"
        self._which = which               # "default" | "first" | "even"
        self._element: Optional[ET.Element] = None
        self._name: str = ""

    # ------------------------------------------------------------- resolution
    def _reference(self) -> Optional[ET.Element]:
        """The ``w:headerReference``/``w:footerReference`` on this section."""
        for child in self._section._sect():
            if _tag(child) == f"{self._kind}Reference":
                declared = _attr(child, "type") or "default"
                if declared == self._which:
                    return child
        return None

    @property
    def is_linked_to_previous(self) -> bool:
        """True when this section has no header/footer part of its own."""
        return self._reference() is None

    def _part_name(self) -> str:
        if self._name:
            return self._name
        reference = self._reference()
        if reference is None:
            return ""
        rel_id = _attr(reference, "id")
        target = self._document._rel_target(rel_id)
        self._name = target
        return target

    def _ensure_part(self) -> str:
        """Create the part (and wire it up) the first time it is written to."""
        name = self._part_name()
        if name:
            if not self._document._has_part(name):
                self._document._add_part(name, self._new_root(name))
            self._element = self._document._part_root(name)
            return name

        name = self._document._free_part_name(self._kind)
        rel_id = self._document._reserve_rel(
            self._kind, self._document._rel_to_main(name))
        reference = ET.Element(self._document._w(f"{self._kind}Reference"))
        reference.set(self._document._w("type"), self._which)
        if self._which == "first":
            # Word only prints a first-page header when the section says so
            if self._section._child("titlePg") is None:
                self._section._sect().append(
                    ET.Element(self._document._w("titlePg")))
                self._section._touch()
        elif self._which == "even":
            self._document.settings.odd_and_even_pages_header_footer = True
        reference.set(self._document._r("id"), rel_id)
        section = self._section._sect()
        section.insert(0, reference)          # references come first in w:sectPr
        self._name = name
        self._element = self._document._add_part(name, self._new_root(name))
        return name

    def _new_root(self, name: str) -> ET.Element:
        """An empty header/footer part.

        The root element is ``w:hdr``/``w:ftr`` — not ``w:header``; a part with
        the wrong root tag is silently not a header to other readers.

        No placeholder paragraph either: ``header.text`` right after
        ``add_paragraph("x")`` should read ``"x"``, not ``"\nx"``. The schema
        needs at least one block element, so ``to_bytes`` fills one in when a
        header was created and left empty.
        """
        root_tag = "hdr" if self._kind == "header" else "ftr"
        return ET.Element(self._document._w(root_tag))

    # --------------------------------------------------------------- reading
    @property
    def _root(self) -> Optional[ET.Element]:
        name = self._part_name()
        if not name:
            return None
        if self._element is None:
            self._element = self._document._part_root(name)
        return self._element

    @property
    def paragraphs(self) -> List[Paragraph]:
        """The header/footer paragraphs (empty when the section has none)."""
        root = self._root
        if root is None:
            return []
        from .views import views_of as _views
        out: List[Paragraph] = []
        index = 0
        for element in root:
            if _tag(element) != "p":
                continue
            out.append(Paragraph(
                text=_views(element).get("accepted") or "", view="accepted",
                kind=self._kind, part_name=self._part_name(), index=index,
                _element=element,
                _root=root, _w_ns=root_ns(root)))
            index += 1
        return out

    @property
    def text(self) -> str:
        return "\n".join(p.text for p in self.paragraphs)

    # --------------------------------------------------------------- writing
    def add_paragraph(self, text: str = "", **fmt) -> Paragraph:
        """Add a paragraph to the header/footer, creating the part if needed."""
        name = self._ensure_part()
        root = self._element
        builder = Builder({"w": root_ns(root), "r": TRANSITIONAL["r"]})
        element = builder.paragraph(text, **fmt)
        root.append(element)
        self._document._mark_changed(name)
        return Paragraph(text=text, view="accepted", kind=self._kind,
                         part_name=name, index=len(self.paragraphs) - 1,
                         _element=element, _root=root, _w_ns=builder.w)

    @property
    def tables(self) -> List[object]:
        """Tables written into this header/footer (python-docx's ``tables``)."""
        root = self._root
        if root is None:
            return []
        from .write import Table
        return [Table(child, root_ns(root)) for child in root
                if _tag(child) == "tbl"]

    def add_table(self, rows: int = 1, cols: int = 1, *, width=None,
                  text: Optional[Sequence[Sequence[str]]] = None):
        """A table in the header/footer, creating the part if it has none."""
        from .write import Table, twips
        name = self._ensure_part()
        root = self._element
        builder = Builder({"w": root_ns(root), "r": TRANSITIONAL["r"]})
        element = builder.table(rows, cols, style="TableGrid",
                                widths=[twips(width)] * cols if width else None,
                                text=text)
        root.append(element)
        self._document._mark_changed(name)
        return Table(element, builder.w)

    def iter_inner_content(self):
        """Paragraphs and tables, in the order they sit in the part."""
        root = self._root
        if root is None:
            return
        from .write import Table
        known = {id(paragraph._element): paragraph for paragraph in self.paragraphs}
        for child in root:
            if id(child) in known:
                yield known[id(child)]
            elif _tag(child) == "tbl":
                yield Table(child, root_ns(root))



    def __repr__(self) -> str:
        state = "linked to previous" if self.is_linked_to_previous \
            else repr(self.text[:30])
        return f"<{self._kind}: {state}>"


# ------------------------------------------------------------------- styles
# The internal names Word writes for its own styles, and the names it shows.
# ``w:name w:val="heading 1"`` is what is in the file; ``style.name`` is what a
# python-docx program expects to read, so the UI spelling is what we return.
_UI_NAMES = {
    "heading 1": "Heading 1", "heading 2": "Heading 2", "heading 3": "Heading 3",
    "heading 4": "Heading 4", "heading 5": "Heading 5", "heading 6": "Heading 6",
    "heading 7": "Heading 7", "heading 8": "Heading 8", "heading 9": "Heading 9",
    "normal": "Normal", "title": "Title", "subtitle": "Subtitle",
    "quote": "Quote", "intense quote": "Intense Quote",
    "list bullet": "List Bullet", "list number": "List Number",
    "list paragraph": "List Paragraph", "no spacing": "No Spacing",
    "table grid": "Table Grid", "normal table": "Normal Table",
    "toc heading": "TOC Heading", "caption": "Caption",
    "header": "Header", "footer": "Footer",
}


def _ui_name(internal: str) -> str:
    """The name Word shows for a built-in style (``heading 1`` -> ``Heading 1``)."""
    if not internal:
        return internal
    return _UI_NAMES.get(internal.strip().lower(), internal)


class Style:
    """One ``w:style``: a name, an id, and what it is a style *of*."""

    def __init__(self, document, style_id: str, name: str = "", kind: str = "paragraph"):
        self._document = document
        self._id = style_id
        self._name = name or style_id
        self._kind = kind

    @property
    def style_id(self) -> str:
        return self._id

    @property
    def name(self) -> str:
        return _ui_name(self._name)

    @name.setter
    def name(self, value: str) -> None:
        root = self._document._styles_root(create=True)
        for style in root:
            if _tag(style) == "style" and _attr(style, "styleId") == self._id:
                for child in style:
                    if _tag(child) == "name":
                        child.set(self._document._w("val"), value)
        self._name = value

    @property
    def type(self) -> str:
        return self._kind

    def __eq__(self, other) -> bool:
        if isinstance(other, Style):
            return self._id == other._id
        return self._id == other

    def __hash__(self) -> int:
        return hash(self._id)

    # ------------------------------------------------------- reading, writing
    def _element_node(self) -> Optional[ET.Element]:
        """The ``w:style`` element itself, so the properties below can write."""
        root = self._document._styles_root()
        if root is None:
            return None
        for element in root:
            if _tag(element) == "style" and _attr(element, "styleId") == self._id:
                return element
        return None

    @property
    def builtin(self) -> bool:
        """True for a style Word ships itself (no ``w:customStyle``)."""
        element = self._element_node()
        if element is None:
            return False
        return _attr(element, "customStyle") not in ("1", "true")

    @property
    def font(self):
        """The style's character formatting, named as ``run.font`` names it."""
        from .write import Font, Run
        element = self._element_node()
        if element is None:
            return None
        return Font(Run(element, self._document._w("").strip("{}")))

    @property
    def paragraph_format(self):
        """The style's paragraph formatting (``w:pPr``), read and writable."""
        from .write import ParagraphFormat
        element = self._element_node()
        if element is None:
            return None
        return ParagraphFormat(element, self._document._w("").strip("{}"))

    def _child_value(self, field: str, value: Optional[str] = None):
        """Read — or, given ``value``, write — a plain child element's w:val."""
        element = self._element_node()
        if element is None:
            return None
        for child in element:
            if _tag(child) != field:
                continue
            if value is None:
                return _attr(child, "val") or None
            child.set(self._document._w("val"), value)
            self._document._touch_styles()
            return value
        if value is None:
            return None
        new_element = ET.SubElement(element, self._document._w(field))
        new_element.set(self._document._w("val"), value)
        self._document._touch_styles()
        return value

    def _style_ref(self, style_id: str) -> Style:
        # Use the core collection directly. A compatibility subclass may expose
        # a stricter, python-docx-shaped ``document.styles.get_by_id`` signature.
        try:
            return Styles(self._document).get_by_id(style_id)
        except KeyError:
            return Style(self._document, style_id, style_id, self._kind)

    @property
    def base_style(self) -> Optional["Style"]:
        """The style this one inherits from (``w:basedOn``)."""
        style_id = self._child_value("basedOn")
        return self._style_ref(style_id) if style_id else None

    @base_style.setter
    def base_style(self, value) -> None:
        self._child_value("basedOn", str(getattr(value, "style_id", value) or ""))

    @property
    def next_paragraph_style(self) -> Optional["Style"]:
        """The style Word moves to when the author presses Enter."""
        style_id = self._child_value("next")
        return self._style_ref(style_id) if style_id else None

    @next_paragraph_style.setter
    def next_paragraph_style(self, value) -> None:
        self._child_value("next", str(getattr(value, "style_id", value) or ""))

    def _flag(self, name: str, value=None):
        element = self._element_node()
        if element is None:
            return None if value is None else value
        for child in list(element):
            if _tag(child) != name:
                continue
            if value is None:
                return True
            element.remove(child)
            break
        if value is None or not value:
            if value is not None:
                self._document._touch_styles()
            return None if value is None else False
        ET.SubElement(element, self._document._w(name))
        self._document._touch_styles()
        return True

    @property
    def hidden(self):
        """True when the style is hidden from Word's UI (``w:hidden``)."""
        return self._flag("hidden")

    @hidden.setter
    def hidden(self, value):
        self._flag("hidden", value)

    @property
    def quick_style(self):
        """True when the style is in Word's gallery (``w:qFormat``)."""
        return self._flag("qFormat")

    @quick_style.setter
    def quick_style(self, value):
        self._flag("qFormat", value)

    @property
    def unhide_when_used(self):
        return self._flag("unhideWhenUsed")

    @unhide_when_used.setter
    def unhide_when_used(self, value):
        self._flag("unhideWhenUsed", value)

    @property
    def locked(self):
        """True when the style cannot be changed in Word's UI."""
        return self._flag("locked")

    @locked.setter
    def locked(self, value):
        self._flag("locked", value)

    @property
    def priority(self) -> Optional[int]:
        """Where the style sorts in Word's list (``w:uiPriority``)."""
        value = self._child_value("uiPriority")
        try:
            return int(value) if value is not None else None
        except ValueError:
            return None

    @priority.setter
    def priority(self, value) -> None:
        if value is None:
            element = self._element_node()
            if element is not None:
                for child in list(element):
                    if _tag(child) == "uiPriority":
                        element.remove(child)
                self._document._touch_styles()
            return
        self._child_value("uiPriority", str(int(value)))

    def delete(self) -> None:
        """Remove the style from the document (python-docx's ``delete``)."""
        root = self._document._styles_root()
        element = self._element_node()
        if root is None or element is None:
            return
        root.remove(element)
        self._document._touch_styles()

    def __repr__(self) -> str:
        return f"Style({self._name!r}, {self._kind})"


@dataclass
class LatentStyle:
    """One ``w:lsdException``: a style Word knows but the document never used."""

    name: str
    locked: bool = False
    semi_hidden: bool = False
    unhide_when_used: bool = False
    ui_priority: int = 0

    def __repr__(self) -> str:
        return f"LatentStyle({self.name!r})"


class LatentStyles:
    """``doc.styles.latent_styles`` — the latent-style block, read and written.

    ``w:latentStyles`` lists the styles Word can offer; a ``w:lsdException``
    overrides its behaviour for one of them. Python code mostly reads the
    defaults (font name and size) and adds an exception for a style it ships.
    """

    def __init__(self, document):
        self._document = document

    # -- the block itself
    def _block(self, create: bool = False) -> Optional[ET.Element]:
        root = self._document._styles_root(create=create)
        if root is None:
            return None
        for child in root:
            if _tag(child) == "latentStyles":
                return child
        if not create:
            return None
        element = ET.Element(self._document._w("latentStyles"))
        element.set(self._document._w("defLockedState"), "0")
        element.set(self._document._w("defUIPriority"), "99")
        element.set(self._document._w("defSemiHidden"), "1")
        element.set(self._document._w("defUnhideWhenUsed"), "1")
        element.set(self._document._w("count"), "0")
        root.insert(0, element)
        self._document._touch_styles()
        return element

    def _doc_defaults(self) -> Optional[ET.Element]:
        root = self._document._styles_root()
        if root is None:
            return None
        for child in root:
            if _tag(child) == "docDefaults":
                return child
        return None

    @property
    def default_font_name(self) -> str:
        """The document's default font, out of ``w:docDefaults``."""
        defaults = self._doc_defaults()
        if defaults is None:
            return ""
        for element in defaults.iter():
            if _tag(element) == "rFonts":
                return _attr(element, "ascii") or _attr(element, "hAnsi")
        return ""

    @default_font_name.setter
    def default_font_name(self, value: str) -> None:
        fonts = self._doc_defaults_fonts(create=True)
        if fonts is None:
            return
        for name in ("ascii", "hAnsi", "cs", "eastAsia"):
            fonts.set(self._document._w(name), str(value))
        self._document._touch_styles()

    @property
    def default_font_size(self):
        """The default size as EMU, the way every other length here is."""
        defaults = self._doc_defaults()
        if defaults is None:
            return None
        for element in defaults.iter():
            if _tag(element) == "sz":
                value = _attr(element, "val")
                if value:
                    return int(value) * 635 * 10        # half-points -> EMU
        return None

    @default_font_size.setter
    def default_font_size(self, value) -> None:
        from .write import half_points
        defaults = self._doc_defaults_rpr(create=True)
        if defaults is None:
            return
        for element in list(defaults):
            if _tag(element) == "sz":
                defaults.remove(element)
        element = ET.SubElement(defaults, self._document._w("sz"))
        element.set(self._document._w("val"), str(half_points(value)))
        self._document._touch_styles()

    def _doc_defaults_rpr(self, create: bool = False) -> Optional[ET.Element]:
        defaults = self._doc_defaults()
        if defaults is None:
            if not create:
                return None
            root = self._document._styles_root(create=True)
            defaults = ET.Element(self._document._w("docDefaults"))
            root.insert(0, defaults)
        for child in defaults:
            if _tag(child) == "rPrDefault":
                for inner in child:
                    if _tag(inner) == "rPr":
                        return inner
                return ET.SubElement(child, self._document._w("rPr"))
        missing = ET.SubElement(defaults, self._document._w("rPrDefault"))
        return ET.SubElement(missing, self._document._w("rPr"))

    def _doc_defaults_fonts(self, create: bool = False) -> Optional[ET.Element]:
        defaults = self._doc_defaults()
        if defaults is None:
            if not create:
                return None
        rpr = self._doc_defaults_rpr(create=create)
        if rpr is None:
            return None
        for child in rpr:
            if _tag(child) == "rFonts":
                return child
        return ET.SubElement(rpr, self._document._w("rFonts"))

    # -- the exceptions
    @property
    def defaults(self) -> Dict[str, object]:
        """The block's own defaults, as a dict, for a quick look."""
        block = self._block()
        if block is None:
            return {}
        w = self._document._w
        return {name: _attr(block, name) for name in
                ("defLockedState", "defUIPriority", "defSemiHidden",
                 "defUnhideWhenUsed", "defQFormat", "count")}

    def __iter__(self):
        block = self._block()
        if block is None:
            return
        for child in block:
            if _tag(child) != "lsdException":
                continue
            yield LatentStyle(
                name=_attr(child, "name"),
                locked=_attr(child, "locked") in ("1", "true"),
                semi_hidden=_attr(child, "semiHidden") in ("1", "true"),
                unhide_when_used=_attr(child, "unhideWhenUsed") in ("1", "true"),
                ui_priority=int(_attr(child, "uiPriority") or 0),
            )

    def __len__(self) -> int:
        return sum(1 for _ in self)

    def __contains__(self, name: str) -> bool:
        return any(style.name == name for style in self)

    def __getitem__(self, name: str) -> LatentStyle:
        for style in self:
            if style.name == name:
                return style
        raise KeyError(name)

    def add_latent_style(self, name: str, locked: bool = False,
                         semi_hidden: bool = False,
                         unhide_when_used: bool = False,
                         ui_priority: int = 0) -> LatentStyle:
        """Declare how Word should treat a style it has not seen used yet."""
        block = self._block(create=True)
        element = ET.Element(self._document._w("lsdException"))
        w = self._document._w
        element.set(w("name"), name)
        element.set(w("uiPriority"), str(int(ui_priority)))
        for attribute, flag in (("locked", locked), ("semiHidden", semi_hidden),
                                ("unhideWhenUsed", unhide_when_used)):
            if flag:
                element.set(w(attribute), "1")
        block.append(element)
        block.set(w("count"), str(len(block)))
        self._document._touch_styles()
        return LatentStyle(name, locked, semi_hidden, unhide_when_used,
                           int(ui_priority))

    def __repr__(self) -> str:
        return f"LatentStyles({len(self)} exceptions)"

class Styles:
    """``doc.styles`` — the styles a document declares, and ``add_style``."""

    def __init__(self, document):
        self._document = document

    # -- discovery
    def _pairs(self) -> List[Tuple[str, str, str]]:
        root = self._document._styles_root()
        out: List[Tuple[str, str, str]] = []
        if root is None:
            return out
        for style in root:
            if _tag(style) != "style":
                continue
            style_id = _attr(style, "styleId")
            name = ""
            for child in style:
                if _tag(child) == "name":
                    name = _attr(child, "val")
            if style_id:
                out.append((style_id, name or style_id,
                            _attr(style, "type") or "paragraph"))
        return out

    def __iter__(self):
        for style_id, name, kind in self._pairs():
            yield Style(self._document, style_id, name, kind)

    def __len__(self) -> int:
        return len(self._pairs())

    def __contains__(self, name: str) -> bool:
        wanted = "".join(str(name).lower().split())
        return any(wanted in ("".join(sid.lower().split()),
                              "".join(label.lower().split()))
                   for sid, label, _ in self._pairs())

    def __getitem__(self, name: str) -> Style:
        wanted = "".join(str(name).lower().split())
        for style_id, label, kind in self._pairs():
            if wanted in ("".join(style_id.lower().split()),
                          "".join(label.lower().split())):
                return Style(self._document, style_id, label, kind)
        raise KeyError(name)

    def default(self, style_type=None) -> Optional[Style]:
        """The document's default style of that type (python-docx's ``default``)."""
        kind = _style_type_name(style_type)
        for style in self:
            if style.type != kind:
                continue
            element = style._element_node()
            if element is not None and _attr(element, "default") in ("1", "true"):
                return style
        for style in self:
            if style.type == kind:
                return style
        return None

    def get_by_id(self, style_id: str) -> Style:
        """Look a style up by its id (``Heading1``), not its name."""
        for style in self:
            if style.style_id == style_id:
                return style
        raise KeyError(f"no style with id {style_id!r}")

    def get_style_id(self, style=None, style_type=None) -> str:
        """The id Word writes for a name, a Style object, or a default."""
        if style is None:
            found = self.default(style_type)
            return found.style_id if found is not None else ""
        wanted = str(getattr(style, "name", style) or "")
        for candidate in self:
            if wanted in (candidate.name, candidate.style_id):
                return candidate.style_id
        if wanted.startswith("Heading") or " " in wanted:
            return wanted.replace(" ", "")
        return wanted

    @property
    def latent_styles(self) -> "LatentStyles":
        """Word's latent styles: the ones it can offer but has not used yet."""
        return LatentStyles(self._document)

    def _format_style(self, element: ET.Element, **fmt) -> None:
        """Write ``w:rPr``/``w:pPr`` onto a style — the same tags a run uses."""
        from .write import half_points
        w = self._document._w

        def child(parent: ET.Element, name: str) -> ET.Element:
            for existing in parent:
                if _tag(existing) == name:
                    return existing
            return ET.SubElement(parent, w(name))

        def flag(parent: ET.Element, name: str, value) -> None:
            wanted = None if value is None else bool(value)
            found = None
            for existing in parent:
                if _tag(existing) == name:
                    found = existing
            if wanted is None:
                if found is not None:
                    parent.remove(found)
                return
            if found is None:
                found = ET.SubElement(parent, w(name))
            found.set(w("val"), "1" if wanted else "0")

        rpr = None
        for key in ("bold", "italic", "underline", "size", "color", "font"):
            if fmt.get(key) is not None:
                rpr = child(element, "rPr")
                break
        if rpr is not None:
            flag(rpr, "b", fmt.get("bold"))
            flag(rpr, "i", fmt.get("italic"))
            flag(rpr, "u", None)                    # placeholders, replaced below
            if fmt.get("underline") is not None:
                for existing in list(rpr):
                    if _tag(existing) == "u":
                        rpr.remove(existing)
                if fmt["underline"]:
                    underline = ET.SubElement(rpr, w("u"))
                    underline.set(w("val"), "single")
            if fmt.get("size") is not None:
                for existing in list(rpr):
                    if _tag(existing) in ("sz", "szCs"):
                        rpr.remove(existing)
                for name in ("sz", "szCs"):
                    size = ET.SubElement(rpr, w(name))
                    size.set(w("val"), str(half_points(fmt["size"])))
            if fmt.get("color") is not None:
                for existing in list(rpr):
                    if _tag(existing) == "color":
                        rpr.remove(existing)
                colour = ET.SubElement(rpr, w("color"))
                colour.set(w("val"), str(fmt["color"]).lstrip("#").upper())
            if fmt.get("font") is not None:
                for existing in list(rpr):
                    if _tag(existing) == "rFonts":
                        rpr.remove(existing)
                fonts = ET.SubElement(rpr, w("rFonts"))
                for attribute in ("ascii", "hAnsi", "cs", "eastAsia"):
                    fonts.set(w(attribute), str(fmt["font"]))
        if fmt.get("align") is not None:
            value = str(getattr(fmt["align"], "value", fmt["align"]))
            ppr = child(element, "pPr")
            for existing in list(ppr):
                if _tag(existing) == "jc":
                    ppr.remove(existing)
            justify = ET.SubElement(ppr, w("jc"))
            justify.set(w("val"), value)

    def add_style(self, name: str, style_type=None, builtin: bool = False, *,
                  base: Optional[str] = None, next_style: Optional[str] = None,
                  bold: Optional[bool] = None, italic: Optional[bool] = None,
                  underline: Optional[bool] = None, size=None, color=None,
                  font: Optional[str] = None,
                  align=None) -> Style:
        """Create a paragraph/character/table style and return it.

        ``style_type`` accepts ``WD_STYLE_TYPE`` members, their ints, or the
        strings "paragraph"/"character"/"table". The keyword arguments are this
        library's addition: they save the two-step "create, then format" that
        python-docx needs, and the call ``add_style(name, style_type)`` behaves
        exactly as python-docx's.
        """
        kind = _style_type_name(style_type)
        style_id = "".join(str(name).split())
        root = self._document._styles_root(create=True)
        for style in root:
            if _tag(style) == "style" and _attr(style, "styleId") == style_id:
                raise ValueError(f"a style called {name!r} already exists")
        element = ET.SubElement(root, self._document._w("style"))
        element.set(self._document._w("type"), kind)
        element.set(self._document._w("styleId"), style_id)
        label = ET.SubElement(element, self._document._w("name"))
        label.set(self._document._w("val"), name)
        if not builtin:
            based = ET.SubElement(element, self._document._w("basedOn"))
            based.set(self._document._w("val"), base or (
                "Normal" if kind == "paragraph" else "DefaultParagraphFont"))
            if kind == "paragraph":
                following = ET.SubElement(element, self._document._w("next"))
                following.set(self._document._w("val"), next_style or "Normal")
        self._format_style(element, bold=bold, italic=italic,
                           underline=underline, size=size, color=color,
                           font=font, align=align)
        self._document._touch_styles()
        return Style(self._document, style_id, name, kind)

    def __repr__(self) -> str:
        return f"Styles({len(self)} styles)"


def _style_type_name(value) -> str:
    if value is None:
        return "paragraph"
    if isinstance(value, str):
        text = value.lower()
        return {"paragraph": "paragraph", "character": "character",
                "table": "table", "numbering": "numbering",
                "list": "numbering"}.get(text, "paragraph")
    number = getattr(value, "value", value)
    return {1: "paragraph", 2: "character", 3: "table", 4: "numbering"}.get(
        int(number) if isinstance(number, int) or str(number).isdigit()
        else 1, "paragraph")


# ------------------------------------------------------------- core properties
class CoreProperties:
    """``docProps/core_properties`` — the package's own metadata."""

    FIELDS = {
        "title": ("dc", "title"), "author": ("dc", "creator"),
        "subject": ("dc", "subject"), "comments": ("dc", "description"),
        "category": ("cp", "category"), "keywords": ("cp", "keywords"),
        "last_modified_by": ("cp", "lastModifiedBy"),
        "revision": ("cp", "revision"), "content_status": ("cp", "contentStatus"),
        "created": ("dcterms", "created"), "modified": ("dcterms", "modified"),
        "identifier": ("dc", "identifier"), "language": ("dc", "language"),
        "last_printed": ("cp", "lastPrinted"), "version": ("cp", "version"),
    }
    NAMESPACES = {
        "dc": "http://purl.org/dc/elements/1.1/",
        "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
        "dcterms": "http://purl.org/dc/terms/",
    }

    def __init__(self, document):
        self._document = document

    def _name(self) -> str:
        for candidate in self._document._package.names():
            if candidate.endswith("docProps/core.xml"):
                return candidate
        return "docProps/core.xml"

    def _root(self, create: bool = False) -> Optional[ET.Element]:
        name = self._name()
        if self._document._has_part(name):
            return self._document._part_root(name)
        if not create:
            return None
        root = ET.Element(f"{{{self.NAMESPACES['cp']}}}coreProperties")
        declaration = ET.Element(f"{{{self.NAMESPACES['dc']}}}title")
        declaration.text = ""
        root.append(declaration)
        self._document._add_part(name, root)
        self._document._extra_content_types.append(
            (f"/{name}", "application/vnd.openxmlformats-package."
                         "core-properties+xml", True))
        self._document._reserve_package_rel("metadata/core-properties", name)
        return root

    def _get(self, field: str) -> str:
        prefix, local = self.FIELDS[field]
        root = self._root()
        if root is None:
            return ""
        for child in root:
            if isinstance(child.tag, str) and \
                    child.tag == f"{{{self.NAMESPACES[prefix]}}}{local}":
                return child.text or ""
        return ""

    def _set(self, field: str, value: str) -> None:
        prefix, local = self.FIELDS[field]
        root = self._root(create=True)
        for child in root:
            if isinstance(child.tag, str) and \
                    child.tag == f"{{{self.NAMESPACES[prefix]}}}{local}":
                child.text = value
                self._document._mark_changed(self._name())
                return
        element = ET.SubElement(root, f"{{{self.NAMESPACES[prefix]}}}{local}")
        element.text = value
        self._document._mark_changed(self._name())

    title = property(lambda self: self._get("title"),
                     lambda self, v: self._set("title", v))
    author = property(lambda self: self._get("author"),
                      lambda self, v: self._set("author", v))
    subject = property(lambda self: self._get("subject"),
                       lambda self, v: self._set("subject", v))
    comments = property(lambda self: self._get("comments"),
                        lambda self, v: self._set("comments", v))
    category = property(lambda self: self._get("category"),
                        lambda self, v: self._set("category", v))
    keywords = property(lambda self: self._get("keywords"),
                        lambda self, v: self._set("keywords", v))
    last_modified_by = property(lambda self: self._get("last_modified_by"),
                                lambda self, v: self._set("last_modified_by", v))
    content_status = property(lambda self: self._get("content_status"),
                              lambda self, v: self._set("content_status", v))
    identifier = property(lambda self: self._get("identifier"),
                          lambda self, v: self._set("identifier", v))
    language = property(lambda self: self._get("language"),
                        lambda self, v: self._set("language", v))
    version = property(lambda self: self._get("version"),
                       lambda self, v: self._set("version", v))

    @property
    def created(self):
        """When the document was created — a ``datetime``, as python-docx has it."""
        return _as_datetime(self._get("created"))

    @created.setter
    def created(self, value) -> None:
        self._set("created", _as_text(value))

    @property
    def modified(self):
        return _as_datetime(self._get("modified"))

    @modified.setter
    def modified(self, value) -> None:
        self._set("modified", _as_text(value))

    @property
    def last_printed(self):
        return _as_datetime(self._get("last_printed"))

    @last_printed.setter
    def last_printed(self, value) -> None:
        self._set("last_printed", _as_text(value))

    revision = property(lambda self: self._get("revision"),
                        lambda self, v: self._set("revision", v))

    def __repr__(self) -> str:
        return f"CoreProperties(title={self.title!r}, author={self.author!r})"


# -------------------------------------------------------------------- media
def media_plan(document, image: Union[str, bytes, "ImageInfo"],
               source_name: str = "") -> Tuple[ImageInfo, bytes, str]:
    """(facts, bytes, filename) for an image given as a path, bytes or facts."""
    if isinstance(image, ImageInfo):
        data = b""
        facts = image
        filename = source_name or f"image.{facts.extension}"
        return facts, data, filename
    if isinstance(image, (bytes, bytearray)):
        data = bytes(image)
        filename = source_name or "image"
    else:
        path = image
        data = open(path, "rb").read() if isinstance(path, (str,)) else path.read()
        filename = source_name or str(path).replace("\\", "/").rsplit("/", 1)[-1]
    facts = detect(data)
    if "." not in filename:
        filename = f"{filename}.{facts.extension}"
    return facts, data, filename


def next_media_name(package: Package, extension: str,
                    reserved: Iterable[str] = ()) -> str:
    """``word/media/image7.png`` — the first free image name."""
    import re
    used = set(package.names()) | set(reserved)
    pattern = re.compile(rf"^word/media/image(\d+)\.{re.escape(extension)}$",
                         re.IGNORECASE)
    highest = 0
    for name in used:
        match = pattern.match(name)
        if match:
            highest = max(highest, int(match.group(1)))
    candidate = f"word/media/image{highest + 1}.{extension}"
    while candidate in used:
        highest += 1
        candidate = f"word/media/image{highest + 1}.{extension}"
    return candidate
