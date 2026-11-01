"""The three views of a Word document, and every paragraph in it.

Word lets a reader choose what tracked changes *mean*:

  * **original** — reject every change (the text before review)
  * **accepted** — accept every change (the text the author intends)
  * **markup**   — show everything, deleted words included

python-docx returns a fourth thing: runs that are direct children of a
paragraph, which drops deleted text (an accepted-view behaviour) *and* drops
inserted text (a rejected-view behaviour) at the same time. For a paragraph
containing both, the result is text Word cannot display at all — `CENSUS.md`
§3.2.

This module implements the three views exactly and walks every part that holds
text: body, tables (nested ones too), text boxes, headers, footers, footnotes,
endnotes, comments and content controls. A paragraph's own text never includes
text boxes anchored inside it — those are surfaced as their own paragraphs with
``kind="textbox"``, the way Word treats shapes.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, Iterator, List, Optional, Sequence, Tuple
from xml.etree import ElementTree as ET

from .ooxml import (STRICT, TRANSITIONAL, DocxError, Package, Part,
                    root_ns)
from .write import Builder, Run

VIEWS = ("original", "accepted", "markup")


@dataclass
class Hyperlink:
    """A ``w:hyperlink``: its text and where it points.

    ``address`` is the relationship target when the link is external, and the
    bookmark name when it is internal — python-docx reports both as ``address``
    because a caller only ever wants "where does this go".
    """

    text: str
    address: str
    paragraph: Optional["Paragraph"] = field(default=None, repr=False,
                                             compare=False)
    _element: Optional[ET.Element] = field(default=None, repr=False,
                                           compare=False)

    @property
    def url(self) -> str:
        """python-docx's older name for ``address``."""
        return self.address

    def __repr__(self) -> str:
        return f"Hyperlink({self.text!r} -> {self.address!r})"

# every part that carries text; "textbox" is not a part but a paragraph-kind
# inside the body, kept in this list so callers can ask for it by name
PART_KINDS = ("document", "header", "footer", "comments", "footnotes", "endnotes",
              "textbox")

# Elements python-docx maps to characters; mirroring them keeps a comparison
# honest (tab-vs-space is not a fidelity bug, a missing word is). ``w:br`` is
# not here: only a *line* break is a character — a page or column break is not
# part of the text (python-docx draws the same line, and a program that joins
# ``paragraph.text`` would otherwise gain a stray "\n").
CHAR_FOR = {"tab": "\t", "ptab": "\t", "cr": "\n",
            "noBreakHyphen": "-", "softHyphen": "\u00ad"}
# Never visible as document text (field instructions, properties, markers)
NOT_VISIBLE = {"instrText", "delInstrText", "rPr", "pPr", "proofErr",
               "commentReference", "footnoteRef", "endnoteRef", "annotationRef",
               "bookmarkStart", "bookmarkEnd", "lastRenderedPageBreak"}


def _tag(el: ET.Element) -> str:
    """Local name of an element; comments and PIs have a non-string tag."""
    if not isinstance(el.tag, str):
        return ""                        # <!-- --> and <? ?> nodes
    return el.tag.rsplit("}", 1)[-1]


@dataclass(frozen=True)
class Piece:
    """One run of text with the revision containers that wrap it."""

    text: str
    inserted: bool = False
    deleted: bool = False
    moved_from: bool = False
    moved_to: bool = False
    symbol_unmapped: bool = False

    def in_original(self) -> bool:
        return not (self.inserted or self.moved_to)

    def in_accepted(self) -> bool:
        return not (self.deleted or self.moved_from)

    def revision_kind(self) -> Optional[str]:
        if self.inserted:
            return "insert"
        if self.deleted:
            return "delete"
        if self.moved_from:
            return "move_from"
        if self.moved_to:
            return "move_to"
        return None


def pieces(el: ET.Element, inserted=False, deleted=False, moved_from=False,
           moved_to=False, skip_textboxes: bool = True) -> Iterator[object]:
    """Every visible text piece under ``el``, tagged with its revision context."""
    for child in el:
        tag = _tag(child)
        if tag in MARKER_TAGS:
            yield Marker(MARKER_TAGS[tag], _attr(child, "id"))
            continue
        if tag in NOT_VISIBLE:
            continue
        if tag == "txbxContent" and skip_textboxes:
            continue
        if tag == "ins":
            yield from pieces(child, True, deleted, moved_from, moved_to, skip_textboxes)
        elif tag == "del":
            yield from pieces(child, inserted, True, moved_from, moved_to, skip_textboxes)
        elif tag == "moveFrom":
            yield from pieces(child, inserted, deleted, True, moved_to, skip_textboxes)
        elif tag == "moveTo":
            yield from pieces(child, inserted, deleted, moved_from, True, skip_textboxes)
        elif tag == "t":
            yield Piece(child.text or "", inserted, deleted, moved_from, moved_to)
        elif tag == "delText":
            yield Piece(child.text or "", inserted, True, moved_from, moved_to)
        elif tag == "br":
            # textWrapping (the default when the attribute is absent) is a line
            # break; page/column breaks are layout, not text
            kind = ""
            for attr, value in child.attrib.items():
                if attr.endswith("}type"):
                    kind = value
            char = "\n" if kind in ("", "textWrapping") else ""
            yield Piece(char, inserted, deleted, moved_from, moved_to)
        elif tag in CHAR_FOR:
            yield Piece(CHAR_FOR[tag], inserted, deleted, moved_from, moved_to)
        elif tag == "sym":
            char = None
            for attr, value in child.attrib.items():
                if attr.endswith("}char"):
                    char = value
            # symbol-font private-use codes (F0xx) map onto a font we do not have
            code = int(char, 16) if char and char[0] not in "Ff" else None
            yield Piece(chr(code) if code else "", inserted, deleted,
                        moved_from, moved_to, symbol_unmapped=code is None)
        else:
            yield from pieces(child, inserted, deleted, moved_from, moved_to,
                              skip_textboxes)


@dataclass(frozen=True)
class Marker:
    """A non-text event inside the paragraph's flow: a comment anchor.

    ``offset`` is filled in by whoever walks the text, because it means
    "position in the view text", and that depends on the view.
    """

    kind: str          # range_start | range_end | reference
    id: str
    offset: int = -1


MARKER_TAGS = {"commentRangeStart": "range_start", "commentRangeEnd": "range_end",
               "commentReference": "reference"}


def _attr(el: ET.Element, name: str) -> str:
    for key, value in el.attrib.items():
        if key.endswith("}" + name):
            return value
    return ""


@dataclass
class Views:
    original: str = ""
    accepted: str = ""
    markup: str = ""
    revisions: List[str] = field(default_factory=list)
    revision_chars: Dict[str, int] = field(default_factory=dict)   # kind -> chars
    unmapped_symbols: int = 0

    def get(self, view: str) -> str:
        if view not in VIEWS:
            raise ValueError(f"view must be one of {VIEWS}, not {view!r}")
        return getattr(self, view)

    @property
    def diverge(self) -> bool:
        """True when accepting or rejecting changes changes the text."""
        return self.original != self.accepted


def views_of(el: ET.Element, skip_textboxes: bool = True) -> Views:
    """The three views of any element (a paragraph, a table, a whole part)."""
    v = Views()
    for piece in pieces(el, skip_textboxes=skip_textboxes):
        if isinstance(piece, Marker):
            continue
        if piece.symbol_unmapped:
            v.unmapped_symbols += 1
        v.markup += piece.text
        if piece.in_original():
            v.original += piece.text
        if piece.in_accepted():
            v.accepted += piece.text
        kind = piece.revision_kind()
        if kind and piece.text:
            v.revisions.append(kind)
            v.revision_chars[kind] = v.revision_chars.get(kind, 0) + len(piece.text)
    return v


def view_and_markers(el: ET.Element, view: str = "accepted",
                     skip_textboxes: bool = True) -> Tuple[str, List[Marker]]:
    """The view's text of ``el``, plus comment anchors as offsets into it."""
    if view not in VIEWS:
        raise ValueError(f"view must be one of {VIEWS}, not {view!r}")
    chunks: List[str] = []
    markers: List[Marker] = []
    pos = 0
    for ev in pieces(el, skip_textboxes=skip_textboxes):
        if isinstance(ev, Marker):
            markers.append(Marker(ev.kind, ev.id, pos))
            continue
        if ev.symbol_unmapped:
            continue
        keep = (ev.in_original() if view == "original"
                else ev.in_accepted() if view == "accepted" else True)
        if keep:
            chunks.append(ev.text)
            pos += len(ev.text)
    return "".join(chunks), markers


# style ids Word's built-ins carry, and the names users see in the UI. Word
# writes both in styles.xml; a file without one is still readable, so this map
# keeps ``paragraph.style.name`` meaningful.
STYLE_NAMES = {
    "Normal": "Normal", "Title": "Title", "Subtitle": "Subtitle", "Quote": "Quote",
    "ListBullet": "List Bullet", "ListNumber": "List Number",
    "TableGrid": "Table Grid", "TableNormal": "Normal Table",
    "Header": "Header", "Footer": "Footer",
    "Emphasis": "Emphasis", "Strong": "Strong", "SubtleEmphasis": "Subtle Emphasis",
    "DefaultParagraphFont": "Default Paragraph Font",
    "Heading1": "Heading 1", "Heading2": "Heading 2", "Heading3": "Heading 3",
    "Heading4": "Heading 4", "Heading5": "Heading 5", "Heading6": "Heading 6",
    "Heading7": "Heading 7", "Heading8": "Heading 8", "Heading9": "Heading 9",
    "Caption": "Caption", "IntenseQuote": "Intense Quote",
    "NoSpacing": "No Spacing", "TOCHeading": "TOC Heading",
}


class AlignRef(str):
    """A justification value (``"center"``) that also answers ``.value``.

    python-docx hands back ``WD_ALIGN_PARAGRAPH.CENTER``; this is a string so
    ``p.alignment == "center"`` works, with ``.value`` for the other style.
    """

    @property
    def value(self) -> str:
        return str(self)

    # python-docx's member names for the same tokens: WD_ALIGN_PARAGRAPH.JUSTIFY
    # is ``w:jc="both"``, so the file's word is not the API's word.
    MEMBERS = {"both": "JUSTIFY", "distribute": "DISTRIBUTE", "left": "LEFT",
               "center": "CENTER", "right": "RIGHT"}

    @property
    def name(self) -> str:
        """python-docx reports the member name (``"JUSTIFY"``); so does this."""
        return self.MEMBERS.get(str(self), str(self).upper())

    def __repr__(self) -> str:
        return f"AlignRef({str(self)!r})"


class StyleRef(str):
    """A paragraph/table style: a string id that also knows its display name.

    It *is* a str, so ``paragraph.style == "Heading1"`` compares and hashes the
    way a string does, and ``paragraph.style.name`` returns the name Word shows
    ("Heading 1") — which is what python-docx code reaches for.
    """

    @property
    def name(self) -> str:
        """The name Word shows (``Heading 1``), not the id (``Heading1``)."""
        return STYLE_NAMES.get(str(self), str(self))

    @property
    def builtin(self) -> bool:
        """:attr:`STYLE_NAMES` holds the built-ins this library ships."""
        return str(self) in STYLE_NAMES

    @property
    def style_id(self) -> str:
        return str(self)

    @property
    def type(self) -> int:
        return 3 if str(self).startswith("Table") else 1

    def __repr__(self) -> str:
        return f"StyleRef({self.name!r})"


@dataclass
class Paragraph:
    """One paragraph, with where it came from and what it looked like."""

    text: str
    view: str
    kind: str = "document"          # document | textbox | header | footnote …
    part_name: str = ""
    index: int = 0                  # position within its part
    in_table: bool = False
    in_textbox: bool = False
    revision_ids: Tuple[str, ...] = ()
    revision_chars: Dict[str, int] = field(default_factory=dict)
    views: Optional[Views] = None   # filled in unless streaming
    # writing: an element-backed paragraph can be edited and saved
    _element: Optional[ET.Element] = field(default=None, repr=False,
                                           compare=False)
    _root: Optional[ET.Element] = field(default=None, repr=False, compare=False)
    _w_ns: str = field(default=TRANSITIONAL["w"], repr=False, compare=False)
    _default_style: Optional[str] = field(default=None, repr=False, compare=False)
    _rels: Dict[str, str] = field(default_factory=dict, repr=False, compare=False)
    _document: Optional[object] = field(default=None, repr=False, compare=False)

    # ------------------------------------------------------------- writing
    @property
    def editable(self) -> bool:
        return self._element is not None

    def _builder(self) -> Builder:
        return Builder({"w": self._w_ns, "r": TRANSITIONAL["r"]})

    def _require_element(self) -> ET.Element:
        if self._element is None:
            raise DocxError(
                "this paragraph came from a streaming walk (doc.stream()), so it "
                "is not attached to an editable tree. Use doc.paragraphs() if you "
                "intend to write.")
        return self._element

    @property
    def runs(self) -> List[Run]:
        """The runs this paragraph shows, in reading order (python-docx shape).

        Deleted runs are not in the accepted view and are not returned; text
        inside a text box belongs to the text box's own paragraphs. Use
        :meth:`runs_in` to ask for another view.
        """
        return self.runs_in(self.view)

    def runs_in(self, view: str = "accepted") -> List[Run]:
        """The runs a reader sees in ``view``, in order."""
        element = self._require_element()
        out: List[Run] = []
        for child in element:
            tag = _tag(child)
            if tag in ("w:del", "del", "moveFrom", "txbxContent", "delText"):
                continue
            if tag == "r":
                out.append(Run(child, self._w_ns, document=self._document,
                               part_name=self.part_name))
            elif tag in ("ins", "moveTo", "hyperlink", "smartTag", "sdt",
                         "sdtContent", "bdo", "dir", "bookmarkStart",
                         "bookmarkEnd", "proofErr", "commentRangeStart",
                         "commentRangeEnd"):
                out.extend(_runs_under(child, self._w_ns, self._document,
                                       self.part_name))
        return out

    def add_run(self, text: str = "", **fmt) -> Run:
        """Append a run of text, with optional bold/italic/size/color/font."""
        element = self._require_element()
        run = self._builder().run(text, **fmt)
        element.append(run)
        self.refresh()
        return Run(run, self._w_ns, refresh=self.refresh,
                   document=self._document, part_name=self.part_name)

    def set_text(self, text: str, **fmt) -> None:
        """Replace everything the paragraph says (formatting included)."""
        element = self._require_element()
        for child in list(element):
            if _tag(child) != "pPr":
                element.remove(child)
        if text:
            element.append(self._builder().run(text, **fmt))
        self.refresh()

    def clear(self) -> None:
        self.set_text("")

    def delete(self) -> bool:
        """Remove this paragraph from its part. True if it was attached."""
        element = self._require_element()
        if self._root is None:
            return False
        parent = _parent_of(self._root, element)
        if parent is None:
            return False
        parent.remove(element)
        self._element = None
        return True

    def insert_after(self, text: str = "", **fmt) -> "Paragraph":
        """Insert a new paragraph right after this one (Word's Enter key)."""
        element = self._require_element()
        parent = _parent_of(self._root, element) if self._root is not None else None
        if parent is None:
            raise DocxError("this paragraph is not attached to a part")
        new = self._builder().paragraph(text, **fmt)
        index = list(parent).index(element)
        parent.insert(index + 1, new)
        return Paragraph(text=text, view=self.view, kind=self.kind,
                         part_name=self.part_name, index=self.index + 1,
                         in_table=self.in_table, in_textbox=self.in_textbox,
                         _element=new, _root=self._root, _w_ns=self._w_ns)

    @property
    def style(self) -> Optional[StyleRef]:
        """The paragraph style, or ``None`` when the document sets none.

        Returns a :class:`StyleRef`: a string id (``"Heading1"``) that also
        answers ``.name`` (``"Heading 1"``), so both styles of code work.
        """
        element = self._element
        if element is not None:
            for child in element:
                if _tag(child) == "pPr":
                    for prop in child:
                        if _tag(prop) == "pStyle":
                            value = _attr(prop, "val")
                            if value:
                                return StyleRef(value)
        return StyleRef(self._default_style) if self._default_style else None

    @style.setter
    def style(self, value: Optional[str]) -> None:
        element = self._require_element()
        props = None
        for child in element:
            if _tag(child) == "pPr":
                props = child
                break
        if props is None:
            props = ET.Element(f"{{{self._w_ns}}}pPr")
            element.insert(0, props)
        for child in list(props):
            if _tag(child) == "pStyle":
                props.remove(child)
        if value:
            name = getattr(value, "name", value)
            style = ET.Element(f"{{{self._w_ns}}}pStyle")
            style.set(f"{{{self._w_ns}}}val", str(name).replace(" ", ""))
            props.insert(0, style)

    @property
    def alignment(self) -> Optional[str]:
        """The OOXML justification value: left, center, right, both, distribute."""
        element = self._element
        if element is None:
            return None
        for child in element:
            if _tag(child) == "pPr":
                for prop in child:
                    if _tag(prop) == "jc":
                        value = _attr(prop, "val")
                        return AlignRef(value) if value else None
        return None

    @alignment.setter
    def alignment(self, value: Optional[str]) -> None:
        element = self._require_element()
        props = None
        for child in element:
            if _tag(child) == "pPr":
                props = child
                break
        if props is None:
            props = ET.Element(f"{{{self._w_ns}}}pPr")
            element.insert(0, props)
        for child in list(props):
            if _tag(child) == "jc":
                props.remove(child)
        if value is not None:
            from .write import token          # local: views is imported by it
            align = ET.Element(f"{{{self._w_ns}}}jc")
            align.set(f"{{{self._w_ns}}}val", token(value))
            props.append(align)

    def text_in(self, view: str) -> str:
        """This paragraph's text in another view (original/accepted/markup)."""
        if self._element is None:
            return ""
        return views_of(self._element).get(view) or ""

    @property
    def original_text(self) -> str:
        """The same paragraph with every tracked change rejected."""
        return self.text_in("original")

    @property
    def markup_text(self) -> str:
        """The same paragraph with insertions and deletions both shown."""
        return self.text_in("markup")

    # --------------------------------------------------- python-docx helpers
    @property
    def paragraph_format(self):
        """``w:pPr`` as an object: indents, spacing, keep-with-next, alignment."""
        from .write import ParagraphFormat
        return ParagraphFormat(self._require_element(), self._w_ns)

    @property
    def contains_page_break(self) -> bool:
        """True when a page break is written inside this paragraph."""
        return self.rendered_page_breaks > 0

    @property
    def rendered_page_breaks(self) -> int:
        """How many page breaks this paragraph holds."""
        element = self._element
        if element is None:
            return 0
        count = 0
        for child in element.iter():
            if not isinstance(child.tag, str):
                continue
            local = _tag(child)
            if local == "br" and _attr(child, "type") == "page":
                count += 1
            elif local == "lastRenderedPageBreak":
                count += 1
        return count

    @property
    def hyperlinks(self) -> List["Hyperlink"]:
        """The links in this paragraph: the text, and where each one goes."""
        element = self._require_element()
        out: List[Hyperlink] = []
        for child in element:
            if _tag(child) != "hyperlink":
                continue
            rel_id = _attr(child, "id")
            address = self._rels.get(rel_id, "") if rel_id else ""
            if not address:
                address = _attr(child, "anchor")
            out.append(Hyperlink(views_of(child).get("accepted") or "", address,
                                 self, child))
        return out

    def iter_inner_content(self) -> Iterator[object]:
        """Runs and hyperlinks, in document order (python-docx's walk)."""
        element = self._require_element()
        links = {id(link._element): link for link in self.hyperlinks}
        for child in element:
            if _tag(child) == "r":
                yield Run(child, self._w_ns, document=self._document,
                          part_name=self.part_name)
            elif id(child) in links:
                yield links[id(child)]

    def insert_paragraph_before(self, text: str = "",
                                style: Optional[str] = None) -> "Paragraph":
        """A new paragraph directly before this one (python-docx's helper)."""
        element = self._require_element()
        parent = self._parent()
        if parent is None:
            raise DocxError("this paragraph is not attached to a container")
        new_element = self._builder().paragraph(
            text, style=self._style_id(style) if style else None)
        parent.insert(list(parent).index(element), new_element)
        return Paragraph(text=text, view=self.view, kind=self.kind,
                         part_name=self.part_name, index=self.index,
                         _element=new_element, _root=self._root,
                         _w_ns=self._w_ns, _default_style=self._default_style,
                         _rels=dict(self._rels))

    def _parent(self) -> Optional[ET.Element]:
        """The element that holds this paragraph, found through the root."""
        if self._root is None or self._element is None:
            return None
        for candidate in self._root.iter():
            for child in candidate:
                if child is self._element:
                    return candidate
        return None

    def _style_id(self, style) -> str:
        """A style name or id -> the id Word writes (python-docx's style_id)."""
        name = str(getattr(style, "name", style) or "")
        for style_id, label in STYLE_NAMES.items():
            if label == name:
                return style_id
        return name.replace(" ", "")

    def refresh(self) -> "Paragraph":
        """Re-read this paragraph from its element after it was edited."""
        element = self._element
        if element is None:
            return self
        v = views_of(element)
        self.views = v
        self.text = v.get(self.view) or ""
        self.revision_ids = tuple(v.revisions)
        self.revision_chars = dict(v.revision_chars)
        return self

    @property
    def is_revised(self) -> bool:
        return bool(self.revision_ids)

    def __str__(self) -> str:       # pragma: no cover - convenience only
        return self.text


# document order, which is what a reader expects



def _containers(el: ET.Element, in_table: bool, in_textbox: bool):
    """Child elements that may hold paragraphs, with their context flags.

    Recurse through everything except paragraphs themselves (and text boxes,
    which are returned as their own container). Footnote, endnote, header,
    footer and comment parts nest their paragraphs under part-specific
    elements (`w:footnote`, `w:hdr`, `w:comment` …), so a fixed container list
    would silently skip whole parts.
    """
    for child in el:
        tag = _tag(child)
        if tag == "p":
            yield child, in_table, in_textbox, "p"
        elif tag == "txbxContent":
            yield child, in_table, True, "txbxContent"
        else:
            table = in_table or tag in ("tbl", "tc")
            yield from _containers(child, table, in_textbox)


def _find_textboxes(el: ET.Element):
    for sub in el.iter():
        if _tag(sub) == "txbxContent":
            yield sub


class _PartWalker:
    """Walks one part and yields (paragraph, xml element) pairs in reading order.

    The element is handed out alongside the paragraph because comment anchors
    need the paragraph's XML and the same index the reader sees.
    """

    def __init__(self, part_name: str, kind: str, view: str,
                 include_textboxes: bool, with_views: bool, owner=None,
                 default_style: Optional[str] = None):
        self.part_name, self.kind, self.view = part_name, kind, view
        self._owner = owner
        self._default_style = default_style
        self.include_textboxes, self.with_views = include_textboxes, with_views
        self.index = 0
        self._root: Optional[ET.Element] = None
        self._w_ns: str = TRANSITIONAL["w"]
        self._revisions: Optional[Dict[int, Tuple[str, ...]]] = None

    def walk(self, root: ET.Element) -> Iterator[Paragraph]:
        self._root = root
        for paragraph, _element in self.events(root):
            yield paragraph

    def events(self, root: ET.Element):
        self._root = root
        self._w_ns = root_ns(root)
        for el, in_table, in_textbox, tag in _containers(root, False, False):
            if tag == "txbxContent":
                if self.include_textboxes:
                    yield from self._events_under(el, in_table, True)
                continue
            # a paragraph: its own text, then any text box anchored in it
            yield self._paragraph(el, in_table, in_textbox), el
            if self.include_textboxes:
                for box in _find_textboxes(el):
                    yield from self._events_under(box, in_table, True)

    def _revision_kinds(self, root: ET.Element, p: ET.Element) -> Tuple[str, ...]:
        """Revision kinds for a paragraph, from a per-part cache.

        The map depends on the XML alone, so it is built once per part and
        reused by the original/accepted/markup walks.
        """
        if self._revisions is None:
            cached = getattr(self._owner, "_revision_map_cache", None)
            if cached is None:
                cached = _revision_map(root)
                try:
                    self._owner._revision_map_cache = cached
                except AttributeError:          # a part that refuses attributes
                    pass
            self._revisions = cached
        return self._revisions.get(id(p), ())

    def _paragraph(self, p: ET.Element, in_table: bool,
                   in_textbox: bool) -> Paragraph:
        v = views_of(p)
        paragraph = Paragraph(
            text=v.get(self.view), view=self.view,
            kind="textbox" if in_textbox else self.kind,
            part_name=self.part_name, index=self.index,
            in_table=in_table, in_textbox=in_textbox,
            revision_ids=tuple(dict.fromkeys(
                tuple(v.revisions) + self._revision_kinds(self._root, p))),
            revision_chars=dict(v.revision_chars),
            views=v if self.with_views else None,
            _element=p, _root=self._root, _w_ns=self._w_ns,
            _default_style=self._default_style,
        )
        self.index += 1
        return paragraph

    def _events_under(self, el: ET.Element, in_table: bool, in_textbox: bool):
        for p in _paragraph_elements(el):
            yield self._paragraph(p, in_table, in_textbox), p


def _paragraph_elements(el: ET.Element) -> Iterator[ET.Element]:
    """Every <w:p> that belongs to this element's own text flow."""
    for child in el:
        tag = _tag(child)
        if tag == "p":
            yield child
        elif tag != "txbxContent":
            yield from _paragraph_elements(child)


REVISION_KINDS = {"ins": "insert", "del": "delete", "moveFrom": "move_from",
                 "moveTo": "move_to"}


def revision_ids_under(el: ET.Element, skip_textboxes: bool = True
                       ) -> List[Tuple[str, str, str, str]]:
    """(kind, id, author, date) for every revision container under ``el``.

    A nested ``w:txbxContent`` is skipped *with its whole subtree*: the text box
    is walked as its own paragraphs, so counting its revisions here as well
    would count them twice.
    """
    out: List[Tuple[str, str, str, str]] = []

    def walk(node: ET.Element) -> None:
        for sub in node:
            tag = _tag(sub)
            if tag == "txbxContent" and skip_textboxes:
                continue
            kind = REVISION_KINDS.get(tag)
            if kind:
                out.append((kind, _attr(sub, "id"), _attr(sub, "author"),
                            _attr(sub, "date")))
            walk(sub)

    walk(el)
    return out


def own_paragraph_under(el: ET.Element) -> bool:
    """Does ``el``'s own text flow hold a paragraph (a block-level revision)?"""
    for sub in el:
        tag = _tag(sub)
        if tag == "txbxContent":
            continue
        if tag == "p" or own_paragraph_under(sub):
            return True
    return False


def _revision_map(root: ET.Element) -> Dict[int, Tuple[str, ...]]:
    """paragraph element id -> revision kinds touching it, in one pass.

    Walking every paragraph's subtree separately is quadratic on a pathological
    document (the 28 MB bomb costs minutes); one depth-first pass over the part
    costs the same as reading it.
    """
    found: Dict[int, list] = {}

    def walk(el: ET.Element, current: Optional[int]) -> None:
        for child in el:
            tag = _tag(child)
            kind = REVISION_KINDS.get(tag)
            if kind and current is not None:
                found.setdefault(current, []).append(kind)
            if tag == "p":
                walk(child, id(child))
            elif tag == "txbxContent":
                walk(child, None)        # a text box is its own paragraphs
            else:
                walk(child, current)

    walk(root, None)
    return {key: tuple(dict.fromkeys(kinds)) for key, kinds in found.items()}


def _parent_of(root: ET.Element, target: ET.Element) -> Optional[ET.Element]:
    for child in root:
        if child is target:
            return root
        found = _parent_of(child, target)
        if found is not None:
            return found
    return None


def _runs_under(el: ET.Element, w_ns: str, document=None,
                part_name: str = "") -> List[Run]:
    out: List[Run] = []
    for child in el:
        tag = _tag(child)
        if tag == "r":
            out.append(Run(child, w_ns, document=document,
                           part_name=part_name))
        elif tag in ("ins", "moveTo", "hyperlink", "smartTag", "sdt",
                     "sdtContent", "bdo", "dir"):
            out.extend(_runs_under(child, w_ns, document, part_name))
    return out


def default_paragraph_style(package) -> Optional[str]:
    """The styleId declared as the default paragraph style, if any.

    Word shows an unstyled paragraph as "Normal" because styles.xml says so;
    reading ``w:default="1"`` is what makes ``paragraph.style`` agree with the
    UI instead of reporting nothing.
    """
    for name in package.names():
        if not name.endswith("styles.xml") or not package.exists(name):
            continue
        try:
            root = package.part(name).root
        except Exception:                                    # noqa: BLE001
            continue
        for style in root:
            if _tag(style) != "style":
                continue
            if _attr(style, "type") not in ("paragraph", ""):
                continue
            if _attr(style, "default") in ("1", "true"):
                style_id = _attr(style, "styleId")
                if style_id:
                    return style_id
    return None


def _structural_revisions(el: ET.Element) -> Tuple[str, ...]:
    """Revision kinds touching this paragraph, whether or not they show text.

    A deleted image, a deleted text box or an emptied run carries revision
    markup without carrying characters, and the view text cannot show it. The
    audit's question — "does this paragraph change under accept/reject?" — has
    to be answered from the markup, not from the text.
    """
    return tuple(dict.fromkeys(kind for kind, _, _, _
                               in revision_ids_under(el, skip_textboxes=True)))


def part_paragraph_elements(part, view: str = "accepted",
                            include_textboxes: bool = True,
                            with_views: bool = False, default_style=None):
    """(Paragraph, xml element) pairs for one part, in reading order."""
    walker = _PartWalker(part.name, part.kind, view, include_textboxes,
                         with_views, owner=part, default_style=default_style)
    yield from walker.events(part.root)


def part_paragraphs(part, view: str = "accepted",
                    include_textboxes: bool = True,
                    with_views: bool = False, default_style=None
                    ) -> Iterator[Paragraph]:
    """Every paragraph of one part, in reading order."""
    if view not in VIEWS:
        raise ValueError(f"view must be one of {VIEWS}, not {view!r}")
    walker = _PartWalker(part.name, part.kind, view, include_textboxes,
                         with_views, owner=part, default_style=default_style)
    yield from walker.walk(part.root)


def iter_paragraphs(package: Package, view: str = "accepted",
                    kinds: Optional[Sequence[str]] = None,
                    with_views: bool = False) -> Iterator[Paragraph]:
    """Paragraphs from every text-bearing part, in reading order."""
    if view not in VIEWS:
        raise ValueError(f"view must be one of {VIEWS}, not {view!r}")
    wanted = set(kinds) if kinds else set(PART_KINDS)
    default_style = default_paragraph_style(package)
    for part in package.text_parts():
        if part.kind not in wanted:
            continue
        yield from part_paragraphs(part, view,
                                   include_textboxes="textbox" in wanted,
                                   with_views=with_views,
                                   default_style=default_style)


def document_text(package: Package, view: str = "accepted",
                  separator: str = "\n") -> str:
    return separator.join(p.text for p in iter_paragraphs(package, view))


def stream_paragraphs(package: Package, view: str = "accepted",
                      part: Optional[str] = None) -> Iterator[Paragraph]:
    """Walk the main part with ``iterparse``, clearing elements as we go.

    The census measured a 0.49 MB document that costs 416 MB of RAM through
    python-docx, because every part is parsed eagerly and kept. This path holds
    one paragraph at a time, so memory stays flat as the document grows.
    """
    if view not in VIEWS:
        raise ValueError(f"view must be one of {VIEWS}, not {view!r}")
    name = part or package.main_part
    info = package._zip.getinfo(name)
    problem = package.limits.check(name, info.compress_size, info.file_size)
    if problem:
        raise PartTooLarge(problem)
    index = 0
    stack: List[str] = []
    # The tree walk emits the anchor paragraph before text-box paragraphs
    # nested in it. iterparse encounters the nested paragraphs first on their
    # end events, so buffer their small text results until the anchor closes.
    # This preserves the tree walk's reading order without retaining XML trees.
    pending_textbox_paragraphs: List[Paragraph] = []
    with package._zip.open(name) as fh:
        for event, el in ET.iterparse(fh, events=("start", "end")):
            tag = _tag(el)
            if event == "start":
                stack.append(tag)
                continue
            stack.pop()
            if tag != "p":
                if tag in ("tbl", "txbxContent"):
                    el.clear()
                continue
            in_box = "txbxContent" in stack
            if in_box:
                box_at = stack.index("txbxContent")
                anchored = "p" in stack[:box_at]
            else:
                anchored = False
            v = views_of(el)
            paragraph = Paragraph(text=v.get(view), view=view,
                                  kind="textbox" if in_box else "document",
                                  part_name=name, index=index,
                                  in_table="tbl" in stack,
                                  in_textbox=in_box,
                                  revision_ids=_structural_revisions(el))
            if anchored:
                pending_textbox_paragraphs.append(paragraph)
            else:
                paragraph.index = index
                index += 1
                yield paragraph
                if not in_box and pending_textbox_paragraphs:
                    for textbox_paragraph in pending_textbox_paragraphs:
                        textbox_paragraph.index = index
                        index += 1
                        yield textbox_paragraph
                    pending_textbox_paragraphs.clear()
            el.clear()
    # A malformed or unusually structured part may contain an unanchored
    # text box after the last normal paragraph. Do not lose its paragraphs.
    for textbox_paragraph in pending_textbox_paragraphs:
        textbox_paragraph.index = index
        index += 1
        yield textbox_paragraph
