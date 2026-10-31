"""``docviews.compat`` — the python-docx-shaped door.

    import docviews.compat as docx        # instead of: import docx

    doc = docx.Document()                # new, empty, saveable
    doc = docx.Document("contract.docx")  # or an opened one you can edit
    doc.add_heading("Q3", 1)
    para = doc.add_paragraph("Revenue grew ")
    para.add_run("18%").bold = True
    doc.tables[0].cell(0, 0).text = "Region"
    doc.save("out.docx")

Nothing is faked: this is the same three-view reader underneath, so every
paragraph also knows its ``original``/``markup`` text, every revision carries a
``scope``, and comments keep their anchor ranges. What this module adds is a
practical compatibility layer for a tested subset of the documented
python-docx API—not a complete import-compatible replacement. The attribute
inventory in ``COMPAT.md`` is not proof of signature or behavioral parity.

Some deliberate differences from python-docx, because the alternative is a lie
about the file:

* ``Document(path)`` does not raise on ISO-Strict packages (python-docx cannot
  open them at all; there are 61 in a public 832-file sample).
* Saving copies every part the caller did not edit byte-for-byte, so footnotes,
  comments, tracked changes, embedded media and extension parts come back
  exactly as they went in.
"""

from __future__ import annotations

import copy
import hashlib
from collections.abc import Sequence
from pathlib import Path
from typing import Iterable, List, Optional, Union
from xml.etree import ElementTree as ET

from ..authoring import LatentStyles as _LatentStyles
from ..authoring import find_body_inline_shapes, find_part_rels
from ..images import detect as _detect_image
from ..document import Document as _Document
from ..ooxml import DocxError
from ..views import (Hyperlink as _Hyperlink, Paragraph,
                     Paragraph as _Paragraph, _attr, _tag, views_of)
from ..write import (Builder, ColorFormat as _ColorFormat, Font as _Font,
                     ParagraphFormat as _ParagraphFormat, Run,
                     Run as _Run, Table, Table as _Table)
from .enum import WD_BREAK, WD_SECTION_START, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from .shared import Cm, Emu, Inches, Length, Mm, Pt, RGBColor, Twips

__all__ = [
    "Document", "DocumentError", "Comment", "Comments", "Inches", "Cm", "Mm",
    "Pt", "Emu", "Twips", "Length", "RGBColor", "Image", "Drawing",
    "RenderedPageBreak",
]


def _style_name(style) -> Optional[str]:
    """A style *name* or Style object -> the string the builders want."""
    if style is None:
        return None
    if isinstance(style, str):
        return style
    name = getattr(style, "name", None)
    return str(name) if name else str(style)


def _align_value(align) -> Optional[str]:
    """Accept ``WD_ALIGN_PARAGRAPH.CENTER``, ``"center"`` or ``None``."""
    from ..write import token
    return token(align)


def _enum_from_token(enum_name: str, value):
    if value is None:
        return None
    from . import enum
    return getattr(enum, enum_name).from_xml(str(value))


class DocumentError(Exception):
    """python-docx's own error type, so ``except docx.DocumentError`` works."""


def _wrap_error(exc: DocxError) -> Exception:
    return DocumentError(str(exc))


def _compat_paragraph(paragraph: Paragraph, document: "Document") -> "_CompatParagraph":
    """Shallow-wrap a core paragraph with python-docx read/write semantics."""
    if isinstance(paragraph, _CompatParagraph):
        return paragraph
    wrapped = object.__new__(_CompatParagraph)
    wrapped.__dict__.update(paragraph.__dict__)
    wrapped._compat_document = document
    # Story-part paragraphs are detached from their owner in the core reader;
    # this owner is required for operations such as run.add_picture().
    wrapped._document = document
    return wrapped


class _CompatColorFormat:
    """Color access with the tuple-shaped ``docx.shared.RGBColor`` value."""

    def __init__(self, color_format):
        self._color_format = color_format

    @property
    def rgb(self):
        value = self._color_format.rgb
        return RGBColor(*value) if value is not None else None

    @rgb.setter
    def rgb(self, value):
        self._color_format.rgb = value

    @property
    def theme_color(self):
        return self._color_format.theme_color

    @theme_color.setter
    def theme_color(self, value):
        self._color_format.theme_color = value

    @property
    def type(self):
        return self._color_format.type


class _CompatFont(_Font):
    """Font proxy whose nested color value matches python-docx's class."""

    @property
    def color(self):
        return _CompatColorFormat(_Font.color.fget(self))

    @color.setter
    def color(self, value):
        _Font.color.fset(self, value)


class _CompatRun(_Run):
    """Run proxy with python-docx's style object and formatting behavior."""

    def __init__(self, run: _Run, document: "Document", paragraph=None):
        super().__init__(run._element, run._w_ns, refresh=run._refresh,
                         document=document, part_name=run._part_name)
        self._compat_document = document
        self._compat_paragraph = paragraph

    @property
    def style(self):
        from .enum import WD_STYLE_TYPE
        style_ref = _Run.style.fget(self)
        if style_ref is None:
            return self._compat_document.styles.default(WD_STYLE_TYPE.CHARACTER)
        style_id = getattr(style_ref, "style_id", str(style_ref))
        return self._compat_document.styles.get_by_id(
            style_id, style_type=WD_STYLE_TYPE.CHARACTER)

    @style.setter
    def style(self, value) -> None:
        if isinstance(value, _ParagraphStyle):
            value = value._style
        _Run.style.fset(self, value)

    @property
    def font(self):
        return _CompatFont(self)

    def add_break(self, break_type=WD_BREAK.LINE):
        return _Run.add_break(self, break_type)

    @property
    def contains_page_break(self):
        return any(_tag(child) == "lastRenderedPageBreak"
                   for child in self._element)

    def iter_inner_content(self):
        """Yield text, drawings, and rendered breaks in run order."""
        text_parts = []
        text_tags = {"br", "cr", "noBreakHyphen", "ptab", "t", "tab"}
        for child in self._element:
            tag = _tag(child)
            if tag in text_tags:
                if tag == "br":
                    text = "\n" if _attr(child, "type") in ("", "textWrapping") else ""
                elif tag == "cr":
                    text = "\n"
                elif tag == "noBreakHyphen":
                    text = "-"
                elif tag == "ptab" or tag == "tab":
                    text = "\t"
                else:
                    text = child.text or ""
                text_parts.append(text)
            elif tag in ("drawing", "lastRenderedPageBreak"):
                if text_parts:
                    yield "".join(text_parts)
                    text_parts.clear()
                if tag == "drawing":
                    yield Drawing(child, self)
                else:
                    yield RenderedPageBreak(child, self)
        if text_parts:
            yield "".join(text_parts)

    def add_picture(self, image_path_or_stream, width=None, height=None):
        return self._compat_document._add_picture_to_run(
            image_path_or_stream, width, height, self)


def _compat_run(run: _Run, document: "Document", paragraph=None) -> _CompatRun:
    if isinstance(run, _CompatRun):
        if paragraph is not None:
            run._compat_paragraph = paragraph
        return run
    return _CompatRun(run, document, paragraph)


class _CompatParagraph(_Paragraph):
    """A paragraph that keeps docviews' reader, with python-docx accessors."""

    def runs_in(self, view: str = "accepted") -> List[Run]:
        return [_compat_run(run, self._compat_document, self)
                for run in _Paragraph.runs_in(self, view)]

    @property
    def hyperlinks(self):
        return [_CompatHyperlink(link, self)
                for link in _Paragraph.hyperlinks.fget(self)]

    def iter_inner_content(self):
        for item in _Paragraph.iter_inner_content(self):
            if isinstance(item, _Run):
                yield _compat_run(item, self._compat_document, self)
            elif isinstance(item, _Hyperlink):
                yield _CompatHyperlink(item, self)
            else:
                yield item

    @property
    def contains_page_break(self):
        return any(run.contains_page_break for run in self.runs)

    @property
    def text(self) -> str:
        if self._element is None:
            return self.__dict__.get("text", "")
        return views_of(self._element).get(self.view) or ""

    @text.setter
    def text(self, value: str) -> None:
        self.set_text(value or "")

    @property
    def alignment(self):
        value = _Paragraph.alignment.fget(self)
        return _enum_from_token("WD_ALIGN_PARAGRAPH", str(value)) \
            if value is not None else None

    @alignment.setter
    def alignment(self, value) -> None:
        _Paragraph.alignment.fset(self, value)

    @property
    def style(self):
        value = _Paragraph.style.fget(self)
        if value is None:
            return None
        from .enum import WD_STYLE_TYPE
        try:
            return self._compat_document.styles.get_by_id(
                value.style_id, WD_STYLE_TYPE.PARAGRAPH)
        except (AttributeError, KeyError):
            return value

    @style.setter
    def style(self, value) -> None:
        _Paragraph.style.fset(self, value)

    @property
    def paragraph_format(self):
        return _CompatParagraphFormat(self._require_element(), self._w_ns)

    def refresh(self):
        if self._element is None:
            return self
        current = views_of(self._element)
        self.views = current
        self.__dict__["text"] = current.get(self.view) or ""
        self.revision_ids = tuple(current.revisions)
        self.revision_chars = dict(current.revision_chars)
        return self

    def add_run(self, text=None, style=None) -> Run:
        """Add a run using python-docx's ``text`` and character ``style`` args."""
        run = _Paragraph.add_run(self, "" if text is None else text)
        if style is not None:
            run.style = style
        return _compat_run(run, self._compat_document, self)

    def insert_after(self, text: str = "", **fmt):
        paragraph = _Paragraph.insert_after(self, text, **fmt)
        return _compat_paragraph(paragraph, self._compat_document)

    def insert_paragraph_before(self, text=None, style=None):
        paragraph = _Paragraph.insert_paragraph_before(
            self, "" if text is None else text, style=style)
        return _compat_paragraph(paragraph, self._compat_document)


class _CompatParagraphFormat(_ParagraphFormat):
    """Paragraph formatting with python-docx's enum-valued alignment."""

    @property
    def alignment(self):
        value = _ParagraphFormat.alignment.fget(self)
        return _enum_from_token("WD_ALIGN_PARAGRAPH", str(value)) \
            if value is not None else None

    @alignment.setter
    def alignment(self, value) -> None:
        _ParagraphFormat.alignment.fset(self, value)

    @property
    def line_spacing(self):
        return _ParagraphFormat.line_spacing.fget(self)

    @line_spacing.setter
    def line_spacing(self, value) -> None:
        if value is None:
            self._set_child("spacing", "line", None)
            self._set_child("spacing", "lineRule", None)
        elif isinstance(value, Length):
            self._set_child("spacing", "line", int(round(value / 635)))
            if self._get_child("spacing", "lineRule") != "atLeast":
                self._set_child("spacing", "lineRule", "exact")
        else:
            # python-docx treats any numeric value (including int) as a
            # multiple of the normal line height, not as a length.
            multiplier = float(value)
            self._set_child("spacing", "line", int(round(multiplier * 240)))
            self._set_child("spacing", "lineRule", "auto")

    @property
    def line_spacing_rule(self):
        from .enum import WD_LINE_SPACING
        rule = self._get_child("spacing", "lineRule")
        if rule is None:
            return None
        value = WD_LINE_SPACING.from_xml(rule)
        line = self._get_child("spacing", "line")
        if value is WD_LINE_SPACING.MULTIPLE and line is not None:
            return {"240": WD_LINE_SPACING.SINGLE,
                    "360": WD_LINE_SPACING.ONE_POINT_FIVE,
                    "480": WD_LINE_SPACING.DOUBLE}.get(line, value)
        return value

    @line_spacing_rule.setter
    def line_spacing_rule(self, value) -> None:
        from .enum import WD_LINE_SPACING
        if value is None:
            self._set_child("spacing", "lineRule", None)
        elif value == WD_LINE_SPACING.SINGLE:
            self.line_spacing = 1.0
        elif value == WD_LINE_SPACING.ONE_POINT_FIVE:
            self.line_spacing = 1.5
        elif value == WD_LINE_SPACING.DOUBLE:
            self.line_spacing = 2.0
        else:
            self._set_child("spacing", "lineRule", value)

    @property
    def tab_stops(self) -> "_CompatTabStops":
        return _CompatTabStops(self)


class Image:
    """Image bytes and dimensions exposed by ``Drawing.image``."""

    def __init__(self, blob: bytes, filename: Optional[str] = None):
        self._blob = bytes(blob)
        self._info = _detect_image(self._blob)
        if filename is None:
            extension = {"jpeg": "jpg"}.get(
                self._info.extension, self._info.extension)
            filename = f"image.{extension}"
        self._filename = filename

    @classmethod
    def from_blob(cls, blob):
        return cls(blob)

    @classmethod
    def from_file(cls, image_descriptor):
        if isinstance(image_descriptor, (str, Path)):
            path = Path(image_descriptor)
            return cls(path.read_bytes(), path.name)
        image_descriptor.seek(0)
        return cls(image_descriptor.read())

    @property
    def blob(self):
        return self._blob

    @property
    def content_type(self):
        return self._info.content_type

    @property
    def ext(self):
        return Path(self._filename).suffix[1:]

    @property
    def filename(self):
        return self._filename

    @property
    def px_width(self):
        return self._info.width_px

    @property
    def px_height(self):
        return self._info.height_px

    @property
    def horz_dpi(self):
        return int(round(self._info.dpi_x))

    @property
    def vert_dpi(self):
        return int(round(self._info.dpi_y))

    @property
    def width(self):
        return Inches(self.px_width / self.horz_dpi)

    @property
    def height(self):
        return Inches(self.px_height / self.vert_dpi)

    def scaled_dimensions(self, width=None, height=None):
        if width is None and height is None:
            return self.width, self.height
        if width is None:
            scaling_factor = float(height) / float(self.height)
            width = round(self.width * scaling_factor)
        if height is None:
            scaling_factor = float(width) / float(self.width)
            height = round(self.height * scaling_factor)
        return Emu(width), Emu(height)

    @property
    def sha1(self):
        return hashlib.sha1(self._blob).hexdigest()


class Drawing:
    """A DrawingML item yielded by :meth:`Run.iter_inner_content`."""

    def __init__(self, element: ET.Element, parent: _CompatRun):
        self._drawing = self._element = element
        self._parent = parent

    def _picture(self):
        for element in self._drawing.iter():
            if _tag(element) == "graphicData":
                for child in element:
                    if _tag(child) == "pic":
                        return child
        return None

    @property
    def has_picture(self):
        return self._picture() is not None

    @property
    def image(self):
        picture = self._picture()
        if picture is None:
            raise ValueError("drawing does not contain a picture")
        rel_id = next((_attr(element, "embed") for element in picture.iter()
                       if _tag(element) == "blip" and _attr(element, "embed")), "")
        if not rel_id:
            raise ValueError("drawing does not contain a picture")

        document = self._parent._compat_document
        part_name = self._parent._part_name or document._package.main_part
        target = ""
        if part_name == document._package.main_part:
            target = document._rel_target(rel_id)
        else:
            for source, reserved_id, _kind, destination in \
                    document._extra_story_rels:
                if source == part_name and reserved_id == rel_id:
                    target = document._part_for_rel(destination, part_name)
                    break
        if not target:
            target = find_part_rels(document._package, part_name).get(rel_id, "")
        if target in document._added_parts:
            blob = document._added_parts[target]
        else:
            blob = document._package.read(target)
        return Image.from_blob(blob)

    def __repr__(self):
        return "<Drawing>"


def _element_parents(root):
    return {child: parent for parent in root.iter() for child in parent}


class RenderedPageBreak:
    """A page break inserted by Word's renderer, not an author's hard break."""

    def __init__(self, element: ET.Element, parent: _CompatRun):
        self._element = element
        self._parent = parent

    def _context(self):
        paragraph = self._parent._compat_paragraph
        parents = _element_parents(paragraph._element)
        run = parents.get(self._element)
        hyperlink = parents.get(run) if run is not None else None
        is_hyperlink = hyperlink is not None and _tag(hyperlink) == "hyperlink"
        return paragraph, run, hyperlink if is_hyperlink else None, parents

    def _is_first_in_paragraph(self, paragraph):
        for child in paragraph._element:
            if _tag(child) == "r":
                for item in child:
                    if _tag(item) == "lastRenderedPageBreak":
                        return item is self._element
            elif _tag(child) == "hyperlink":
                for run in child:
                    if _tag(run) == "r":
                        for item in run:
                            if _tag(item) == "lastRenderedPageBreak":
                                return item is self._element
        return False

    def _at_paragraph_boundary(self, paragraph, run, hyperlink, side):
        if hyperlink is not None or run is None:
            return False
        children = list(paragraph._element)
        runs = [child for child in children if _tag(child) == "r"]
        if not runs:
            return False
        run_children = list(run)
        index = run_children.index(self._element)
        if side == "preceding":
            return run is runs[0] and not any(
                _tag(child) in {"br", "cr", "drawing", "noBreakHyphen",
                                "ptab", "t", "tab"}
                for child in run_children[:index])
        return run is runs[-1] and not any(
            _tag(child) in {"br", "cr", "drawing", "noBreakHyphen",
                            "ptab", "t", "tab"}
            for child in run_children[index + 1:])

    def _fragment(self, side):
        paragraph, run, hyperlink, parents = self._context()
        if not self._is_first_in_paragraph(paragraph):
            raise ValueError("only defined on first rendered page-break in paragraph")
        if self._at_paragraph_boundary(paragraph, run, hyperlink, side):
            return None

        # Work on a clone so fragment inspection never mutates the source file.
        original = paragraph._element
        path = []
        node = self._element
        while node is not original:
            parent = parents.get(node)
            if parent is None:
                raise ValueError("rendered page-break is not attached to a paragraph")
            path.append(list(parent).index(node))
            node = parent
        fragment = copy.deepcopy(original)
        node = fragment
        for index in reversed(path):
            node = list(node)[index]
        clone_parents = _element_parents(fragment)
        clone_run = clone_parents[node]
        clone_hyperlink = clone_parents.get(clone_run)
        if _tag(clone_hyperlink) != "hyperlink":
            clone_hyperlink = None

        if clone_hyperlink is not None:
            sibling = clone_hyperlink
            if side == "preceding":
                after = False
                for child in list(fragment):
                    if child is sibling:
                        after = True
                    elif after:
                        fragment.remove(child)
                clone_run.remove(node)
            else:
                before = True
                for child in list(fragment):
                    if child is sibling:
                        before = False
                        continue
                    if before and _tag(child) != "pPr":
                        fragment.remove(child)
                if sibling in list(fragment):
                    fragment.remove(sibling)
        else:
            if side == "preceding":
                after = False
                for child in list(fragment):
                    if child is clone_run:
                        after = True
                    elif after:
                        fragment.remove(child)
                run_children = list(clone_run)
                index = run_children.index(node)
                for child in run_children[index + 1:]:
                    clone_run.remove(child)
            else:
                before = True
                for child in list(fragment):
                    if child is clone_run:
                        before = False
                        continue
                    if before and _tag(child) != "pPr":
                        fragment.remove(child)
                run_children = list(clone_run)
                index = run_children.index(node)
                for child in run_children[:index]:
                    if _tag(child) != "rPr":
                        clone_run.remove(child)
            clone_run.remove(node)

        document = paragraph._compat_document
        part_name = paragraph.part_name
        rels = find_part_rels(document._package, part_name) if part_name else {}
        core = _Paragraph(
            text=views_of(fragment).get("accepted") or "", view="accepted",
            kind=paragraph.kind, part_name=part_name, _element=fragment,
            _w_ns=paragraph._w_ns, _rels=rels, _document=document,
        )
        return _compat_paragraph(core, document)

    @property
    def preceding_paragraph_fragment(self):
        return self._fragment("preceding")

    @property
    def following_paragraph_fragment(self):
        return self._fragment("following")

    def __repr__(self):
        return "<RenderedPageBreak>"


class _CompatHyperlink:
    """Python-docx's reading interface for one hyperlink in a paragraph."""

    def __init__(self, hyperlink: _Hyperlink, paragraph: _CompatParagraph):
        self._hyperlink = hyperlink
        self._paragraph = paragraph
        self._element = hyperlink._element

    @property
    def address(self) -> str:
        # An internal bookmark is represented by w:anchor, not a relationship
        # address. The core view keeps that target in ``address`` for convenience.
        return self._hyperlink.address if _attr(self._element, "id") else ""

    @property
    def contains_page_break(self) -> bool:
        return any(_tag(element) == "lastRenderedPageBreak"
                   for element in self._element.iter())

    @property
    def fragment(self) -> str:
        return _attr(self._element, "anchor") or ""

    @property
    def runs(self):
        return [_compat_run(_Run(element, self._paragraph._w_ns,
                                 document=self._paragraph._document,
                                 part_name=self._paragraph.part_name),
                            self._paragraph._compat_document,
                            self._paragraph)
                for element in self._element
                if _tag(element) == "r"]

    @property
    def text(self) -> str:
        return self._hyperlink.text

    @property
    def url(self) -> str:
        address, fragment = self.address, self.fragment
        if not address:
            return ""
        return f"{address}#{fragment}" if fragment else address

    @property
    def paragraph(self):
        return self._paragraph

    def __getattr__(self, name):
        return getattr(self._hyperlink, name)

    def __repr__(self) -> str:
        return f"Hyperlink({self.text!r} -> {self.url!r})"


class _CompatTabStops:
    """Editable tab-stop sequence backed by a paragraph's ``w:tabs`` element."""

    def __init__(self, paragraph_format: _CompatParagraphFormat):
        self._format = paragraph_format

    def _tabs(self):
        return self._format._holder("tabs")

    def __len__(self) -> int:
        tabs = self._tabs()
        return 0 if tabs is None else sum(1 for child in tabs if _tag(child) == "tab")

    def __iter__(self):
        tabs = self._tabs()
        if tabs is not None:
            for element in tabs:
                if _tag(element) == "tab":
                    yield _CompatTabStop(element, self._format)

    def __getitem__(self, index):
        tabs = self._tabs()
        if tabs is None:
            raise IndexError("TabStops object is empty")
        stops = [child for child in tabs if _tag(child) == "tab"]
        try:
            return _CompatTabStop(stops[index], self._format)
        except IndexError:
            raise IndexError("list index out of range") from None

    def __delitem__(self, index) -> None:
        tabs = self._tabs()
        if tabs is None:
            raise IndexError("tab index out of range")
        elements = [child for child in tabs if _tag(child) == "tab"]
        try:
            tabs.remove(elements[index])
        except IndexError:
            raise IndexError("tab index out of range") from None
        if not any(_tag(child) == "tab" for child in tabs):
            ppr = self._format._ppr()
            if ppr is not None:
                ppr.remove(tabs)

    def add_tab_stop(self, position,
                     alignment=WD_TAB_ALIGNMENT.LEFT,
                     leader=WD_TAB_LEADER.SPACES):
        from ..write import token

        alignment = WD_TAB_ALIGNMENT(alignment)
        leader = WD_TAB_LEADER.SPACES if leader is None else WD_TAB_LEADER(leader)
        tabs = self._format._holder("tabs", create=True)
        element = ET.Element(f"{{{self._format._w_ns}}}tab")
        element.set(f"{{{self._format._w_ns}}}val", token(alignment))
        element.set(f"{{{self._format._w_ns}}}leader", token(leader))
        pos = int(round(int(position) / 635))
        element.set(f"{{{self._format._w_ns}}}pos", str(pos))
        for index, current in enumerate(tabs):
            if (_tag(current) == "tab" and
                    pos < int(current.get(f"{{{self._format._w_ns}}}pos", "0"))):
                tabs.insert(index, element)
                break
        else:
            tabs.append(element)
        return _CompatTabStop(element, self._format)

    def clear_all(self) -> None:
        ppr = self._format._ppr()
        if ppr is not None:
            tabs = next((child for child in ppr if _tag(child) == "tabs"), None)
            if tabs is not None:
                ppr.remove(tabs)


class _CompatTabStop:
    """One writable Word tab stop."""

    def __init__(self, element: ET.Element, paragraph_format: _CompatParagraphFormat):
        self._element = element
        self._format = paragraph_format

    @property
    def alignment(self):
        return _enum_from_token("WD_TAB_ALIGNMENT",
                                self._element.get(f"{{{self._format._w_ns}}}val")
                                or "left")

    @alignment.setter
    def alignment(self, value) -> None:
        from ..write import token
        self._element.set(f"{{{self._format._w_ns}}}val",
                          token(WD_TAB_ALIGNMENT(value)))

    @property
    def leader(self):
        return _enum_from_token("WD_TAB_LEADER",
                                self._element.get(f"{{{self._format._w_ns}}}leader")
                                or "spaces")

    @leader.setter
    def leader(self, value) -> None:
        from ..write import token
        value = WD_TAB_LEADER.SPACES if value is None else WD_TAB_LEADER(value)
        self._element.set(f"{{{self._format._w_ns}}}leader", token(value))

    @property
    def position(self):
        value = self._element.get(f"{{{self._format._w_ns}}}pos")
        return Length(int(value) * 635) if value is not None else None

    @position.setter
    def position(self, value) -> None:
        tabs = self._format._holder("tabs")
        old = self._element
        tabs.remove(old)
        position = int(round(int(value) / 635))
        old.set(f"{{{self._format._w_ns}}}pos", str(position))
        for index, current in enumerate(tabs):
            if (_tag(current) == "tab" and
                    position < int(current.get(f"{{{self._format._w_ns}}}pos", "0"))):
                tabs.insert(index, old)
                break
        else:
            tabs.append(old)


class Document(_Document):
    """A python-docx-shaped document: create, open, edit, save."""

    def __init__(self, docx: Union[str, Path, bytes, None] = None):
        if docx is None:
            super().__init__(type(self)._new_bytes())
        else:
            try:
                super().__init__(docx)
            except DocxError as exc:
                raise _wrap_error(exc) from exc

    @staticmethod
    def _new_bytes() -> bytes:
        from ..write import new_package_bytes
        return new_package_bytes()

    # ------------------------------------------------------------ python-docx API
    def add_paragraph(self, text: str = "", style=None) -> Paragraph:
        """Add a paragraph with python-docx's text and style parameters."""
        paragraph = super().add_paragraph(text, style=_style_name(style))
        return _compat_paragraph(paragraph, self)

    def add_heading(self, text: str = "", level: int = 1) -> Paragraph:
        return _compat_paragraph(super().add_heading(text, level), self)

    def add_page_break(self) -> Paragraph:
        return _compat_paragraph(super().add_page_break(), self)

    def add_section(self, start_type=WD_SECTION_START.NEW_PAGE):
        """Start a section using python-docx's default ``NEW_PAGE`` break."""
        section = super().add_section()
        section.start_type = start_type
        return _Section(self, section)

    def add_table(self, rows: int, cols: int, style=None):
        table = super().add_table(rows, cols, style=_style_name(style))
        return _CompatTable(self, table)

    @property
    def tables(self):
        return [_CompatTable(self, table)
                for table in _Document.tables.fget(self)]

    @property
    def paragraphs(self):                     # noqa: F811  (python-docx shape)
        """Direct body paragraphs, matching python-docx's collection boundary.

        Paragraphs in tables, content controls, drawings, headers, footers,
        footnotes, endnotes and comments remain available through
        ``iter_paragraphs()``. The compatibility property intentionally mirrors
        python-docx's ``CT_Body.paragraphs`` selector (direct ``w:p`` children).
        """
        direct = {id(child) for child in self._body() if _tag(child) == "p"}
        return [_compat_paragraph(para, self)
                for para in super().paragraphs()
                if para.part_name == self._package.main_part
                and id(para._element) in direct]

    def iter_inner_content(self):
        """Yield body paragraphs and tables in document order, like python-docx."""
        paragraphs = {id(para._element): para for para in self.paragraphs}
        w_ns = self._package.ns["w"]
        for child in self._body():
            if _tag(child) == "p":
                para = paragraphs.get(id(child))
                if para is not None:
                    yield para
            elif _tag(child) == "tbl":
                yield _CompatTable(self, Table(child, w_ns))

    @property
    def _sections(self):
        """The library's own sections — what the compat wrappers delegate to."""
        return super().sections

    @property
    def sections(self):
        return _Sections(self)

    @property
    def core_properties(self):
        return _CoreProperties(self)

    @property
    def inline_shapes(self):
        return _InlineShapes(self)

    @property
    def styles(self):
        return _Styles(self)

    def _settings_accessor(self):
        return _Document.settings.fget(self)

    @property
    def settings(self) -> "_Settings":
        """python-docx's ``doc.settings``."""
        return _Settings(self)

    def _add_unanchored_comment(self, text: str = "", author: str = "",
                                initials: Optional[str] = "") -> "Comment":
        """Create the collection-level comment supported by python-docx 1.2."""
        from ..document import _utc_now

        part_name, root = self._comments_part(create=True)
        comment_id = self._next_comment_id(root)
        w_ns = self._package.ns["w"]
        comment_element = ET.SubElement(root, f"{{{w_ns}}}comment")
        for name, value in (("id", str(comment_id)), ("author", author),
                            ("date", _utc_now())):
            comment_element.set(f"{{{w_ns}}}{name}", value)
        if initials is not None:
            comment_element.set(f"{{{w_ns}}}initials", initials)

        builder = Builder(self._package.ns)
        lines = (text or "").split("\n")
        paragraphs = []
        for line in lines:
            paragraph = builder.paragraph(line, style="CommentText")
            comment_element.append(paragraph)
            paragraphs.append(paragraph)

        # The annotation reference belongs to the comment body even before a
        # body run is marked as the comment's anchor, matching python-docx.
        reference = builder.el("w:r")
        properties = builder.el("w:rPr", reference)
        builder.el("w:rStyle", properties, w_val="CommentReference")
        builder.el("w:annotationRef", reference)
        paragraphs[0].append(reference)
        self._mark_changed(part_name)
        return Comment(self, str(comment_id))

    @property
    def comments(self) -> "Comments":
        """python-docx's collection API, with docviews' additional anchors."""
        return Comments(self)

    def add_comment(self, runs, text: str = "", author: str = "",
                    initials: str = "") -> "Comment":
        """Add a comment anchored to runs and return its python-docx-shaped proxy."""
        comment = _Document.add_comment(self, runs, text=text, author=author,
                                        initials=initials)
        return Comment(self, comment)

    def add_picture(self, image_path_or_stream, width=None, height=None):
        """Add an inline picture to the document body."""
        return _InlineShape(super().add_picture(
            image_path_or_stream, width, height), self)

    def _add_picture_to_run(self, image_path_or_stream, width, height, run):
        return _InlineShape(super().add_picture(
            image_path_or_stream, width, height, run=run), self)

    def save(self, path_or_stream) -> None:
        """Save to a path or writable stream, returning python-docx's ``None``."""
        if path_or_stream is not None and hasattr(path_or_stream, "write"):
            data = self.to_bytes()
            if hasattr(path_or_stream, "seek"):
                path_or_stream.seek(0)
            if hasattr(path_or_stream, "truncate"):
                path_or_stream.truncate(0)
            path_or_stream.write(data)
            return None
        super().save(path_or_stream)
        return None


class Comments:
    """Collection of comments with the python-docx 1.2.0 API."""

    def __init__(self, document: Document):
        self._document = document

    def __iter__(self):
        for comment in _Document.comments(self._document):
            yield Comment(self._document, comment)

    def __len__(self) -> int:
        return len(_Document.comments(self._document))

    def __getitem__(self, index):
        return list(self)[index]

    def add_comment(self, text: str = "", author: str = "",
                    initials: Optional[str] = "") -> "Comment":
        """Add an unanchored comment, which can later be marked on a run."""
        return self._document._add_unanchored_comment(text, author, initials)

    def get(self, comment_id: int) -> Optional["Comment"]:
        wanted = str(comment_id)
        for comment in self:
            if str(comment.comment_id) == wanted:
                return comment
        return None


class Comment:
    """A python-docx-shaped comment with docviews' anchor information."""

    def __init__(self, document: Document, comment):
        object.__setattr__(self, "_document", document)
        object.__setattr__(self, "_id", str(getattr(comment, "comment_id", comment)))
        object.__setattr__(self, "_cached", comment if hasattr(comment, "comment_id")
                           else None)

    def _record(self):
        for record in _Document.comments(self._document):
            if str(record.comment_id) == self._id:
                object.__setattr__(self, "_cached", record)
                return record
        return self._cached

    def _elements(self):
        part_name, root = self._document._comments_part(create=False)
        if root is None:
            raise KeyError(f"comment {self._id} is no longer in the document")
        id_name = self._document._w("id")
        for element in root:
            if _tag(element) == "comment" and element.get(id_name) == self._id:
                return part_name, root, element
        raise KeyError(f"comment {self._id} is no longer in the document")

    @property
    def author(self) -> str:
        _part, _root, element = self._elements()
        return element.get(self._document._w("author"), "")

    @author.setter
    def author(self, value: str) -> None:
        part, _root, element = self._elements()
        element.set(self._document._w("author"), value)
        self._document._mark_changed(part)

    @property
    def comment_id(self) -> int:
        return int(self._id)

    @property
    def initials(self) -> Optional[str]:
        _part, _root, element = self._elements()
        return element.get(self._document._w("initials"))

    @initials.setter
    def initials(self, value: Optional[str]) -> None:
        part, _root, element = self._elements()
        key = self._document._w("initials")
        if value is None:
            element.attrib.pop(key, None)
        else:
            element.set(key, value)
        self._document._mark_changed(part)

    @property
    def timestamp(self):
        _part, _root, element = self._elements()
        value = element.get(self._document._w("date"))
        if not value:
            return None
        from datetime import datetime, timezone
        try:
            result = datetime.fromisoformat(value.replace("Z", "+00:00"))
        except ValueError:
            return None
        return result if result.tzinfo is not None else result.replace(tzinfo=timezone.utc)

    @property
    def text(self) -> str:
        return "\n".join(paragraph.text for paragraph in self.paragraphs)

    @property
    def paragraphs(self) -> List[Paragraph]:
        part, root, comment = self._elements()
        w_ns = self._document._package.ns["w"]
        out = []
        for index, element in enumerate(child for child in comment
                                       if _tag(child) == "p"):
            paragraph = _Paragraph(
                text=views_of(element).get("accepted") or "", view="accepted",
                kind="comment", part_name=part, index=index, _element=element,
                _root=root, _w_ns=w_ns, _document=self._document)
            out.append(_compat_paragraph(paragraph, self._document))
        return out

    @property
    def tables(self):
        part, _root, comment = self._elements()
        del part
        w_ns = self._document._package.ns["w"]
        return [_CompatTable(self._document, _Table(element, w_ns))
                for element in comment if _tag(element) == "tbl"]

    def add_paragraph(self, text: str = "", style=None) -> Paragraph:
        part, _root, comment = self._elements()
        w_ns = self._document._package.ns["w"]
        builder = Builder(self._document._package.ns)
        paragraph_element = builder.paragraph(
            "" if text is None else text,
            style=_style_name(style) if style is not None else "CommentText")
        comment.append(paragraph_element)
        self._document._mark_changed(part)
        paragraph = _Paragraph(
            text=views_of(paragraph_element).get("accepted") or "",
            view="accepted", kind="comment", part_name=part,
            index=len(self.paragraphs) - 1, _element=paragraph_element,
            _root=_root, _w_ns=w_ns, _document=self._document)
        return _compat_paragraph(paragraph, self._document)

    def add_table(self, rows: int, cols: int, width) -> "_CompatTable":
        part, _root, comment = self._elements()
        w_ns = self._document._package.ns["w"]
        per_column = int(width) / cols
        builder = Builder(self._document._package.ns)
        element = builder.table(rows, cols, widths=[per_column] * cols)
        comment.append(element)
        self._document._mark_changed(part)
        return _CompatTable(self._document, _Table(element, w_ns))

    def iter_inner_content(self):
        part, root, comment = self._elements()
        w_ns = self._document._package.ns["w"]
        for element in comment:
            if _tag(element) == "p":
                paragraph = _Paragraph(
                    text=views_of(element).get("accepted") or "",
                    view="accepted", kind="comment", part_name=part,
                    index=len([p for p in comment[:list(comment).index(element)]
                               if _tag(p) == "p"]),
                    _element=element, _root=root, _w_ns=w_ns,
                    _document=self._document)
                yield _compat_paragraph(paragraph, self._document)
            elif _tag(element) == "tbl":
                yield _CompatTable(self._document, _Table(element, w_ns))

    @property
    def anchors(self):
        record = self._record()
        return record.anchors if record is not None else []

    @property
    def anchor_text(self) -> str:
        record = self._record()
        return record.anchor_text if record is not None else ""

    def __getattr__(self, name):
        record = self._record()
        if record is not None and hasattr(record, name):
            return getattr(record, name)
        raise AttributeError(name)

    def __repr__(self) -> str:
        return f"<Comment {self.comment_id} by {self.author!r}>"


class _Settings:
    """``doc.settings`` — python-docx's window on ``word/settings.xml``."""

    def __init__(self, document: Document):
        self._document = document

    odd_and_even_pages_header_footer = property(
        lambda self: self._document._settings_accessor().odd_and_even_pages_header_footer,
        lambda self, v: setattr(self._document._settings_accessor(),
                                "odd_and_even_pages_header_footer", v))

    def __repr__(self) -> str:
        return repr(self._document._settings_accessor())

class _Sections(Sequence):
    """``doc.sections`` — every section, with python-docx's names.

    Delegates to the library's own sections, so a multi-section document is not
    flattened to its first section the way the earlier shim had to.
    """

    def __init__(self, document: Document):
        self._document = document

    def __len__(self) -> int:
        return len(self._document._sections)

    def __iter__(self):
        for index in range(len(self)):
            yield self[index]

    def __getitem__(self, index):
        if isinstance(index, slice):
            return [self[position]
                    for position in range(*index.indices(len(self)))]
        try:
            section = self._document._sections[index]
        except IndexError as exc:
            raise IndexError(str(exc)) from None
        return _Section(self._document, section)

    def __repr__(self) -> str:
        return f"<Sections({len(self)})>"


class _Section:
    """One section: geometry in EMU, plus ``header`` and ``footer``."""

    def __init__(self, document: Document, section):
        self._document = document
        self._section = section

    # -- geometry (python-docx returns Lengths; these are the same ints)
    page_width = property(lambda self: self._section.page_width,
                          lambda self, v: setattr(self._section, "page_width", v))
    page_height = property(lambda self: self._section.page_height,
                           lambda self, v: setattr(self._section, "page_height", v))
    left_margin = property(lambda self: self._section.left_margin,
                           lambda self, v: setattr(self._section, "left_margin", v))
    right_margin = property(lambda self: self._section.right_margin,
                            lambda self, v: setattr(self._section, "right_margin", v))
    top_margin = property(lambda self: self._section.top_margin,
                          lambda self, v: setattr(self._section, "top_margin", v))
    bottom_margin = property(lambda self: self._section.bottom_margin,
                             lambda self, v: setattr(self._section, "bottom_margin", v))
    header_distance = property(lambda self: self._section.header_distance,
                               lambda self, v: setattr(self._section, "header_distance", v))
    footer_distance = property(lambda self: self._section.footer_distance,
                               lambda self, v: setattr(self._section, "footer_distance", v))

    different_first_page_header_footer = property(
        lambda self: self._section.different_first_page_header_footer,
        lambda self, v: setattr(self._section,
                                "different_first_page_header_footer", v))

    @property
    def first_page_header(self) -> "_HeaderFooter":
        return _HeaderFooter(self._section.first_page_header, True)

    @property
    def first_page_footer(self) -> "_HeaderFooter":
        return _HeaderFooter(self._section.first_page_footer, False)

    @property
    def even_page_header(self) -> "_HeaderFooter":
        return _HeaderFooter(self._section.even_page_header, True)

    @property
    def even_page_footer(self) -> "_HeaderFooter":
        return _HeaderFooter(self._section.even_page_footer, False)

    orientation = property(lambda self: self._section.orientation,
                           lambda self, v: setattr(self._section, "orientation", v))
    start_type = property(lambda self: self._section.start_type,
                          lambda self, v: setattr(self._section, "start_type", v))
    gutter = property(lambda self: self._section.gutter,
                      lambda self, v: setattr(self._section, "gutter", v))

    def iter_inner_content(self):
        yield self.header
        yield self.footer

    @property
    def header(self) -> "_HeaderFooter":
        return _HeaderFooter(self._section.header, True)

    @property
    def footer(self) -> "_HeaderFooter":
        return _HeaderFooter(self._section.footer, False)

    def __repr__(self) -> str:
        return f"<Section({len(self._section._document._sections)} total)>"


class _HeaderFooter:
    """``section.header`` / ``section.footer``, python-docx-shaped."""

    def __init__(self, part, is_header: bool):
        self._part = part
        self._document = part._document
        self._is_header = is_header

    @property
    def is_linked_to_previous(self) -> bool:
        return self._part.is_linked_to_previous

    def _started(self) -> None:
        """A header/footer this door creates starts the way python-docx's does.

        python-docx's new header part is not empty: it holds one paragraph
        carrying the built-in *Header* style (Word creates it the same way). So
        ``header.add_paragraph("Page header")`` leaves **two** paragraphs in
        the file, and a program that counts them sees the same count after the
        switch. The core API still creates parts with no invented content.
        """
        if self._part.is_linked_to_previous:
            self._part.add_paragraph(
                "", style="Header" if self._is_header else "Footer")

    @property
    def paragraphs(self) -> List[Paragraph]:
        """python-docx creates the part the moment a header is looked at, so
        ``header.paragraphs[0].text = ...`` works on a document with none.
        Doing the same here keeps that idiom working; the core API keeps the
        honest answer (``is_linked_to_previous``, no invented parts)."""
        self._started()
        return [_compat_paragraph(p, self._document)
                for p in self._part.paragraphs]

    @property
    def text(self) -> str:
        """Not in python-docx — the whole header as one string."""
        return self._part.text

    @property
    def tables(self):
        return [_CompatTable(self._document, table) for table in self._part.tables]

    def add_table(self, rows: int, cols: int, width):
        """python-docx's ``header.add_table(rows, cols, width)``."""
        self._started()
        return _CompatTable(self._document,
                            self._part.add_table(rows, cols, width=width))

    def iter_inner_content(self):
        for item in self._part.iter_inner_content():
            if isinstance(item, _Paragraph):
                yield _compat_paragraph(item, self._document)
            elif isinstance(item, _Table):
                yield _CompatTable(self._document, item)

    def add_paragraph(self, text: str = "", style=None):
        self._started()
        paragraph = self._part.add_paragraph(text, style=_style_name(style))
        return _compat_paragraph(paragraph, self._document)

    def __repr__(self) -> str:
        kind = "header" if self._is_header else "footer"
        return f"<{kind}: {self._part.text[:30]!r}>"


class _CoreProperties:
    """``doc.core_properties`` — the docProps/core.xml fields."""

    # every field python-docx's CoreProperties has, by python-docx's name
    FIELDS = ("title", "author", "subject", "comments", "category", "keywords",
              "last_modified_by", "revision", "content_status", "identifier",
              "language", "version", "created", "modified", "last_printed")

    def __init__(self, document: Document):
        self._document = document

    def _properties(self):
        # the library's own object, not this wrapper: reading a package without
        # core properties creates nothing, writing creates the part and states
        # its content type and its package relationship
        return _Document.core_properties.fget(self._document)

    def __getattr__(self, name: str):
        if name in _CoreProperties.FIELDS:
            return getattr(self._properties(), name)
        raise AttributeError(name)

    def __setattr__(self, name: str, value) -> None:
        if name in _CoreProperties.FIELDS:
            setattr(self._properties(), name, value)
            return
        object.__setattr__(self, name, value)


class _InlineShape:
    """One inline drawing, with python-docx's type and editable dimensions."""

    def __init__(self, shape, document: Optional[Document] = None):
        self._shape = shape
        self._document = document

    def _length(self, value):
        from ..units import Length
        return Length(value) if value is not None else None

    def _set_extent(self, dimension: str, value) -> None:
        element = self._shape._element
        if element is None:
            raise AttributeError("this inline shape is not attached to an editable document")
        amount = int(value)
        key = "cx" if dimension == "width" else "cy"
        outer = False
        for child in element.iter():
            local = _tag(child)
            if local == "extent":
                child.set(key, str(amount))
                outer = True
            elif local == "ext" and self.type in (
                    self._enum().PICTURE, self._enum().LINKED_PICTURE):
                child.set(key, str(amount))
        if not outer:
            raise DocumentError("inline drawing has no wp:extent")
        setattr(self._shape, dimension, amount)
        if self._document is not None:
            self._document._mark_changed(self._shape.part_name)

    @staticmethod
    def _enum():
        from .enum import WD_INLINE_SHAPE_TYPE
        return WD_INLINE_SHAPE_TYPE

    @property
    def width(self):
        return self._length(self._shape.width)

    @width.setter
    def width(self, value):
        self._set_extent("width", value)

    @property
    def height(self):
        return self._length(self._shape.height)

    @height.setter
    def height(self, value):
        self._set_extent("height", value)

    name = property(lambda self: self._shape.name)

    @property
    def type(self):
        enum = self._enum()
        uri = self._shape.graphic_uri.rstrip("/").lower()
        if uri.endswith("/picture"):
            if self._shape.media.startswith(("http://", "https://")):
                return enum.LINKED_PICTURE
            return enum.PICTURE
        if uri.endswith("/chart"):
            return enum.CHART
        if uri.endswith("/diagram"):
            return enum.SMART_ART
        return enum.NOT_IMPLEMENTED

    def __repr__(self) -> str:
        return repr(self._shape)


class _InlineShapes:
    """python-docx's body-only inline drawing collection."""

    def __init__(self, document: Document):
        self._document = document

    def _shapes(self):
        shapes = find_body_inline_shapes(
            self._document._package, self._document._package.main_part,
            self._document._body())
        reserved = {rel_id: self._document._part_for_rel(target)
                    for rel_id, _kind, target in self._document._extra_rels}
        for shape in shapes:
            if not shape.media and shape.embed in reserved:
                shape.media = reserved[shape.embed]
        return [_InlineShape(shape, self._document) for shape in shapes]

    def __len__(self) -> int:
        return len(self._shapes())

    def __iter__(self):
        return iter(self._shapes())

    def __getitem__(self, index):
        try:
            return self._shapes()[index]
        except IndexError:
            raise IndexError(f"no inline shape {index}") from None


class _ParagraphStyle:
    """One style, python-docx-shaped: ``.name``, ``.style_id``, ``.type``."""

    def __init__(self, document: Document, style):
        self._document = document
        self._style = style

    name = property(lambda self: self._style.name,
                    lambda self, v: setattr(self._style, "name", v))
    style_id = property(lambda self: self._style.style_id)

    @property
    def type(self):
        from .enum import WD_STYLE_TYPE
        kind = (self._style.type or "paragraph").upper()
        return getattr(WD_STYLE_TYPE, kind, WD_STYLE_TYPE.PARAGRAPH)

    @property
    def paragraph_format(self):
        return self._style.paragraph_format

    @property
    def builtin(self) -> bool:
        return self._style.builtin

    @property
    def base_style(self):
        found = self._style.base_style
        return _ParagraphStyle(self._document, found) if found else None

    @property
    def next_paragraph_style(self):
        found = self._style.next_paragraph_style
        return _ParagraphStyle(self._document, found) if found else None

    quick_style = property(lambda self: self._style.quick_style,
                           lambda self, v: setattr(self._style, "quick_style", v))
    priority = property(lambda self: self._style.priority,
                        lambda self, v: setattr(self._style, "priority", v))
    hidden = property(lambda self: self._style.hidden,
                      lambda self, v: setattr(self._style, "hidden", v))
    locked = property(lambda self: self._style.locked,
                      lambda self, v: setattr(self._style, "locked", v))
    unhide_when_used = property(
        lambda self: self._style.unhide_when_used,
        lambda self, v: setattr(self._style, "unhide_when_used", v))

    def delete(self) -> None:
        self._style.delete()

    @property
    def font(self):
        font = self._style.font
        return _CompatFont(font._run) if font is not None else None

    def __repr__(self) -> str:
        return repr(self._style)


_LATENT_NAME_ALIASES = {
    "Caption": "caption",
    "Footer": "footer",
    "Header": "header",
    **{f"Heading {level}": f"heading {level}" for level in range(1, 10)},
}
_LATENT_NAME_REVERSE = {value: key for key, value in _LATENT_NAME_ALIASES.items()}


def _latent_internal_name(name: str) -> str:
    return _LATENT_NAME_ALIASES.get(name, name)


def _latent_ui_name(name: str) -> str:
    return _LATENT_NAME_REVERSE.get(name, name)


class _CompatLatentStyles:
    """The documented python-docx interface over latent-style XML."""

    def __init__(self, document: Document):
        self._document = document
        self._latent = _LatentStyles(document)

    def _block(self, create: bool = False):
        return self._latent._block(create=create)

    def _get_int(self, name):
        block = self._block()
        value = _attr(block, name) if block is not None else ""
        return int(value) if value else None

    def _set_value(self, name, value):
        block = self._block(create=value is not None)
        if block is None:
            return
        attribute = self._document._w(name)
        if value is None:
            block.attrib.pop(attribute, None)
        else:
            block.set(attribute, str(int(value)))
        self._document._touch_styles()

    def _get_bool(self, name):
        block = self._block()
        value = _attr(block, name) if block is not None else ""
        return bool(value) and value not in ("0", "false", "off")

    def _set_bool(self, name, value):
        block = self._block(create=True)
        attribute = self._document._w(name)
        block.set(attribute, "1" if bool(value) else "0")
        self._document._touch_styles()

    @property
    def default_priority(self):
        return self._get_int("defUIPriority")

    @default_priority.setter
    def default_priority(self, value):
        self._set_value("defUIPriority", value)

    @property
    def default_to_hidden(self):
        return self._get_bool("defSemiHidden")

    @default_to_hidden.setter
    def default_to_hidden(self, value):
        self._set_bool("defSemiHidden", value)

    @property
    def default_to_locked(self):
        return self._get_bool("defLockedState")

    @default_to_locked.setter
    def default_to_locked(self, value):
        self._set_bool("defLockedState", value)

    @property
    def default_to_quick_style(self):
        return self._get_bool("defQFormat")

    @default_to_quick_style.setter
    def default_to_quick_style(self, value):
        self._set_bool("defQFormat", value)

    @property
    def default_to_unhide_when_used(self):
        return self._get_bool("defUnhideWhenUsed")

    @default_to_unhide_when_used.setter
    def default_to_unhide_when_used(self, value):
        self._set_bool("defUnhideWhenUsed", value)

    @property
    def load_count(self):
        return self._get_int("count")

    @load_count.setter
    def load_count(self, value):
        self._set_value("count", value)

    def add_latent_style(self, name):
        block = self._block(create=True)
        element = ET.Element(self._document._w("lsdException"))
        element.set(self._document._w("name"), _latent_internal_name(name))
        block.append(element)
        self._document._touch_styles()
        return _CompatLatentStyle(self, element)

    def __iter__(self):
        block = self._block()
        if block is not None:
            for element in block:
                if _tag(element) == "lsdException":
                    yield _CompatLatentStyle(self, element)

    def __len__(self):
        return sum(1 for _ in self)

    def __getitem__(self, key):
        wanted = _latent_internal_name(key)
        block = self._block()
        if block is not None:
            for element in block:
                if (_tag(element) == "lsdException" and
                        _attr(element, "name") == wanted):
                    return _CompatLatentStyle(self, element)
        raise KeyError(f"no latent style with name '{key}'")

    def __getattr__(self, name):
        # Keep docviews' additional latent-font inspection available without
        # claiming those names are part of python-docx's documented API.
        return getattr(self._latent, name)

    def __repr__(self):
        return f"LatentStyles({len(self)} exceptions)"


class _CompatLatentStyle:
    """Proxy for one documented ``_LatentStyle`` exception."""

    def __init__(self, latent_styles: _CompatLatentStyles, element: ET.Element):
        self._latent_styles = latent_styles
        self._document = latent_styles._document
        self._element = element

    def _ensure_element(self):
        if self._element is None:
            raise AttributeError("this latent style has been deleted")
        return self._element

    def _get_bool(self, name):
        value = _attr(self._ensure_element(), name)
        return None if not value else value not in ("0", "false", "off")

    def _set_bool(self, name, value):
        element = self._ensure_element()
        attribute = self._document._w(name)
        if value is None:
            element.attrib.pop(attribute, None)
        else:
            element.set(attribute, "1" if bool(value) else "0")
        self._document._touch_styles()

    @property
    def name(self):
        value = _attr(self._ensure_element(), "name") or ""
        return _latent_ui_name(value)

    @property
    def priority(self):
        value = _attr(self._ensure_element(), "uiPriority")
        return int(value) if value else None

    @priority.setter
    def priority(self, value):
        element = self._ensure_element()
        attribute = self._document._w("uiPriority")
        if value is None:
            element.attrib.pop(attribute, None)
        else:
            element.set(attribute, str(int(value)))
        self._document._touch_styles()

    @property
    def hidden(self):
        return self._get_bool("semiHidden")

    @hidden.setter
    def hidden(self, value):
        self._set_bool("semiHidden", value)

    @property
    def locked(self):
        return self._get_bool("locked")

    @locked.setter
    def locked(self, value):
        self._set_bool("locked", value)

    @property
    def quick_style(self):
        return self._get_bool("qFormat")

    @quick_style.setter
    def quick_style(self, value):
        self._set_bool("qFormat", value)

    @property
    def unhide_when_used(self):
        return self._get_bool("unhideWhenUsed")

    @unhide_when_used.setter
    def unhide_when_used(self, value):
        self._set_bool("unhideWhenUsed", value)

    def delete(self):
        element = self._ensure_element()
        block = self._latent_styles._block()
        if block is not None and element in list(block):
            block.remove(element)
            self._document._touch_styles()
        self._element = None

    def __repr__(self):
        return f"LatentStyle({self.name!r})"


class _Styles:
    """``doc.styles`` — lookup, iteration, ``add_style``."""

    def __init__(self, document: Document):
        self._document = document

    def _styles(self):
        return _Document.styles.fget(self._document)

    def __len__(self) -> int:
        return len(self._styles())

    def __iter__(self):
        for style in self._styles():
            yield _ParagraphStyle(self._document, style)

    def __contains__(self, name) -> bool:
        return name in self._styles()

    def __getitem__(self, name) -> _ParagraphStyle:
        return _ParagraphStyle(self._document, self._styles()[name])

    def default(self, style_type) -> "_ParagraphStyle":
        """python-docx's ``styles.default`` — the default style of that type."""
        found = self._styles().default(style_type)
        return _ParagraphStyle(self._document, found) if found else None

    def get_by_id(self, style_id: Optional[str], style_type):
        """Look up by ID and type, returning that type's default on mismatch."""
        if style_type is None:
            if style_id is None:
                return None
            found = self._styles().get_by_id(style_id)
        else:
            if style_id is None:
                return self.default(style_type)
            try:
                found = self._styles().get_by_id(style_id)
            except KeyError:
                return self.default(style_type)
            from ..write import token
            expected = _enum_from_token("WD_STYLE_TYPE", token(style_type))
            actual = _enum_from_token("WD_STYLE_TYPE", found.type)
            if expected != actual:
                return self.default(style_type)
        return _ParagraphStyle(self._document, found)

    def get_style_id(self, style_or_name, style_type) -> str:
        if isinstance(style_or_name, _ParagraphStyle):
            style_or_name = style_or_name._style
        return self._styles().get_style_id(style_or_name, style_type)

    @property
    def latent_styles(self):
        return _CompatLatentStyles(self._document)

    def add_style(self, name: str, style_type,
                  builtin: bool = False) -> _ParagraphStyle:
        """``doc.styles.add_style(name, style_type, builtin=False)``."""
        style = self._styles().add_style(name, style_type, builtin=builtin)
        return _ParagraphStyle(self._document, style)


class _CompatTable:
    """Table proxy that adapts enum-valued properties and child collections."""

    def __init__(self, document: Document, table: Table):
        self._document, self._table = document, table

    def __getattr__(self, name):
        return getattr(self._table, name)

    def __setattr__(self, name, value):
        if name in ("_document", "_table") or hasattr(type(self), name):
            object.__setattr__(self, name, value)
        else:
            setattr(self._table, name, value)

    @property
    def style(self):
        value = self._table.style
        if not value:
            from .enum import WD_STYLE_TYPE
            try:
                return self._document.styles.default(WD_STYLE_TYPE.TABLE)
            except (KeyError, AttributeError):
                return value
        from .enum import WD_STYLE_TYPE
        try:
            return self._document.styles.get_by_id(
                value.style_id, WD_STYLE_TYPE.TABLE)
        except (AttributeError, KeyError):
            return value

    @style.setter
    def style(self, value):
        self._table.style = value

    @property
    def alignment(self):
        # python-docx 1.2.0's getter returns the shared CT_Jc member type
        # (WD_PARAGRAPH_ALIGNMENT), although its setter is documented with
        # WD_TABLE_ALIGNMENT.
        return _enum_from_token("WD_ALIGN_PARAGRAPH", self._table.alignment)

    @alignment.setter
    def alignment(self, value):
        self._table.alignment = value

    @property
    def rows(self):
        return _CompatRows(self)

    @property
    def columns(self):
        return _CompatColumns(self)

    def cell(self, row_idx: int, col_idx: int):
        return _CompatCell(self._document, self,
                           self._table.cell(row_idx, col_idx))

    def row_cells(self, row_idx: int):
        return [_CompatCell(self._document, self, cell)
                for cell in self._table.row_cells(row_idx)]

    def column_cells(self, column_idx: int):
        return [_CompatCell(self._document, self, cell)
                for cell in self._table.column_cells(column_idx)]

    def add_row(self):
        return _CompatRow(self._document, self, self._table.add_row())

    def add_column(self, width):
        return _CompatColumn(self._document, self,
                             self._table.add_column(width))

    def __repr__(self):
        return repr(self._table)


class _CompatRows:
    """Sequence of rows with the same table back-reference as python-docx."""

    def __init__(self, table: _CompatTable):
        self._table = table

    @property
    def table(self):
        return self._table

    def __len__(self) -> int:
        return len(self._table._table.rows)

    def __iter__(self):
        for row in self._table._table.rows:
            yield _CompatRow(self._table._document, self._table, row)

    def __getitem__(self, index):
        rows = list(self)
        return rows[index]


class _CompatColumns:
    """Sequence of columns with the same table back-reference as python-docx."""

    def __init__(self, table: _CompatTable):
        self._table = table

    @property
    def table(self):
        return self._table

    def __len__(self) -> int:
        return len(self._table._table.columns)

    def __iter__(self):
        for column in self._table._table.columns:
            yield _CompatColumn(self._table._document, self._table, column)

    def __getitem__(self, index):
        columns = list(self)
        try:
            return columns[index]
        except IndexError:
            raise IndexError(f"column index [{index}] is out of range") from None


class _CompatRow:
    def __init__(self, document: Document, table: _CompatTable, row):
        self._document, self._table, self._row = document, table, row

    def __getattr__(self, name):
        return getattr(self._row, name)

    def __setattr__(self, name, value):
        if name in ("_document", "_table", "_row") or hasattr(type(self), name):
            object.__setattr__(self, name, value)
        else:
            setattr(self._row, name, value)

    @property
    def table(self):
        return self._table

    @property
    def cells(self):
        return [_CompatCell(self._document, self._table, cell)
                for cell in self._row.cells]

    @property
    def height_rule(self):
        return _enum_from_token("WD_ROW_HEIGHT_RULE", self._row.height_rule)

    @height_rule.setter
    def height_rule(self, value):
        self._row.height_rule = value


class _CompatCell:
    def __init__(self, document: Document, table: _CompatTable, cell):
        self._document, self._table, self._cell = document, table, cell

    def __getattr__(self, name):
        return getattr(self._cell, name)

    def __setattr__(self, name, value):
        if name in ("_document", "_table", "_cell") or hasattr(type(self), name):
            object.__setattr__(self, name, value)
        else:
            setattr(self._cell, name, value)

    @property
    def table(self):
        return self._table

    @property
    def paragraphs(self):
        return [_compat_paragraph(p, self._document)
                for p in self._cell.paragraphs]

    @property
    def tables(self):
        return [_CompatTable(self._document, table)
                for table in self._cell.tables]

    @property
    def vertical_alignment(self):
        return _enum_from_token("WD_CELL_VERTICAL_ALIGNMENT",
                                self._cell.vertical_alignment)

    @vertical_alignment.setter
    def vertical_alignment(self, value):
        self._cell.vertical_alignment = value

    def add_paragraph(self, text: str = "", style=None):
        paragraph = self._cell.add_paragraph(text)
        if style is not None:
            paragraph.style = _style_name(style)
        return _compat_paragraph(paragraph, self._document)

    def add_table(self, rows: int, cols: int):
        return _CompatTable(self._document,
                            self._cell.add_table(rows, cols))

    def merge(self, other_cell):
        core = (other_cell._cell if isinstance(other_cell, _CompatCell)
                else other_cell)
        return _CompatCell(self._document, self._table,
                           self._cell.merge(core))

    def iter_inner_content(self):
        for item in self._cell.iter_inner_content():
            if isinstance(item, _Paragraph):
                yield _compat_paragraph(item, self._document)
            elif isinstance(item, _Table):
                yield _CompatTable(self._document, item)


class _CompatColumn:
    def __init__(self, document: Document, table: _CompatTable, column):
        self._document, self._table, self._column = document, table, column

    def __getattr__(self, name):
        return getattr(self._column, name)

    def __setattr__(self, name, value):
        if name in ("_document", "_table", "_column") or hasattr(type(self), name):
            object.__setattr__(self, name, value)
        else:
            setattr(self._column, name, value)

    @property
    def table(self):
        return self._table

    @property
    def cells(self):
        return [_CompatCell(self._document, self._table, cell)
                for cell in self._column.cells]


# ---------------------------------------------------------------- conveniences
def document_from_template(path: Union[str, Path]) -> Document:
    """Open an existing file as a template (a copy of python-docx's habit)."""
    return Document(path)
