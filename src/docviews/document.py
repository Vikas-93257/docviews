"""The document object: a thin, honest façade over the package and the views."""

from __future__ import annotations

import copy
import os
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, Iterator, List, Optional, Sequence, Union

from xml.etree import ElementTree as ET

from .authoring import (CoreProperties, InlineShape, Section, Sections, Style,
                        Styles, find_inline_shapes, find_part_rels, media_plan,
                        next_media_name)
from .images import ImageError, ImageInfo, detect
from .ooxml import DocxError, Limits, Package, Part, parse_xml, root_ns
from .write import (HEADER_FOOTER_TYPES, IMAGE_CONTENT_TYPES, Builder, Cell,
                    Row, Table, build_package, ensure_relationship,
                    new_package_bytes, numbering_xml, same_meaning, serialize,
                    token, styles_xml)
from .comments import Anchor, Comment, collect_comments
from .views import (REVISION_KINDS, VIEWS, Paragraph, _attr, _tag,
                    iter_paragraphs, own_paragraph_under, part_paragraphs,
                    revision_ids_under, stream_paragraphs)

NUMBER = re.compile(r"\d[\d,]*\.?\d*")


@dataclass
class PartInfo:
    name: str
    kind: str
    paragraphs: int
    chars: Dict[str, int]              # per view
    revision_chars: Dict[str, int] = field(default_factory=dict)
    uncompressed_bytes: int = 0

    @property
    def diverge(self) -> bool:
        return self.chars.get("original", "") != self.chars.get("accepted", "")


@dataclass
class RevisionInfo:
    kind: str                          # insert | delete | move_from | move_to
    id: str
    author: str
    date: str
    part: str
    paragraph_index: int               # -1 for rows, marks and block content
    scope: str = "text"                # text | paragraph-mark | row | block


@dataclass
class NumberToken:
    value: str
    part: str
    paragraph_index: int
    offset: int                        # char offset within the paragraph


def _same_xml(original: bytes, current: bytes) -> bool:
    """Is a rewritten part still the same document, element for element?

    Prefix names are normalised first: this library adopts the file's own
    prefixes, but a part that only differs in prefix spelling was not edited.
    """
    return same_meaning(original, current)


def _normalise(name: str) -> str:
    return "".join(name.lower().split())


_BUILTIN_BY_ID = {
    "Normal": "Normal", "Title": "Title", "Subtitle": "Subtitle",
    "Quote": "Quote", "ListBullet": "List Bullet", "ListNumber": "List Number",
    "TableGrid": "Table Grid", "TableNormal": "Normal Table",
    "Header": "Header", "Footer": "Footer",
}


def _keep_valid(root: ET.Element) -> None:
    """Repair the one shape a caller can leave invalid.

    ``w:hdr``/``w:ftr`` must hold at least one block element, and a header the
    library created may still be empty if nothing was written into it.
    """
    local = _tag(root) if isinstance(root.tag, str) else ""
    if local not in ("hdr", "ftr", "header", "footer"):
        return
    for child in root:
        if isinstance(child.tag, str) and _tag(child) in ("p", "tbl", "sdt"):
            return
    namespace = root.tag.split("}")[0].lstrip("{") if "}" in root.tag else ""
    root.append(ET.Element(f"{{{namespace}}}p"))




def _utc_now() -> str:
    """The timestamp format ``w:date`` wants: UTC, ISO-8601, seconds."""
    from datetime import datetime, timezone
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


class DocumentSettings:
    """``doc.settings`` — a window onto ``word/settings.xml``.

    Only the flag people ask for is named (odd/even page headers); the rest of
    the part is preserved byte-for-byte by the writer, so naming every element
    in it would add surface without adding capability.
    """

    def __init__(self, document):
        self._document = document

    def _name(self) -> str:
        main = self._document._package.main_part
        base = main.rsplit("/", 1)[0] if "/" in main else ""
        return f"{base}/settings.xml" if base else "settings.xml"

    def _root(self, create: bool = False) -> Optional[ET.Element]:
        name = self._name()
        if self._document._has_part(name):
            return self._document._part_root(name)
        if not create:
            return None
        root = ET.Element(self._document._w("settings"))
        self._document._add_part(name, root)
        self._document._reserve_rel("settings",
                                    self._document._rel_to_main(name))
        return root

    @property
    def odd_and_even_pages_header_footer(self) -> bool:
        """True when even pages get a header/footer of their own."""
        root = self._root()
        if root is None:
            return False
        for child in root:
            if _tag(child) == "evenAndOddHeaders":
                return child.get(self._document._w("val")) not in                     ("0", "false", "off")
        return False

    @odd_and_even_pages_header_footer.setter
    def odd_and_even_pages_header_footer(self, value) -> None:
        root = self._root(create=bool(value))
        if root is None:
            return
        for child in list(root):
            if _tag(child) == "evenAndOddHeaders":
                root.remove(child)
        if value:
            root.insert(0, ET.Element(self._document._w("evenAndOddHeaders")))
        self._document._mark_changed(self._name())

    def __repr__(self) -> str:
        return (f"DocumentSettings(odd_and_even="
                f"{self.odd_and_even_pages_header_footer})")




def _count_paragraphs(root: ET.Element) -> int:
    """Every ``w:p`` under *root*, text boxes and tables included.

    One pass over the tree, no Paragraph objects: the walk costs the same as a
    plain iteration and the result is the count the reader's index is built
    from.
    """
    return sum(1 for el in root.iter() if isinstance(el.tag, str)
               and el.tag.endswith("}p"))


class Document:
    """A .docx you can read in three views."""

    # ------------------------------------------------------------------ writing
    @classmethod
    def new(cls, strict: bool = False, limits: Optional[Limits] = None) -> "Document":
        """A new, empty document you can write into and save."""
        return cls(new_package_bytes(strict=strict), limits=limits)

    @property
    def editable(self) -> bool:
        """True when this document came from a package this library can save."""
        return self._package.exists(self._package.main_part)

    def _body(self) -> ET.Element:
        root = self._package.part(self._package.main_part).root
        for child in root:
            if _tag(child) == "body":
                return child
        raise DocxError("the main part has no w:body")

    def _main_builder(self) -> Builder:
        return Builder({"w": root_ns(self._package.part(self._package.main_part).root),
                        "r": self._package.ns["r"]})

    def _appended_index(self, element: ET.Element) -> int:
        """Reading-order index of a paragraph just appended to the body.

        Counting ``w:p`` in the XML is exact and does not build a Paragraph (or
        run ``views_of``) for everything already in the document — the previous
        way took 2.7 s for 500 paragraphs, this takes 0.2 s. The count is then
        carried over to the next append, which changes the body by exactly one
        child and puts *this* element at the tail (``w:sectPr`` stays last, so
        the tail is ``-1`` or ``-2``). Anything else that changes the body — a
        table, a deletion — fails the identity check and is counted again.
        """
        root = self._body()
        cache = self._para_cache
        if cache is not None and cache[0] == id(root) and len(root) == cache[1] + 1:
            # the body grew by exactly one child and it is this paragraph (at
            # the end, or just before the trailing w:sectPr)
            at_tail = len(root) > 0 and root[-1] is element
            if not at_tail and len(root) > 1:
                at_tail = root[-2] is element
            if at_tail:
                self._para_cache = (id(root), len(root), cache[2] + 1, element)
                return cache[2]
        count = _count_paragraphs(root)
        self._para_cache = (id(root), len(root), count, element)
        return count - 1

    def _append(self, element: ET.Element):
        """Append to the body, keeping w:sectPr last (Word requires it)."""
        body = self._body()
        section = None
        if len(body) and _tag(body[-1]) == "sectPr":
            section = body[-1]
        if section is not None:
            body.insert(list(body).index(section), element)
        else:
            body.append(element)
        return element

    # ------------------------------------------------------ package plumbing
    def _w(self, local: str) -> str:
        """A wordprocessingml tag in this document's own namespace."""
        return f"{{{root_ns(self._package.part(self._package.main_part).root)}}}{local}"

    def _r(self, local: str) -> str:
        return f"{{{self._package.ns['r']}}}{local}"

    def _sectPr(self) -> ET.Element:
        """The body's final ``w:sectPr``, created when a document lacks one."""
        body = self._body()
        if len(body) and _tag(body[-1]) == "sectPr":
            return body[-1]
        element = ET.SubElement(body, self._w("sectPr"))
        element.append(self._pgSz())
        element.append(self._pgMar())
        return element

    def _pgSz(self) -> ET.Element:
        """Letter, the same default python-docx has (a missing sectPr is rare)."""
        from xml.etree import ElementTree as _ET
        element = _ET.Element(self._w("pgSz"))
        element.set(self._w("w"), "12240")
        element.set(self._w("h"), "15840")
        return element

    def _pgMar(self) -> ET.Element:
        from xml.etree import ElementTree as _ET
        element = _ET.Element(self._w("pgMar"))
        for name, value in (("top", 1440), ("right", 1800), ("bottom", 1440),
                            ("left", 1800), ("header", 720), ("footer", 720),
                            ("gutter", 0)):
            element.set(self._w(name), str(value))
        return element

    @staticmethod
    def _relationship_part_name(part_name: str) -> str:
        base, tail = part_name.rsplit("/", 1) if "/" in part_name \
            else ("", part_name)
        return f"{base}/_rels/{tail}.rels" if base else f"_rels/{tail}.rels"

    def _main_rels_name(self) -> str:
        return self._relationship_part_name(self._package.main_part)

    def _main_rels_root(self) -> ET.Element:
        name = self._main_rels_name()
        if name in self._added_parts:
            return self._added_parts[name]
        if self._package.exists(name):
            return parse_xml(self._package._zip.read(name))
        root = ET.Element("{http://schemas.openxmlformats.org/package/2006/"
                          "relationships}Relationships")
        self._added_parts[name] = root
        return root

    def _rel_target(self, rel_id: str) -> str:
        """Resolve a relationship id on the main part to the part it names.

        Relationships reserved since the document was opened count too: a
        header's ``r:id`` is written into the ``w:sectPr`` before the package
        has been saved, and reading it back has to work in the meantime.
        """
        for reserved_id, _kind, target in self._extra_rels:
            if reserved_id == rel_id:
                return self._part_for_rel(target)
        for rel in self._main_rels_root():
            if rel.get("Id") == rel_id:
                target = (rel.get("Target") or "").lstrip("/")
                base = self._package.main_part.rsplit("/", 1)[0] \
                    if "/" in self._package.main_part else ""
                if target.startswith("word/") or not base:
                    return target
                return f"{base}/{target}"
        return ""

    def _reserve_rel(self, kind: str, target: str) -> str:
        """Reserve a relationship id on the main document part."""
        used = {rel.get("Id", "") for rel in self._main_rels_root()}
        used |= self._reserved_rels
        index = 1
        while f"rId{index}" in used:
            index += 1
        rel_id = f"rId{index}"
        self._reserved_rels.add(rel_id)
        self._extra_rels.append((rel_id, kind, target))
        return rel_id

    def _reserve_rel_for_part(self, part_name: str, kind: str,
                              target: str) -> str:
        """Reserve a relationship on a story part such as a header or footer."""
        rels_name = self._relationship_part_name(part_name)
        if rels_name in self._added_parts:
            root = self._added_parts[rels_name]
        elif self._package.exists(rels_name):
            root = parse_xml(self._package._zip.read(rels_name))
        else:
            root = ET.Element("{http://schemas.openxmlformats.org/package/"
                              "2006/relationships}Relationships")
        # Relationship IDs are scoped to their source part. Do not use the
        # document-wide reservation set here; a footer may validly use rId1
        # even when the body and header already do.
        used = {rel.get("Id", "") for rel in root}
        used |= {rel_id for source, rel_id, _rel_kind, _target
                 in self._extra_story_rels if source == part_name}
        index = 1
        while f"rId{index}" in used:
            index += 1
        rel_id = f"rId{index}"
        self._extra_story_rels.append((part_name, rel_id, kind, target))
        return rel_id

    def _reserve_package_rel(self, kind: str, target: str) -> str:
        """Reserve a relationship id on the package itself (``_rels/.rels``).

        The core properties hang off the package, not off the document part, so
        they need an id of their own here.
        """
        root = parse_xml(self._package._zip.read("_rels/.rels"))
        used = {rel.get("Id", "") for rel in root} | self._reserved_rels
        index = 1
        while f"rId{index}" in used:
            index += 1
        rel_id = f"rId{index}"
        self._reserved_rels.add(rel_id)
        self._extra_package_rels.append((rel_id, kind, target))
        return rel_id

    def _rel_to_part(self, target_part: str, source_part: str) -> str:
        """A target part as one story part sees it in a relationship."""
        import posixpath
        base = source_part.rsplit("/", 1)[0] if "/" in source_part else "."
        return posixpath.relpath(target_part, base or ".")

    def _rel_to_main(self, part_name: str) -> str:
        """A part name as the main part sees it: relationships are relative."""
        return self._rel_to_part(part_name, self._package.main_part)

    def _part_for_rel(self, target: str, source_part: Optional[str] = None) -> str:
        """The inverse: a relationship target resolved to a package part name."""
        import posixpath
        if target.startswith(("http://", "https://")):
            return target
        source = source_part or self._package.main_part
        base = source.rsplit("/", 1)[0] if "/" in source else "."
        if target.startswith("/"):
            return target.lstrip("/")
        name = target
        if base != "." and name.startswith(f"{base}/"):
            return name
        resolved = posixpath.normpath(posixpath.join(base, name))
        return resolved[2:] if resolved.startswith("./") else resolved

    def _has_part(self, name: str) -> bool:
        return name in self._added_parts or self._package.exists(name)

    def _add_part(self, name: str, root: ET.Element) -> ET.Element:
        self._added_parts[name] = root
        return root

    def _part_root(self, name: str) -> ET.Element:
        """The live tree of a part: an added one, or the parsed original.

        Parsing here does not commit to rewriting: ``to_bytes`` compares the
        serialised tree with the original bytes and leaves untouched parts alone.
        """
        if name in self._added_parts:
            return self._added_parts[name]
        if name not in self._edited_roots:
            if not self._package.exists(name):
                raise DocxError(f"this package has no part {name}")
            self._edited_roots[name] = self._package.part(name).root
        return self._edited_roots[name]

    def _mark_changed(self, name: str) -> None:
        if name not in self._added_parts:
            self._changed_parts.add(name)

    def _free_part_name(self, kind: str) -> str:
        """``word/header4.xml`` — the first free name for a header/footer part."""
        import re
        base = self._package.main_part.rsplit("/", 1)[0] \
            if "/" in self._package.main_part else "word"
        used = set(self._package.names()) | set(self._added_parts)
        pattern = re.compile(rf"^{base}/{kind}(\d+)\.xml$")
        highest = 0
        for name in used:
            match = pattern.match(name)
            if match:
                highest = max(highest, int(match.group(1)))
        candidate = f"{base}/{kind}{highest + 1}.xml"
        while candidate in used:
            highest += 1
            candidate = f"{base}/{kind}{highest + 1}.xml"
        return candidate

    def _styles_name(self) -> str:
        """The document's styles part, found the way Word means it.

        Through the main part's relationships: a package may hold more than one
        ``styles.xml`` (a glossary has its own), and only the one the document
        part is related to is the document's.
        """
        for candidate in self._added_parts:
            if candidate.rsplit("/", 1)[-1] == "styles.xml":
                return candidate
        for rel in self._main_rels_root():
            kind = rel.get("Type") or ""
            if kind.endswith("/styles"):
                target = self._rel_target(rel.get("Id", ""))
                if target:
                    return target
        base = self._package.main_part.rsplit("/", 1)[0] \
            if "/" in self._package.main_part else ""
        candidate = f"{base}/styles.xml" if base else "styles.xml"
        return candidate if self._package.exists(candidate) else ""

    def _styles_root(self, create: bool = False) -> Optional[ET.Element]:
        """The live styles tree; with ``create`` a stub part is added on demand."""
        name = self._styles_name()
        if name:
            return self._part_root(name)
        if not create:
            return None
        base = self._package.main_part.rsplit("/", 1)[0] \
            if "/" in self._package.main_part else "word"
        name = f"{base}/styles.xml"
        root = parse_xml(styles_xml({"w": root_ns(
            self._package.part(self._package.main_part).root),
            "r": self._package.ns["r"]}))
        self._added_parts[name] = root
        return root

    def _touch_styles(self) -> None:
        name = self._styles_name()
        if name:
            self._mark_changed(name)

    def _needs_part(self, stem: str) -> bool:
        """Register a stub part (styles/numbering) the document must grow."""
        for name in list(self._package.names()) + list(self._added_parts):
            if name.rsplit("/", 1)[-1] == f"{stem}.xml":
                return False
        self._must_add.add(stem)
        return True

    # ------------------------------------------------------------- paragraphs
    def add_paragraph(self, text: str = "", *, style: Optional[str] = None,
                      align: Optional[str] = None, **fmt) -> Paragraph:
        """Append a paragraph. ``style`` takes a name ("Heading 1") or an id."""
        builder = self._main_builder()
        style_id = self._style_id(style) if style else None
        element = builder.paragraph(text, style=style_id,
                                    align=token(align), **fmt)
        self._append(element)
        self._touch_styles() if style_id else None
        # the new paragraph is the last one in the part, so its reading-order
        # index is (paragraphs in the part) - 1. Counting the XML directly is
        # exact and does not build a Paragraph (or run views_of) for every
        # paragraph already in the document — the difference between adding
        # 500 paragraphs in 0.1 s and in 2.7 s.
        index = self._appended_index(element)
        return Paragraph(text=text, view="accepted", kind="document",
                         part_name=self._package.main_part, index=index,
                         _element=element, _root=self._body(),
                         _w_ns=builder.w, _default_style=self._default_style(),
                         _document=self)

    def _default_style(self) -> Optional[str]:
        from .views import default_paragraph_style
        return default_paragraph_style(self._package) or "Normal"

    def add_heading(self, text: str = "", level: int = 1) -> Paragraph:
        """A heading, python-docx style: ``level=0`` is the Title style."""
        if not 0 <= level <= 9:
            raise ValueError("level must be 0..9")
        if level == 0:
            return self.add_paragraph(text, style="Title")
        self._needs_part("styles")
        return self.add_paragraph(text, style=f"Heading{level}")

    def add_list_item(self, text: str = "", kind: str = "bullet",
                      level: int = 0) -> Paragraph:
        """A bullet or numbered list item ("bullet", "number")."""
        if kind not in ("bullet", "number"):
            raise ValueError('kind must be "bullet" or "number"')
        self._needs_part("numbering")
        style = "ListBullet" if kind == "bullet" else "ListNumber"
        paragraph = self.add_paragraph(text, style=style)
        if level:
            element = paragraph._element
            props = element.find(self._w("pPr"))
            numbering = ET.Element(self._w("numPr"))
            for name, value in (("ilvl", level),
                                ("numId", 1 if kind == "bullet" else 2)):
                child = ET.Element(self._w(name))
                child.set(self._w("val"), str(value))
                numbering.append(child)
            props.insert(0, numbering)
        return paragraph

    def add_page_break(self) -> Paragraph:
        """A paragraph holding one page break."""
        builder = self._main_builder()
        element = self._append(builder.page_break_paragraph())
        return Paragraph(text="", view="accepted", kind="document",
                         part_name=self._package.main_part,
                         _element=element, _root=self._body(),
                         _w_ns=builder.w)

    def add_table(self, rows: int = 1, cols: int = 1, *,
                  style: str = "Table Grid",
                  text: Optional[Iterable[Iterable[str]]] = None) -> Table:
        """A table; pass ``text`` and it is filled in, row by row."""
        if rows < 1 or cols < 1:
            raise ValueError("a table needs at least one row and one column")
        if text is not None:
            data = [list(row) for row in text]
            rows = max(rows, len(data))
            cols = max([cols] + [len(row) for row in data])
        else:
            data = []
        self._needs_part("styles")
        builder = self._main_builder()
        element = self._append(builder.table(rows, cols,
                                             style=self._style_id(style),
                                             text=data))
        return Table(element, builder.w)

    # ------------------------------------------------------------------ styles
    @property
    def settings(self) -> "DocumentSettings":
        """``word/settings.xml`` — the flags a document author actually changes."""
        return DocumentSettings(self)



    @property
    def styles(self) -> Styles:
        """The document's styles: iterate, look up by name, or ``add_style``."""
        return Styles(self)

    def _style_id(self, name: str) -> str:
        """Map a style name ("Heading 1") onto the id the document declares.

        Word writes names like "heading 1" and ids like "Heading1"; python-docx
        hides that difference, so this does too.
        """
        if not name:
            return name
        wanted = _normalise(name)
        root = self._styles_root()
        if root is not None:
            for style in root:
                if _tag(style) != "style":
                    continue
                style_id = _attr(style, "styleId") or ""
                label = ""
                for child in style:
                    if _tag(child) == "name":
                        label = _attr(child, "val") or ""
                if wanted in (_normalise(style_id), _normalise(label)):
                    return style_id
        for style_id, label in _BUILTIN_BY_ID.items():
            if wanted in (_normalise(label), _normalise(style_id)):
                return style_id
        return name.replace(" ", "")

    # ----------------------------------------------------------------- images
    @property
    def core_properties(self) -> "CoreProperties":
        """docProps/core.xml: title, author, subject, keywords, and the rest."""
        return CoreProperties(self)

    def _inline_shapes(self) -> List[InlineShape]:
        """The pictures. Core code calls this, never the property below.

        A subclass may substitute its own ``inline_shapes`` (the compat shim
        does), and core code must keep working underneath that.
        """
        roots = dict(self._edited_roots)
        roots[self._package.main_part] = self._part_root(self._package.main_part)
        roots.update({name: value for name, value in self._added_parts.items()
                      if isinstance(value, ET.Element)
                      and not name.endswith(".rels")})
        shapes = find_inline_shapes(self._package, roots=roots)
        reserved_main = {rel_id: self._part_for_rel(target)
                         for rel_id, _kind, target in self._extra_rels}
        reserved_story = {(source, rel_id): self._part_for_rel(target, source)
                          for source, rel_id, _kind, target
                          in self._extra_story_rels}
        for shape in shapes:
            if not shape.media and shape.embed in reserved_main:
                shape.media = reserved_main[shape.embed]
            if not shape.media and (shape.part_name, shape.embed) in reserved_story:
                shape.media = reserved_story[(shape.part_name, shape.embed)]
        return shapes

    def _pictures(self) -> List[InlineShape]:
        """All current inline drawings, including unsaved new story parts."""
        roots = dict(self._edited_roots)
        roots.update({name: value for name, value in self._added_parts.items()
                      if isinstance(value, ET.Element)
                      and not name.endswith(".rels")})
        shapes = find_inline_shapes(self._package, roots=roots)
        seen = {(shape.part_name, shape.embed) for shape in shapes}
        shapes.extend(shape for shape in self._pending_shapes
                      if (shape.part_name, shape.embed) not in seen)
        return shapes

    @property
    def inline_shapes(self) -> List[InlineShape]:
        """Every picture in the package, with its relationship resolved.

        Pictures inserted since the document was opened are included, in the
        order they appear in the file — ``len(doc.inline_shapes)`` is true as
        soon as ``add_picture`` returns.
        """
        return self._inline_shapes()

    def add_picture(self, image_path_or_stream, width: Optional[int] = None,
                    height: Optional[int] = None, *, run: Optional[Run] = None
                    ) -> InlineShape:
        """Insert a picture, creating the media part and its relationship.

        ``image_path_or_stream`` is a path, readable stream, or image bytes;
        ``width``/``height`` are EMU
        (``Inches(4)``), and give one to keep the aspect ratio. Without a
        measurement the image cannot be placed, so an unreadable format raises
        rather than guessing.
        """
        facts, data, filename = media_plan(self, image_path_or_stream)
        media = next_media_name(self._package, facts.extension,
                                reserved=self._added_parts)
        cx, cy = facts.scaled(width, height)
        source_part = (getattr(run, "_part_name", "") or self._package.main_part) \
            if run is not None else self._package.main_part
        target = self._rel_to_part(media, source_part)
        if source_part == self._package.main_part:
            rel_id = self._reserve_rel("image", target)
        else:
            if not self._has_part(source_part):
                raise DocxError(f"run belongs to unknown story part {source_part!r}")
            rel_id = self._reserve_rel_for_part(source_part, "image", target)
        self._extra_defaults.append((facts.extension,
                                     IMAGE_CONTENT_TYPES[facts.extension]))
        if data:
            self._added_parts[media] = data          # bytes, not XML
        story_root = self._part_root(source_part)
        builder = Builder({"w": root_ns(story_root), "r": self._package.ns["r"]})
        shape_name = f"Picture {len(self._pictures()) + 1}"
        drawing = builder.drawing(rel_id, cx, cy, name=shape_name,
                                  docpr_id=self._next_docpr_id(), filename=filename)
        if run is not None:
            run._element.append(drawing)
        else:
            paragraph_element = builder.el("w:p")
            paragraph_run = builder.el("w:r", paragraph_element)
            paragraph_run.append(drawing)
            self._append(paragraph_element)
        inline = next((element for element in drawing.iter()
                       if _tag(element) == "inline"), None)
        shape = InlineShape(part_name=source_part, name=shape_name,
                            width=cx, height=cy, embed=rel_id, media=media,
                            filename=filename,
                            graphic_uri="http://schemas.openxmlformats.org/"
                                       "drawingml/2006/picture",
                            _element=inline)
        self._pending_shapes.append(shape)
        return shape

    # ------------------------------------------------------------------ comments
    def _comments_part(self, create: bool = False):
        """The comments part as ``(name, root)``, created on demand.

        Found the way Word means it — through the main part's relationships —
        and only then by name, because a glossary can carry a comments part of
        its own that is not the document's.
        """
        for rel_id, kind, target in self._extra_rels:
            if kind != "comments":
                continue
            name = self._part_for_rel(target)
            if self._has_part(name):
                return name, self._part_root(name)
        for rel in self._main_rels_root():
            if (rel.get("Type") or "").endswith("/comments"):
                name = self._rel_target(rel.get("Id", ""))
                if name and self._has_part(name):
                    return name, self._part_root(name)
        for candidate in list(self._package.names()) + list(self._added_parts):
            if candidate.endswith("comments.xml") and "/glossary/" not in candidate:
                if self._has_part(candidate):
                    return candidate, self._part_root(candidate)
        if not create:
            return None, None
        base = self._package.main_part.rsplit("/", 1)[0] \
            if "/" in self._package.main_part else ""
        name = f"{base}/comments.xml" if base else "comments.xml"
        root = ET.Element(self._w("comments"))
        self._add_part(name, root)
        self._reserve_rel("comments", self._rel_to_main(name))
        return name, root

    def _added_comment_parts(self) -> List[Part]:
        """Comment parts created since the document was opened."""
        out = []
        for name, root in self._added_parts.items():
            if isinstance(root, ET.Element) and \
                    name.endswith("comments.xml") and "/glossary/" not in name:
                out.append(Part(name=name, kind="comments", _root=root))
        return out

    @staticmethod
    def _next_comment_id(root: ET.Element) -> int:
        """max(id) + 1, which is what Word and python-docx both do."""
        highest = -1
        for child in root:
            if not isinstance(child.tag, str):
                continue
            for key, value in child.attrib.items():
                if key.endswith("}id"):
                    try:
                        highest = max(highest, int(value))
                    except ValueError:
                        pass
        return highest + 1

    def add_comment(self, runs, text: str = "", author: str = "",
                    initials: str = "", date: Optional[str] = None):
        """Attach a comment to one or more runs (python-docx's ``add_comment``).

        Returns a :class:`~docviews.comments.Comment` whose anchor already
        points at the runs — the range python-docx 1.2.0 documents as not
        implemented.
        """
        wanted = list(runs) if isinstance(runs, (list, tuple)) else [runs]
        if not wanted:
            raise DocxError("a comment needs at least one run to point at")
        name, root = self._comments_part(create=True)
        comment_id = self._next_comment_id(root)
        stamp = date or _utc_now()
        element = ET.SubElement(root, self._w("comment"))
        for attribute, value in (("id", str(comment_id)), ("author", author),
                                 ("initials", initials), ("date", stamp)):
            if value or attribute in ("id", "author"):
                element.set(self._w(attribute), value)
        paragraph = ET.SubElement(element, self._w("p"))
        properties = ET.SubElement(paragraph, self._w("pPr"))
        style = ET.SubElement(properties, self._w("pStyle"))
        style.set(self._w("val"), "CommentText")
        run = ET.SubElement(paragraph, self._w("r"))
        text_element = ET.SubElement(run, self._w("t"))
        text_element.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        text_element.text = text
        marker = ET.SubElement(paragraph, self._w("r"))
        marker_properties = ET.SubElement(marker, self._w("rPr"))
        marker_style = ET.SubElement(marker_properties, self._w("rStyle"))
        marker_style.set(self._w("val"), "CommentReference")
        ET.SubElement(marker, self._w("annotationRef"))
        self._mark_changed(name)
        located = self._mark_comment_range(comment_id, wanted[0], wanted[-1])
        comment = Comment(id=str(comment_id), author=author, initials=initials,
                          date=stamp, text=text, part=name)
        if located:
            part_name, paragraph_index, start, end = located
            comment.anchors.append(Anchor(
                part=part_name, paragraph_index=paragraph_index, start=start,
                end=end, text="".join(run.text for run in wanted)))
        return comment

    def _live_roots(self):
        """Every part tree this document can put text in, by part name."""
        roots = {self._package.main_part: self._body_root()}
        for name, root in self._edited_roots.items():
            roots[name] = root
        for name, root in self._added_parts.items():
            if isinstance(root, ET.Element):
                roots.setdefault(name, root)
        return roots

    @staticmethod
    def _paragraph_index(paragraph: ET.Element) -> int:
        """Where this paragraph sits among the paragraphs of its part."""
        root = paragraph
        while True:                            # walk up to the part's root
            parent = None
            for candidate in root.iter():
                for child in candidate:
                    if child is root:
                        parent = candidate
                        break
                if parent is not None:
                    break
            if parent is None:
                break
            root = parent
        paragraphs = [element for element in root.iter()
                      if isinstance(element.tag, str) and _tag(element) == "p"]
        for index, element in enumerate(paragraphs):
            if element is paragraph:
                return index
        return len(paragraphs)

    @staticmethod
    def _offset(children, element, end_of: bool = False) -> int:
        """The character offset of ``element`` in its paragraph's text."""
        offset = 0
        for child in children:
            if child is element:
                if end_of:
                    return offset + Document._run_text(child)
                return offset
            offset += Document._run_text(child)
        return offset

    @staticmethod
    def _run_text(element: ET.Element) -> int:
        return sum(len(node.text or "") for node in element.iter()
                   if isinstance(node.tag, str) and node.tag.endswith("}t"))

    def _body_root(self) -> ET.Element:
        return self._package.part(self._package.main_part).root

    def _mark_comment_range(self, comment_id: int, first, last):
        """Write the range markers around ``first`` … ``last`` (inclusive).

        Returns ``(part name, paragraph index, start, end)`` — where in the
        document the range sits, which is what a reader of the file would work
        out from the markers.
        """
        first_element, last_element = first._element, last._element
        target = None
        for part_name, root in self._live_roots().items():
            for element in root.iter():
                if not isinstance(element.tag, str) or _tag(element) != "p":
                    continue
                children = list(element)
                if any(child is first_element for child in children) and \
                        any(child is last_element for child in children):
                    target = (part_name, element)
                    break
            if target is not None:
                break
        if target is None:
            raise DocxError(
                "the runs being commented are not in this document's body, "
                "header or footer")
        part_name, paragraph = target
        children = list(paragraph)
        start_index = next(i for i, c in enumerate(children) if c is first_element)
        end_index = next(i for i, c in enumerate(children) if c is last_element)
        start = ET.Element(self._w("commentRangeStart"))
        start.set(self._w("id"), str(comment_id))
        end = ET.Element(self._w("commentRangeEnd"))
        end.set(self._w("id"), str(comment_id))
        reference = ET.Element(self._w("r"))
        properties = ET.SubElement(reference, self._w("rPr"))
        style = ET.SubElement(properties, self._w("rStyle"))
        style.set(self._w("val"), "CommentReference")
        anchor = ET.SubElement(reference, self._w("commentReference"))
        anchor.set(self._w("id"), str(comment_id))
        paragraph[:] = (children[:start_index] + [start]
                        + children[start_index:end_index + 1]
                        + [end, reference] + children[end_index + 1:])
        self._mark_changed(part_name)
        return (part_name, self._paragraph_index(paragraph),
                self._offset(children, first_element),
                self._offset(children, last_element, end_of=True))

    def _next_docpr_id(self) -> int:
        """``wp:docPr id`` must be unique across all story parts."""
        roots = {part.name: self._edited_roots.get(part.name, part.root)
                 for part in self._package.text_parts()}
        roots.update({name: value for name, value in self._added_parts.items()
                      if isinstance(value, ET.Element)
                      and not name.endswith(".rels")})
        highest = 0
        for root in roots.values():
            for element in root.iter():
                if isinstance(element.tag, str) and \
                        element.tag.endswith("}docPr"):
                    try:
                        highest = max(highest, int(element.get("id") or 0))
                    except ValueError:
                        continue
        return highest + 1

    # --------------------------------------------------------------- sections
    @property
    def sections(self) -> Sections:
        """Page settings, in document order; ``sections[0].header`` and friends."""
        return Sections(self)

    def add_section(self) -> Section:
        """Start a new section (Word's Layout → Breaks → Next Page).

        The current final ``w:sectPr`` moves into the last paragraph, which is
        how Word records a section break, and a fresh copy becomes the new
        section's properties.
        """
        body = self._body()
        current = self._sectPr()
        body.remove(current)
        builder = self._main_builder()
        paragraph = builder.el("w:p")
        props = builder.el("w:pPr", paragraph)
        props.append(current)                 # the finished section's properties
        body.append(paragraph)
        fresh = copy.deepcopy(current)
        for child in list(fresh):
            if _tag(child) in ("headerReference", "footerReference"):
                fresh.remove(child)           # a new section starts linked
        body.append(fresh)
        return Section(self, fresh)

    # ------------------------------------------------------------------ tables
    @property
    def tables(self) -> List[Table]:
        """Top-level tables of the body, in order."""
        ns = root_ns(self._package.part(self._package.main_part).root)
        out = []
        for child in self._body():
            if _tag(child) == "tbl":
                out.append(Table(child, ns))
        return out

    # -------------------------------------------------------------------- save
    def to_bytes(self) -> bytes:
        """The document as a .docx package.

        Changed parts are re-serialised with the prefixes they came with, new
        parts are added with their content-type entries and relationships, and
        everything else is copied byte-for-byte out of the original package.
        """
        package = self._package
        main = package.main_part
        replacements: Dict[str, bytes] = {}
        additions: Dict[str, bytes] = {}

        main_root = package.part(main).root
        current = serialize(main_root, package.part(main).xml)
        if not _same_xml(package.part(main).xml, current):
            replacements[main] = current

        for name, root in list(self._edited_roots.items()):
            if name in self._added_parts or name == main:
                continue
            if name == self._main_rels_name():
                continue                 # relationships are rebuilt below
            if not package.exists(name):
                continue
            original = package._zip.read(name)
            rewritten = serialize(root, original)
            if not _same_xml(original, rewritten):
                replacements[name] = rewritten

        ns = dict(package.ns)
        for stem in sorted(self._must_add):
            name = f"{package.main_part.rsplit('/', 1)[0]}/{stem}.xml" \
                if "/" in package.main_part else f"{stem}.xml"
            if package.exists(name) or name in self._added_parts:
                continue
            if stem == "styles":
                additions[name] = styles_xml(ns)
            elif stem == "numbering":
                additions[name] = numbering_xml(ns)

        for name, value in self._added_parts.items():
            if isinstance(value, (bytes, bytearray)):
                additions[name] = bytes(value)
            elif name == self._main_rels_name():
                continue
            else:
                _keep_valid(value)
                additions[name] = serialize(value)
        rels_root = self._main_rels_root()
        if rels_root is not None and self._extra_rels:
            additions[self._main_rels_name()] = serialize(rels_root)

        story_rels: Dict[str, bytes] = {}
        empty_rels = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
            '2006/relationships"></Relationships>').encode()
        for source_part, rel_id, kind, target in self._extra_story_rels:
            rels_name = self._relationship_part_name(source_part)
            if rels_name in story_rels:
                original = story_rels[rels_name]
            elif rels_name in additions:
                original = additions[rels_name]
            elif rels_name in replacements:
                original = replacements[rels_name]
            elif rels_name in self._added_parts and isinstance(
                    self._added_parts[rels_name], ET.Element):
                original = serialize(self._added_parts[rels_name])
            elif package.exists(rels_name):
                original = package._zip.read(rels_name)
            else:
                original = empty_rels
            story_rels[rels_name] = ensure_relationship(
                original, package.ns["r"], kind, target, rel_id)
        for rels_name, data in story_rels.items():
            if package.exists(rels_name):
                replacements[rels_name] = data
            else:
                additions[rels_name] = data

        return build_package(package, replacements, additions,
                             extra_rels=self._extra_rels,
                             extra_content_types=self._extra_content_types,
                             extra_defaults=self._extra_defaults,
                             package_rels=self._extra_package_rels)

    def save(self, path: Optional[Union[str, Path]] = None) -> str:
        """Write the package back out. Returns the path written.

        The write is atomic: the bytes land in a temporary file next to the
        target and are moved into place, so a failure halfway through cannot
        leave a half-written document behind.
        """
        if path is None:
            if self.path in ("", "<bytes>"):
                raise DocxError("no path to save to: pass one, or use to_bytes()")
            target = Path(self.path)
        else:
            target = Path(path)
        data = self.to_bytes()
        tmp = target.with_name(target.name + ".docviews-tmp")
        try:
            tmp.write_bytes(data)
            os.replace(tmp, target)
        finally:
            if tmp.exists():
                tmp.unlink()
        return str(target)


    # ------------------------------------------------------------------ basics
    @property
    def warnings(self) -> List[str]:
        return self._package.warnings

    @property
    def strict(self) -> bool:
        """ISO/IEC 29500 Strict package (python-docx raises on these)."""
        return self._package.strict

    @property
    def main_part(self) -> str:
        return self._package.main_part

    # ------------------------------------------------------------------- views
    def iter_paragraphs(self, view: str = "accepted",
                        kinds: Optional[Sequence[str]] = None,
                        with_views: bool = False) -> Iterator[Paragraph]:
        """Lazy walk of every text-bearing part — bounded memory, no list."""
        rels = {name: find_part_rels(self._package, name)
                for name in {part.name for part in self._package.text_parts()}}
        for paragraph in iter_paragraphs(self._package, view=view, kinds=kinds,
                                         with_views=with_views):
            paragraph._rels = rels.get(paragraph.part_name, {})
            paragraph._document = self
            yield paragraph

    def paragraphs(self, view: str = "accepted",
                   kinds: Optional[Sequence[str]] = None,
                   with_views: bool = False) -> List[Paragraph]:
        """Every paragraph, in reading order, as a list — like python-docx's.

        For a document big enough that materialising the list is the problem,
        ``iter_paragraphs()`` streams the same walk and ``stream()`` walks the
        main part alone at a flat ~15 MB (census §3.6).
        """
        return list(self.iter_paragraphs(view, kinds, with_views))

    def text(self, view: str = "accepted", separator: str = "\n",
             kinds: Optional[Sequence[str]] = None) -> str:
        return separator.join(p.text for p in self.iter_paragraphs(view, kinds))

    # backwards-compatible aliases, because the names people reach for first
    @property
    def original(self) -> str:
        return self.text("original")

    @property
    def accepted(self) -> str:
        return self.text("accepted")

    def stream(self, view: str = "accepted") -> Iterator[Paragraph]:
        """Bounded-memory walk of the main part (see census §3.6)."""
        return stream_paragraphs(self._package, view=view)

    """A .docx you can read in three views."""

    def __init__(self, source: Union[str, Path, bytes], limits: Optional[Limits] = None):
        self._package = Package(source, limits=limits)
        self.path = str(source) if not isinstance(source, bytes) else "<bytes>"
        self._must_add: set = set()                 # stub parts to create (styles…)
        self._added_parts: Dict[str, ET.Element] = {}   # new parts, by name
        self._changed_parts: set = set()            # existing parts to rewrite
        self._extra_rels: List[tuple] = []          # (id, kind, target) on the main part
        self._extra_story_rels: List[tuple] = []    # (part, id, kind, target)
        self._extra_package_rels: List[tuple] = []   # (id, kind, target) on the package
        self._extra_content_types: List[tuple] = []  # (part name, content type)
        self._extra_defaults: List[tuple] = []      # (extension, content type)
        self._reserved_rels: set = set()            # relationship ids spoken for
        self._edited_roots: Dict[str, ET.Element] = {}   # parsed roots we may write
        self._pending_shapes: List[InlineShape] = []
        self._para_cache = None                     # (id(root), len, count, element)

    # ------------------------------------------------------------------- views
    def parts(self) -> List[PartInfo]:
        out: List[PartInfo] = []
        for part in self._package.text_parts():
            counts: Dict[str, int] = {}
            revised: Dict[str, int] = {}
            n_paras = 0
            for view in VIEWS:
                paras = list(part_paragraphs(part, view))
                counts[view] = len("\n".join(p.text for p in paras))
                n_paras = max(n_paras, len(paras))
                if view == "accepted":
                    for p in paras:
                        for kind, n in p.revision_chars.items():
                            revised[kind] = revised.get(kind, 0) + n
            out.append(PartInfo(
                name=part.name, kind=part.kind, paragraphs=n_paras, chars=counts,
                revision_chars=revised,
                uncompressed_bytes=self._package.expanded_bytes(part.name)))
        return out

    # --------------------------------------------------------------- metadata
    def revisions(self) -> List[RevisionInfo]:
        """Every tracked change in every part, with author, date and position.

        Revisions do not only sit around runs. Word also marks a *deleted table
        row* (``w:del`` inside ``w:trPr``) and a *deleted paragraph mark*
        (``w:del`` inside ``w:pPr/w:rPr``), and both change what the document
        says — the row goes away, two paragraphs merge — while no Python text
        reader reports either. ``scope`` says which kind of revision this is,
        so a caller can tell "the words changed" from "the paragraph went away".
        """
        out: List[RevisionInfo] = []
        for part in self._package.text_parts():
            counter = [0]

            def alloc() -> int:
                index = counter[0]
                counter[0] += 1
                return index

            def walk(el, path: tuple, p_index: int) -> None:
                for child in el:
                    tag = _tag(child)
                    kind = REVISION_KINDS.get(tag)
                    if kind:
                        if path[-2:] == ("pPr", "rPr") or path[-1:] == ("pPr",):
                            scope = "paragraph-mark"
                        elif path[-1:] == ("trPr",):
                            scope = "row"
                        elif own_paragraph_under(child):
                            scope = "block"
                        else:
                            scope = "text"
                        out.append(RevisionInfo(
                            kind, _attr(child, "id"), _attr(child, "author"),
                            _attr(child, "date"), part.name,
                            p_index if scope == "text" and p_index >= 0 else -1,
                            scope))
                        walk(child, path + (tag,), p_index)
                    elif tag == "p":
                        walk(child, path + (tag,), alloc())
                    else:
                        walk(child, path + (tag,), p_index)

            walk(part.root, (), -1)
        return out

    def comments(self, view: str = "accepted") -> List[Comment]:
        """Every comment, with anchor ranges, replies and resolved state.

        python-docx 1.2.0 gives you comment text and documents that it cannot
        give you the range — this returns the range.
        """
        return collect_comments(self, view=view,
                                extra=self._added_comment_parts())

    def numbers(self, view: str = "accepted", kinds=None) -> List[NumberToken]:
        """Every number in the document, with where it came from."""
        out: List[NumberToken] = []
        for p in self.iter_paragraphs(view, kinds):
            for m in NUMBER.finditer(p.text):
                out.append(NumberToken(value=m.group().replace(",", "").strip("."),
                                       part=p.part_name, paragraph_index=p.index,
                                       offset=m.start()))
        return out

    # ------------------------------------------------------------------ housekeeping
    def close(self) -> None:
        self._package.close()

    def __enter__(self) -> "Document":
        return self

    def __exit__(self, *exc) -> None:
        self.close()


def open_docx(source: Union[str, Path, bytes],
              limits: Optional[Limits] = None) -> Document:
    """Open a .docx (or an in-memory package) for reading."""
    return Document(source, limits=limits)
