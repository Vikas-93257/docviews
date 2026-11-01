"""Writing: create .docx files, and edit opened ones without destroying them.

Two rules shape this module.

**Rule 1 — a reader that can write must not be a writer that deletes.** When an
opened document is saved, every part the caller did not touch is copied
byte-for-byte out of the original package. Footnotes, comments, tracked-change
markup, custom XML, embedded media, unsupported extension parts: none of them
pass through the parser, so none of them can be damaged by it. Only the parts
that actually changed are re-serialised, and those are re-serialised with the
prefixes, comments and processing instructions the original used.

**Rule 2 — nothing is invented silently.** The output is a real package:
`[Content_Types].xml` entries and relationships are added only when the part
they describe is added, and the parts this library could not write are reported
by the reader's `warnings`, not hidden.
"""

from __future__ import annotations

import re
import zipfile
from importlib.resources import files
from io import BytesIO
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

from xml.etree import ElementTree as ET

from .units import Length, RGBColor
from .ooxml import PREFIXES, STRICT, TRANSITIONAL, DocxError, Package, parse_xml

XML_NS = "http://www.w3.org/XML/1998/namespace"

# --------------------------------------------------------------------- units
EMU_PER_PT = 12700
EMU_PER_INCH = 914400
EMU_PER_CM = 360000
EMU_PER_MM = 36000


def token(value) -> Optional[str]:
    """The OOXML token for an enum member, a string or a number.

    ``WD_ALIGN_PARAGRAPH.CENTER`` is ``1`` (python-docx's value) but Word's XML
    wants ``"center"``, so every value that reaches the file goes through here.
    """
    if value is None:
        return None
    missing = object()
    ooxml = getattr(value, "ooxml", missing)
    if ooxml is not missing:
        return ooxml if isinstance(ooxml, str) and ooxml else None
    if isinstance(value, str):
        return value
    return str(getattr(value, "value", value))


def half_points(value: Optional[float]) -> int:
    """EMU (python-docx's unit) -> half-points, what ``w:sz`` wants."""
    return int(round(float(value) / (EMU_PER_PT / 2)))


def twips(value: Optional[float]) -> int:
    """EMU -> twentieths of a point, what ``w:tw``/``w:tcW`` want."""
    return int(round(float(value) / (EMU_PER_PT / 20)))


def clamp(value: int, low: int, high: int) -> int:
    return max(low, min(high, value))


# ---------------------------------------------------------------- serialising
_FIRST_TAG = re.compile(rb"<[A-Za-z_][A-Za-z0-9_.:-]*(?:\s[^<>]*?)?/?>", re.S)
_XMLNS = re.compile(rb'xmlns:([A-Za-z_][A-Za-z0-9_.-]*)="([^"]*)"')
_XMLNS_ANY = re.compile(rb'\s+xmlns:([A-Za-z_][A-Za-z0-9_.-]*)=("[^"]*"|\'[^\']*\')')


def same_meaning(one: bytes, two: bytes) -> bool:
    """C14N equality with prefixes normalised: is this the same document?"""
    try:
        return (ET.canonicalize(one, with_comments=True, rewrite_prefixes=True)
                == ET.canonicalize(two, with_comments=True,
                                   rewrite_prefixes=True))
    except Exception:                                        # noqa: BLE001
        return one == two


def root_declarations(xml: bytes) -> Dict[str, str]:
    """``xmlns:prefix`` declarations written on the root element."""
    match = _FIRST_TAG.search(xml[:8000])
    if not match:
        return {}
    return {prefix.decode(): uri.decode()
            for prefix, uri in _XMLNS.findall(match.group(0))}


def serialize(root: ET.Element, original: Optional[bytes] = None,
              declaration: bool = True) -> bytes:
    """Serialise a tree the way the file it came from was written.

    ElementTree invents prefixes (``ns0``) for namespaces it has not registered
    and drops declarations it does not use — and ``mc:Ignorable="w14 w15"`` is an
    attribute that *names prefixes*. So: register every OOXML prefix, then put
    back any declaration the original root carried and the output lost.
    """
    if original:
        # Adopt the prefixes the file itself uses *before* serialising, or
        # ElementTree renames anything it has never seen to ns0, ns1 … and
        # mc:Ignorable="w14 w16cex" stops naming anything real.
        for prefix, uri in root_declarations(original).items():
            try:
                ET.register_namespace(prefix, uri)
            except ValueError:            # reserved names like ns0
                continue
    body = ET.tostring(root, encoding="utf-8")
    if original:
        wanted = root_declarations(original)
        if wanted:
            head = _FIRST_TAG.search(body)
            if head:
                tag = head.group(0)
                present = {prefix.decode()
                           for prefix, _ in _XMLNS.findall(tag)}
                extra = "".join(
                    f' xmlns:{prefix}="{uri}"' for prefix, uri in wanted.items()
                    if prefix not in present)
                if extra:
                    fixed = (tag[:-2] + extra + ">" if tag.endswith(b"/>")
                             else tag[:-1] + extra.encode() + b">")
                    body = fixed + body[len(tag):]
    if not declaration:
        return body
    return (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            + body.lstrip())


# -------------------------------------------------------------------- builders
class Builder:
    """Builds OOXML elements in one document's namespace.

    The namespace comes from the document being edited, so an edit to an
    ISO-Strict file stays Strict instead of silently mixing two vocabularies.
    """

    def __init__(self, ns: Optional[Dict[str, str]] = None):
        ns = dict(ns or TRANSITIONAL)
        self.w = ns.get("w", TRANSITIONAL["w"])
        self.r = ns.get("r", TRANSITIONAL["r"])

    # -- primitives
    def tag(self, name: str, prefix: str = "w") -> str:
        uri = self.w if prefix == "w" else self.r if prefix == "r" else PREFIXES[prefix]
        return f"{{{uri}}}{name}"

    def el(self, name: str, parent: Optional[ET.Element] = None, **attrs):
        """``el("w:p")`` -> element; ``el("w:jc", pPr, w_val="center")``."""
        prefix, _, local = name.partition(":")
        element = ET.Element(self.tag(local, prefix=prefix or "w"))
        for key, value in attrs.items():
            if value is None:
                continue
            prefix, _, local = key.partition("_")
            if prefix == "w":
                name = self.tag(local, prefix="w")
            elif prefix == "r":
                name = self.tag(local, prefix="r")
            elif prefix == "xml":
                name = f"{{{XML_NS}}}{local}"
            else:
                name = f"{{{PREFIXES[prefix]}}}{local}"
            element.set(name, token(value))
        if parent is not None:
            parent.append(element)
        return element

    # -- runs
    def run(self, text: str = "", *, bold=None, italic=None, underline=None,
            size=None, color=None, font=None, strike=None, highlight=None):
        r = self.el("w:r")
        rpr = None
        if any(v is not None for v in (bold, italic, underline, size, color,
                                       font, strike, highlight)):
            rpr = self.el("w:rPr", r)
            if font:
                self.el("w:rFonts", rpr, w_ascii=font, w_hAnsi=font,
                        w_eastAsia=font, w_cs=font)
            if bold is not None:
                self.el("w:b", rpr, w_val="1" if bold else "0")
            if italic is not None:
                self.el("w:i", rpr, w_val="1" if italic else "0")
            if underline is not None:
                self.el("w:u", rpr,
                        w_val="single" if underline in (True, "single") else
                        (underline or "none"))
            if strike is not None:
                self.el("w:strike", rpr, w_val="1" if strike else "0")
            if color:
                self.el("w:color", rpr, w_val=color)
            if highlight:
                self.el("w:highlight", rpr, w_val=highlight)
            if size:
                self.el("w:sz", rpr, w_val=half_points(size))
                self.el("w:szCs", rpr, w_val=half_points(size))
        attrs = {"xml_space": "preserve"} if text != text.strip() else {}
        t = self.el("w:t", r, **attrs)
        t.text = text
        return r

    # -- paragraphs
    def paragraph(self, text: Optional[str] = None, *, style=None, align=None,
                  bold=None, italic=None, underline=None, size=None, color=None,
                  font=None, flush=None):
        p = self.el("w:p")
        props = None
        if style or align:
            props = self.el("w:pPr", p)
            if style:
                self.el("w:pStyle", props, w_val=style)
            if align:
                self.el("w:jc", props, w_val=align)
        if text:
            p.append(self.run(text, bold=bold, italic=italic, underline=underline,
                             size=size, color=color, font=font))
        return p

    def drawing(self, embed: str, cx: int, cy: int, *, name: str = "Picture",
                docpr_id: int = 1, descr: str = "",
                filename: str = "image.png") -> ET.Element:
        """A ``w:drawing`` holding one inline picture, in Word's own shape."""
        def tag(ns: str, local: str) -> str:
            return f"{{{ns}}}{local}"

        drawing = ET.Element(self.tag("drawing"))
        inline = ET.SubElement(drawing, tag(WP_NS, "inline"))
        for side in ("distT", "distB", "distL", "distR"):
            inline.set(side, "0")
        extent = ET.SubElement(inline, tag(WP_NS, "extent"))
        extent.set("cx", str(int(cx)))
        extent.set("cy", str(int(cy)))
        effect = ET.SubElement(inline, tag(WP_NS, "effectExtent"))
        for side in ("l", "t", "r", "b"):
            effect.set(side, "0")
        doc_pr = ET.SubElement(inline, tag(WP_NS, "docPr"))
        doc_pr.set("id", str(int(docpr_id)))
        doc_pr.set("name", name)
        doc_pr.set("descr", descr)
        locks = ET.SubElement(inline, tag(WP_NS, "cNvGraphicFramePr"))
        frame_locks = ET.SubElement(locks, tag(DRAWING_NS, "graphicFrameLocks"))
        frame_locks.set("noChangeAspect", "1")

        graphic = ET.SubElement(inline, tag(DRAWING_NS, "graphic"))
        graphic_data = ET.SubElement(graphic, tag(DRAWING_NS, "graphicData"))
        graphic_data.set("uri", PIC_NS)
        pic = ET.SubElement(graphic_data, tag(PIC_NS, "pic"))
        nv = ET.SubElement(pic, tag(PIC_NS, "nvPicPr"))
        c_nv_pr = ET.SubElement(nv, tag(PIC_NS, "cNvPr"))
        c_nv_pr.set("id", str(int(docpr_id)))
        c_nv_pr.set("name", filename)
        ET.SubElement(nv, tag(PIC_NS, "cNvPicPr"))
        fill = ET.SubElement(pic, tag(PIC_NS, "blipFill"))
        blip = ET.SubElement(fill, tag(DRAWING_NS, "blip"))
        blip.set(self.tag("embed", "r"), embed)
        stretch = ET.SubElement(fill, tag(DRAWING_NS, "stretch"))
        ET.SubElement(stretch, tag(DRAWING_NS, "fillRect"))
        shape_props = ET.SubElement(pic, tag(PIC_NS, "spPr"))
        transform = ET.SubElement(shape_props, tag(DRAWING_NS, "xfrm"))
        offset = ET.SubElement(transform, tag(DRAWING_NS, "off"))
        offset.set("x", "0")
        offset.set("y", "0")
        ext = ET.SubElement(transform, tag(DRAWING_NS, "ext"))
        ext.set("cx", str(int(cx)))
        ext.set("cy", str(int(cy)))
        geometry = ET.SubElement(shape_props, tag(DRAWING_NS, "prstGeom"))
        geometry.set("prst", "rect")
        ET.SubElement(geometry, tag(DRAWING_NS, "avLst"))
        return drawing

    def picture_run(self, embed: str, cx: int, cy: int, **kwargs) -> ET.Element:
        """A run containing one inline picture."""
        run = self.el("w:r")
        run.append(self.drawing(embed, cx, cy, **kwargs))
        return run

    def page_break_paragraph(self):
        p = self.el("w:p")
        r = self.el("w:r", p)
        self.el("w:br", r, w_type="page")
        return p

    # -- tables
    def table(self, rows: int, cols: int, *, style: str = "TableGrid",
              widths: Optional[Sequence[float]] = None,
              text: Optional[Iterable[Iterable[str]]] = None) -> ET.Element:
        tbl = self.el("w:tbl")
        props = self.el("w:tblPr", tbl)
        if style:
            self.el("w:tblStyle", props, w_val=style)
        self.el("w:tblW", props, w_w="0", w_type="auto")
        self.el("w:tblLook", props, w_val="04A0", w_firstRow="1",
                w_lastRow="0", w_firstColumn="1", w_lastColumn="0",
                w_noHBand="0", w_noVBand="0")
        grid = self.el("w:tblGrid", tbl)
        for index in range(cols):
            width = twips(widths[index]) if widths else 2400
            self.el("w:gridCol", grid, w_w=width)
        data = [list(row) for row in text] if text else []
        for row_index in range(rows):
            tr = self.el("w:tr", tbl)
            for col_index in range(cols):
                tc = self.el("w:tc", tr)
                tcpr = self.el("w:tcPr", tc)
                self.el("w:tcW", tcpr,
                        w_w=twips(widths[col_index]) if widths else 2400,
                        w_type="dxa")
                cell_text = ""
                if row_index < len(data) and col_index < len(data[row_index]):
                    cell_text = data[row_index][col_index] or ""
                tc.append(self.paragraph(cell_text))
        return tbl


# ------------------------------------------------------------- blank document
_NS = TRANSITIONAL

CONTENT_TYPE_DOCUMENT = ("application/vnd.openxmlformats-officedocument."
                         "wordprocessingml.document.main+xml")
CONTENT_TYPES = {
    "styles": "application/vnd.openxmlformats-officedocument."
              "wordprocessingml.styles+xml",
    "numbering": "application/vnd.openxmlformats-officedocument."
                 "wordprocessingml.numbering+xml",
    "settings": "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.settings+xml",
    "webSettings": "application/vnd.openxmlformats-officedocument."
                   "wordprocessingml.webSettings+xml",
    "fontTable": "application/vnd.openxmlformats-officedocument."
                 "wordprocessingml.fontTable+xml",
    "theme": "application/vnd.openxmlformats-officedocument.theme+xml",
    "comments": "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.comments+xml",
}

REL_TYPES = {
    "styles": "styles", "numbering": "numbering", "settings": "settings",
    "webSettings": "webSettings", "fontTable": "fontTable", "theme": "theme",
    "comments": "comments",
}

# media is referenced by relationship, never by a content-type override
IMAGE_CONTENT_TYPES = {
    "png": "image/png", "jpeg": "image/jpeg", "jpg": "image/jpeg",
    "gif": "image/gif", "bmp": "image/bmp", "tif": "image/tiff",
    "tiff": "image/tiff",
}
HEADER_FOOTER_TYPES = {
    "header": "application/vnd.openxmlformats-officedocument."
              "wordprocessingml.header+xml",
    "footer": "application/vnd.openxmlformats-officedocument."
              "wordprocessingml.footer+xml",
}
DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
WP_NS = ("http://schemas.openxmlformats.org/drawingml/2006/"
         "wordprocessingDrawing")


def _rel_base(strict: bool) -> str:
    """Relationship-type base for the package's own vocabulary."""
    return STRICT["r"] if strict else TRANSITIONAL["r"]


PKG_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def package_rel_base(kind: str, strict: bool = False) -> str:
    """Which vocabulary a relationship declared in ``_rels/.rels`` speaks.

    Both flavours live in the same file: the document is related with the
    *officeDocument* namespace, while core properties and thumbnails use the
    *package* one. Word writes it that way and python-docx reads it that way —
    putting core properties in the officeDocument namespace makes the title
    invisible to python-docx, which then reports its own default instead.
    """
    if kind.startswith("metadata/"):
        return PKG_NS
    return _rel_base(strict)


def _ct_part() -> bytes:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        f'<Default Extension="png" ContentType="image/png"/>'
        f'<Default Extension="jpeg" ContentType="image/jpeg"/>'
        f'<Override PartName="/word/document.xml" ContentType="{CONTENT_TYPE_DOCUMENT}"/>'
        f'<Override PartName="/word/styles.xml" ContentType="{CONTENT_TYPES["styles"]}"/>'
        f'<Override PartName="/word/numbering.xml" ContentType="{CONTENT_TYPES["numbering"]}"/>'
        "</Types>"
    ).encode()


def _root_rels(rel_base: str) -> bytes:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{rel_base}/officeDocument" Target="word/document.xml"/>'
        "</Relationships>"
    ).encode()


def _document_rels(rel_base: str) -> bytes:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{rel_base}/styles" Target="styles.xml"/>'
        f'<Relationship Id="rId2" Type="{rel_base}/numbering" Target="numbering.xml"/>'
        "</Relationships>"
    ).encode()


def styles_xml(ns: Optional[Dict[str, str]] = None) -> bytes:
    """Return the full python-docx 1.2.0 default style part.

    Keeping the complete template style collection gives new documents the
    same built-in styles, defaults, and latent-style metadata as python-docx.
    The packaged XML is static, so loading it adds no runtime dependency. Only
    the core OOXML namespace URIs are rewritten for strict-conformance output.
    """
    ns = ns or _NS
    xml = files("docviews").joinpath("default_styles.xml").read_bytes()
    for prefix in ("w", "r", "wp"):
        old_uri = TRANSITIONAL.get(prefix)
        new_uri = ns.get(prefix, old_uri)
        if old_uri and new_uri and old_uri != new_uri:
            xml = xml.replace(old_uri.encode(), new_uri.encode())
    return xml


_BUILTIN_STYLE_NAMES = {
    "Normal": "Normal", "Title": "Title", "Subtitle": "Subtitle",
    "Quote": "Quote", "IntenseQuote": "Intense Quote",
    "ListBullet": "List Bullet", "ListNumber": "List Number",
    "TableGrid": "Table Grid", "TableNormal": "Normal Table",
    "Emphasis": "Emphasis", "Strong": "Strong",
    "SubtleEmphasis": "Subtle Emphasis",
    "DefaultParagraphFont": "Default Paragraph Font", "NoList": "No List",
    "Header": "Header", "Footer": "Footer",
}


def numbering_xml(ns: Optional[Dict[str, str]] = None) -> bytes:
    """Bullet and decimal lists, three levels each."""
    ns = ns or _NS
    w = ns["w"]
    bullets = [("Symbol", "\u2022"), ("Courier New", "o"), ("Wingdings", "\u00a7")]
    levels = ""
    for level in range(3):
        font, char = bullets[level]
        indent = 720 + 360 * level
        levels += (
            f'<w:lvl w:ilvl="{level}"><w:start w:val="1"/>'
            f'<w:numFmt w:val="bullet"/><w:lvlText w:val="{char}"/>'
            f'<w:lvlJc w:val="left"/>'
            f'<w:pPr><w:ind w:left="{indent}" w:hanging="360"/></w:pPr>'
            f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:hint="default"/>'
            f'</w:rPr></w:lvl>')
    decimal = ""
    for level in range(3):
        indent = 720 + 360 * level
        decimal += (
            f'<w:lvl w:ilvl="{level}"><w:start w:val="1"/>'
            f'<w:numFmt w:val="decimal"/>'
            f'<w:lvlText w:val="%{level + 1}."/>'
            f'<w:lvlJc w:val="left"/>'
            f'<w:pPr><w:ind w:left="{indent}" w:hanging="360"/></w:pPr></w:lvl>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:numbering xmlns:w="{w}">'
        f'<w:abstractNum w:abstractNumId="1">'
        f'<w:multiLevelType w:val="hybridMultilevel"/>{levels}</w:abstractNum>'
        f'<w:abstractNum w:abstractNumId="2">'
        f'<w:multiLevelType w:val="hybridMultilevel"/>{decimal}</w:abstractNum>'
        '<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>'
        '<w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>'
        '</w:numbering>'
    ).encode()


def _sect_pr() -> str:
    # US Letter with 1in / 1.25in margins: the geometry python-docx's own blank
    # document has, so a migrating program that never touches page setup sees
    # the page it saw before.
    return ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
            '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" '
            'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')


def blank_document_xml(ns: Optional[Dict[str, str]] = None) -> bytes:
    ns = ns or _NS
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{ns["w"]}" xmlns:r="{ns["r"]}">'
        f'<w:body>{_sect_pr()}</w:body></w:document>'
    ).encode()


def new_package_bytes(strict: bool = False) -> bytes:
    """A minimal but complete, Word-openable empty .docx."""
    ns = dict(STRICT if strict else TRANSITIONAL)
    rel_base = _rel_base(strict)
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _ct_part())
        zf.writestr("_rels/.rels", _root_rels(rel_base))
        zf.writestr("word/document.xml", blank_document_xml(ns))
        zf.writestr("word/_rels/document.xml.rels", _document_rels(rel_base))
        zf.writestr("word/styles.xml", styles_xml(ns))
        zf.writestr("word/numbering.xml", numbering_xml(ns))
    return buf.getvalue()


# --------------------------------------------------------------- part surgery
def ensure_content_type(xml: bytes, part_name: str, content_type: str) -> bytes:
    root = parse_xml(xml)
    for override in root:
        if override.get("PartName") == part_name:
            if override.get("ContentType") != content_type:
                override.set("ContentType", content_type)
            return serialize(root, xml)
    element = ET.SubElement(root, "{http://schemas.openxmlformats.org/package/2006/content-types}Override")
    element.set("PartName", part_name)
    element.set("ContentType", content_type)
    return serialize(root, xml)


def ensure_default_content_type(xml: bytes, extension: str,
                                content_type: str) -> bytes:
    """Add ``<Default Extension="png" ContentType="image/png"/>`` if missing."""
    root = parse_xml(xml)
    tag = ("{http://schemas.openxmlformats.org/package/2006/content-types}"
           "Default")
    for element in root:
        if element.tag == tag and \
                (element.get("Extension") or "").lower() == extension.lower():
            return xml                              # already declared: leave it
    element = ET.SubElement(root, tag)
    element.set("Extension", extension)
    element.set("ContentType", content_type)
    return serialize(root, xml)


def remove_relationship(xml: bytes, rel_id: str) -> bytes:
    """Remove one relationship id while preserving the rest of its ``.rels``."""
    root = parse_xml(xml)
    removed = False
    for element in list(root):
        if element.get("Id") == rel_id:
            root.remove(element)
            removed = True
    return serialize(root, xml) if removed else xml


def ensure_relationship(xml: bytes, rel_base: str, kind: str, target: str,
                        rel_id: str = "") -> bytes:
    """Add a relationship to a .rels part.

    With a ``rel_id`` the caller has already named it (an image's ``r:embed``
    has to point at it before the package is written), so exactly that id is
    written. Without one an id is allocated, and — since the caller cannot know
    an id yet — a relationship of the same kind already present is reused, the
    way the styles and numbering parts do.
    """
    root = parse_xml(xml)
    wanted = f"{rel_base}/{kind}"
    used = set()
    for rel in root:
        used.add(rel.get("Id", ""))
        if rel.get("Id") == rel_id:
            return serialize(root, xml)          # already there: leave it
        if not rel_id and rel.get("Type") == wanted:
            return serialize(root, xml)
    if rel_id:
        element = ET.SubElement(root, "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship")
        element.set("Id", rel_id)
        element.set("Type", wanted)
        element.set("Target", target)
        return serialize(root, xml)
    index = 1
    while f"rId{index}" in used:
        index += 1
    element = ET.SubElement(root, "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship")
    element.set("Id", f"rId{index}")
    element.set("Type", wanted)
    element.set("Target", target)
    return serialize(root, xml)


def build_package(package: Package, replacements: Dict[str, bytes],
                  additions: Dict[str, bytes],
                  extra_rels: Sequence[Tuple[str, str, str]] = (),
                  extra_content_types: Sequence[Tuple[str, str, bool]] = (),
                  extra_defaults: Sequence[Tuple[str, str]] = (),
                  package_rels: Sequence[Tuple[str, str, str]] = (),
                  removed_rels: Sequence[str] = ()) -> bytes:
    """Copy the original package, replacing only what changed.

    ``replacements`` are parts that already existed; ``additions`` are new ones.
    ``extra_rels`` are ``(relationship id, kind, target)`` triples for the main
    part — the id may be empty, in which case one is allocated — and
    ``package_rels`` are ``(id, kind, target)`` triples for ``_rels/.rels``,
    ``extra_content_types`` are ``(part name, content type, is_override)``
    triples and ``extra_defaults`` are ``(extension, content type)`` pairs —
    stated by the caller, because only the caller knows whether a new part is a
    header, an image, or something else entirely. ``removed_rels`` are ids to
    drop from the main document part (used when a header/footer is unlinked).
    """
    names = list(package.names())
    rel_base = _rel_base(package.strict)
    base_dir = package.main_part.rsplit("/", 1)[0] if "/" in package.main_part \
        else ""
    main_name = package.main_part.rsplit("/", 1)[-1]
    rels_name = (f"{base_dir}/_rels/{main_name}.rels" if base_dir
                 else f"_rels/{main_name}.rels")

    content_types = package._zip.read("[Content_Types].xml")
    rels = None
    if rels_name in names:
        rels = package._zip.read(rels_name)
    else:
        rels = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            "</Relationships>").encode()
    for rel_id in removed_rels:
        rels = remove_relationship(rels, rel_id)

    for extension, content_type in extra_defaults:
        content_types = ensure_default_content_type(content_types, extension,
                                                    content_type)
    reserved_targets = {target.rsplit("/", 1)[-1]
                        for _rel_id, _kind, target in extra_rels}
    for name, data in additions.items():
        tail = name.rsplit("/", 1)[-1]
        stem = tail.split(".")[0]
        stem_key = stem.rstrip("0123456789") or stem
        content_type = CONTENT_TYPES.get(stem_key) or \
            HEADER_FOOTER_TYPES.get(stem_key)
        if content_type:
            content_types = ensure_content_type(content_types, f"/{name}",
                                                content_type)
        kind = REL_TYPES.get(stem_key)
        if kind and "/" not in name[len(base_dir) + 1:] \
                and tail not in reserved_targets:
            rels = ensure_relationship(rels, rel_base, kind, tail)
    for part_name, content_type, _is_override in extra_content_types:
        content_types = ensure_content_type(content_types, part_name,
                                            content_type)
    for rel_id, kind, target in extra_rels:
        rels = ensure_relationship(rels, rel_base, kind, target, rel_id)

    package_rels_xml = None
    if package_rels:
        package_rels_xml = package._zip.read("_rels/.rels")
        for rel_id, kind, target in package_rels:
            package_rels_xml = ensure_relationship(
                package_rels_xml, package_rel_base(kind, package.strict),
                kind, target, rel_id)

    plan = list(names)
    for extra in ("[Content_Types].xml", "_rels/.rels"):
        if extra not in plan:
            plan.insert(0, extra)            # a package always has both

    buf = BytesIO()
    written = set()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as out:
        for name in plan:
            written.add(name)
            if name == "[Content_Types].xml":
                out.writestr(name, content_types)
            elif name.endswith("/"):
                # a directory entry: keep it, so the package's entry list is
                # exactly the one it came with, in the order it came in
                out.writestr(package._zip.getinfo(name), b"")
            elif name == rels_name:
                out.writestr(name, rels)
            elif name == "_rels/.rels" and package_rels_xml is not None:
                out.writestr(name, package_rels_xml)
            elif name in additions:
                out.writestr(name, additions[name])
            elif name in replacements:
                out.writestr(name, replacements[name])
            else:
                out.writestr(name, package._zip.read(name))
        for name, data in additions.items():
            if name not in written:                 # parts the package did not have
                out.writestr(name, data)
    return buf.getvalue()


# ------------------------------------------------------- element-backed proxies
def _local(el: ET.Element) -> str:
    return el.tag.rsplit("}", 1)[-1] if isinstance(el.tag, str) else ""


class Run:
    """One ``w:r``: its text and its character formatting, readable and settable.

    Text is what this run shows in the accepted view; deleted text inside
    ``w:del`` belongs to its own (deleted) run and is not part of this one.
    """

    def __init__(self, element: ET.Element, w_ns: str = TRANSITIONAL["w"],
                 refresh=None, document=None, part_name: str = ""):
        self._element = element
        self._w_ns = w_ns
        self._refresh = refresh
        self._document = document
        self._part_name = part_name

    # ------------------------------------------------------------------ reading
    def _rpr(self, create: bool = False) -> Optional[ET.Element]:
        for child in self._element:
            if _local(child) == "rPr":
                return child
        if not create:
            return None
        rpr = ET.Element(f"{{{self._w_ns}}}rPr")
        if _local(self._element) == "style":
            self._element.append(rpr)      # a w:style keeps w:rPr near the end
        else:
            self._element.insert(0, rpr)
        return rpr

    def _flag(self, name: str) -> Optional[bool]:
        rpr = self._rpr()
        if rpr is None:
            return None
        for child in rpr:
            if _local(child) == name:
                value = child.get(f"{{{self._w_ns}}}val")
                if value is None:
                    return True
                return value not in ("0", "false", "off")
        return None

    def _set_flag(self, name: str, value: Optional[bool]) -> None:
        rpr = self._rpr(create=value is not None)
        if rpr is None:
            return
        for child in list(rpr):
            if _local(child) == name:
                rpr.remove(child)
        if value is None:
            return
        element = ET.Element(f"{{{self._w_ns}}}{name}")
        element.set(f"{{{self._w_ns}}}val", "1" if value else "0")
        rpr.insert(self._placement(rpr, name), element)

    # The w: flags python-docx gives a name to, and the tag each one writes.
    FLAGS = {
        "all_caps": "caps", "small_caps": "smallCaps", "strike": "strike",
        "double_strike": "dstrike", "outline": "outline", "shadow": "shadow",
        "emboss": "emboss", "imprint": "imprint", "hidden": "vanish",
        "web_hidden": "webHidden", "no_proof": "noProof", "rtl": "rtl",
        "complex_script": "cs", "cs_bold": "bCs", "cs_italic": "iCs",
        "math": "oMath", "snap_to_grid": "snapToGrid",
        "spec_vanish": "specVanish",
    }

    def flag(self, name: str) -> Optional[bool]:
        """Read one of :attr:`FLAGS` — ``run.font.all_caps`` and friends."""
        return self._flag(self.FLAGS[name])

    def set_flag(self, name: str, value: Optional[bool]) -> None:
        self._set_flag(self.FLAGS[name], value)

    @property
    def vert_align(self) -> Optional[str]:
        """``"subscript"``, ``"superscript"``, ``"baseline"`` or None."""
        rpr = self._rpr()
        if rpr is None:
            return None
        for child in rpr:
            if _local(child) == "vertAlign":
                return child.get(f"{{{self._w_ns}}}val") or None
        return None

    @vert_align.setter
    def vert_align(self, value) -> None:
        rpr = self._rpr(create=value is not None)
        if rpr is None:
            return
        for child in list(rpr):
            if _local(child) == "vertAlign":
                rpr.remove(child)
        if value is None:
            return
        element = ET.Element(f"{{{self._w_ns}}}vertAlign")
        element.set(f"{{{self._w_ns}}}val", token(value))
        rpr.insert(self._placement(rpr, "vertAlign"), element)

    # --------------------------------------------------------- page furniture
    @property
    def contains_page_break(self) -> bool:
        for child in self._element:
            if _local(child) == "br" and \
                    child.get(f"{{{self._w_ns}}}type") == "page":
                return True
        return False

    def iter_inner_content(self):
        """A run is its own content — python-docx yields the run itself."""
        yield self

    def add_picture(self, image_path_or_stream, width=None, height=None):
        """A picture inside this run, python-docx's ``run.add_picture``."""
        if self._document is None:
            raise DocxError(
                "this run is not attached to a document, so a picture cannot "
                "be given a relationship. Use doc.add_picture(...) instead.")
        return self._document.add_picture(image_path_or_stream, width, height,
                                          run=self)

    # ------------------------------------------------------------------ editing
    def mark_comment_range(self, last_run: Run,
                           comment_id: int) -> None:
        """Mark the runs from this one to ``last_run`` for a comment.

        python-docx's signature, and python-docx's effect: a
        ``w:commentRangeStart`` before this run, and a ``w:commentRangeEnd``
        plus a reference run after the last one.
        """
        if self._document is None:
            raise DocxError(
                "this run is not attached to a document, so a comment range "
                "cannot be written. Use doc.add_comment(runs, text).")
        self._document._mark_comment_range(comment_id, self, last_run)

    def add_text(self, text: str) -> None:
        """Append characters to this run (python-docx's ``add_text``)."""
        element = ET.SubElement(self._element, f"{{{self._w_ns}}}t")
        element.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        element.text = text

    def add_tab(self) -> None:
        ET.SubElement(self._element, f"{{{self._w_ns}}}tab")

    def add_break(self, break_type=None, *, kind=None) -> None:
        """Add a line break; ``WD_BREAK.PAGE`` writes a page break instead.

        ``break_type`` matches python-docx's keyword. ``kind`` remains as a
        backwards-compatible alias for early docviews callers.
        """
        if kind is not None:
            if break_type is not None:
                raise TypeError("pass either break_type or kind, not both")
            break_type = kind
        names_by_value = {
            6: "LINE", 7: "PAGE", 8: "COLUMN", 9: "LINE_CLEAR_LEFT",
            10: "LINE_CLEAR_RIGHT", 11: "LINE_CLEAR_ALL",
        }
        names_by_token = {
            "TEXTWRAPPING": "LINE", "LINE": "LINE", "PAGE": "PAGE",
            "COLUMN": "COLUMN",
            "LINECLEARLEFT": "LINE_CLEAR_LEFT",
            "LINECLEARRIGHT": "LINE_CLEAR_RIGHT",
            "LINECLEARALL": "LINE_CLEAR_ALL",
        }
        if break_type is None:
            name = "LINE"
        else:
            name = getattr(break_type, "name", None)
            if name is None:
                value = getattr(break_type, "value", break_type)
                try:
                    name = names_by_value.get(int(value))
                except (TypeError, ValueError):
                    name = None
                if name is None:
                    name = names_by_token.get(str(value).replace("_", "").upper())
        break_specs = {
            "LINE": (None, None),
            "PAGE": ("page", None),
            "COLUMN": ("column", None),
            "LINE_CLEAR_LEFT": (None, "left"),
            "LINE_CLEAR_RIGHT": (None, "right"),
            "LINE_CLEAR_ALL": (None, "all"),
        }
        if name not in break_specs:
            raise KeyError(break_type)
        break_type_name, clear = break_specs[name]
        element = ET.SubElement(self._element, f"{{{self._w_ns}}}br")
        if break_type_name is not None:
            element.set(f"{{{self._w_ns}}}type", break_type_name)
        if clear is not None:
            element.set(f"{{{self._w_ns}}}clear", clear)

    def clear(self) -> None:
        """Take this run back to nothing: no text, no formatting."""
        for child in list(self._element):
            self._element.remove(child)

    def __repr__(self) -> str:
        return f"Run({self.text!r})"

    @staticmethod
    def _placement(rpr: ET.Element, name: str) -> int:
        order = ("rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps",
                 "strike", "dstrike", "outline", "shadow", "emboss", "imprint",
                 "noProof", "snapToGrid", "vanish", "webHidden", "color",
                 "spacing", "w", "kern", "position", "sz", "szCs", "highlight",
                 "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl",
                 "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath")
        try:
            wanted = order.index(name)
        except ValueError:
            return len(rpr)
        for index, child in enumerate(rpr):
            try:
                if order.index(_local(child)) > wanted:
                    return index
            except ValueError:
                continue
        return len(rpr)

    @property
    def text(self) -> str:
        from .views import views_of                # function-level: no cycle
        return views_of(self._element).get("accepted") or ""

    @text.setter
    def text(self, value: str) -> None:
        for child in list(self._element):
            if _local(child) in ("t", "delText", "br", "cr", "tab", "noBreakHyphen",
                                 "ptab"):
                self._element.remove(child)
        rpr = self._rpr()
        index = list(self._element).index(rpr) + 1 if rpr is not None else 0
        attrs = ({"{http://www.w3.org/XML/1998/namespace}space": "preserve"}
                 if value != value.strip() else {})
        element = ET.Element(f"{{{self._w_ns}}}t", attrs)
        element.text = value
        self._element.insert(index, element)
        if self._refresh:
            self._refresh()

    @property
    def bold(self) -> Optional[bool]:
        return self._flag("b")

    @bold.setter
    def bold(self, value: Optional[bool]) -> None:
        self._set_flag("b", value)

    @property
    def italic(self) -> Optional[bool]:
        return self._flag("i")

    @italic.setter
    def italic(self, value: Optional[bool]) -> None:
        self._set_flag("i", value)

    @property
    def underline(self):
        rpr = self._rpr()
        if rpr is None:
            return None
        for child in rpr:
            if _local(child) == "u":
                value = child.get(f"{{{self._w_ns}}}val")
                if value in (None, "single"):
                    return True
                if value == "none":
                    return False
                return value
        return None

    @underline.setter
    def underline(self, value) -> None:
        rpr = self._rpr(create=value is not None)
        if rpr is None:
            return
        for child in list(rpr):
            if _local(child) == "u":
                rpr.remove(child)
        if value is None:
            return
        element = ET.Element(f"{{{self._w_ns}}}u")
        if value is True:
            underline = "single"
        elif value is False:
            underline = "none"
        else:
            underline = token(value)
            if underline is None and hasattr(value, "xml_value"):
                # Match BaseXmlEnum.to_xml(): INHERITED has no representation
                # and must not degrade into an invalid numeric XML value.
                underline = type(value).to_xml(value)
        element.set(f"{{{self._w_ns}}}val", underline)
        rpr.insert(self._placement(rpr, "u"), element)

    @property
    def size(self) -> Optional[int]:
        """Font size in EMU (12700 per point), like python-docx's Pt()."""
        rpr = self._rpr()
        if rpr is None:
            return None
        for child in rpr:
            if _local(child) == "sz":
                half = child.get(f"{{{self._w_ns}}}val")
                return Length(int(half) * 6350) if half else None
        return None

    @size.setter
    def size(self, value) -> None:
        rpr = self._rpr(create=value is not None)
        if rpr is None:
            return
        for child in list(rpr):
            if _local(child) in ("sz", "szCs"):
                rpr.remove(child)
        if value is None:
            return
        for name in ("sz", "szCs"):
            element = ET.Element(f"{{{self._w_ns}}}{name}")
            element.set(f"{{{self._w_ns}}}val", str(half_points(value)))
            rpr.append(element)

    @property
    def color(self) -> Optional[str]:
        """Hex RGB string, e.g. ``"FF0000"``."""
        rpr = self._rpr()
        if rpr is None:
            return None
        for child in rpr:
            if _local(child) == "color":
                return child.get(f"{{{self._w_ns}}}val")
        return None

    @color.setter
    def color(self, value) -> None:
        text = str(value) if value is not None else None
        rpr = self._rpr(create=text is not None)
        if rpr is None:
            return
        for child in list(rpr):
            if _local(child) == "color":
                rpr.remove(child)
        if text:
            element = ET.Element(f"{{{self._w_ns}}}color")
            element.set(f"{{{self._w_ns}}}val", text.lstrip("#").upper())
            rpr.append(element)

    @property
    def font_name(self) -> Optional[str]:
        rpr = self._rpr()
        if rpr is None:
            return None
        for child in rpr:
            if _local(child) == "rFonts":
                return child.get(f"{{{self._w_ns}}}ascii")
        return None

    @font_name.setter
    def font_name(self, value: Optional[str]) -> None:
        rpr = self._rpr(create=bool(value))
        if rpr is None:
            return
        for child in list(rpr):
            if _local(child) == "rFonts":
                rpr.remove(child)
        if value:
            element = ET.Element(f"{{{self._w_ns}}}rFonts")
            for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
                element.set(f"{{{self._w_ns}}}{attr}", value)
            rpr.insert(0, element)

    @property
    def style(self):
        """The character style (``w:rStyle``) this run sets, or None."""
        from .views import StyleRef
        rpr = self._rpr()
        if rpr is None:
            return None
        for child in rpr:
            if _local(child) == "rStyle":
                value = child.get(f"{{{self._w_ns}}}val") or None
                return StyleRef(value) if value else None
        return None

    @style.setter
    def style(self, value) -> None:
        style_id = (getattr(value, "style_id", None) if value is not None
                    else None)
        name = (None if value is None else
                str(style_id or getattr(value, "name", value) or ""))
        rpr = self._rpr(create=name is not None)
        if rpr is None:
            return
        for child in list(rpr):
            if _local(child) == "rStyle":
                rpr.remove(child)
        if name:
            element = ET.Element(f"{{{self._w_ns}}}rStyle")
            element.set(f"{{{self._w_ns}}}val", name.replace(" ", ""))
            rpr.insert(0, element)

    @property
    def font(self) -> "Font":
        """Character formatting, python-docx style: ``run.font.bold = True``."""
        return Font(self)

    def __repr__(self) -> str:
        return f"Run({self.text[:40]!r})"


class Rgb:
    """Deprecated value wrapper retained for source compatibility."""

    def __init__(self, hex_digits: str):
        self.rgb = hex_digits

    def __str__(self) -> str:
        return self.rgb

    def __repr__(self) -> str:
        return f"Rgb({self.rgb})"


class ColorFormat:
    """The python-docx-shaped color accessor: RGB, theme, or inherited.

    The proxy is returned even when a run has no direct color. Consequently
    ``run.font.color.rgb`` safely returns ``None`` for an unformatted run, as it
    does in python-docx, while setting ``rgb`` or ``theme_color`` edits the
    underlying ``w:rPr/w:color`` element.
    """

    def __init__(self, run: "Run"):
        self._run = run

    def _element(self, create: bool = False) -> Optional[ET.Element]:
        rpr = self._run._rpr(create=create)
        if rpr is None:
            return None
        for child in rpr:
            if _local(child) == "color":
                return child
        if not create:
            return None
        element = ET.Element(f"{{{self._run._w_ns}}}color")
        rpr.insert(self._run._placement(rpr, "color"), element)
        return element

    @property
    def rgb(self):
        element = self._element()
        if element is None:
            return None
        value = element.get(f"{{{self._run._w_ns}}}val")
        if not value or value.lower() == "auto":
            return None
        try:
            raw = bytes.fromhex(value)
        except ValueError as exc:
            raise ValueError(f"invalid RGB color value {value!r}") from exc
        if len(raw) != 3:
            raise ValueError(f"RGB color must be six hex digits, got {value!r}")
        return RGBColor(raw[0], raw[1], raw[2])

    @rgb.setter
    def rgb(self, value) -> None:
        element = self._element()
        if value is None:
            if element is not None:
                self._run._rpr().remove(element)
            return
        raw = getattr(value, "rgb", value)
        text = str(raw).lstrip("#").upper()
        try:
            decoded = bytes.fromhex(text)
        except ValueError as exc:
            raise ValueError(f"invalid RGB color value {text!r}") from exc
        if len(decoded) != 3:
            raise ValueError(f"RGB color must be six hex digits, got {text!r}")
        element = self._element(create=True)
        element.set(f"{{{self._run._w_ns}}}val", text)
        element.attrib.pop(f"{{{self._run._w_ns}}}themeColor", None)
        element.attrib.pop(f"{{{self._run._w_ns}}}themeTint", None)
        element.attrib.pop(f"{{{self._run._w_ns}}}themeShade", None)

    @property
    def theme_color(self):
        element = self._element()
        if element is None:
            return None
        value = element.get(f"{{{self._run._w_ns}}}themeColor")
        if value is None:
            return None
        from .compat.enum import MSO_THEME_COLOR_INDEX
        return MSO_THEME_COLOR_INDEX.from_xml(value)

    @theme_color.setter
    def theme_color(self, value) -> None:
        element = self._element()
        if value is None:
            if element is not None:
                self._run._rpr().remove(element)
            return
        xml_value = token(value)
        if xml_value is None and hasattr(value, "xml_value"):
            xml_value = type(value).to_xml(value)
        element = self._element(create=True)
        element.set(f"{{{self._run._w_ns}}}themeColor", xml_value)

    @property
    def type(self):
        element = self._element()
        if element is None:
            return None
        from .compat.enum import MSO_COLOR_TYPE
        if element.get(f"{{{self._run._w_ns}}}themeColor") is not None:
            return MSO_COLOR_TYPE.THEME
        if (element.get(f"{{{self._run._w_ns}}}val") or "").lower() == "auto":
            return MSO_COLOR_TYPE.AUTO
        return MSO_COLOR_TYPE.RGB

    def __repr__(self) -> str:
        return f"ColorFormat(rgb={self.rgb!r}, theme_color={self.theme_color!r})"


class Font:
    """``run.font`` — name, bold, italic, underline, size, colour, and the rest.

    Every ``w:`` character flag python-docx names has a property here, generated
    from :attr:`Run.FLAGS` so the two stay in step: ``font.all_caps``,
    ``font.small_caps``, ``font.strike``, ``font.subscript``, ``font.superscript``,
    ``font.hidden``, ``font.rtl``, ``font.emboss``, …
    """

    def __init__(self, run: "Run"):
        self._run = run

    def _flag_property(self, name: str):
        return property(lambda self: self._run.flag(name),
                        lambda self, value: self._run.set_flag(name, value))

    @property
    def subscript(self) -> Optional[bool]:
        value = self._run.vert_align
        return None if value is None else value == "subscript"

    @subscript.setter
    def subscript(self, value) -> None:
        if value is None:
            self._run.vert_align = None
        elif value:
            self._run.vert_align = "subscript"
        elif self._run.vert_align == "subscript":
            self._run.vert_align = None

    @property
    def superscript(self) -> Optional[bool]:
        value = self._run.vert_align
        return None if value is None else value == "superscript"

    @superscript.setter
    def superscript(self, value) -> None:
        if value is None:
            self._run.vert_align = None
        elif value:
            self._run.vert_align = "superscript"
        elif self._run.vert_align == "superscript":
            self._run.vert_align = None

    name = property(lambda self: self._run.font_name,
                    lambda self, v: setattr(self._run, "font_name", v))
    bold = property(lambda self: self._run.bold,
                    lambda self, v: setattr(self._run, "bold", v))
    italic = property(lambda self: self._run.italic,
                      lambda self, v: setattr(self._run, "italic", v))
    underline = property(lambda self: self._run.underline,
                         lambda self, v: setattr(self._run, "underline", v))
    size = property(lambda self: self._run.size,
                    lambda self, v: setattr(self._run, "size", v))

    @property
    def color(self) -> ColorFormat:
        return ColorFormat(self._run)

    @color.setter
    def color(self, value) -> None:
        # Assignment is a convenience beyond python-docx's API; the documented
        # interface is ``run.font.color.rgb = RGBColor(...)``.
        ColorFormat(self._run).rgb = value

    @property
    def highlight_color(self):
        """Read and write python-docx's ``font.highlight_color``."""
        rpr = self._run._rpr()
        if rpr is not None:
            for child in rpr:
                if _local(child) == "highlight":
                    return child.get(f"{{{self._run._w_ns}}}val")
        return None

    @highlight_color.setter
    def highlight_color(self, value) -> None:
        text = token(value)
        if text is None and value is not None and hasattr(value, "xml_value"):
            text = type(value).to_xml(value)
        rpr = self._run._rpr(create=text is not None)
        if rpr is None:
            return
        for child in list(rpr):
            if _local(child) == "highlight":
                rpr.remove(child)
        if text:
            element = ET.Element(f"{{{self._run._w_ns}}}highlight")
            element.set(f"{{{self._run._w_ns}}}val", text)
            rpr.insert(self._run._placement(rpr, "highlight"), element)


for _flag_name in Run.FLAGS:
    setattr(Font, _flag_name, Font._flag_property(Font, _flag_name))


class ParagraphFormat:
    """``paragraph.paragraph_format`` — the ``w:pPr`` a paragraph carries.

    Lengths are EMU (``Pt(6)``, ``Inches(0.5)``); ``line_spacing`` is a multiple
    of single spacing when given as a float, or a length when given as EMU —
    which is python-docx's rule, and the one Word uses.
    """

    # python-docx names -> the w: tag each one reads and writes
    FLAGS = {"keep_together": "keepLines", "keep_with_next": "keepNext",
             "page_break_before": "pageBreakBefore", "widow_control": "widowControl"}
    LENGTHS = {"left_indent": ("ind", "left"), "right_indent": ("ind", "right")}

    def __init__(self, element: ET.Element, w_ns: str = TRANSITIONAL["w"]):
        self._element = element
        self._w_ns = w_ns

    # ------------------------------------------------------------------ plumbing
    def _ppr(self, create: bool = False) -> Optional[ET.Element]:
        for child in self._element:
            if _local(child) == "pPr":
                return child
        if not create:
            return None
        ppr = ET.Element(f"{{{self._w_ns}}}pPr")
        if _local(self._element) == "style":
            # w:pPr comes after the name/basedOn/… block inside w:style
            self._element.append(ppr)
        else:
            self._element.insert(0, ppr)
        return ppr

    # the order Word's schema wants these children in
    ORDER = ("pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
             "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd",
             "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
             "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN",
             "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind",
             "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
             "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl",
             "divId", "cnfStyle", "rPr", "sectPr", "pPrChange")

    def _holder(self, name: str, create: bool = False) -> Optional[ET.Element]:
        """The ``w:pPr/w:<name>`` element (``w:spacing``, ``w:ind``, …)."""
        ppr = self._ppr(create=create)
        if ppr is None:
            return None
        for child in ppr:
            if _local(child) == name:
                return child
        if not create:
            return None
        element = ET.Element(f"{{{self._w_ns}}}{name}")
        try:
            position = self.ORDER.index(name)
        except ValueError:
            position = len(ppr)
        index = len(ppr)
        for existing_index, child in enumerate(ppr):
            try:
                if self.ORDER.index(_local(child)) > position:
                    index = existing_index
                    break
            except ValueError:
                continue
        ppr.insert(index, element)
        return element

    def _set_child(self, name: str, attribute: str, value) -> None:
        """Set one attribute of ``w:pPr/w:<name>``; drop the element if empty."""
        holder = self._holder(name, create=value is not None)
        if holder is None:
            return
        if value is None:
            holder.attrib.pop(f"{{{self._w_ns}}}{attribute}", None)
        else:
            holder.set(f"{{{self._w_ns}}}{attribute}", token(value))
        if not holder.attrib:
            self._ppr().remove(holder)

    def _get_child(self, name: str, attribute: str) -> Optional[str]:
        holder = self._holder(name)
        return None if holder is None else holder.get(f"{{{self._w_ns}}}{attribute}")

    # ------------------------------------------------------------------ flags
    def _flag_property(self, name: str):
        tag = self.FLAGS[name]

        def get(self):
            ppr = self._ppr()
            if ppr is None:
                return None
            for child in ppr:
                if _local(child) == tag:
                    value = child.get(f"{{{self._w_ns}}}val")
                    if value is None:
                        return True
                    return value not in ("0", "false", "off")
            return None

        def set_(self, value):
            ppr = self._ppr(create=value is not None)
            if ppr is None:
                return
            for child in list(ppr):
                if _local(child) == tag:
                    ppr.remove(child)
            if value is None:
                return
            element = ET.Element(f"{{{self._w_ns}}}{tag}")
            element.set(f"{{{self._w_ns}}}val", "1" if value else "0")
            ppr.append(element)

        return property(get, set_)

    # ------------------------------------------------------------------ lengths
    def _length_property(self, name: str):
        parent, attribute = self.LENGTHS[name]

        def get(self):
            value = self._get_child(parent, attribute)
            return Length(int(value) * 635) if value else None

        def set_(self, value):
            # ``w:left`` and ``w:hanging`` are independent OOXML values. A
            # negative left indent is legal and assigning it must not clear a
            # separately-set first-line/hanging indent.
            twips_value = None if value is None else twips(value)
            self._set_child(parent, attribute, twips_value)

        return property(get, set_)

    # ------------------------------------------------------------------ spacing
    @property
    def space_before(self) -> Optional[int]:
        value = self._get_child("spacing", "before")
        return Length(int(value) * 635) if value else None

    @space_before.setter
    def space_before(self, value) -> None:
        self._set_child("spacing", "before",
                        None if value is None else twips(value))

    @property
    def space_after(self) -> Optional[int]:
        value = self._get_child("spacing", "after")
        return Length(int(value) * 635) if value else None

    @space_after.setter
    def space_after(self, value) -> None:
        self._set_child("spacing", "after",
                        None if value is None else twips(value))

    @property
    def line_spacing(self):
        """A float multiple of single spacing, or a length when in EMU."""
        line = self._get_child("spacing", "line")
        if not line:
            return None
        rule = self._get_child("spacing", "lineRule") or "auto"
        if rule in ("auto", "atLeast"):
            return int(line) / 240
        return Length(int(line) * 635)

    @line_spacing.setter
    def line_spacing(self, value) -> None:
        if value is None:
            self._set_child("spacing", "line", None)
            self._set_child("spacing", "lineRule", None)
            return
        if isinstance(value, float):
            self._set_child("spacing", "line", int(round(value * 240)))
            self._set_child("spacing", "lineRule", "auto")
        else:
            self._set_child("spacing", "line",
                            int(round(int(value) / 635)))
            self._set_child("spacing", "lineRule", "atLeast")

    @property
    def first_line_indent(self) -> Optional[int]:
        for attribute in ("firstLine", "hanging"):
            value = self._get_child("ind", attribute)
            if value:
                return Length(int(value) * 635 * (-1 if attribute == "hanging" else 1))
        return None

    @first_line_indent.setter
    def first_line_indent(self, value) -> None:
        self._set_child("ind", "firstLine", None)
        self._set_child("ind", "hanging", None)
        if value is None:
            return
        if value < 0:
            self._set_child("ind", "hanging", str(twips(-value)))
        else:
            self._set_child("ind", "firstLine", str(twips(value)))

    @property
    def alignment(self):
        return self._get_child("jc", "val")

    @alignment.setter
    def alignment(self, value) -> None:
        self._set_child("jc", "val", token(value))

    def __repr__(self) -> str:
        parts = []
        for name in ("alignment", "space_before", "space_after", "line_spacing",
                     "left_indent", "first_line_indent"):
            value = getattr(self, name)
            if value is not None:
                parts.append(f"{name}={value!r}")
        return f"ParagraphFormat({', '.join(parts)})"


for _name in ParagraphFormat.FLAGS:
    setattr(ParagraphFormat, _name, ParagraphFormat._flag_property(ParagraphFormat, _name))
for _name in ParagraphFormat.LENGTHS:
    setattr(ParagraphFormat, _name, ParagraphFormat._length_property(ParagraphFormat, _name))


class Row:
    def __init__(self, element: ET.Element, w_ns: str):
        self._element, self._w_ns = element, w_ns
        self._table: Optional["Table"] = None

    @property
    def cells(self) -> List["Cell"]:
        """Cells in grid order.

        A ``w:gridSpan`` cell occupies several grid columns; python-docx reports
        it once per column it covers, and so does this — otherwise a merged
        table's ``row.cells[1]`` would address the wrong column.
        """
        out: List[Cell] = []
        for child in self._element:
            if _local(child) != "tc":
                continue
            cell = Cell(child, self._w_ns)
            cell._table_element = getattr(self, "_table_element", None)
            cell._row_element = self._element
            span = cell.grid_span
            out.extend([cell] * span if span > 1 else [cell])
        return out

    def _elements(self) -> List[ET.Element]:
        return [child for child in self._element if _local(child) == "tc"]

    # --------------------------------------------------- geometry, like Word
    @property
    def table(self) -> Optional["Table"]:
        """The table this row belongs to, when it is known."""
        return self._table

    @property
    def height(self) -> Optional[int]:
        """The declared row height in EMU, or None when the row has none."""
        props = self._trpr()
        if props is None:
            return None
        for child in props:
            if _local(child) == "trHeight":
                value = child.get(f"{{{self._w_ns}}}val")
                return Length(int(value) * 635) if value else None
        return None

    @height.setter
    def height(self, value) -> None:
        self._set_trprop("trHeight", "val",
                         None if value is None else twips(value))
        if value is None and self.height_rule is None:
            return

    @property
    def height_rule(self) -> Optional[str]:
        """``"auto"``, ``"atLeast"`` or ``"exact"`` (the ``w:hRule`` tokens)."""
        props = self._trpr()
        if props is None:
            return None
        for child in props:
            if _local(child) == "trHeight":
                return child.get(f"{{{self._w_ns}}}hRule")
        return None

    @height_rule.setter
    def height_rule(self, value) -> None:
        self._set_trprop("trHeight", "hRule", value)

    def _trpr(self, create: bool = False) -> Optional[ET.Element]:
        for child in self._element:
            if _local(child) == "trPr":
                return child
        if not create:
            return None
        props = ET.Element(f"{{{self._w_ns}}}trPr")
        self._element.insert(0, props)
        return props

    def _set_trprop(self, name: str, attribute: str, value) -> None:
        props = self._trpr(create=value is not None)
        if props is None:
            return
        holder = None
        for child in props:
            if _local(child) == name:
                holder = child
        if holder is None:
            if value is None:
                return
            holder = ET.SubElement(props, f"{{{self._w_ns}}}{name}")
        holder.set(f"{{{self._w_ns}}}{attribute}",
                   "" if value is None else token(value))
        if value is None:
            holder.attrib.pop(f"{{{self._w_ns}}}{attribute}", None)
        if not holder.attrib:
            props.remove(holder)

    @property
    def grid_cols_before(self) -> int:
        """Count of unpopulated layout-grid cells before this row's first cell."""
        props = self._trpr()
        if props is None:
            return 0
        for child in props:
            if _local(child) == "gridBefore":
                return int(child.get(f"{{{self._w_ns}}}val") or 0)
        return 0

    @property
    def grid_cols_after(self) -> int:
        """Count of unpopulated layout-grid cells after this row's last cell."""
        props = self._trpr()
        if props is None:
            return 0
        for child in props:
            if _local(child) == "gridAfter":
                return int(child.get(f"{{{self._w_ns}}}val") or 0)
        return 0


class Cell:
    """One ``w:tc``. ``text`` reads like python-docx's (paragraphs joined by
    newlines) and sets the same way: first paragraph keeps the text, the rest go.

    ``merge`` is implemented the way Word records it: a horizontal merge writes
    ``w:gridSpan`` on the first cell, a vertical one writes ``w:vMerge``
    (``restart`` on top, an empty value below), and the text of every cell that
    disappears is kept, in order, in the surviving cell.
    """

    def __init__(self, element: ET.Element, w_ns: str):
        self._element, self._w_ns = element, w_ns

    # ------------------------------------------------------------------ props
    def _tcpr(self, create: bool = False) -> Optional[ET.Element]:
        for child in self._element:
            if _local(child) == "tcPr":
                return child
        if not create:
            return None
        props = ET.Element(f"{{{self._w_ns}}}tcPr")
        self._element.insert(0, props)
        return props

    @property
    def grid_span(self) -> int:
        """How many grid columns this cell covers (1 unless merged)."""
        props = self._tcpr()
        if props is None:
            return 1
        for child in props:
            if _local(child) == "gridSpan":
                try:
                    return max(1, int(child.get(f"{{{self._w_ns}}}val") or 1))
                except ValueError:
                    return 1
        return 1

    @property
    def vertical_merge(self) -> str:
        """``""``, ``"restart"`` (top of a vertical merge) or ``"continue"``."""
        props = self._tcpr()
        if props is None:
            return ""
        for child in props:
            if _local(child) == "vMerge":
                return child.get(f"{{{self._w_ns}}}val") or "continue"
        return ""

    def _set(self, name: str, value: Optional[str]) -> None:
        props = self._tcpr(create=value is not None)
        if props is None:
            return
        for child in list(props):
            if _local(child) == name:
                props.remove(child)
        if value is None:
            return
        element = ET.Element(f"{{{self._w_ns}}}{name}")
        if value != "":
            element.set(f"{{{self._w_ns}}}val", value)
        props.append(element)

    def _table(self) -> Optional[ET.Element]:
        return getattr(self, "_table_element", None)

    def _row(self) -> Optional[ET.Element]:
        return getattr(self, "_row_element", None)

    # ----------------------------------------------------------------- merging
    def merge(self, other: "Cell") -> "Cell":
        """Merge a rectangular region, rejecting cells from another table."""
        if not isinstance(other, Cell):
            raise TypeError("other must be a Cell")
        if self._table() is None or other._table() is not self._table():
            raise ValueError("cells can only be merged inside one table")
        return self._merge(other)

    def _merge(self, other: "Cell", *, allow_cross_table: bool = False) -> "Cell":
        """Implementation shared with the python-docx compatibility proxy."""
        from .exceptions import InvalidSpanError

        if not isinstance(other, Cell):
            raise TypeError("other must be a Cell")
        table = self._table()
        if table is None or other._table() is None:
            raise InvalidSpanError("requested span not rectangular")
        if other._table() is not table and not allow_cross_table:
            raise ValueError("cells can only be merged inside one table")
        rows = [element for element in table if _local(element) == "tr"]

        def table_rows(table_element: ET.Element) -> list[ET.Element]:
            return [element for element in table_element
                    if _local(element) == "tr"]

        def row_index(cell: "Cell") -> int:
            cell_rows = table_rows(cell._table())
            row = cell._row()
            if row not in cell_rows:
                raise InvalidSpanError("requested span not rectangular")
            return cell_rows.index(row)

        def cell_at(row_element: ET.Element, grid_offset: int,
                    table_element: ET.Element) -> Optional["Cell"]:
            row = Row(row_element, self._w_ns)
            offset = row.grid_cols_before
            for element in row_element:
                if _local(element) != "tc":
                    continue
                cell = Cell(element, self._w_ns)
                cell._row_element = row_element
                cell._table_element = table_element
                if offset == grid_offset:
                    return cell
                offset += cell.grid_span
            return None

        def bounds(cell: "Cell") -> tuple[int, int, int, int]:
            cell_table = cell._table()
            cell_rows = table_rows(cell_table)
            row_number = row_index(cell)
            left = cell._column_index()
            right = left + cell.grid_span
            top, bottom = row_number, row_number + 1

            if cell.vertical_merge == "continue":
                for above_index in range(row_number - 1, -1, -1):
                    above = cell_at(cell_rows[above_index], left, cell_table)
                    if above is None or above._column_index() != left \
                            or above.grid_span != cell.grid_span:
                        break
                    if above.vertical_merge == "restart":
                        top = above_index
                        break
                    if above.vertical_merge != "continue":
                        break
            if cell.vertical_merge in ("restart", "continue"):
                for below_index in range(row_number + 1, len(cell_rows)):
                    below = cell_at(cell_rows[below_index], left, cell_table)
                    if below is None or below._column_index() != left \
                            or below.grid_span != cell.grid_span \
                            or below.vertical_merge != "continue":
                        break
                    bottom = below_index + 1
            return top, left, bottom, right

        a_top, a_left, a_bottom, a_right = bounds(self)
        b_top, b_left, b_bottom, b_right = bounds(other)
        if a_top == b_top and a_bottom != b_bottom:
            raise InvalidSpanError("requested span not rectangular")
        if a_left == b_left and a_right != b_right:
            raise InvalidSpanError("requested span not rectangular")

        if a_top < b_top and a_bottom > b_bottom \
                or b_top < a_top and b_bottom > a_bottom:
            raise InvalidSpanError("requested span not rectangular")
        if a_left < b_left and a_right > b_right \
                or b_left < a_left and b_right > a_right:
            raise InvalidSpanError("requested span not rectangular")

        top = min(a_top, b_top)
        left = min(a_left, b_left)
        bottom = max(a_bottom, b_bottom)
        right = max(a_right, b_right)
        height, width = bottom - top, right - left
        if height < 1 or width < 1:
            raise InvalidSpanError("requested span not rectangular")

        top_cell = cell_at(rows[top], left, table)
        if top_cell is None:
            raise InvalidSpanError("not enough grid columns")

        # Grow one row at a time. This is the same order used by python-docx:
        # each row's content is collected left-to-right, then rows top-to-bottom.
        for row_number in range(top, bottom):
            row_element = rows[row_number]
            row_cell = cell_at(row_element, left, table)
            if row_cell is None:
                raise InvalidSpanError("not enough grid columns")
            direct_cells = [element for element in row_element
                            if _local(element) == "tc"]
            try:
                cell_index = direct_cells.index(row_cell._element)
            except ValueError as exc:
                raise InvalidSpanError("requested span not rectangular") from exc

            row_span = row_cell.grid_span
            if row_span > width:
                raise InvalidSpanError("span is not rectangular")
            if row_number != top:
                row_cell._move_content_to(top_cell)
            while row_span < width:
                next_index = cell_index + 1
                if next_index >= len(direct_cells):
                    raise InvalidSpanError("not enough grid columns")
                next_cell = Cell(direct_cells[next_index], self._w_ns)
                next_cell._row_element = row_element
                next_cell._table_element = table
                if row_span + next_cell.grid_span > width:
                    raise InvalidSpanError("span is not rectangular")

                if next_cell.width and row_cell.width:
                    row_cell.width = row_cell.width + next_cell.width
                next_cell._move_content_to(top_cell)
                row_element.remove(next_cell._element)
                direct_cells.pop(next_index)
                row_span += next_cell.grid_span

            row_cell._set("gridSpan", str(width) if width > 1 else None)
            row_cell._set("vMerge", "restart" if row_number == top and height > 1
                          else "" if height > 1 else None)

        top_cell._row_element = rows[top]
        top_cell._table_element = table
        return top_cell

    @property
    def width(self) -> Optional[int]:
        """The cell's declared width in EMU (``w:tcW``), or None."""
        props = self._tcpr()
        if props is None:
            return None
        for child in props:
            if _local(child) == "tcW":
                value = child.get(f"{{{self._w_ns}}}w")
                if child.get(f"{{{self._w_ns}}}type") == "dxa" and value:
                    return Length(int(value) * 635)
        return None

    @width.setter
    def width(self, value) -> None:
        props = self._tcpr(create=value is not None)
        if props is None:
            return
        for child in list(props):
            if _local(child) == "tcW":
                props.remove(child)
        if value is None:
            return
        element = ET.Element(f"{{{self._w_ns}}}tcW")
        element.set(f"{{{self._w_ns}}}w", str(twips(value)))
        element.set(f"{{{self._w_ns}}}type", "dxa")
        props.insert(0, element)

    @property
    def vertical_alignment(self):
        """``"top"``, ``"center"`` or ``"bottom"``, as python-docx reports it."""
        props = self._tcpr()
        if props is None:
            return None
        for child in props:
            if _local(child) == "vAlign":
                return child.get(f"{{{self._w_ns}}}val")
        return None

    @vertical_alignment.setter
    def vertical_alignment(self, value) -> None:
        text = token(value)
        props = self._tcpr(create=text is not None)
        if props is None:
            return
        for child in list(props):
            if _local(child) == "vAlign":
                props.remove(child)
        if text:
            element = ET.Element(f"{{{self._w_ns}}}vAlign")
            element.set(f"{{{self._w_ns}}}val", text)
            props.append(element)

    @property
    def table(self) -> Optional[Table]:
        element = self._table()
        return None if element is None else Table(element, self._w_ns)

    @property
    def tables(self) -> List[Table]:
        """Tables nested inside this cell (a cell can hold its own document)."""
        return [Table(child, self._w_ns) for child in self._element
                if _local(child) == "tbl"]

    def add_table(self, rows: int = 1, cols: int = 1, *, width=None,
                  text=None) -> Table:
        """Put a table inside this cell."""
        builder = Builder({"w": self._w_ns})
        widths = None if width is None else (
            [int(width) // cols] * cols if cols else [])
        element = builder.table(rows, cols, style="TableGrid",
                                widths=widths, text=text)
        self._element.append(element)
        return Table(element, self._w_ns)

    def iter_inner_content(self):
        """The paragraphs and tables in this cell, in document order."""
        for child in self._element:
            if _local(child) == "p":
                yield from [p for p in self.paragraphs
                            if p._element is child]
            elif _local(child) == "tbl":
                yield Table(child, self._w_ns)

    def _absorb(self, other: "Cell") -> None:
        """Move another cell's block content into this one."""
        other._move_content_to(self)

    def _move_content_to(self, other: "Cell") -> None:
        """Append this cell's content to ``other`` and leave a valid empty cell."""
        if self._element is other._element:
            return
        blocks = [child for child in self._element
                  if _local(child) in ("p", "tbl")]
        if len(blocks) == 1 and _local(blocks[0]) == "p" \
                and not any(_local(child) == "r" for child in blocks[0]):
            # python-docx treats a cell containing a single run-free paragraph
            # as empty and leaves it in place instead of adding a blank line.
            return

        target_blocks = [child for child in other._element
                         if _local(child) in ("p", "tbl")]
        if target_blocks and _local(target_blocks[-1]) == "p" \
                and not any(_local(child) == "r" for child in target_blocks[-1]):
            other._element.remove(target_blocks[-1])

        for block in blocks:
            self._element.remove(block)
            other._element.append(block)
        # A table cell must retain at least one block-level paragraph, including
        # when it becomes a vertical-merge continuation cell.
        self._element.append(ET.Element(f"{{{self._w_ns}}}p"))

    def _column_index(self) -> int:
        """Grid column of this cell, counting omitted and spanned cells."""
        row = self._row()
        if row is None:
            return 0
        index = Row(row, self._w_ns).grid_cols_before
        for child in row:
            if _local(child) != "tc":
                continue
            if child is self._element:
                return index
            index += Cell(child, self._w_ns).grid_span
        return index

    def _merge_origin(self) -> Optional["Cell"]:
        """For a vertically merged cell: the restart cell holding the text.

        Word writes the same content once, in the top cell; readers that do not
        follow ``w:vMerge`` back up show blanks in every row below it.
        """
        if self.vertical_merge != "continue":
            return None
        row, table = self._row(), self._table()
        if row is None or table is None:
            return None
        column = self._column_index()
        rows = [element for element in table if _local(element) == "tr"]
        try:
            index = rows.index(row)
        except ValueError:
            return None
        for above in reversed(rows[:index]):
            cells = [element for element in above if _local(element) == "tc"]
            offset = Row(above, self._w_ns).grid_cols_before
            found = None
            for element in cells:
                if offset == column:
                    found = element
                    break
                offset += Cell(element, self._w_ns).grid_span
            if found is None:
                continue
            cell = Cell(found, self._w_ns)
            cell._row_element = above
            cell._table_element = table
            if cell.vertical_merge == "restart":
                return cell
            if cell.vertical_merge != "continue":
                return None
        return None

    def _paragraph_elements(self) -> List[ET.Element]:
        return [child for child in self._element if _local(child) == "p"]

    @property
    def text(self) -> str:
        from .views import views_of
        origin = self._merge_origin()
        if origin is not None:
            return origin.text
        return "\n".join((views_of(p).get("accepted") or "")
                         for p in self._paragraph_elements())

    @text.setter
    def text(self, value: str) -> None:
        from .views import Paragraph, _tag
        paragraphs = self._paragraph_elements()
        if not paragraphs:
            builder = Builder({"w": self._w_ns})
            self._element.append(builder.paragraph(""))
            paragraphs = self._paragraph_elements()
        for extra in paragraphs[1:]:
            self._element.remove(extra)
        first = Paragraph(text="", view="accepted", kind="document",
                          _element=paragraphs[0], _w_ns=self._w_ns)
        first.set_text(value or "")

    @property
    def paragraphs(self) -> List["Paragraph"]:
        from .views import Paragraph, views_of
        out = []
        for index, element in enumerate(self._paragraph_elements()):
            out.append(Paragraph(text=views_of(element).get("accepted") or "",
                                 view="accepted", kind="document", index=index,
                                 in_table=True, _element=element,
                                 _w_ns=self._w_ns))
        return out

    def add_paragraph(self, text: str = "", **fmt):
        from .views import Paragraph, views_of
        element = Builder({"w": self._w_ns}).paragraph(text, **fmt)
        self._element.append(element)
        return Paragraph(text=views_of(element).get("accepted") or "",
                         view="accepted", kind="document",
                         index=len(self._paragraph_elements()) - 1,
                         in_table=True, _element=element, _w_ns=self._w_ns)


class Column:
    """One grid column: its width, and the cells under it."""

    def __init__(self, table: "Table", index: int):
        self._table = table
        self._index = index

    @property
    def width(self) -> Optional[int]:
        """The declared width in EMU (``w:tblGrid``), or None."""
        columns = self._grid()
        if self._index >= len(columns):
            return None
        value = columns[self._index].get(f"{{{self._table._w_ns}}}w")
        return Length(int(value) * 635) if value else None

    @width.setter
    def width(self, value) -> None:
        columns = self._grid()
        if self._index < len(columns):
            columns[self._index].set(f"{{{self._table._w_ns}}}w", str(twips(value)))

    def _grid(self) -> List[ET.Element]:
        for child in self._table._element:
            if _local(child) == "tblGrid":
                return [c for c in child if _local(c) == "gridCol"]
        return []

    @property
    def cells(self) -> List["Cell"]:
        """The cell in each row that sits in this column, merged cells included."""
        return [row.cells[self._index] for row in self._table.rows
                if self._index < len(row.cells)]

    @property
    def table(self) -> "Table":
        return self._table

    def __repr__(self) -> str:
        return f"Column({self._index})"


class Table:
    """A ``w:tbl`` you can edit: rows, cells, text, new rows."""

    def __init__(self, element: ET.Element, w_ns: str):
        self._element, self._w_ns = element, w_ns

    # ------------------------------------------------------------- geometry
    def _tblpr(self, create: bool = False) -> Optional[ET.Element]:
        for child in self._element:
            if _local(child) == "tblPr":
                return child
        if not create:
            return None
        props = ET.Element(f"{{{self._w_ns}}}tblPr")
        self._element.insert(0, props)
        return props

    def _prop(self, name: str, attribute: str = "val") -> Optional[str]:
        props = self._tblpr()
        if props is None:
            return None
        for child in props:
            if _local(child) == name:
                return child.get(f"{{{self._w_ns}}}{attribute}")
        return None

    def _set_prop(self, name: str, attribute: str, value) -> None:
        props = self._tblpr(create=value is not None)
        if props is None:
            return
        for child in list(props):
            if _local(child) == name:
                props.remove(child)
        if value is None:
            return
        element = ET.Element(f"{{{self._w_ns}}}{name}")
        element.set(f"{{{self._w_ns}}}{attribute}", token(value))
        props.append(element)

    @property
    def alignment(self):
        """``LEFT``/``CENTER``/``RIGHT``, the value python-docx reports."""
        return self._prop("jc")

    @alignment.setter
    def alignment(self, value) -> None:
        self._set_prop("jc", "val", token(value))

    @property
    def autofit(self) -> bool:
        """True unless the table explicitly declares fixed-width layout."""
        layout = self._prop("tblLayout", "type")
        return True if layout is None else layout != "fixed"

    @autofit.setter
    def autofit(self, value) -> None:
        self._set_prop("tblLayout", "type", None if value is None
                       else ("autofit" if value else "fixed"))

    @property
    def table_direction(self) -> Optional[str]:
        """``"rtl"``, ``"ltr"`` or None when the table declares neither."""
        props = self._tblpr()
        if props is None:
            return None
        for child in props:
            if _local(child) == "bidiVisual":
                return "ltr" if child.get(f"{{{self._w_ns}}}val") in ("0", "false") \
                    else "rtl"
        return None

    @table_direction.setter
    def table_direction(self, value) -> None:
        props = self._tblpr(create=value is not None)
        if props is None:
            return
        for child in list(props):
            if _local(child) == "bidiVisual":
                props.remove(child)
        if value is None:
            return
        element = ET.Element(f"{{{self._w_ns}}}bidiVisual")
        element.set(f"{{{self._w_ns}}}val", "0" if value == "ltr" else "1")
        props.append(element)

    @property
    def table(self) -> "Table":
        """python-docx answers this from a row, a cell or a column; so do we."""
        return self

    def _grid_columns(self) -> List[ET.Element]:
        for child in self._element:
            if _local(child) == "tblGrid":
                return [c for c in child if _local(c) == "gridCol"]
        return []

    @property
    def columns(self) -> List[Column]:
        """One :class:`Column` per grid column, left to right."""
        count = len(self._grid_columns())
        if not count:
            count = max([0] + [len(row.cells) for row in self.rows])
        return [Column(self, index) for index in range(count)]

    def column_cells(self, index: int) -> List[Cell]:
        """Every cell in one column, in row order."""
        return [row.cells[index] for row in self.rows if index < len(row.cells)]

    def row_cells(self, index: int) -> List[Cell]:
        """Every cell in one row — python-docx's name for ``rows[i].cells``."""
        return self.rows[index].cells

    def add_column(self, width=None) -> Column:
        """Add a column on the right, with an empty cell in every row."""
        grid = None
        for child in self._element:
            if _local(child) == "tblGrid":
                grid = child
        if grid is None:
            grid = ET.Element(f"{{{self._w_ns}}}tblGrid")
            self._element.insert(0, grid)
        columns = [child for child in grid if _local(child) == "gridCol"]
        reference = columns[-1].get(f"{{{self._w_ns}}}w") if columns else "2400"
        gridcol = ET.SubElement(grid, f"{{{self._w_ns}}}gridCol")
        gridcol.set(f"{{{self._w_ns}}}w",
                    str(twips(width)) if width is not None else reference)
        builder = Builder({"w": self._w_ns})
        for row_element in self._element:
            if _local(row_element) != "tr":
                continue
            cell = ET.SubElement(row_element, f"{{{self._w_ns}}}tc")
            props = ET.SubElement(cell, f"{{{self._w_ns}}}tcPr")
            width_element = ET.SubElement(props, f"{{{self._w_ns}}}tcW")
            width_element.set(f"{{{self._w_ns}}}w",
                              gridcol.get(f"{{{self._w_ns}}}w"))
            width_element.set(f"{{{self._w_ns}}}type", "dxa")
            cell.append(builder.paragraph(""))
        return Column(self, len(columns))

    @property
    def rows(self) -> List[Row]:
        out = []
        for child in self._element:
            if _local(child) != "tr":
                continue
            row = Row(child, self._w_ns)
            row._table_element = self._element
            row._table = self
            out.append(row)
        return out

    def cell(self, row: int, col: int) -> Cell:
        return self.rows[row].cells[col]

    def add_row(self) -> Row:
        rows = self.rows
        builder = Builder({"w": self._w_ns})
        columns = len(rows[0].cells) if rows else 1
        tr = ET.Element(f"{{{self._w_ns}}}tr")
        for index in range(columns):
            tc = ET.Element(f"{{{self._w_ns}}}tc")
            tc.append(builder.paragraph(""))
            tr.append(tc)
        self._element.append(tr)
        return Row(tr, self._w_ns)

    @property
    def style(self):
        """The table style: ``""``, or a StyleRef with ``.name``."""
        from .views import StyleRef
        for child in self._element:
            if _local(child) == "tblPr":
                for prop in child:
                    if _local(prop) == "tblStyle":
                        value = prop.get(f"{{{self._w_ns}}}val")
                        return StyleRef(value) if value else ""
        return ""

    @style.setter
    def style(self, value) -> None:
        """Set the table style, creating ``w:tblPr`` when the table lacks one."""
        props = None
        for child in self._element:
            if _local(child) == "tblPr":
                props = child
                break
        if props is None:
            props = ET.Element(f"{{{self._w_ns}}}tblPr")
            self._element.insert(0, props)
        for child in list(props):
            if _local(child) == "tblStyle":
                props.remove(child)
        if value:
            name = getattr(value, "name", value)
            element = ET.Element(f"{{{self._w_ns}}}tblStyle")
            element.set(f"{{{self._w_ns}}}val", str(name).replace(" ", ""))
            props.insert(0, element)

    def text(self) -> List[List[str]]:
        return [[cell.text for cell in row.cells] for row in self.rows]

    def __repr__(self) -> str:
        rows = self.rows
        return f"Table({len(rows)}x{len(rows[0].cells) if rows else 0})"


class Styles:
    """The style names a document declares, plus the built-ins this library writes."""

    def __init__(self, names: Optional[Iterable[str]] = None):
        self._names = list(names or [])
        for builtin in _BUILTIN_STYLE_NAMES.values():
            if builtin not in self._names:
                self._names.append(builtin)

    def __iter__(self):
        return iter(self._names)

    def __contains__(self, name: str) -> bool:
        return name in self._names

    def __getitem__(self, name: str):
        if name not in self._names:
            raise KeyError(name)
        return Style(name)

    def __len__(self) -> int:
        return len(self._names)

    def __repr__(self) -> str:
        return f"Styles({len(self._names)}: {', '.join(self._names[:6])}…)"


class Style:
    def __init__(self, name: str):
        self.name = name

    @property
    def style_id(self) -> str:
        for style_id, label in _BUILTIN_STYLE_NAMES.items():
            if label == self.name:
                return style_id
        return self.name.replace(" ", "")

    @property
    def type(self) -> int:
        return 2 if self.name.startswith("Table") or self.name == "Table Grid" else 1

    def __repr__(self) -> str:
        return f"Style({self.name!r})"
