"""Low-level OOXML package access.

Nothing here depends on python-docx or lxml: a .docx is a zip with XML in it,
and both the ISO/IEC 29500 *Transitional* and *Strict* flavours can be read with
the standard library. python-docx hardcodes the Transitional relationship type
and therefore raises on Strict packages (60 of Microsoft's 378 conformance
documents in the project census).

Design rules:
  * never silently drop anything — every anomaly lands in ``warnings``
  * every limit is explicit and adjustable (a 0.5 MB upload can hide 28 MB of XML)
"""

from __future__ import annotations

import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterator, List, Optional, Union
from xml.etree import ElementTree as ET

# --------------------------------------------------------------------- namespaces
# Strict (ISO/IEC 29500-1 second edition) moved every namespace onto
# purl.oclc.org; Transitional is what Office writes by default. Same element
# names, different URIs — so one mapping each way is enough.
TRANSITIONAL = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}
STRICT = {
    "w": "http://purl.oclc.org/ooxml/wordprocessingml/main",
    "r": "http://purl.oclc.org/ooxml/officeDocument/relationships",
    "wp": "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}

# relationship types that carry text, in the order a Word user would meet them
# Every prefix Office writes, so a re-serialised part keeps the prefixes the
# file already used (mc:Ignorable refers to them by name, so inventing ns0
# would break the attribute's meaning).
PREFIXES = {
    "w": TRANSITIONAL["w"], "r": TRANSITIONAL["r"], "wp": TRANSITIONAL["wp"],
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2010/11/wordml",
    "w16cex": "http://schemas.microsoft.com/office/word/2018/wordml/cex",
    "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "wpi": "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
    "wne": "http://schemas.microsoft.com/office/word/2006/wordml",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "o": "urn:schemas-microsoft-com:office:office",
    "v": "urn:schemas-microsoft-com:vml",
    "w10": "urn:schemas-microsoft-com:office:word",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "sl": "http://schemas.openxmlformats.org/schemaLibrary/2006/main",
    "ct": TRANSITIONAL["ct"], "pr": TRANSITIONAL["pr"],
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
    "dcmitype": "http://purl.org/dc/dcmitype/",
    "xsi": "http://www.w3.org/2001/XMLSchema-instance",
}

for _prefix, _uri in PREFIXES.items():
    ET.register_namespace(_prefix, _uri)


def root_ns(root: ET.Element) -> str:
    """The wordprocessingml namespace a part actually uses."""
    tag = root.tag
    if isinstance(tag, str) and tag.startswith("{"):
        return tag[1:].split("}", 1)[0]
    return TRANSITIONAL["w"]


def parse_xml(data: bytes) -> ET.Element:
    """Parse XML keeping comments and processing instructions.

    A round-trip has to preserve what it does not understand, and ElementTree
    drops comments by default — which would quietly rewrite 7 of the 832 census
    documents.
    """
    parser = ET.XMLParser(target=ET.TreeBuilder(insert_comments=True,
                                                insert_pis=True))
    parser.feed(data)
    return parser.close()


REL_SUFFIXES = {
    "officeDocument": ("document",),
    "footnotes": ("footnotes",),
    "endnotes": ("endnotes",),
    "header": ("header",),
    "footer": ("footer",),
    "comments": ("comments",),
}


@dataclass
class Limits:
    """Guardrails against hostile or broken packages (see census §3.6)."""

    max_part_bytes: int = 64 * 1024 * 1024      # a single XML part
    max_expansion_ratio: float = 400.0          # uncompressed / compressed
    max_total_uncompressed: int = 512 * 1024 * 1024

    def check(self, name: str, compressed: int, uncompressed: int) -> Optional[str]:
        if uncompressed > self.max_part_bytes:
            return (f"{name} is {uncompressed / 1e6:.1f} MB uncompressed "
                    f"(limit {self.max_part_bytes / 1e6:.0f} MB)")
        if compressed and uncompressed / compressed > self.max_expansion_ratio:
            return (f"{name} expands {uncompressed / compressed:.0f}x "
                    f"(limit {self.max_expansion_ratio:.0f}x)")
        return None


class DocxError(Exception):
    """The package cannot be read at all."""


class PartTooLarge(DocxError):
    """A part exceeds the configured limits; raise the limit to read it."""


@dataclass
class Part:
    name: str
    kind: str                                   # document | footnotes | header | …
    xml: bytes = b""
    _root: Optional[ET.Element] = field(default=None, repr=False)

    @property
    def root(self) -> ET.Element:
        if self._root is None:
            self._root = parse_xml(self.xml)
        return self._root

    def release(self) -> None:
        """Drop the parsed tree (used by the streaming path)."""
        self._root = None
        self.xml = b""


class Package:
    """A .docx opened as a zip. Detects Strict vs Transitional automatically."""

    def __init__(self, source: Union[str, Path, bytes], limits: Optional[Limits] = None):
        self.limits = limits or Limits()
        self.warnings: List[str] = []
        try:
            self._zip = (zipfile.ZipFile(source) if not isinstance(source, bytes)
                         else zipfile.ZipFile(_BytesReader(source)))
        except zipfile.BadZipFile as exc:
            raise DocxError(f"not a zip package: {exc}") from exc

        infos = self._zip.infolist()
        total = sum(i.file_size for i in infos)
        if total > self.limits.max_total_uncompressed:
            self.warnings.append(
                f"package expands to {total / 1e6:.0f} MB "
                f"(limit {self.limits.max_total_uncompressed / 1e6:.0f} MB)")
        self._parts: Dict[str, Part] = {}
        self.ns = dict(TRANSITIONAL)
        self.strict = False
        self.main_part = self._find_main_part()
        self._load_relationships()

    # ------------------------------------------------------------------ discovery
    def _find_main_part(self) -> str:
        """Find the document part via the package relationships.

        Both flavours are accepted; this is the step python-docx skips (it looks
        for the Transitional relationship type only).
        """
        try:
            rels = ET.fromstring(self._zip.read("_rels/.rels"))
        except KeyError as exc:
            raise DocxError("package has no _rels/.rels") from exc
        except ET.ParseError as exc:
            raise DocxError(f"_rels/.rels is not XML: {exc}") from exc
        for rel in rels:
            rtype = rel.get("Type", "")
            target = rel.get("Target", "")
            if rtype.endswith("/officeDocument") and target:
                self.strict = "purl.oclc.org" in rtype
                self.ns = dict(STRICT if self.strict else TRANSITIONAL)
                return target.lstrip("/")
        raise DocxError("package has no officeDocument relationship")

    def _load_relationships(self) -> None:
        """Map part name -> relationship type, so parts are found by kind."""
        self.rels: Dict[str, str] = {}
        base = self.main_part.rsplit("/", 1)[0] if "/" in self.main_part else ""
        rels_name = f"{base}/_rels/{self.main_part.rsplit('/', 1)[-1]}.rels"
        try:
            rels = ET.fromstring(self._zip.read(rels_name))
        except (KeyError, ET.ParseError) as exc:
            self.warnings.append(f"no readable relationships part ({rels_name}): {exc}")
            return
        for rel in rels:
            target = rel.get("Target", "")
            if rel.get("TargetMode") == "External" or not target:
                continue
            name = target.lstrip("/")
            if not name.startswith("word/") and base:
                name = f"{base}/{name}"
            self.rels[name] = rel.get("Type", "")

    def kind_of(self, part_name: str) -> str:
        rtype = self.rels.get(part_name, "")
        for suffix, (kind,) in REL_SUFFIXES.items():
            if rtype.endswith("/" + suffix):
                return kind
        tail = part_name.rsplit("/", 1)[-1].lower()
        for suffix, (kind,) in REL_SUFFIXES.items():
            if tail.startswith(kind):
                return kind
        return "other"

    # ---------------------------------------------------------------------- parts
    def names(self) -> List[str]:
        return [i.filename for i in self._zip.infolist()]

    def exists(self, name: str) -> bool:
        try:
            self._zip.getinfo(name)
            return True
        except KeyError:
            return False

    def read(self, name: str) -> bytes:
        info = self._zip.getinfo(name)
        problem = self.limits.check(name, info.compress_size, info.file_size)
        if problem:
            raise PartTooLarge(
                f"{problem}. Raise Limits.max_part_bytes / max_expansion_ratio "
                f"if this document is trusted.")
        return self._zip.read(name)

    def part(self, name: str) -> Part:
        if name not in self._parts:
            self._parts[name] = Part(name=name, kind=self.kind_of(name),
                                     xml=self.read(name))
        return self._parts[name]

    def text_parts(self) -> List[Part]:
        """Every part that carries visible text, in reading order.

        Body first, then the parts python-docx 1.2.0 cannot reach at all.
        """
        wanted: List[str] = [self.main_part]
        for kind in ("header", "footer", "comments", "footnotes", "endnotes"):
            wanted += sorted(n for n in self.rels if self.kind_of(n) == kind)
        parts = []
        for name in wanted:
            if not self.exists(name):
                if name == self.main_part:
                    raise DocxError(f"main part {name} is missing")
                self.warnings.append(f"relationship points at missing part {name}")
                continue
            try:
                parts.append(self.part(name))
            except PartTooLarge as exc:
                if name == self.main_part:
                    # never hand back an empty document: the guards exist to
                    # stop a bomb, not to silently drop the text
                    raise
                self.warnings.append(f"skipped {name}: {exc}")
            except ET.ParseError as exc:
                self.warnings.append(f"skipped {name}: not well-formed XML ({exc})")
        return parts

    def expanded_bytes(self, name: str) -> int:
        try:
            return self._zip.getinfo(name).file_size
        except KeyError:
            return 0

    def close(self) -> None:
        self._zip.close()

    def __enter__(self) -> "Package":
        return self

    def __exit__(self, *exc) -> None:
        self.close()


class _BytesReader:
    """zipfile needs a seekable file object; bytes are already in memory."""

    def __init__(self, data: bytes):
        import io
        self._io = io.BytesIO(data)

    def read(self, *a):
        return self._io.read(*a)

    def seek(self, *a):
        return self._io.seek(*a)

    def tell(self):
        return self._io.tell()

    def seekable(self):
        return True

    def close(self):
        return self._io.close()
