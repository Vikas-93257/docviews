"""Relationship integrity: does every r:id in a package still resolve?

A .docx is a package of parts plus relationships. If a relationship is dropped
but the part still references it (``r:id="rId1"``), the document is *broken*:
Word either repairs the file or silently drops the content that pointed there
(a hyperlink, an image, a footnote reference).

    python3 check_integrity.py [limit]        # python-docx vs docviews saves
"""
from __future__ import annotations

import io
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET

CORPUS = Path(__file__).resolve().parent / "corpus"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
STRICT_REL_NS = "http://purl.oclc.org/ooxml/officeDocument/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
REF = re.compile(rb'r:(?:id|embed|link|pict|dm|lo|qs|cs|href)="([^"]+)"')


def rels_for(part: str) -> str:
    if "/" in part:
        folder, name = part.rsplit("/", 1)
        return f"{folder}/_rels/{name}.rels"
    return f"_rels/{part}.rels"


def dangling(zf: zipfile.ZipFile) -> list:
    """Parts that reference a relationship id their .rels does not define."""
    broken = []
    for name in zf.namelist():
        if not name.endswith(".xml") or name.endswith(".rels"):
            continue
        data = zf.read(name)
        if b"r:" not in data:
            continue
        ids = {m.decode() for m in REF.findall(data)}
        if not ids:
            continue
        available = set()
        rels_name = rels_for(name)
        if rels_name in zf.namelist():
            try:
                root = ET.fromstring(zf.read(rels_name))
                available = {rel.get("Id") for rel in root}
            except ET.ParseError:
                pass
        missing = sorted(ids - available)
        if missing:
            broken.append((name, missing))
    return broken


def save_python_docx(raw: bytes) -> bytes:
    import docx
    document = docx.Document(io.BytesIO(raw))
    out = io.BytesIO()
    document.save(out)
    return out.getvalue()


def save_docviews(raw: bytes) -> bytes:
    sys.path.insert(0, "src")
    from docviews import open_docx
    return open_docx(raw).to_bytes()


def main(limit=10 ** 9):
    files = sorted(CORPUS.glob("*.docx"))[:limit]
    for label, saver in (("python-docx 1.2.0", save_python_docx),
                         ("docviews 0.2.0", save_docviews)):
        stats = Counter()
        examples = []
        for path in files:
            raw = path.read_bytes()
            try:
                saved = saver(raw)
            except Exception:                                  # noqa: BLE001
                stats["could not open"] += 1
                continue
            stats["saved"] += 1
            with zipfile.ZipFile(io.BytesIO(raw)) as before:
                bad_before = dangling(before)
            with zipfile.ZipFile(io.BytesIO(saved)) as after:
                bad_after = dangling(after)
            if bad_after:
                stats["files with dangling refs after save"] += 1
                stats["dangling refs after save"] += sum(len(m) for _, m in bad_after)
                if len(examples) < 4:
                    examples.append((path.name, bad_after[:2]))
            if bad_before and not bad_after:
                stats["refs repaired"] += 1
            if bad_before:
                stats["already dangling before"] += 1
        print(f"\n{label}")
        for key, value in stats.most_common():
            print(f"  {key:38s} {value:,}")
        for name, bad in examples:
            print(f"    {name[:46]:46s} {bad}")


if __name__ == "__main__":
    main(int(sys.argv[1]) if len(sys.argv) > 1 else 10 ** 9)
