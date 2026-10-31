"""Helpers for building tiny .docx packages in memory, in both flavours."""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

import pytest

TRANSITIONAL_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
STRICT_W = "http://purl.oclc.org/ooxml/wordprocessingml/main"
TRANSITIONAL_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
STRICT_REL = "http://purl.oclc.org/ooxml/officeDocument/relationships"

CORPUS = Path(__file__).resolve().parent.parent / "corpus"


def make_docx(body_xml: str, *, strict: bool = False,
              extra_parts: dict | None = None) -> bytes:
    """Build a minimal but valid .docx and return its bytes."""
    ns = STRICT_W if strict else TRANSITIONAL_W
    rel_base = STRICT_REL if strict else TRANSITIONAL_REL
    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{ns}" xmlns:r="{rel_base}">'
        f"<w:body>{body_xml}</w:body></w:document>"
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.'
        'openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        "</Types>"
    )
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{rel_base}/officeDocument" '
        'Target="word/document.xml"/></Relationships>'
    )
    doc_rels_items = []
    for i, (name, rtype) in enumerate({
        "footnotes.xml": "footnotes",
        "endnotes.xml": "endnotes",
        "header1.xml": "header",
        "comments.xml": "comments",
    }.items()):
        if any(p.endswith(name) for p in (extra_parts or {})):
            doc_rels_items.append(
                f'<Relationship Id="rId{i + 10}" Type="{rel_base}/{rtype}" '
                f'Target="{name}"/>')
    doc_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(doc_rels_items) + "</Relationships>"
    )

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", root_rels)
        zf.writestr("word/document.xml", document)
        zf.writestr("word/_rels/document.xml.rels", doc_rels)
        for name, xml in (extra_parts or {}).items():
            path = name if name.startswith("word/") else f"word/{name}"
            zf.writestr(path, xml.format(ns=ns) if "{ns}" in xml else xml)
    return buf.getvalue()


@pytest.fixture
def para():
    def _para(*runs: str) -> str:
        return "".join(runs)

    return _para


@pytest.fixture
def corpus_path() -> Path:
    if not CORPUS.exists():
        pytest.skip("census corpus is not present")
    return CORPUS
