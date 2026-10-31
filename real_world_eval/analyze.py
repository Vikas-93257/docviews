"""Compare package-open/save behaviour on the external real-world DOCX set.

This is an exploratory compatibility smoke test, not a Word-rendered truth set.
Run from the project root with Python-docx 1.2.0 installed::

    PYTHONPATH=src python3 real_world_eval/analyze.py
"""
from __future__ import annotations

import hashlib
import io
import json
import re
import sys
import time
import zipfile
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parent
PROJECT = ROOT.parent
sys.path.insert(0, str(PROJECT / "src"))
from docviews import open_docx  # noqa: E402


def package_diff(before: bytes, after: bytes) -> dict:
    with zipfile.ZipFile(io.BytesIO(before)) as a, zipfile.ZipFile(io.BytesIO(after)) as b:
        names_a = {n for n in a.namelist() if not n.endswith("/")}
        names_b = {n for n in b.namelist() if not n.endswith("/")}
        common = names_a & names_b
        identical = sum(a.read(n) == b.read(n) for n in common)
        changed = sorted(n for n in common if a.read(n) != b.read(n))
        return {
            "entries_before": len(names_a),
            "entries_after": len(names_b),
            "entries_lost": sorted(names_a - names_b),
            "entries_added": sorted(names_b - names_a),
            "common_parts": len(common),
            "parts_byte_identical": identical,
            "parts_changed": changed,
        }


def package_facts(data: bytes) -> dict:
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        names = [n for n in z.namelist() if not n.endswith("/")]
        text_parts = [n for n in names if n.startswith("word/") and n.endswith(".xml")]
        strict = False
        for name in text_parts:
            try:
                root = ET.fromstring(z.read(name))
            except ET.ParseError:
                continue
            if any("purl.oclc.org/ooxml" in tag for tag in root.iter()):
                strict = True
                break
        return {
            "entries": len(names),
            "uncompressed_bytes": sum(z.getinfo(n).file_size for n in names),
            "word_xml_parts": len(text_parts),
            "media_parts": sum(n.startswith("word/media/") for n in names),
            "custom_xml_parts": sum(n.startswith("customXml/") for n in names),
            "has_comments_part": "word/comments.xml" in names,
            "has_footnotes_part": "word/footnotes.xml" in names,
            "has_endnotes_part": "word/endnotes.xml" in names,
            "strict_namespace_seen": strict,
        }


def docviews_probe(data: bytes) -> dict:
    t0 = time.perf_counter()
    doc = open_docx(data)
    opened = time.perf_counter()
    text = doc.text("accepted")
    paragraphs = doc.paragraphs()
    infos = doc.parts()
    saved = doc.to_bytes()
    completed = time.perf_counter()
    return {
        "open_seconds": round(opened - t0, 4),
        "open_save_seconds": round(completed - t0, 4),
        "paragraphs": len(paragraphs),
        "accepted_chars_all_parts": len(text),
        "text_parts": len(infos),
        "warnings": list(doc.warnings),
        "package_after_save": package_diff(data, saved),
    }


def _body_text(doc) -> str:
    bits = [p.text for p in doc.paragraphs]
    seen_cells = set()
    kept = []

    def visit_table(table):
        for row in table.rows:
            for cell in row.cells:
                tc = cell._tc
                if tc in seen_cells:
                    continue
                seen_cells.add(tc)
                kept.append(tc)
                bits.extend(p.text for p in cell.paragraphs)
                for nested in cell.tables:
                    visit_table(nested)

    for table in doc.tables:
        visit_table(table)
    return "\n".join(bits)


def python_docx_probe(data: bytes) -> dict:
    import docx

    t0 = time.perf_counter()
    doc = docx.Document(io.BytesIO(data))
    opened = time.perf_counter()
    body_text = _body_text(doc)
    out = io.BytesIO()
    doc.save(out)
    completed = time.perf_counter()
    return {
        "version": docx.__version__,
        "open_seconds": round(opened - t0, 4),
        "open_save_seconds": round(completed - t0, 4),
        "body_paragraphs": len(doc.paragraphs),
        "body_and_table_chars": len(body_text),
        "tables": len(doc.tables),
        "sections": len(doc.sections),
        "package_after_save": package_diff(data, out.getvalue()),
    }


def run() -> dict:
    manifest = json.loads((ROOT / "manifest.json").read_text())
    rows = []
    version = None
    for source in manifest:
        path = ROOT / "files" / source["local_file"]
        data = path.read_bytes()
        if hashlib.sha256(data).hexdigest() != source["sha256"]:
            raise RuntimeError(f"sha256 mismatch: {path.name}")
        row = {
            "id": source["id"],
            "filename": path.name,
            "title": source["title"],
            "publisher": source["publisher"],
            "kind": source["kind"],
            "sha256": source["sha256"],
            "bytes": len(data),
            "package": package_facts(data),
        }
        try:
            row["docviews"] = docviews_probe(data)
        except Exception as exc:
            row["docviews"] = {"error": f"{type(exc).__name__}: {exc}"}
        try:
            row["python_docx"] = python_docx_probe(data)
            version = row["python_docx"]["version"]
        except Exception as exc:
            row["python_docx"] = {"error": f"{type(exc).__name__}: {exc}"}
        rows.append(row)
    return {
        "created_utc": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "python_docx_version": version,
        "purpose": "Exploratory real-world package/open-save smoke test; not an independent Word-rendered fidelity benchmark.",
        "files": rows,
    }


def main() -> None:
    result = run()
    out = ROOT / "analysis.json"
    out.write_text(json.dumps(result, indent=2, ensure_ascii=False) + "\n")
    print(f"Python-docx: {result['python_docx_version']}")
    print(f"Documents: {len(result['files'])}")
    print("\n{:<39} {:>10} {:>10} {:>13} {:>13} {:>12}".format(
        "file", "dv-open", "pd-open", "dv-roundtrip", "pd-lost", "pd-open?"))
    for row in result["files"]:
        dv = row["docviews"]
        pd = row["python_docx"]
        dv_open = dv.get("open_seconds", "FAIL")
        pd_open = pd.get("open_seconds", "FAIL")
        dv_diff = dv.get("package_after_save", {})
        dv_round = "exact" if dv_diff and not dv_diff["entries_lost"] and not dv_diff["entries_added"] and not dv_diff["parts_changed"] else "changed/fail"
        pd_diff = pd.get("package_after_save", {})
        pd_lost = len(pd_diff.get("entries_lost", [])) if pd_diff else "—"
        print("{:<39} {:>10} {:>10} {:>13} {:>13} {:>12}".format(
            row["filename"][:39], dv_open, pd_open, dv_round, pd_lost,
            "yes" if "error" not in pd else "no"))
    print(f"\nWrote {out}")


if __name__ == "__main__":
    main()
