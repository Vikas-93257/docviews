"""Head-to-head: open every corpus file, save it unchanged, compare packages.

    python3 compare_roundtrip.py [limit]

Reports, per reader: how many files open, how many come back with the same
package entries, how many parts came back byte-identical, and how many files
survive the trip with no part rewritten at all.
"""
from __future__ import annotations

import io
import sys
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET

sys.path.insert(0, "src")
from docviews import open_docx                                   # noqa: E402

CORPUS = Path(__file__).resolve().parent / "corpus"


def compare(data_a: bytes, data_b: bytes) -> tuple:
    """(entries_same, parts_byte_identical, parts_meaning_same, total_parts)"""
    try:
        with zipfile.ZipFile(io.BytesIO(data_a)) as a, \
                zipfile.ZipFile(io.BytesIO(data_b)) as b:
            names_a, names_b = sorted(a.namelist()), sorted(b.namelist())
            entries_same = names_a == names_b
            identical = same_meaning = total = 0
            for part in set(names_a) & set(names_b):
                if part.endswith("/"):
                    continue
                total += 1
                da, db = a.read(part), b.read(part)
                if da == db:
                    identical += 1
                    same_meaning += 1
                    continue
                try:
                    if (ET.canonicalize(da, with_comments=True, rewrite_prefixes=True) ==
                            ET.canonicalize(db, with_comments=True, rewrite_prefixes=True)):
                        same_meaning += 1
                except ET.ParseError:
                    pass
            return entries_same, identical, same_meaning, total
    except Exception:                                          # noqa: BLE001
        return False, 0, 0, 0


def probe_docviews(path: Path) -> dict:
    original = path.read_bytes()
    doc = open_docx(original)
    rewritten = doc.to_bytes()
    entries, identical, meaning, total = compare(original, rewritten)
    return {"opened": True, "entries_same": entries, "identical": identical,
            "meaning": meaning, "total": total}


def probe_python_docx(path: Path) -> dict:
    import docx
    buffer = io.BytesIO(path.read_bytes())
    document = docx.Document(buffer)
    out = io.BytesIO()
    document.save(out)
    entries, identical, meaning, total = compare(buffer.getvalue(), out.getvalue())
    return {"opened": True, "entries_same": entries, "identical": identical,
            "meaning": meaning, "total": total}


def main():
    limit = int(sys.argv[1]) if len(sys.argv) > 1 else 10 ** 9
    files = sorted(CORPUS.glob("*.docx"))[:limit]
    totals = {"docviews": Counter(), "python-docx": Counter()}
    failures = []
    for path in files:
        for label, probe in (("docviews", probe_docviews),
                             ("python-docx", probe_python_docx)):
            stats = totals[label]
            stats["files"] += 1
            try:
                row = probe(path)
            except Exception as exc:                           # noqa: BLE001
                stats["would_not_open"] += 1
                if label == "docviews":
                    failures.append((path.name, type(exc).__name__))
                continue
            stats["opened"] += 1
            stats["entries_same"] += bool(row["entries_same"])
            stats["all_parts_identity"] += row["identical"] == row["total"]
            stats["parts_identical"] += row["identical"]
            stats["parts_meaning_same"] += row["meaning"]
            stats["parts_total"] += row["total"]
    for label, stats in totals.items():
        opened, total = stats["opened"], stats["files"]
        print(f"\n{label}")
        print(f"  opened                 {opened:,} / {total:,}")
        print(f"  package entries kept   {stats['entries_same']:,}")
        print(f"  every part unchanged    {stats['all_parts_identity']:,}")
        print(f"  parts byte-identical   {stats['parts_identical']:,} / {stats['parts_total']:,}")
        print(f"  parts meaning-same     {stats['parts_meaning_same']:,} / {stats['parts_total']:,}")
    if failures:
        print(f"\ndocviews could not open: {failures}")


if __name__ == "__main__":
    main()
