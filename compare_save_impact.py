"""What does a save actually cost you? Measured in document content.

For every corpus file that python-docx can open: python-docx opens it, saves it
unchanged, and docviews then reads both packages and compares what a user cares
about — the text, the tracked changes, the comments, the footnotes.

Same loop for docviews' own save, as the control.
"""
from __future__ import annotations

import io
import sys
import zipfile
from collections import Counter
from pathlib import Path

sys.path.insert(0, "src")
from docviews import open_docx                                    # noqa: E402

CORPUS = Path(__file__).resolve().parent / "corpus"


def fingerprint(doc) -> dict:
    return {
        "text": doc.text("accepted"),
        "original": doc.text("original"),
        "chars": sum(len(p.text) for p in doc.iter_paragraphs("accepted")),
        "revisions": len(doc.revisions()),
        "comments": len(doc.comments()),
        "anchors": sum(len(c.anchors) for c in doc.comments()),
        "unreachable": sum(len(p.text.strip()) for p in doc.iter_paragraphs(
            "accepted", kinds=["footnotes", "endnotes"])),
        "numbers": len(doc.numbers()),
    }


def compare_fingerprints(before: dict, after: dict) -> list:
    problems = []
    for key in ("text", "original"):
        if before[key] != after[key]:
            problems.append(f"{key} changed "
                            f"({len(before[key])} -> {len(after[key])} chars)")
    for key in ("chars", "revisions", "comments", "anchors", "unreachable",
                "numbers"):
        if before[key] != after[key]:
            problems.append(f"{key} {before[key]} -> {after[key]}")
    return problems


def save_with_python_docx(path: Path) -> bytes:
    import docx
    buffer = io.BytesIO(path.read_bytes())
    document = docx.Document(buffer)
    out = io.BytesIO()
    document.save(out)
    return out.getvalue()


def save_with_docviews(path: Path) -> bytes:
    return open_docx(path.read_bytes()).to_bytes()


def run(label, saver, files, out_path: Path):
    stats = Counter()
    rows = []
    for path in files:
        stats["files"] += 1
        try:
            saved = saver(path)
        except Exception:                                      # noqa: BLE001
            stats["could_not_open"] += 1
            continue
        stats["saved"] += 1
        try:
            before = fingerprint(open_docx(path.read_bytes()))
            after = fingerprint(open_docx(saved))
        except Exception as exc:                               # noqa: BLE001
            stats["unreadable_after"] += 1
            rows.append({"file": path.name, "label": label,
                         "problems": [f"unreadable after save: {exc}"]})
            continue
        problems = compare_fingerprints(before, after)
        if problems:
            stats["content_changed"] += 1
            rows.append({"file": path.name, "label": label, "problems": problems})
        # package-level damage
        with zipfile.ZipFile(path) as a, zipfile.ZipFile(io.BytesIO(saved)) as b:
            names_a, names_b = set(a.namelist()), set(b.namelist())
            if names_a - names_b:
                stats["package_entries_lost"] += 1
                rows.append({"file": path.name, "label": label,
                             "problems": [f"entries lost: "
                                          f"{sorted(n for n in names_a - names_b)[:3]}"]})
    with out_path.open("a") as fh:
        import json
        for row in rows:
            fh.write(json.dumps(row) + "\n")
    return stats, rows


def main():
    files = sorted(CORPUS.glob("*.docx"))
    if len(sys.argv) > 1 and sys.argv[1].isdigit():
        files = files[:int(sys.argv[1])]
    for label, saver, out in (("python-docx 1.2.0", save_with_python_docx,
                               Path("impact_python_docx.jsonl")),
                              ("docviews 0.2.0", save_with_docviews,
                               Path("impact_docviews.jsonl"))):
        stats, rows = run(label, saver, files, out)
        changed = [r for r in rows if any("entries" not in p for p in r["problems"])]
        print(f"\n{label}")
        print(f"  files                       {stats['files']:,}")
        print(f"  could not open              {stats['could_not_open']:,}")
        print(f"  saved                       {stats['saved']:,}")
        print(f"  content changed by the save {stats['content_changed']:,}")
        print(f"  package entries lost        {stats['package_entries_lost']:,}")
        for row in changed[:6]:
            print(f"    {row['file'][:48]:48s} {row['problems'][:2]}")


if __name__ == "__main__":
    main()
