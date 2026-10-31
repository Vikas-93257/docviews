"""Audit every document in corpus/ and print the census numbers docviews claims.

    python3 audit_corpus.py [--out audit_corpus.json] [--start N] [--end N] [--append]

Batched runs (fresh process per --start/--end slice) keep peak memory flat on a
2 GB box; every slice appends to the same JSONL so a crash still leaves output.
"""
from __future__ import annotations

import argparse
import gc
import json
import sys
from collections import Counter
from dataclasses import asdict
from pathlib import Path

sys.path.insert(0, "src")                     # noqa: E402
from docviews.audit import audit_file         # noqa: E402

CORPUS = Path(__file__).resolve().parent / "corpus"


def run(out_path: Path, start: int, end: int, append: bool, batch: int = 50) -> None:
    files = sorted(CORPUS.glob("*.docx"))
    files = files[start:end]
    mode = "a" if append else "w"
    with out_path.open(mode) as fh:
        for i in range(0, len(files), batch):
            for path in files[i:i + batch]:
                row = audit_file(path)
                fh.write(json.dumps(asdict(row)) + "\n")
            fh.flush()
            gc.collect()
    print(f"wrote {len(files)} rows to {out_path} ({start}..{end})", flush=True)


def summarise(path: Path) -> None:
    rows = [json.loads(line) for line in path.open()]
    verdicts = Counter(r["verdict"] for r in rows)
    print(f"\nfiles: {len(rows)}")
    for verdict, n in verdicts.most_common():
        print(f"  {verdict:10s} {n:4d}  ({n / len(rows):5.1%})")

    strict = sum(1 for r in rows if r.get("strict"))
    print(f"  strict: {strict}")

    rev = sum(r.get("revision_paragraphs", 0) for r in rows)
    rev_docs = sum(1 for r in rows if r.get("revision_paragraphs"))
    print(f"revision paragraphs: {rev} in {rev_docs} documents")

    unreach = Counter()
    docs = Counter()
    for r in rows:
        for kind, chars in (r.get("unreachable_chars") or {}).items():
            unreach[kind] += chars
            docs[kind] += 1
    for kind, chars in unreach.most_common():
        print(f"unreachable {kind}: {chars:,} chars in {docs[kind]} documents")
    print(f"numbers at risk: {sum(r.get('numbers_at_risk', 0) for r in rows):,}")

    kinds = Counter()
    for r in rows:
        for kind, n in (r.get("revision_aliases") or {}).items():
            kinds[kind] += n
    print("revisions by kind:", dict(kinds))

    for view in ("accepted", "original", "markup"):
        total = sum((r.get("chars") or {}).get(view, 0) for r in rows)
        print(f"chars[{view}]: {total:,}")

    biggest = max(rows, key=lambda r: r.get("largest_part_mb", 0))
    print(f"largest part: {biggest.get('largest_part_mb', 0):.1f} MB "
          f"({biggest['path'].split('/')[-1][:50]}, expansion "
          f"{biggest.get('expansion_ratio', 0):.0f}x)")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="audit_corpus.jsonl")
    ap.add_argument("--start", type=int, default=0)
    ap.add_argument("--end", type=int, default=10 ** 9)
    ap.add_argument("--append", action="store_true")
    ap.add_argument("--summarise", action="store_true")
    args = ap.parse_args()
    out = Path(args.out)
    if args.summarise:
        summarise(out)
        return 0
    run(out, args.start, args.end, args.append)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
