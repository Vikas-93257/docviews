"""Turn the census JSONL into the verdict table.

Run: python docviews/census_report2.py docviews/census2.jsonl docviews/corpus/manifest.json
"""

from __future__ import annotations

import json
import statistics
import sys
from collections import Counter, defaultdict
from pathlib import Path
from typing import Dict, List


def load(path: str) -> List[Dict[str, object]]:
    rows = []
    for line in Path(path).read_text().splitlines():
        line = line.strip()
        if line:
            row = json.loads(line)
            if isinstance(row, dict):
                rows.append(row)
    return rows


def main() -> int:
    rows = load(sys.argv[1])
    manifest = Path(sys.argv[2]) if len(sys.argv) > 2 else None
    source_of = {}
    if manifest and manifest.exists():
        source_of = {m["file"]: m["source"] for m in
                     json.loads(manifest.read_text())}

    skipped = [r for r in rows if "skipped" in r]
    measured = [r for r in rows if "skipped" not in r]
    n = len(measured)
    too_big = [r for r in skipped if "uncompressed" in str(r.get("skipped"))]
    unopenable = [r for r in skipped if "uncompressed" not in str(r.get("skipped"))]

    def docs_with(key) -> List[Dict[str, object]]:
        return [r for r in measured if int(r.get(key) or 0) > 0]

    def line(label: str, k: int, total: int = n, extra: str = "") -> str:
        return (f"  {label:<52}{k:5d}  {k / max(1, total):6.1%}  {extra}")

    print("=" * 84)
    print(f"CENSUS — {n} documents measured, {len(skipped)} unmeasurable")
    print("=" * 84)

    rev_docs = docs_with("revision_paragraphs")
    rev_paras = sum(int(r["revision_paragraphs"]) for r in measured)
    a_docs = docs_with("ins_only_dropped")
    c_docs = docs_with("matches_neither_view")
    d_docs = docs_with("differs_from_accepted")
    num_docs = [r for r in measured if int(r["numbers_lost"]) or int(r["numbers_invented"])]
    empty_docs = docs_with("reported_empty_but_has_text")

    print("\n  A. TRACKED CHANGES — the reader silently returns the wrong text")
    print(line("documents containing revision markup", len(rev_docs)))
    print(line("  ...where reader text == ORIGINAL view (draft text)",
               len(a_docs), n, f"from {sum(int(r['ins_only_dropped']) for r in measured):,} paragraphs"))
    print(line("  ...where reader text matches NEITHER view",
               len(c_docs), n, f"from {sum(int(r['matches_neither_view']) for r in measured):,} paragraphs"))
    print(line("  ...where reader text != what Word displays (accepted)",
               len(d_docs), n, f"from {sum(int(r['differs_from_accepted']) for r in measured):,} paragraphs"))
    print(line("  ...with numbers lost or invented",
               len(num_docs), n,
               f"{sum(int(r['numbers_lost']) for r in measured):,} lost / "
               f"{sum(int(r['numbers_invented']) for r in measured):,} invented"))
    print(line("  ...where reader reports EMPTY but XML holds text",
               len(empty_docs), n, f"from {sum(int(r['reported_empty_but_has_text']) for r in measured):,} paragraphs"))
    print(f"      revision paragraphs in total: {rev_paras:,}")

    print("\n  B. TEXT NO python-docx API CAN REACH")
    unr_docs = docs_with("unreachable_total")
    unr_chars = sum(int(r["unreachable_total"]) for r in measured)
    body_chars = sum(int(r["paragraphs"]) for r in measured)
    print(line("documents with unreachable text", len(unr_docs)))
    print(f"      {unr_chars:,} characters in footnotes / endnotes / text boxes "
          f"(no API returns them)")
    tb = sum(r["unreachable"]["textboxes"] for r in measured)
    fn = sum(r["unreachable"]["footnotes"] for r in measured)
    en = sum(r["unreachable"]["endnotes"] for r in measured)
    print(f"      text boxes {tb:,} · footnotes {fn:,} · endnotes {en:,}")

    print("\n  C. RUNS THE .runs PROPERTY NEVER YIELDS")
    tot_runs = sum(int(r["runs_total"]) for r in measured)
    reach_runs = sum(int(r["runs_reachable"]) for r in measured)
    run_docs = [r for r in measured if int(r["runs_total"]) > int(r["runs_reachable"])]
    print(line("documents with unreachable runs", len(run_docs)))
    print(f"      {tot_runs - reach_runs:,} of {tot_runs:,} <w:r> runs invisible "
          f"({(tot_runs - reach_runs) / max(1, tot_runs):.1%})")

    print("\n  D. UNMEASURABLE (the reader cannot even open these)")
    print(line("files python-docx raises on", len(unopenable), len(rows)))
    reasons = Counter(str(r.get("skipped", ""))[:46] for r in unopenable)
    for reason, k in reasons.most_common(4):
        print(f"        {k:4d}  {reason}")
    print(line("files with a >12 MB XML part (skipped by the probe)",
               len(too_big), len(rows)))

    print("\n  E. BY SOURCE (is the revision rate an artefact of one test suite?)")
    groups: Dict[str, Dict[str, int]] = defaultdict(
        lambda: {"n": 0, "rev": 0, "num": 0, "unr": 0, "para": 0})
    for r in measured:
        group = source_of.get(r["file"], "?").split("/")[0]
        g = groups[group]
        g["n"] += 1
        g["rev"] += int(r["revision_paragraphs"]) > 0
        g["num"] += bool(int(r["numbers_lost"]) or int(r["numbers_invented"]))
        g["unr"] += int(r["unreachable_total"]) > 0
        g["para"] += int(r["paragraphs"])
    print(f"  {'source':<26}{'docs':>6}{'revisions':>11}{'numbers wrong':>15}"
          f"{'unreachable':>13}")
    for name, g in sorted(groups.items(), key=lambda kv: -kv[1]["n"]):
        print(f"  {name[:25]:<26}{g['n']:>6}{g['rev'] / g['n']:>10.1%}"
              f"{g['num'] / g['n']:>14.1%}{g['unr'] / g['n']:>12.1%}")

    print("\n  F. WORST CASES — numbers the reader gets wrong")
    worst = sorted(measured, key=lambda r: -(int(r["numbers_lost"])
                                             + int(r["numbers_invented"])))[:8]
    for r in worst:
        if not (int(r["numbers_lost"]) or int(r["numbers_invented"])):
            break
        print(f"  {str(r['file'])[:50]:<50} lost={r['numbers_lost']:>4} "
              f"invented={r['numbers_invented']:>3} "
              f"paras_affected={r['differs_from_accepted']:>4}")
        for ex in (r["examples"] or [])[:1]:
            print(f"      reader reports : {ex['reader'][:88]!r}")
            print(f"      document shows : {ex['accepted'][:88]!r}")

    print("\n  " + "=" * 82)
    revisit = len(rev_docs) / n
    print(f"  VERDICT INPUTS: revision markup in {len(rev_docs)}/{n} = "
          f"{revisit:.1%} of documents")
    print(f"                  wrong-numbers in {len(num_docs)}/{n} = {len(num_docs) / n:.1%}")
    print(f"                  unreachable text in {len(unr_docs)}/{n} = "
          f"{len(unr_docs) / n:.1%}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
