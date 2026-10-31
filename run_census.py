"""Run the census to completion on a memory-tight machine.

Each batch runs in its own interpreter (results streamed to JSONL), so a batch
that is killed by the OOM killer costs at most one file. The driver re-runs
whatever is still missing until the corpus is fully covered.

Run: python docviews/run_census.py --corpus docviews/corpus --out docviews/census2.jsonl
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path


def processed(out: Path) -> set:
    if not out.exists():
        return set()
    names = set()
    for line in out.read_text().splitlines():
        line = line.strip()
        if line:
            try:
                names.add(json.loads(line)["file"])
            except Exception:
                pass
    return names


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--corpus", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--batch", type=int, default=40)
    args = ap.parse_args()

    corpus = {p.name for p in Path(args.corpus).glob("*.docx")}
    out = Path(args.out)
    first = True
    for attempt in range(6):
        done = processed(out)
        missing = sorted(corpus - done)
        print(f"attempt {attempt + 1}: {len(done)}/{len(corpus)} done, "
              f"{len(missing)} missing")
        if not missing:
            break
        for i in range(0, len(missing), args.batch):
            chunk = missing[i:i + args.batch]
            cmd = [sys.executable, str(Path(__file__).parent / "census2.py"),
                   "--corpus", args.corpus, "--out", str(out),
                   "--only", ",".join(chunk)]
            if not first or out.exists():
                cmd.append("--append")
            first = False
            r = subprocess.run(cmd, capture_output=True, text=True)
            tail = (r.stdout.strip().splitlines() or [""])[-1]
            print(f"   batch {i // args.batch + 1}: {tail}"
                  + (f"  [KILLED rc={r.returncode}]" if r.returncode else ""))
    done = processed(out)
    print(f"final: {len(done)}/{len(corpus)} files measured")
    return 0 if done == corpus else 1


if __name__ == "__main__":
    raise SystemExit(main())
