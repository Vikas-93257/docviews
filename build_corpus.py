"""Day-1 corpus builder: dedupe public .docx into corpus/ with a manifest.

Sources are the open-source test-data sets where real user documents accumulate
(bug reports, conformance files, library fixtures) — the closest thing to a
public sample of documents people actually process.

Run: python docviews/build_corpus.py --src /tmp/gh --src /tmp/src_pkgs --out docviews/corpus
"""

from __future__ import annotations

import argparse
import hashlib
import json
import shutil
from pathlib import Path


def sha1(path: Path) -> str:
    h = hashlib.sha1()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", action="append", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--max-bytes", type=int, default=70 * 1024 * 1024)
    args = ap.parse_args()

    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    seen: dict[str, str] = {}
    manifest = []
    total = 0
    for src in args.src:
        for path in sorted(Path(src).rglob("*.docx")):
            if path.name.startswith("~$"):
                continue                       # Word lock files
            try:
                digest = sha1(path)
            except OSError:
                continue
            if digest in seen:
                continue
            size = path.stat().st_size
            if total + size > args.max_bytes:
                continue
            seen[digest] = path.name
            # keep the source tree visible in the filename: collisions are common
            rel = path.relative_to(Path(src))
            name = f"{path.parent.parent.name[:18]}_{path.parent.name[:18]}_{path.name}"
            name = "".join(c if c.isalnum() or c in "._-" else "_" for c in name)
            shutil.copy2(path, out / name)
            manifest.append({"file": name, "sha1": digest, "bytes": size,
                             "source": str(rel)})
            total += size

    (out / "manifest.json").write_text(json.dumps(manifest, indent=1))
    print(f"corpus: {len(manifest)} unique documents, {total / 1e6:.1f} MB")
    print(f"manifest: {out / 'manifest.json'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
