"""Print the exact, hand-checkable evidence for a document.

For every paragraph whose text the incumbent reader gets wrong, show what the
reader returns, what Word shows (accepted), what the original said, and the XML
that explains the difference. No aggregation, no metrics — just the bytes.

Run: python docviews/exhibits.py docviews/corpus/<file>.docx [--max 4]
"""

from __future__ import annotations

import argparse
import re
import zipfile
from pathlib import Path

import docx
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
from census2 import views  # noqa: E402


def short(xml: str, limit: int = 420) -> str:
    return re.sub(r">\s*<", "><", xml)[:limit]


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("file")
    ap.add_argument("--max", type=int, default=4)
    args = ap.parse_args()
    path = Path(args.file)

    document = docx.Document(str(path))
    shown = 0
    total = wrong = 0
    for p in document.paragraphs:
        original, accepted, rev = views(p._p)
        total += 1
        if not rev:
            continue
        try:
            reader = p.text
        except Exception:
            continue
        if reader == accepted:
            continue
        wrong += 1
        if shown >= args.max:
            continue
        shown += 1
        verdict = ("matches NEITHER view" if reader not in (original, accepted)
                   else "returns the ORIGINAL (draft) text"
                   if reader == original else "other")
        print("=" * 88)
        print(f"paragraph {total}   ->   {verdict}")
        print(f"  reader (python-docx .text) : {reader[:160]!r}")
        print(f"  Word, changes accepted     : {accepted[:160]!r}")
        if original != accepted:
            print(f"  Word, changes rejected     : {original[:160]!r}")
        print(f"  XML: {short(ET.tostring(p._p, encoding='unicode'))}")
    print("=" * 88)
    print(f"{path.name}: {wrong} of {total} top-level paragraphs come back wrong "
          f"({wrong / max(1, total):.1%})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
