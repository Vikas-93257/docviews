"""What exactly does python-docx change when it saves an opened file?"""
from __future__ import annotations

import io
import sys
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET

CORPUS = Path(__file__).resolve().parent / "corpus"


def main(limit=40):
    import docx
    dropped = Counter()
    changed = Counter()
    examples = []
    notes = Counter()
    files = sorted(CORPUS.glob("*.docx"))[:limit]
    for path in files:
        try:
            buffer = io.BytesIO(path.read_bytes())
            document = docx.Document(buffer)
            out = io.BytesIO()
            document.save(out)
        except Exception:                                      # noqa: BLE001
            notes["would not open"] += 1
            continue
        with zipfile.ZipFile(io.BytesIO(buffer.getvalue())) as a, \
                zipfile.ZipFile(io.BytesIO(out.getvalue())) as b:
            names_a, names_b = set(a.namelist()), set(b.namelist())
            for name in sorted(names_a - names_b):
                dropped[name.rsplit("/", 1)[-1]] += 1
            for name in sorted(names_b - names_a):
                notes[f"added {name.rsplit('/', 1)[-1]}"] += 1
            for part in sorted(names_a & names_b):
                da, db = a.read(part), b.read(part)
                if da == db:
                    continue
                try:
                    if (ET.canonicalize(da, with_comments=True, rewrite_prefixes=True) ==
                            ET.canonicalize(db, with_comments=True, rewrite_prefixes=True)):
                        notes["same meaning, different bytes"] += 1
                        continue
                except ET.ParseError:
                    pass
                changed[part.rsplit("/", 1)[-1]] += 1
                if len(examples) < 3 and part.endswith("document.xml"):
                    examples.append((path.name, da, db))
    print("entries python-docx DROPS on save:", dict(dropped.most_common(8)))
    print("parts whose MEANING changes:    ", dict(changed.most_common(8)))
    print("other notes:                    ", dict(notes.most_common(8)))
    for name, da, db in examples:
        print(f"\n--- {name}")
        for label, data in (("before", da), ("after", db)):
            root = ET.fromstring(data)
            decls = {k for k in root.attrib if k.endswith("Ignorable")}
            print(f"  {label}: root attribs={sorted(decls)} "
                  f"comment_nodes={sum(1 for _ in root.iter() if not isinstance(_.tag, str))}")
        a_root, b_root = ET.fromstring(da), ET.fromstring(db)
        a_tags = Counter(_.tag for _ in a_root.iter())
        b_tags = Counter(_.tag for _ in b_root.iter())
        lost = [(t.rsplit('}',1)[-1], n - b_tags.get(t, 0)) for t, n in a_tags.items()
                if b_tags.get(t, 0) < n]
        gained = [(t.rsplit('}',1)[-1], n - a_tags.get(t, 0)) for t, n in b_tags.items()
                  if a_tags.get(t, 0) < n]
        print(f"  elements lost: {lost[:6]}")
        print(f"  elements gained: {gained[:6]}")


if __name__ == "__main__":
    main(int(sys.argv[1]) if len(sys.argv) > 1 else 40)
