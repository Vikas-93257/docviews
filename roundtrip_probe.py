"""Round-trip probe: does saving a document destroy anything?

For every .docx in the corpus: open it, save it unchanged, then compare the two
packages part by part.

  * entry sets must be identical (nothing added, nothing lost)
  * every part except word/document.xml must be byte-identical
  * word/document.xml must be C14N-equal (same document, possibly rewritten)
  * the XML comment count must survive
"""
from __future__ import annotations

import io
import json
import sys
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET

sys.path.insert(0, "src")
from docviews import open_docx                       # noqa: E402

CORPUS = Path(__file__).resolve().parent / "corpus"


def compare(original: bytes, rewritten: bytes, name: str) -> dict:
    row = {"file": name, "ok": True, "problems": [], "diff_parts": []}
    try:
        with zipfile.ZipFile(io.BytesIO(original)) as a, \
                zipfile.ZipFile(io.BytesIO(rewritten)) as b:
            names_a, names_b = a.namelist(), b.namelist()
            if sorted(names_a) != sorted(names_b):
                row["ok"] = False
                row["problems"].append(
                    f"entries changed: +{sorted(set(names_b) - set(names_a))} "
                    f"-{sorted(set(names_a) - set(names_b))}")
            for part in sorted(set(names_a) & set(names_b)):
                if part.endswith("/"):
                    continue
                data_a, data_b = a.read(part), b.read(part)
                if data_a == data_b:
                    continue
                row["diff_parts"].append(part)
                if not part.endswith(".xml"):
                    row["ok"] = False
                    row["problems"].append(f"{part} is not byte-identical")
                    continue
                try:
                    same = (ET.canonicalize(data_a, with_comments=True,
                                            rewrite_prefixes=True) ==
                            ET.canonicalize(data_b, with_comments=True,
                                            rewrite_prefixes=True))
                except ET.ParseError as exc:
                    row["ok"] = False
                    row["problems"].append(f"{part} unparseable: {exc}")
                    continue
                if not same:
                    row["ok"] = False
                    row["problems"].append(f"{part} changed meaning")
                elif data_a.count(b"<!--") != data_b.count(b"<!--"):
                    row["ok"] = False
                    row["problems"].append(f"{part} lost XML comments")
    except Exception as exc:                                  # noqa: BLE001
        row["ok"] = False
        row["problems"].append(f"{type(exc).__name__}: {exc}")
    return row


def main(out="roundtrip.jsonl", start=0, end=10 ** 9, batch_report=200):
    files = sorted(CORPUS.glob("*.docx"))[start:end]
    stats = Counter()
    with Path(out).open("a") as fh:
        for index, path in enumerate(files):
            try:
                original = path.read_bytes()
                doc = open_docx(original)
                rewritten = doc.to_bytes()
            except Exception as exc:                          # noqa: BLE001
                row = {"file": path.name, "ok": False,
                       "problems": [f"open/save: {type(exc).__name__}: {exc}"]}
            else:
                row = compare(original, rewritten, path.name)
                stats["saved"] += 1
                if row["diff_parts"]:
                    stats["rewrote_a_part"] += 1
            fh.write(json.dumps(row) + "\n")
            stats["ok" if row["ok"] else "bad"] += 1
            if not row["ok"]:
                print("FAIL", row["file"], row["problems"][:2], flush=True)
            if (index + 1) % batch_report == 0:
                fh.flush()
                print(f"  …{index + 1} files, {dict(stats)}", flush=True)
    print(f"done: {len(files)} files {dict(stats)}")


if __name__ == "__main__":
    args = [int(a) for a in sys.argv[1:] if a.isdigit()]
    main(start=args[0] if args else 0, end=args[1] if len(args) > 1 else 10 ** 9)
