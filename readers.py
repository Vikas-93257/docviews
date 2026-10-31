"""Do the alternative readers lose anything? Measure four of them on the corpus.

Every Python .docx pipeline picks one of these and trusts it. This script reads
each document with four different readers and compares the NUMBERS they report
against the numbers in the document's own XML (accepted view = what Word shows
with changes accepted), across every part of the document.

  python-docx   doc.paragraphs + tables + headers/footers
  mammoth       the converter most "docx to text" pipelines use
  docx2python   the parser most "give me everything" pipelines use
  raw OOXML     the reference: what the file actually contains

Run: python docviews/readers.py --corpus docviews/corpus --out docviews/readers.jsonl
"""

from __future__ import annotations

import argparse
import json
import re
import signal
import zipfile
from collections import Counter
from pathlib import Path
from typing import Dict, List, Optional
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
NUMBER = re.compile(r"\d[\d,]*\.?\d*")
PER_FILE_SECONDS = 45


class _Timeout(Exception):
    pass


def _alarm(signum, frame):  # noqa: ARG001
    raise _Timeout()


BASE64 = re.compile(r"data:[^;,)]{0,60};base64,[^)\s]{40,}|base64,[^)\s]{40,}")
# reader-internal markup that is not document content. Each is documented so the
# comparison cannot be accused of hiding real disagreements:
MARKUP = {
    # images (alt text comes from image metadata, the payload is image data);
    # then markdown escaping, which splits "58.6" into "58" and "6"
    "mammoth": [(re.compile(r"!\[[^\]]*\]\([^)]*\)"), " "),
                (re.compile(r"\\(.)"), r"\1")],
    # its own part separators, image labels and comment markers
    "docx2python": [(re.compile(r"-{3,}[A-Za-z][^-]{0,40}-{3,}"), " "),
                    (re.compile(r"<----[^>]*---->"), " "),
                    (re.compile(r"\bComment\d+\b"), " ")],
}


def plain(text: str) -> str:
    return "".join(c if c.isdigit() or c in ",." else " " for c in text)


def clean(text: str, reader: str) -> str:
    """Strip everything the reader added; keep the document's own text."""
    text = BASE64.sub(" ", text or "")
    for pattern, replacement in MARKUP.get(reader, []):
        text = pattern.sub(replacement, text)
    return plain(text) if reader == "python-docx" else text


def numbers(text: str) -> Counter:
    """Digit strings found in text runs. Reader-internal markup is stripped by
    the caller (see `clean`); base64 image payloads alone would otherwise invent
    hundreds of thousands of numbers that were never written in the document."""
    return Counter(n.replace(",", "").strip(".") for n in NUMBER.findall(text or ""))


# --------------------------------------------------------------- the truth

def xml_buckets(path: Path, view: str = "accepted") -> Dict[str, str]:
    """Text split by whether any python-docx API can reach that part."""
    from census2 import views

    zf = zipfile.ZipFile(path)
    reach, unreach = [], []
    try:
        root = ET.fromstring(zf.read("word/document.xml"))
    except Exception:
        root = None
    if root is not None:                      # text boxes live inside document.xml
        boxed = {id(p) for box in root.iter(f"{W}txbxContent")
                 for p in box.iter(f"{W}p")}
        for p in root.iter(f"{W}p"):
            text = views(p)[1 if view == "accepted" else 0]
            (unreach if id(p) in boxed else reach).append(text)
    for name in zf.namelist():
        if re.match(r"word/(header|footer)\d*\.xml", name) or name == "word/comments.xml":
            try:
                sub = ET.fromstring(zf.read(name))
            except Exception:
                continue
            reach += [views(p)[1 if view == "accepted" else 0]
                      for p in sub.iter(f"{W}p")]
        elif name in ("word/footnotes.xml", "word/endnotes.xml"):
            try:
                sub = ET.fromstring(zf.read(name))
            except Exception:
                continue
            unreach += [views(p)[1 if view == "accepted" else 0]
                        for p in sub.iter(f"{W}p")]
    return {"reachable": "\n".join(reach), "unreachable": "\n".join(unreach)}


def xml_text(path: Path, view: str = "accepted") -> str:
    """All text in the document, in every part, in the requested view."""
    from census2 import views

    parts = ["word/document.xml", "word/footnotes.xml", "word/endnotes.xml",
             "word/comments.xml"]
    zf = zipfile.ZipFile(path)
    parts += [n for n in zf.namelist()
              if re.match(r"word/(header|footer)\d*\.xml", n)]
    chunks: List[str] = []
    for name in parts:
        try:
            root = ET.fromstring(zf.read(name))
        except Exception:
            continue
        for p in root.iter(f"{W}p"):
            original, accepted, _ = views(p)
            chunks.append(accepted if view == "accepted" else original)
    return "\n".join(chunks)


# ------------------------------------------------------------ the readers

def read_python_docx(path: Path) -> str:
    import docx

    doc = docx.Document(str(path))
    out = [p.text for p in doc.paragraphs]
    seen_cells = set()          # a merged cell is returned once per grid position;
    #                             keep strong refs: lxml reuses ids of freed proxies
    kept = []

    def cell_paras(table):
        for row in table.rows:
            for cell in row.cells:
                tc = cell._tc
                if tc in seen_cells:
                    continue
                seen_cells.add(tc)
                kept.append(tc)
                for p in cell.paragraphs:
                    out.append(p.text)
                for t in cell.tables:
                    cell_paras(t)

    for t in doc.tables:
        cell_paras(t)
    seen = set()                            # a header shared by 40 sections must
    for s in doc.sections:                  # be read once, not 40 times
        for part in (s.header, s.footer, s.first_page_header, s.first_page_footer,
                     s.even_page_header, s.even_page_footer):
            try:
                key = part.part.partname
                if key in seen:
                    continue
                seen.add(key)
                out += [p.text for p in part.paragraphs]
                for t in part.tables:
                    for row in t.rows:
                        for cell in row.cells:
                            out += [p.text for p in cell.paragraphs]
            except Exception:
                pass
    try:                                    # comments API, new in 1.2.0
        for c in doc.comments:
            out += [p.text for p in c.paragraphs]
    except Exception:
        pass
    return "\n".join(out)


def read_mammoth(path: Path) -> str:
    import mammoth

    with path.open("rb") as fh:
        return mammoth.convert_to_markdown(fh).value


def read_docx2python(path: Path) -> str:
    from docx2python import docx2python

    with docx2python(str(path)) as content:
        return str(content.text)


READERS = {"python-docx": read_python_docx, "mammoth": read_mammoth,
           "docx2python": read_docx2python}


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--corpus", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--only-interesting", action="store_true",
                    help="files the census flagged: revisions or unreachable text")
    ap.add_argument("--limit", type=int, default=0)
    args = ap.parse_args()

    files = sorted(Path(args.corpus).glob("*.docx"))
    if args.only_interesting:
        census = {}
        for line in Path("docviews/census2.jsonl").read_text().splitlines():
            if line.strip():
                row = json.loads(line)
                census[row["file"]] = row
        keep = []
        for f in files:
            row = census.get(f.name)
            if row and ("skipped" not in row) and (
                    int(row["revision_paragraphs"]) or int(row["unreachable_total"])):
                keep.append(f)
        files = keep
    if args.limit:
        files = files[:args.limit]

    signal.signal(signal.SIGALRM, _alarm)
    rows = []
    with Path(args.out).open("w") as fh:
        for i, path in enumerate(files, 1):
            row: Dict[str, object] = {"file": path.name}
            try:
                signal.alarm(PER_FILE_SECONDS)
                buckets = xml_buckets(path, "accepted")
                truth_acc = numbers(buckets["reachable"]) + numbers(buckets["unreachable"])
                row["numbers_reachable_parts"] = sum(numbers(buckets["reachable"]).values())
                row["numbers_unreachable_parts"] = sum(numbers(buckets["unreachable"]).values())
                row["truth_numbers"] = sum(truth_acc.values())
            except Exception as exc:
                row["error"] = f"truth: {type(exc).__name__}: {exc}"[:100]
                fh.write(json.dumps(row) + "\n")
                continue
            finally:
                signal.alarm(0)
            for name, fn in READERS.items():
                try:
                    signal.alarm(PER_FILE_SECONDS)
                    text = clean(fn(path), name)
                    got = numbers(text)
                    missing = truth_acc - got
                    invented = got - truth_acc
                    row[name] = {
                        "chars": len(text),
                        "numbers": sum(got.values()),
                        "numbers_missing": sum(missing.values()),
                        "numbers_invented": sum(invented.values()),
                        "missing_sample": [k for k, _ in missing.most_common(4)],
                    }
                except _Timeout:
                    row[name] = {"error": "timeout"}
                except Exception as exc:
                    row[name] = {"error": f"{type(exc).__name__}: {exc}"[:90]}
                finally:
                    signal.alarm(0)
            row["revision_chars"] = 0
            try:
                acc, orig = xml_text(path, "accepted"), xml_text(path, "original")
                row["revision_chars"] = abs(len(acc) - len(orig))
            except Exception:
                pass
            fh.write(json.dumps(row) + "\n")
            fh.flush()
            if i % 20 == 0:
                print(f"  ... {i}/{len(files)}", flush=True)
    print(f"done: {len(files)} files -> {args.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
