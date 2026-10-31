"""Census v2 — element-exact measurement of what the standard reader loses.

v1 compared paragraph lists by position, which silently misaligned paragraphs
nested inside tables. v2 compares every python-docx Paragraph against the
revision views of its OWN XML element (`p._p`), so there is no alignment to get
wrong, and it separates four different kinds of loss:

  A. insertions silently dropped   -> reader returns the ORIGINAL view
  B. deletions correctly dropped   -> reader happens to return the ACCEPTED view
  C. text matching NEITHER view    -> no Word setting shows this text
  D. numbers lost / invented       -> what corrupts a contract, invoice or lab value

plus text that no python-docx API can reach at all (footnotes, endnotes, text
boxes) and runs that .runs never yields.

Run: python docviews/census2.py --corpus docviews/corpus --out docviews/census2.jsonl
"""

from __future__ import annotations

import argparse
import json
import re
import signal
from collections import Counter, defaultdict
from pathlib import Path
from typing import Dict, Iterator, List, Optional, Tuple
from xml.etree import ElementTree as ET

import docx  # the incumbent, measured

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
NUMBER = re.compile(r"\d[\d,]*\.?\d*")
PER_FILE_SECONDS = 90
MAX_PART_BYTES = 12 * 1024 * 1024   # a 0.5 MB upload can hide 28 MB of XML


class _Timeout(Exception):
    pass


def _alarm(signum, frame):  # noqa: ARG001
    raise _Timeout()


# ---------------------------------------------------------------- XML views

# python-docx maps these elements to characters inside Paragraph.text
CONTROL = {"tab": "\t", "ptab": "\t", "br": "\n", "cr": "\n", "noBreakHyphen": "-"}

# containers that hide a run from python-docx. Paragraph.text sees direct <w:r>
# children plus runs inside <w:hyperlink>; everything else is invisible to it.
HIDING = {"ins": "inserted run (w:ins)", "del": "deleted run (w:del)",
          "moveFrom": "moved-from run (w:moveFrom)", "moveTo": "moved-to run (w:moveTo)",
          "sdt": "content control (w:sdt)", "smartTag": "smart tag (w:smartTag)",
          "customXmlInsRangeStart": "custom-xml range", "fldSimple": "simple field (w:fldSimple)",
          "txbxContent": "text box"}


def _walk(el: ET.Element, ins=False, dele=False, mfrom=False, mto=False):
    """Every text node under `el`, tagged with the revision containers around it."""
    out = []
    for child in el:
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "ins":
            out += _walk(child, True, dele, mfrom, mto)
        elif tag == "del":
            out += _walk(child, ins, True, mfrom, mto)
        elif tag == "moveFrom":
            out += _walk(child, ins, dele, True, mto)
        elif tag == "moveTo":
            out += _walk(child, ins, dele, mfrom, True)
        elif tag == "t":
            out.append((ins, dele, mfrom, mto, child.text or ""))
        elif tag == "delText":
            out.append((ins, True, mfrom, mto, child.text or ""))
        elif tag in CONTROL:
            out.append((ins, dele, mfrom, mto, CONTROL[tag]))
        else:
            out += _walk(child, ins, dele, mfrom, mto)
    return out


def views(el: ET.Element) -> Tuple[str, str, bool]:
    """(original, accepted, has_revisions) text for one XML element."""
    nodes = _walk(el)
    original, accepted, rev = [], [], False
    for ins, dele, mfrom, mto, text in nodes:
        if ins or dele or mfrom or mto:
            rev = True
        if not (ins or mto):          # untouched by "accept all" -> in original
            original.append(text)
        if not (dele or mfrom):       # untouched by "reject all" -> in accepted
            accepted.append(text)
    return "".join(original), "".join(accepted), rev


def nums(text: str) -> Counter:
    return Counter(n.replace(",", "").strip(".") for n in NUMBER.findall(text))


# ------------------------------------------------------- paragraph walkers

def table_paragraphs(table) -> Iterator[object]:
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                yield p
            for t in cell.tables:          # nested tables
                yield from table_paragraphs(t)


def reachable_paragraphs(document) -> Tuple[List[object], List[str]]:
    """Every paragraph any python-docx API can hand you, plus the API used."""
    paras, how = [], []
    for p in document.paragraphs:
        paras.append(p)
        how.append("doc.paragraphs")
    for t in document.tables:
        for p in table_paragraphs(t):
            paras.append(p)
            how.append("doc.tables")
    seen_parts = set()                       # synthetic files can have 1000s of
    for section in document.sections:        # sections; each header part is parsed
        for part, name in (                     # once, not once per section
                (section.header, "header"), (section.footer, "footer"),
                (section.first_page_header, "first_page_header"),
                (section.first_page_footer, "first_page_footer"),
                (section.even_page_header, "even_page_header"),
                (section.even_page_footer, "even_page_footer")):
            try:
                key = part.part.partname
                if key in seen_parts:
                    continue
                seen_parts.add(key)
                for p in part.paragraphs:
                    paras.append(p)
                    how.append(name)
                for t in part.tables:
                    for p in table_paragraphs(t):
                        paras.append(p)
                        how.append(name)
            except Exception:
                pass
    try:                                    # comments API (new in python-docx 1.2.0)
        for c in document.comments:
            for p in c.paragraphs:
                paras.append(p)
                how.append("doc.comments")
    except Exception:
        pass
    return paras, how


def unreachable_chars(path: Path, root: ET.Element) -> Dict[str, int]:
    """Text that lives in parts python-docx 1.2.0 has no reader for."""
    import zipfile

    out = {"footnotes": 0, "endnotes": 0, "textboxes": 0}
    # text boxes: inside document.xml itself, but no API walks them
    for box in root.iter(f"{W}txbxContent"):
        for p in box.iter(f"{W}p"):
            out["textboxes"] += len(views(p)[1])
    zf = zipfile.ZipFile(path)
    for part, key in (("word/footnotes.xml", "footnotes"),
                      ("word/endnotes.xml", "endnotes")):
        try:
            sub = ET.fromstring(zf.read(part))
        except KeyError:
            continue
        for p in sub.iter(f"{W}p"):
            out[key] += len(views(p)[1])
        del sub
    return out


# ------------------------------------------------------------------ per file

def analyse(path: Path) -> Dict[str, object]:
    signal.signal(signal.SIGALRM, _alarm)
    signal.alarm(PER_FILE_SECONDS)
    try:
        return _analyse(path)
    except _Timeout:
        return {"file": path.name, "skipped": "timeout"}
    except Exception as exc:
        return {"file": path.name, "skipped": f"{type(exc).__name__}: {exc}"[:140]}
    finally:
        signal.alarm(0)


def _analyse(path: Path) -> Dict[str, object]:
    import gc
    import zipfile

    # --- pass 1: raw XML (one parse; freed before the reader runs) ----------
    zf = zipfile.ZipFile(path)
    names = zf.namelist()
    info = zf.getinfo("word/document.xml")
    if info.file_size > MAX_PART_BYTES:
        zf.close()
        return {"file": path.name, "bytes": path.stat().st_size,
                "skipped": f"document.xml is {info.file_size / 1e6:.1f} MB uncompressed",
                "document_xml_mb": round(info.file_size / 1e6, 2)}
    raw = zf.read("word/document.xml")
    zf.close()
    root = ET.fromstring(raw)
    del raw
    unreachable = unreachable_chars(path, root)
    total_runs = sum(1 for _ in root.iter(f"{W}r"))
    for n in names:                        # headers/footers/footnotes also carry runs
        if n == "word/document.xml" or not n.endswith(".xml"):
            continue
        if not (re.match(r"word/(header|footer)\d*\.xml", n)
                or n in ("word/footnotes.xml", "word/endnotes.xml")):
            continue
        try:
            sub = ET.fromstring(zipfile.ZipFile(path).read(n))
        except Exception:
            continue
        total_runs += sum(1 for _ in sub.iter(f"{W}r"))
        del sub
    # every <w:r> in the text-bearing parts, attributed to whatever hides it
    runs_by_wrapper: Counter = Counter()

    def tally(el, cause="visible via .text"):
        for child in el:
            tag = child.tag.rsplit("}", 1)[-1]
            if tag == "r":
                runs_by_wrapper[cause] += 1
            elif tag in HIDING:
                tally(child, HIDING[tag])
            elif tag == "hyperlink":
                runs_by_wrapper["hyperlink run (not in .runs)"] += sum(
                    1 for _ in child.iter(f"{W}r"))
            else:
                tally(child, cause)

    tally(root)
    del root, names
    gc.collect()

    # --- pass 2: the incumbent reader ---------------------------------------
    document = docx.Document(str(path))
    paras, how = reachable_paragraphs(document)

    a = b = c = d = 0                       # the four loss kinds
    rev_paras = 0
    lost = invented = 0
    empty_with_text = 0                     # reported empty, XML holds text
    examples: List[Dict[str, str]] = []
    by_api: Counter = Counter()

    for p, api in zip(paras, how):
        try:
            reader = p.text
        except Exception:
            continue
        original, accepted, rev = views(p._p)
        if not rev:
            continue
        rev_paras += 1
        by_api[api] += 1
        if not reader.strip() and accepted.strip():
            empty_with_text += 1
        same_orig = reader == original
        same_acc = reader == accepted
        if same_orig and not same_acc:
            a += 1
        elif same_acc and not same_orig:
            b += 1
        elif not same_orig and not same_acc:
            c += 1
        if reader != accepted:
            d += 1
        # numbers, compared against what the document's own accepted view says
        na, nr = nums(accepted), nums(reader)
        l = sum(max(0, v - nr.get(k, 0)) for k, v in na.items())
        g = sum(max(0, v - na.get(k, 0)) for k, v in nr.items())
        lost += l
        invented += g
        if (l or g) and len(examples) < 6:
            examples.append({"api": api, "reader": reader[:300],
                             "accepted": accepted[:300], "original": original[:300],
                             "numbers_lost": str(l), "numbers_invented": str(g)})

    reachable_runs = 0
    for p in paras:
        try:
            reachable_runs += len(p.runs)
        except Exception:
            pass

    return {
        "file": path.name,
        "bytes": path.stat().st_size,
        "bytes_by_view": None,
        "paragraphs": len(paras),
        "revision_paragraphs": rev_paras,
        "revision_paragraphs_by_api": dict(by_api),
        "ins_only_dropped": a,               # reader == original, Word shows accepted
        "del_only_ok": b,                    # reader == accepted, by luck
        "matches_neither_view": c,           # the impossible text
        "differs_from_accepted": d,          # what Word shows vs what you get
        "reported_empty_but_has_text": empty_with_text,
        "numbers_lost": lost,
        "numbers_invented": invented,
        "unreachable": unreachable,
        "unreachable_total": sum(unreachable.values()),
        "runs_total": total_runs,
        "runs_reachable": reachable_runs,
        "runs_by_wrapper": dict(runs_by_wrapper),
        "examples": examples,
    }


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--corpus", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--start", type=int, default=0)
    ap.add_argument("--end", type=int, default=None)
    ap.add_argument("--append", action="store_true")
    ap.add_argument("--only", default=None,
                    help="comma-separated filenames to process (for retries)")
    args = ap.parse_args()
    files = sorted(Path(args.corpus).glob("*.docx"))[args.start:args.end]
    if args.only:
        wanted = set(args.only.split(","))
        files = [p for p in files if p.name in wanted] or [
            Path(args.corpus) / n for n in sorted(wanted)]
    mode = "a" if args.append else "w"
    with Path(args.out).open(mode) as fh:
        for i, path in enumerate(files, 1):
            row = analyse(path)
            fh.write(json.dumps(row) + "\n")
            fh.flush()
    print(f"done: {len(files)} files -> {args.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
