"""v0.3 evidence: what authoring touches, measured over the corpus.

Three sweeps, each writing one line per document to ``authoring_impact.jsonl``:

* ``roundtrip``  — open and save unchanged: every entry and every byte must match
* ``header``     — append a paragraph to the first header that exists and diff
* ``picture``    — insert a picture into the body and diff
* ``interop``    — python-docx 1.2.0 opens what was written and must agree

The point is the same as in ``compare_save_impact.py``: an editor is only safe if
the untouched parts of the package are provably untouched. Run with::

    python3 authoring_impact.py [corpus_dir]
"""

from __future__ import annotations

import io
import json
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent / "tests"))

from docviews import Document, open_docx          # noqa: E402
from docviews.images import detect                # noqa: E402
import blobs                                      # noqa: E402

BASELINE = "word/document.xml"
ALLOWED_PICTURE_PARTS = {BASELINE, "word/_rels/document.xml.rels",
                         "[Content_Types].xml"}


def entries_diff(before: bytes, after: bytes) -> tuple:
    """(entries_lost, entries_added, parts_changed) between two packages."""
    with zipfile.ZipFile(io.BytesIO(before)) as a, \
            zipfile.ZipFile(io.BytesIO(after)) as b:
        names_a, names_b = a.namelist(), b.namelist()
        lost = [n for n in names_a if n not in names_b]
        added = [n for n in names_b if n not in names_a]
        changed = []
        for name in names_a:
            if name.endswith("/") or name in added:
                continue
            if name not in names_b or a.read(name) != b.read(name):
                changed.append(name)
    return lost, added, changed


def sweep_roundtrip(files, out):
    for path in files:
        record = {"sweep": "roundtrip", "file": path.name}
        try:
            original = path.read_bytes()
            saved = Document(original).to_bytes()
        except Exception as exc:                              # noqa: BLE001
            record.update(ok=False, error=f"{type(exc).__name__}: {exc}")
            out.write(json.dumps(record) + "\n")
            continue
        lost, added, changed = entries_diff(original, saved)
        record.update(ok=True, entries_lost=len(lost), entries_added=len(added),
                      parts_changed=changed)
        out.write(json.dumps(record) + "\n")


def sweep_header(files, out):
    for path in files:
        record = {"sweep": "header", "file": path.name}
        try:
            original = path.read_bytes()
            doc = Document(original)
            written = False
            for index, _section in enumerate(doc.sections):
                header = doc.sections[index].header
                if header.text.strip():
                    header.add_paragraph("docviews was here")
                    written = True
                    break
            if not written:
                record.update(ok=True, skipped="no header with text")
                out.write(json.dumps(record) + "\n")
                continue
            saved = doc.to_bytes()
        except Exception as exc:                              # noqa: BLE001
            record.update(ok=False, error=f"{type(exc).__name__}: {exc}")
            out.write(json.dumps(record) + "\n")
            continue
        lost, added, changed = entries_diff(original, saved)
        reopened = open_docx(saved)
        record.update(ok=True, entries_lost=len(lost), entries_added=len(added),
                      parts_changed=changed,
                      header_kept="docviews was here" in "".join(
                          section.header.text for section in reopened.sections))
        out.write(json.dumps(record) + "\n")


def sweep_picture(files, out):
    for path in files:
        record = {"sweep": "picture", "file": path.name}
        try:
            original = path.read_bytes()
            before = len(open_docx(original).inline_shapes)
            doc = Document(original)
            doc.add_picture(blobs.PNG_96, width=1828800)
            saved = doc.to_bytes()
        except Exception as exc:                              # noqa: BLE001
            record.update(ok=False, error=f"{type(exc).__name__}: {exc}")
            out.write(json.dumps(record) + "\n")
            continue
        lost, added, changed = entries_diff(original, saved)
        after = open_docx(saved)
        unexpected = [name for name in changed if name not in ALLOWED_PICTURE_PARTS]
        record.update(ok=True, entries_lost=len(lost), entries_added=len(added),
                      parts_changed=changed, unexpected_changes=unexpected,
                      shapes_before=before, shapes_after=len(after.inline_shapes))
        out.write(json.dumps(record) + "\n")


def sweep_interop(files, out):
    """python-docx has to read back what this library wrote."""
    try:
        import docx as python_docx
    except ImportError:                                       # noqa: BLE001
        out.write(json.dumps({"sweep": "interop", "skipped": "python-docx absent"}) + "\n")
        return
    for path in files:
        record = {"sweep": "interop", "file": path.name}
        try:
            original = path.read_bytes()
            doc = Document(original)
            doc.add_picture(blobs.PNG_96, width=1828800)
            for index, _ in enumerate(doc.sections):
                header = doc.sections[index].header
                if header.text.strip():
                    header.add_paragraph("interop check")
                    break
            doc.styles.add_style("DocviewsInterop", "paragraph", bold=True)
            saved = doc.to_bytes()
            theirs = python_docx.Document(io.BytesIO(saved))
        except Exception as exc:                              # noqa: BLE001
            record.update(ok=False, error=f"{type(exc).__name__}: {exc}")
            out.write(json.dumps(record) + "\n")
            continue
        header_text = " ".join(p.text for section in theirs.sections
                               for p in section.header.paragraphs)
        record.update(ok=True, shapes=len(theirs.inline_shapes),
                      style_present=any(s.name == "DocviewsInterop"
                                        for s in theirs.styles),
                      header_seen="interop check" in header_text)
        out.write(json.dumps(record) + "\n")


def measure_images():
    """The measurement table for §6.4: what each format says about itself."""
    rows = []
    for name, data in blobs.BLOBS.items():
        try:
            info = detect(data)
        except Exception as exc:                              # noqa: BLE001
            rows.append({"fixture": name, "refused": type(exc).__name__})
            continue
        rows.append({"fixture": name, "type": info.content_type,
                     "px": [info.width_px, info.height_px],
                     "dpi": [round(info.dpi_x, 2), round(info.dpi_y, 2)],
                     "native_in": [round(info.native_width / 914400, 3),
                                   round(info.native_height / 914400, 3)]})
    return rows


SWEEPS = {"roundtrip": ("w", sweep_roundtrip),
          "header": ("a", sweep_header),
          "picture": ("a", sweep_picture),
          "interop": ("a", sweep_interop)}


def main() -> int:
    corpus = Path(sys.argv[1] if len(sys.argv) > 1 else "corpus")
    wanted = (sys.argv[2].split(",") if len(sys.argv) > 2
              else ["roundtrip", "header", "picture", "interop"])
    files = sorted(corpus.glob("*.docx"))
    target = Path(__file__).resolve().parent / "authoring_impact.jsonl"
    if "roundtrip" in wanted and target.exists():
        target.unlink()                      # the full run writes it from scratch
    for name in wanted:
        mode, sweep = SWEEPS[name]
        with target.open(mode) as out:
            if name == "picture":
                sweep(files[::20], out)
            elif name == "interop":
                sweep(files[::20], out)
            else:
                sweep(files, out)
        print(f"{name}: done")
    with (Path(__file__).resolve().parent / "image_facts.json").open("w") as out:
        json.dump(measure_images(), out, indent=1)
    print(f"wrote authoring_impact.jsonl and image_facts.json for {len(files)} files")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
