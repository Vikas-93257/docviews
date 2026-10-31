"""`docviews audit` — the measurement, as a product.

Answers one question per document: **what is my current reader losing, and
why?** Everything reported here is attributable to a cause (a part nobody
reads, a revision container, a package flavour, an expansion ratio), because
"the text looks wrong" is not actionable and "these 1,204 numbers sit in parts
python-docx has no API for" is.

If python-docx is installed, ``compare="python-docx"`` measures the loss
directly against it instead of describing it in the abstract.
"""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple, Union

from .document import Document, open_docx
from .ooxml import DocxError, Limits, PartTooLarge
from .views import VIEWS, part_paragraphs

SEVERITY = {"clean": 0, "lossy": 1, "unreadable": 2}


@dataclass
class FileAudit:
    path: str
    openable: bool = True
    error: str = ""
    strict: bool = False
    verdict: str = "clean"
    reasons: List[str] = field(default_factory=list)

    # what the file contains
    paragraphs: int = 0
    chars: Dict[str, int] = field(default_factory=dict)     # per view
    parts: List[Tuple[str, str, int, bool]] = field(default_factory=list)
    #   (name, kind, chars_accepted, diverges)
    revision_paragraphs: int = 0
    revision_aliases: Dict[str, int] = field(default_factory=dict)
    revision_scopes: Dict[str, int] = field(default_factory=dict)
    revision_chars: Dict[str, int] = field(default_factory=dict)
    comments: int = 0
    comment_anchors: int = 0
    unreachable_chars: Dict[str, int] = field(default_factory=dict)
    #   what a python-docx-style reader cannot see: footnotes/endnotes/textboxes
    numbers_at_risk: int = 0
    expansion_ratio: float = 1.0
    largest_part_mb: float = 0.0
    warnings: List[str] = field(default_factory=list)

    # what a compared reader loses (only with compare=)
    compare: str = ""
    compare_error: str = ""
    compare_numbers_missing: int = 0
    compare_numbers_invented: int = 0
    compare_chars: int = 0

    @property
    def views_diverge(self) -> bool:
        return self.chars.get("original", "") != self.chars.get("accepted", "")

    @property
    def hidden_revisions(self) -> int:
        """Revisions that change content no text reader would notice."""
        return sum(self.revision_scopes.get(scope, 0)
                   for scope in ("row", "paragraph-mark", "block"))


def _numbers(text: str) -> Counter:
    return Counter(n.replace(",", "").strip(".")
                   for n in re.findall(r"\d[\d,]*\.?\d*", text))


def audit_file(path: Union[str, Path], limits: Optional[Limits] = None,
               compare: str = "") -> FileAudit:
    p = Path(path)
    row = FileAudit(path=str(p))
    try:
        doc = open_docx(p, limits=limits)
    except (DocxError, PartTooLarge) as exc:
        row.openable = False
        row.verdict = "unreadable"
        row.error = f"{type(exc).__name__}: {exc}"
        return row

    with doc:
        row.strict = doc.strict
        row.warnings = list(doc.warnings)
        for view in VIEWS:
            paras = list(doc.iter_paragraphs(view))
            row.chars[view] = sum(len(x.text) for x in paras)
            if view == "accepted":
                row.paragraphs = len(paras)
                row.revision_paragraphs = sum(1 for x in paras if x.is_revised)

        for part in doc.parts():
            row.parts.append((part.name, part.kind, part.chars.get("accepted", 0),
                              part.chars.get("original", "") !=
                              part.chars.get("accepted", "")))

        for rev in doc.revisions():
            row.revision_aliases[rev.kind] = row.revision_aliases.get(rev.kind, 0) + 1
            row.revision_scopes[rev.scope] = row.revision_scopes.get(rev.scope, 0) + 1
        for part in doc.parts():
            for kind, chars in part.revision_chars.items():
                row.revision_chars[kind] = row.revision_chars.get(kind, 0) + chars
        comments = doc.comments()
        row.comments = len(comments)
        row.comment_anchors = sum(len(c.anchors) for c in comments)

        # Text a python-docx-style reader has no API for at all. Empty
        # separator paragraphs are not text, so only stripped content counts —
        # otherwise every document with a footnote *part* looks like a loss.
        for kind in ("footnotes", "endnotes"):
            chars = sum(len(p.text.strip())
                        for p in doc.iter_paragraphs("accepted", kinds=[kind]))
            if chars:
                row.unreachable_chars[kind] = chars
        textbox_chars = sum(len(p.text.strip())
                            for p in doc.iter_paragraphs("accepted")
                            if p.in_textbox)
        if textbox_chars:
            row.unreachable_chars["textboxes"] = textbox_chars

        unreachable_text = "\n".join(
            [doc.text("accepted", kinds=[kind]) for kind in ("footnotes", "endnotes")]
            + [p.text for p in doc.iter_paragraphs("accepted") if p.in_textbox])
        row.numbers_at_risk = sum(_numbers(unreachable_text).values())

    # size profile: the DoS column from census §3.6
    try:
        import zipfile
        zf = zipfile.ZipFile(p)
        sizes = [(i.file_size, i.compress_size, i.filename) for i in zf.infolist()]
        if sizes:
            biggest = max(sizes)
            row.largest_part_mb = biggest[0] / 1e6
            compressed = sum(s[1] for s in sizes) or 1
            row.expansion_ratio = sum(s[0] for s in sizes) / compressed
        zf.close()
    except Exception as exc:            # pragma: no cover - defensive
        row.warnings.append(f"size profile failed: {exc}")

    if compare:
        row.compare = compare
        _compare_row(row, p, compare)

    # verdict
    if not row.openable:
        row.verdict = "unreadable"
    else:
        if row.revision_paragraphs and row.views_diverge:
            inserted = row.revision_chars.get("insert", 0)
            deleted = row.revision_chars.get("delete", 0)
            change = ", ".join(bit for bit in (
                f"+{inserted:,} inserted" if inserted else "",
                f"-{deleted:,} deleted" if deleted else "") if bit)
            row.reasons.append(
                f"{row.revision_paragraphs} paragraphs change under acceptance"
                + (f" ({change} characters)" if change else ""))
        if row.comments:
            row.reasons.append(
                f"{row.comments} comment{'' if row.comments == 1 else 's'} "
                f"({row.comment_anchors} anchors): python-docx returns the text "
                "and documents that it cannot return the range")
        rows = row.revision_scopes.get("row", 0)
        if rows:
            row.reasons.append(
                f"{rows} table row{'' if rows == 1 else 's'} marked deleted "
                "(no text reader reports this)")
        marks = row.revision_scopes.get("paragraph-mark", 0)
        if marks:
            row.reasons.append(
                f"{marks} paragraph mark{'' if marks == 1 else 's'} inserted or "
                "deleted (paragraphs merge; no text reader reports this)")
        if sum(row.unreachable_chars.values()):
            row.reasons.append(
                "text no python-docx API returns: " +
                ", ".join(f"{k} {v:,} chars" for k, v in row.unreachable_chars.items()))
        if row.strict:
            row.reasons.append("ISO/IEC 29500 Strict package (python-docx raises on these)")
        if row.expansion_ratio > 20 or row.largest_part_mb > 5:
            row.reasons.append(
                f"expands {row.expansion_ratio:.0f}x ({row.largest_part_mb:.1f} MB part)")
        if row.compare_numbers_missing:
            row.reasons.append(
                f"{compare} does not return {row.compare_numbers_missing:,} numbers "
                f"that are in the document")
        if row.compare_error:
            row.reasons.append(f"{compare} cannot open this file: {row.compare_error}")
        row.verdict = "lossy" if row.reasons else "clean"
    return row


def _compare_row(row: FileAudit, path: Path, compare: str) -> None:
    """Measure a competing reader on the same file, if it is installed."""
    try:
        if compare == "python-docx":
            text = _read_python_docx(path)
        elif compare == "mammoth":
            import mammoth
            with path.open("rb") as fh:
                text = mammoth.convert_to_markdown(fh).value
        elif compare == "docx2python":
            from docx2python import docx2python
            text = str(docx2python(str(path)).text)
        else:
            row.compare_error = f"unknown reader {compare!r}"
            return
    except ImportError:
        row.compare_error = f"{compare} is not installed"
        return
    except Exception as exc:
        row.compare_error = f"{type(exc).__name__}: {exc}"[:120]
        return

    row.compare_chars = len(text)
    try:
        mine = open_docx(path)
        with mine:
            truth = _numbers(mine.text("accepted")) + _numbers(
                mine.text("accepted", kinds=["footnote", "endnote", "textbox"]))
    except DocxError:
        truth = _numbers(text)
    got = _numbers(text)
    row.compare_numbers_missing = sum((truth - got).values())
    row.compare_numbers_invented = sum((got - truth).values())


def _read_python_docx(path: Path) -> str:
    import docx
    d = docx.Document(str(path))
    out = [p.text for p in d.paragraphs]
    seen = set()
    kept = []

    def cells(table):
        for r in table.rows:
            for c in r.cells:
                if c._tc in seen:
                    continue
                seen.add(c._tc)
                kept.append(c._tc)
                out.extend(p.text for p in c.paragraphs)
                for t in c.tables:
                    cells(t)

    for t in d.tables:
        cells(t)
    seen_parts = set()
    for s in d.sections:
        for part in (s.header, s.footer, s.first_page_header, s.first_page_footer,
                     s.even_page_header, s.even_page_footer):
            try:
                if part.part.partname in seen_parts:
                    continue
                seen_parts.add(part.part.partname)
                out.extend(p.text for p in part.paragraphs)
            except Exception:
                pass
    return "\n".join(out)


# ----------------------------------------------------------------- reporting

def audit_paths(paths: Iterable[Union[str, Path]], compare: str = "",
                limits: Optional[Limits] = None) -> List[FileAudit]:
    files: List[Path] = []
    for raw in paths:
        p = Path(raw)
        if p.is_dir():
            files += sorted(p.rglob("*.docx"))
        elif p.suffix.lower() == ".docx":
            files.append(p)
    return [audit_file(f, limits=limits, compare=compare) for f in files]


def table(rows: Sequence[FileAudit], show: int = 25) -> str:
    """The human report: one line per document that needs attention."""
    lossy = [r for r in rows if r.verdict == "lossy"]
    bad = [r for r in rows if r.verdict == "unreadable"]
    lines = []
    head = (f"{'file':<44}{'verdict':<11}{'revs':>5}{'unreach':>9}"
            f"{'expand':>8}  why")
    lines.append(head)
    lines.append("-" * len(head))
    for r in sorted(rows, key=lambda r: -SEVERITY[r.verdict])[:show]:
        if r.verdict == "clean":
            continue
        name = Path(r.path).name[:43]
        why = "; ".join(r.reasons)[:64] if r.reasons else r.error[:64]
        lines.append(f"{name:<44}{r.verdict:<11}{r.revision_paragraphs:>5}"
                     f"{sum(r.unreachable_chars.values()):>9}"
                     f"{r.expansion_ratio:>7.0f}x  {why}")
    lines.append("")
    lines.append(
        f"{len(rows)} files: {len(lossy)} lossy ({len(lossy) / max(1, len(rows)):.1%}), "
        f"{len(bad)} unreadable ({len(bad) / max(1, len(rows)):.1%}), "
        f"{sum(r.numbers_at_risk for r in rows):,} numbers in unreachable text, "
        f"{sum(r.compare_numbers_missing for r in rows):,} numbers a compared reader drops")
    return "\n".join(lines)
