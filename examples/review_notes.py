"""Two things every other Python reader gets wrong, on real documents.

Run from the repository root:

    python examples/review_notes.py

Both files are part of the census corpus, so the numbers printed here are the
same numbers `CENSUS.md` reports.
"""

from collections import Counter
from pathlib import Path

from docviews import open_docx

CORPUS = Path(__file__).resolve().parent.parent / "corpus"

COMMENTS = "steps_test_files_comments-rich-para.docx"
ROWS = "scripts_layout-probes_word-tracked-tables.docx"
MARKS = "layout-probes_stops2_word-stops-tracked2.docx"


def show_comments() -> None:
    doc = open_docx(CORPUS / COMMENTS)
    print(f"— {COMMENTS}")
    print(f"  {len(doc.comments())} comments, each with the characters it is about:")
    for comment in doc.comments():
        flags = []
        if comment.resolved:
            flags.append("resolved")
        if comment.is_reply:
            flags.append(f"reply to {comment.parent_id}")
        if comment.reply_ids:
            flags.append(f"{len(comment.reply_ids)} replies")
        if comment.is_orphan:
            flags.append("orphan: its anchor was deleted")
        suffix = f"  [{', '.join(flags)}]" if flags else ""
        print(f"    {comment.author or 'unknown':<14} {comment.anchor_text!r}{suffix}")
        print(f"      {'':<14} on: {comment.text.splitlines()[0][:60]!r}")
    print("  python-docx 1.2.0 gives you the second line, never the first.\n")


def show_hidden_revisions() -> None:
    for name, label in ((ROWS, "deleted table rows"), (MARKS, "paragraph marks")):
        doc = open_docx(CORPUS / name)
        scopes = Counter(r.scope for r in doc.revisions())
        print(f"— {name}")
        print(f"  revisions by scope: {dict(scopes)}")
        hidden = scopes["row"] + scopes["paragraph-mark"] + scopes["block"]
        print(f"  {hidden} of {sum(scopes.values())} are invisible to every "
              f"Python text reader ({label})")
        for reason in _audit_reasons(CORPUS / name):
            print(f"    audit says: {reason}")
        print()


def _audit_reasons(path: Path):
    from docviews.audit import audit_file
    return audit_file(path).reasons


def main() -> int:
    show_comments()
    show_hidden_revisions()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
