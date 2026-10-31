"""Command line: `docviews audit`, `docviews text`, `docviews revisions`, `docviews parts`."""

from __future__ import annotations

import argparse
import json
import sys
from dataclasses import asdict
from pathlib import Path

from . import __version__, open_docx
from .audit import SEVERITY, audit_paths, table
from .ooxml import DocxError, Limits
from .views import VIEWS as VIEWS_LIST

CONDITIONS = ("revisions", "unreachable", "strict", "memory", "compare")


def _add_limits(ap: argparse.ArgumentParser) -> None:
    ap.add_argument("--max-part-mb", type=float, default=64.0,
                    help="refuse a single XML part larger than this (MB)")
    ap.add_argument("--max-expansion", type=float, default=400.0,
                    help="refuse parts that expand more than this factor")


def _limits(args) -> Limits:
    return Limits(max_part_bytes=int(args.max_part_mb * 1e6),
                  max_expansion_ratio=args.max_expansion)


def cmd_audit(args) -> int:
    args.fail_on = [c.strip() for c in args.fail_on.split(",") if c.strip()]
    unknown = [c for c in args.fail_on if c not in CONDITIONS]
    if unknown:
        print(f"unknown --fail-on condition(s): {', '.join(unknown)}; "
              f"valid: {', '.join(CONDITIONS)}", file=sys.stderr)
        return 1
    rows = audit_paths(args.paths, compare=args.compare, limits=_limits(args))
    if not rows:
        print("no .docx files found", file=sys.stderr)
        return 1
    if args.json:
        print(json.dumps([asdict(r) for r in rows], indent=1, default=str))
    else:
        print(table(rows, show=args.show))
    if args.ci:
        fired = []
        for r in rows:
            if "revisions" in args.fail_on and (
                    (r.revision_paragraphs and r.views_diverge) or r.hidden_revisions):
                fired.append(r)
            elif "unreachable" in args.fail_on and sum(r.unreachable_chars.values()):
                fired.append(r)
            elif "strict" in args.fail_on and r.strict:
                fired.append(r)
            elif "memory" in args.fail_on and (r.expansion_ratio > 20 or r.largest_part_mb > 5):
                fired.append(r)
            elif "compare" in args.fail_on and (r.compare_numbers_missing or r.compare_error):
                fired.append(r)
        if fired:
            print(f"\n--ci: {len(fired)} file(s) failed "
                  f"[{','.join(args.fail_on)}]", file=sys.stderr)
            return 2
    return 0


def cmd_text(args) -> int:
    with open_docx(args.file, limits=_limits(args)) as doc:
        sys.stdout.write(doc.text(args.view, separator=args.separator) + "\n")
    return 0


def cmd_comments(args) -> int:
    """Print every comment with the characters it is attached to."""
    with open_docx(args.file, limits=_limits(args)) as doc:
        comments = doc.comments(args.view)
        if args.json:
            print(json.dumps([asdict(c) for c in comments], indent=1))
            return 0
        print(f"{len(comments)} comments in {Path(args.file).name} "
              f"({args.view} view)")
        by_id = {c.id: c for c in comments}
        for c in comments:
            if args.roots_only and c.is_reply:
                continue
            flags = []
            if c.resolved:
                flags.append("resolved")
            if c.is_reply:
                flags.append(f"reply to {c.parent_id}")
            if c.reply_ids:
                flags.append(f"{len(c.reply_ids)} repl"
                             + ("y" if len(c.reply_ids) == 1 else "ies"))
            if c.is_orphan:
                flags.append("no anchor in the document")
            head = f"  [{c.id}] {c.author[:28] or 'unknown'}"
            if flags:
                head += f" ({', '.join(flags)})"
            print(head)
            quote = c.anchor_text.replace("\n", " ")[:64]
            if quote:
                print(f"      on: {quote!r}")
            elif c.is_orphan:
                print("      on: (nothing — the anchor was deleted)")
            for anchor in c.anchors:
                if anchor.is_point:
                    print(f"      point anchor at {anchor.part.split('/')[-1]}"
                          f" paragraph {anchor.paragraph_index} "
                          f"offset {anchor.start}")
                elif anchor.part != doc.main_part or len(c.anchors) > 1:
                    print(f"      anchor: {anchor.part.split('/')[-1]} paragraph "
                          f"{anchor.paragraph_index} [{anchor.start}:{anchor.end}]")
            if c.is_reply and c.parent_id in by_id and args.threads:
                print(f"      (reply to: {by_id[c.parent_id].text[:52]!r})")
    return 0


def cmd_revisions(args) -> int:
    with open_docx(args.file, limits=_limits(args)) as doc:
        revs = doc.revisions()
        if args.json:
            print(json.dumps([asdict(r) for r in revs], indent=1))
        else:
            print(f"{len(revs)} tracked changes in {Path(args.file).name}")
            for r in revs[:args.show]:
                where = (f"paragraph={r.paragraph_index}"
                         if r.scope == "text" else f"scope={r.scope}")
                print(f"  {r.kind:<10} {r.author[:18]:<18} {r.date[:16]:<17} "
                      f"id={r.id:<6} {where}")
    return 0


def cmd_parts(args) -> int:
    with open_docx(args.file, limits=_limits(args)) as doc:
        rows = doc.parts()
        print(f"{'part':<34}{'kind':<10}{'paras':>7}{'chars':>9}{'diverges':>10}")
        for p in rows:
            print(f"{p.name[:33]:<34}{p.kind:<10}{p.paragraphs:>7}"
                  f"{p.chars.get('accepted', 0):>9}{str(p.diverge):>10}")
        if doc.strict:
            print("\nISO/IEC 29500 Strict package (python-docx raises on these)")
        for w in doc.warnings:
            print(f"warning: {w}")
    return 0


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="docviews",
                                 description="Read .docx the way Word does.")
    ap.add_argument("--version", action="version", version=f"docviews {__version__}")
    sub = ap.add_subparsers(dest="command", required=True)

    a = sub.add_parser("audit", help="report what a reader loses, per file")
    a.add_argument("paths", nargs="+")
    a.add_argument("--json", action="store_true")
    a.add_argument("--show", type=int, default=25)
    a.add_argument("--compare", default="",
                   help="reader to measure against: python-docx, mammoth, docx2python")
    a.add_argument("--ci", action="store_true", help="exit 2 when --fail-on matches")
    a.add_argument("--fail-on", default="revisions",
                   help="comma list of: " + ", ".join(CONDITIONS) + " (default: revisions)")
    _add_limits(a)
    a.set_defaults(func=cmd_audit)

    t = sub.add_parser("text", help="print the document text in one view")
    t.add_argument("file")
    t.add_argument("--view", default="accepted", choices=["original", "accepted", "markup"])
    t.add_argument("--separator", default="\n")
    _add_limits(t)
    t.set_defaults(func=cmd_text)

    r = sub.add_parser("revisions", help="list tracked changes")
    r.add_argument("file")
    r.add_argument("--json", action="store_true")
    r.add_argument("--show", type=int, default=40)
    _add_limits(r)
    r.set_defaults(func=cmd_revisions)

    k = sub.add_parser("comments", help="list comments with the text they are on")
    k.add_argument("file")
    k.add_argument("--view", choices=VIEWS_LIST, default="accepted")
    k.add_argument("--json", action="store_true")
    k.add_argument("--roots-only", action="store_true",
                   help="skip replies and show only threads' first comments")
    k.add_argument("--threads", action="store_true",
                   help="show the comment each reply answers")
    _add_limits(k)
    k.set_defaults(func=cmd_comments)

    p = sub.add_parser("parts", help="list the parts that carry text")
    p.add_argument("file")
    _add_limits(p)
    p.set_defaults(func=cmd_parts)

    args = ap.parse_args(argv)
    try:
        return args.func(args)
    except (DocxError, OSError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
