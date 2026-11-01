# docviews — positioning (Day 4)

Decided 2026-10-14, after the census (`CENSUS.md`) and the upstream check (§Day 3 below).
This file is the build spec: v0.1 gets implemented straight from here.

---

## Day 3 result — the gap is not closing

| fact | evidence |
|---|---|
| python-docx is dormant | **last commit 2025-06-16** (the v1.2.0 release). 16 months of no commits as of today. Last release 1.2.0, 2025-06-16. |
| the fix PRs are stuck | **#1534** "Add revision management capability" (2,131 additions, 19 files) open since 2026-01-07, last touched 2026-02-03. **#1538** "native support for tracked changes (w:ins, w:del)" open since 2026-01-21, never commented. Both `mergeable_state: clean`, neither merged. |
| the request is 9.8 years old | **issue #340 "revisions/track changes" open since 2016-12-12**, 10 comments, last activity 2026-08-28. In it, a reviewer of #1534 asks for `Paragraph.final_text` — i.e. the accepted view — and gets nothing. |
| the specialist cannot fix it | `docx-revisions` 0.1.5 (61,690 downloads/month, MIT, active 2026-04-14) **imports python-docx** (`RevisionDocument` wraps `docx.Document`). Grep of its source: no `footnote`, no `endnote`, no `txbx`/`textbox`, no `moveFrom`/`moveTo`, comments only as an argument when *writing* a replacement. It inherits every reader-level gap — including the corrupt text: **its own test fixture `OXML_TrackChanges_Test.docx` is one of the 10 documents in my census where python-docx returns text matching neither Word view.** |
| the pain is producing new entrants | **KitchenSink4Word** (Word MCP server, "tracked-changes reading AND writing … 160+ tools", free for noncommercial), repo `KitchenSink4AI/KitchenSink4Word`, pushed 2026-10-13, **9 stars**. Breadth-first server, not a fidelity library — but it says the agent era is now paying for tracked-changes tooling. |
| neighbours are alive but not aiming here | mammoth 1.13.0 (2026-09-26) 15.9M/mo · docx2python 3.7.1 (2026-08-06) 274k/mo · redlines 0.6.2 (2026-08-28) 338k/mo (write-side revisions) · aspose-words 26.10.0 (2026-10-13) 348k/mo · Spire.Doc 14.9.1 (2026-09-29) 66k/mo. None of them offers a view-selectable, full-part, auditable reader. |

**Conclusion:** the window is open, the incumbent is quiet, and the closest specialist is structurally unable to close it (it depends on the thing that is broken).

## 1. Name and one-liner

**Name: `docviews`** (PyPI free as of 2026-10-14; fallbacks already screened: `docxlens`, `ooxmlkit`, `trueview`, `docfidelity`, `unrevision`).

> **docviews — read .docx the way Word does. All three views, every part, zero dependencies.**
> Tracked changes, comments and footnotes aren't edge cases, they're where your text goes missing.

One-line pitch for a pipeline owner:

> **`docviews audit` tells you exactly what your current reader is losing, file by file, before you change a line of code.**

## 2. What v0.1 is (and is not)

**Is:** a read-only, dependency-free OOXML reader with view primitives, every text-bearing part, revision/comment metadata, and an audit CLI. Stdlib only (`zipfile`, `xml.etree`) so it installs anywhere and can be vendored.

**Is not:** a writer. No `add_paragraph`, no `save`. Writing is a different failure surface; `redlines` and `docx-revisions` already do it and interop beats competition here (`docviews` reads what they write — and proves it, see §6).

Why read-only is the right v0.1: it is the entire census result. Every measured loss happens on the reading side.

## 3. API spec (v0.1)

```python
from docviews import open_docx

doc = open_docx("contract.docx")

doc.text(view="accepted")          # what Word shows with changes accepted
doc.text(view="original")          # what Word shows with changes rejected
doc.text(view="markup")            # everything the file contains, deleted text included

for p in doc.paragraphs(view="accepted"):
    p.text, p.part, p.revision_ids, p.part_name

for r in doc.revisions():          # every w:ins/w:del/w:moveFrom/w:moveTo
    r.kind, r.author, r.date, r.id, r.text, r.paragraph_index, r.is_move

for c in doc.comments():           # anchor offsets included — the gap python-docx documents
    c.id, c.author, c.date, c.text, c.resolved, c.parent_id,
    c.anchor_start, c.anchor_end, c.anchor_text

doc.parts()                        # body, footnotes, endnotes, header/footer[i], comments, textboxes
doc.numbers(view="accepted")       # [(value, part, char_offset)] — for diffing reader vs document
doc.warnings                       # unreadable parts, ISO-Strict detection, size ratios
doc.stream()                       # yields paragraphs without holding the tree (the 400 MB fix)
```

Rules the implementation must hold:

1. **No silent omission.** Anything not read is reported in `doc.warnings`, never dropped quietly. That is the whole complaint against the incumbent.
2. **Views are exact.** `original` = reject-all, `accepted` = accept-all, `markup` = everything; `text()` for any view is a pure function of the file.
3. **Every text-bearing part is read**: body, tables (incl. nested), headers, footers, footnotes, endnotes, text boxes, comments, and `w:sdt` content controls.
4. **ISO-Strict detected, not crashed on.** `purl.oclc.org/ooxml` namespaces are read through a namespace map (61/832 of my corpus, including 60 Microsoft conformance files, are unreadable by python-docx today).
5. **Bounded memory.** `stream()` + a documented expansion guard (`zip` ratio, part size) so a 0.5 MB upload cannot cost 400 MB.

## 4. `docviews audit` — the wedge

```
docviews audit PATH [--json] [--ci] [--deep] [--compare python-docx,mammoth,docx2python]
                    [--fail-on revisions,footnotes,textboxes,strict,memory]
```

PATH is a file or a directory tree. Default output is a human table; `--json` for pipelines; `--ci` sets exit code 2 when a `--fail-on` condition is met (so a team can gate deployments on it).

Per file the audit reports:

| column | meaning |
|---|---|
| `openable` | can it be read at all, and by whom (`--compare`) |
| `parts` | which text-bearing parts exist, and their character counts |
| `revision_text` | characters that exist only in `w:ins`/`w:del` — invisible to `.runs`/`.text` today |
| `views_diverge` | original vs accepted character delta; `impossible` flag when a reader returns neither |
| `numbers_at_risk` | numbers present in the document that a compared reader does not return |
| `unreachable` | chars in parts the compared reader has no API for |
| `expand_ratio` | uncompressed/compressed + peak RSS estimate (the DoS column) |
| `strict` | ISO-Strict package |
| `verdict` | `clean` / `lossy` / `unreadable` + why |

Aggregate line for a directory: *"412 files: 38 lossy (9.2 %), 4 unreadable (1.0 %), 1,204 numbers at risk."* That sentence is the product: it is the number a team cannot get today without writing this tool themselves.

## 5. Fidelity benchmark (publishable, rerunnable)

`benchmarks/fidelity.py` — same corpus, every reader, one CSV, method in the repo:

| dimension | docviews | python-docx 1.2.0 | mammoth | docx2python | docx-revisions | aspose-words | Spire.Doc |
|---|---|---|---|---|---|---|---|
| three views (original/accepted/markup) | ✓ | partial (accidental mix) | ✗ | ✗ | accept/reject only | ✓ (API) | ✓ (API) |
| revision metadata (author/date/id/range) | ✓ | ✗ | ✗ | ✗ | ✓ | ✓ | ✓ |
| comment anchor ranges | ✓ | ✗ (documented) | ✗ | ✗ | ✗ | ✓ | ✓ |
| footnotes / endnotes | ✓ | ✗ | ✓ | ✓ | ✗ | ✓ | ✓ |
| text boxes | ✓ | ✗ | ✓ | partial | ✗ | ✓ | ✓ |
| headers / footers | ✓ | ✓ | ✓ | ✓ | ✗ | ✓ | ✓ |
| ISO-Strict packages | ✓ | ✗ (raises) | ? | ? | ✗ | ? | ? |
| dependencies | none | lxml | many | many | python-docx | binary | binary |
| license / price | MIT | MIT | BSD-2 | MIT | MIT | $1,199+/yr | $599+ |

The commercial columns are measured from the same corpus via their pip packages (`aspose-words`, `Spire.Doc`) in an isolated env — trial builds are enough to answer "does the text come out", and the method will be published so the row can be challenged. **The point of the table is not that free beats paid: it is that no reader in Python gives you the views, the parts and the audit in one place, and the ones that come closest are the ones that cost money.**

## 6. Where docviews sits next to the neighbours

* **python-docx** — the thing being replaced, at the reading layer; the API is deliberately familiar so migration is a rename and a view argument.
* **docx-revisions / redlines** — write-side. They generate revisions; docviews reads any document they produce and can *verify* it, including the moves and footnote text they cannot see. Integration, not competition: a `docviews` check in their test suites would have caught the fixture that currently returns impossible text.
* **mammoth / docx2python** — converters. Different job (render to HTML/text). When you need the document's text as data and you need to know what you dropped, that's docviews.
* **Aspose / Spire** — paid, binary, Windows-friendly, and genuinely capable in this space. If your constraint is "must be pure Python, MIT, auditable, and stdlib-only", they are not an option, and that constraint is common in the places where documents are sensitive.

## 7. Monetization (unchanged from the memo, now with a product to sell)

Library stays MIT and free. Money comes from the work the library makes possible:

1. **Integrity reports** — run the audit across a client's document estate (contracts, filings, HR records, protocols) and deliver the loss map + remediation. Priced against the incumbents' per-developer fees ($599–$1,199/yr), not against free OSS.
2. **Support / migration** — "your pipeline is dropping footnote text" engagements: measure, fix, prove with the same tool.
3. **Sponsorship/funding** — the census and the benchmark are the credibility assets that make that ask land.

## 8. Launch narrative

> python-docx is the most-installed document library in Python (95M downloads/month) and it has not had a commit in sixteen months. In the documents that contain tracked changes, it returns text that exists in no version of the document — Word has no setting that produces it. I measured 832 public documents: 9.0 % cannot be opened by it at all, 20.2 % hold text no API of its can reach, and a 0.49 MB upload costs 416 MB of RAM. Here is the proof, the mechanism, and `docviews`: three views, every part, zero dependencies — plus an audit that tells you what your own pipeline is losing.

Assets, in order: `CENSUS.md` (evidence) → audit output on a real file (demo) → fidelity table (comparison) → library (adoption).

## 9. v0.1 build order (~3 weeks of evenings, compressed if full-time)

1. `ooxml.py` — namespace map (Transitional + Strict), part loader, zip guard. *Tests: strict file opens.*
2. `views.py` — the three-view walker with exact python-docx semantics for `w:tab`/`w:br`/`w:noBreakHyphen`; runs, tables, nested tables, `w:sdt`. *Tests: the 10 impossible-text documents from the census become correct.*
3. `revisions.py`, `comments.py` — metadata, anchor offsets, resolved/threads; footnotes/endnotes/textboxes as parts. *Tests: counts match `census2.jsonl`.*
4. `audit.py` + CLI — the table above, `--json`, `--ci`, `--compare`. *Tests: exit codes; report matches census numbers on the same corpus.*
5. `benchmarks/fidelity.py` — the published table.
6. Packaging: stdlib-only, `py.typed`, Python 3.9+, no deps. Then launch.

**Definition of done for v0.1:** running `docviews audit` on the 832-document corpus reproduces the numbers in `CENSUS.md` (that is the regression test), and every one of the 10 impossible-text documents returns a text that matches one of the two views Word can show.
