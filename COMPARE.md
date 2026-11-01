# docviews vs python-docx — the measured comparison

*Status updated 2026-11-01. The main census and head-to-head numbers below are
from the 832-file fixture corpus and were produced by scripts in this repository;
each script is named next to its measurement. The support/adoption scorecard
remains a 2026-10-31 snapshot. Documented-API parity and stable version metadata
were updated on 2026-11-01; the selected 1.0.0 distribution remains unpublished.
An external public-DOCX smoke set was added on 2026-10-29 (§10). Where a number
favours python-docx, it is the first one in the row.*

## The short version

| | python-docx 1.2.0 | docviews 1.0.0 metadata (unpublished) |
|---|---|---|
| What it is | the standard Python `.docx` library, 10 years old | a read-fidelity OOXML library that also writes, with a compat door for python-docx code |
| Downloads | **95.4M / month** (2026-10-14) | 1.0.0 metadata selected; distribution remains unpublished |
| Licence | MIT | MIT |
| API coverage | — | **578 tested documented behaviors**, 0 partial/not-assessed in-scope entries, 6 internal/scope exclusions; the separate sample audit records 276/276 names, 78/78 call shapes, and 471/471 runtime enum members |
| Open → save unchanged (fixture corpus) | 762 / 832 files open; **7** keep every part | **831 / 832** open; **13,912 / 13,912 parts byte-identical** |
| Text a reader sees | loses **15.9 %** of the numbers in the 167 tricky documents (revisions or text outside the body) | reference reader (0 by construction) |
| Dependencies | lxml + typing_extensions | **none** |
| Install | **252,987 B wheel** + 12 MB lxml | **about 144 KiB wheel** |

For the twelve documented recipes in `examples/docs_recipes.py`, both
libraries produced byte-identical description transcripts after save/reopen.
That validates those flows, not all python-docx signatures or behaviors. A
migration changes the relevant import paths (shown below) and still needs
application-specific testing. Code that uses python-docx as a thin veneer over
lxml requires a rewrite. Everything in between is below.

---

## 1. The API: what a switch actually costs

`python3 compat_coverage.py` builds corresponding samples with both libraries,
then compares non-underscore, non-structural public names plus selected method
and constructor parameter shapes.

| sampled objects | python-docx attributes | answered by docviews |
|---|---:|---:|
| Document | 17 | 17 |
| Paragraph / ParagraphFormat | 12 / 13 | 12 / 13 |
| TabStops / TabStop | 2 / 3 | 2 / 3 |
| Run / Font / ColorFormat | 14 / 27 / 3 | 14 / 27 / 3 |
| Table / Row / Rows / Cell / Column / Columns | 12 / 6 / 1 / 10 / 3 / 1 | 12 / 6 / 1 / 10 / 3 / 1 |
| Section / Sections / Header / Footer | 19 / 2 / 6 / 6 | 19 / 2 / 6 / 6 |
| Settings / Styles / Style | 1 / 5 / 14 | 1 / 5 / 14 |
| LatentStyles / LatentStyle | 7 / 7 | 7 / 7 |
| InlineShape / InlineShapes / Drawing / Image | 3 / 0 / 2 / 13 | 3 / 0 / 2 / 13 |
| Length / RGBColor / CoreProperties | 17 / 3 / 15 | 17 / 3 / 15 |
| Comments / Comment / CommentParagraph | 2 / 10 / 12 | 2 / 10 / 12 |
| Hyperlink / RenderedPageBreak | 6 / 2 | 6 / 2 |
| **total** | **276** | **276 (100 %)** |

On these samples, **78/78 method and constructor parameter shapes** (names,
order, kinds, defaults) also match; annotations/return annotations are not
compared. Focused behavior regressions now cover latent-style default and item
properties, run content iteration and rendered-page-break fragments, Drawing/Image
access, and shared unit/color values. This remains a bounded inventory, not full
documented-API parity: special methods are excluded from the attribute total,
not every documented object or behavior is sampled, and `docx.oxml`/lxml/OPC
APIs remain outside the chosen target.
`COMPAT.md` holds the detailed scope and `tests/test_compat_parity.py` contains
focused regression checks.

And the names imported *beside* a document, which is where most migrations
actually break:

| module | classes | members |
|---|---:|---:|
| `docx.enum.text` | 11 / 11 | 117 / 117 |
| `docx.enum.style` | 5 / 5 | 268 / 268 |
| `docx.enum.section` | 7 / 7 | 20 / 20 |
| `docx.enum.table` | 8 / 8 | 19 / 19 |
| `docx.enum.shape` | 2 / 2 | 10 / 10 |
| `docx.enum.dml` | 5 / 5 | 37 / 37 |
| `docx.enum.base` | 2 / 2 | — |
| `docx.shared` | 10 / 10 | — |

The enum classes are real enums built on python-docx's own two base classes, so
the whole documented protocol works — `list()`, `WD_ALIGN_PARAGRAPH(1)`,
`__members__`, `isinstance`, `.xml_value`, `from_xml()`, `to_xml()` — and
`value`, `str()` and `repr()` are identical to python-docx's for **all 471
members** (`tests/test_enums.py`, 0 unexplained divergences).

The switch, verbatim:

```diff
- import docx
- from docx.shared import Pt, Inches
- from docx.enum.text import WD_ALIGN_PARAGRAPH
+ import docviews.compat as docx
+ from docviews.compat.shared import Pt, Inches
+ from docviews.compat.enum.text import WD_ALIGN_PARAGRAPH
```

The program logic in this example does not change; imported module paths do.
`examples/docs_recipes.py` runs twelve recipes from python-docx's own
documentation under both libraries, then re-opens and describes styles, text,
alignment, run formatting, table and section geometry, header/footer paragraphs,
core properties, comments, and inline shapes. **Those two transcripts are
byte-identical.** This is a focused recipe check, not proof of equivalence for
all inputs or call sequences.

## 2. Reading: what each one can see

`readers.py --only-interesting` reads the **167** corpus files the census
flagged — those with tracked changes or text outside the document body — with
four readers, and compares the *numbers* each reports against the numbers in the
file's own XML: 92,158 of them, counted across every text-bearing part.

| reader | numbers lost | numbers invented | files affected |
|---|---:|---:|---:|
| **docviews** (reference) | 0 | 0 | 0 / 167 |
| python-docx | **14,640 (15.9 %)** | 0 | 131 / 167 |
| mammoth | 1,742 (1.9 %) | 5,118 | 150 / 167 |
| docx2python | 189 (0.2 %) | 3,227 | 149 / 167 |

python-docx loses nothing by reading badly; it loses by **never looking**:
footnotes, endnotes, text boxes, and text inside tracked changes. In the worst
file it reports 72 of 1,802 numbers. The census counted **11,883 endnote
characters** in 832 documents that python-docx cannot reach at all.

Tracked changes are the sharpest difference, because a change is a *view*:

```
the same paragraph, three views (docx-revisions fixture):
  docviews  accepted : 'We prefer  color spelling.'      ← what Word shows
  docviews  original : 'We prefer colour  spelling.'     ← before the edit
  docviews  markup   : 'We prefer colour color spelling.' ← both, in order
  python-docx        : 'We prefer   spelling.'           ← both words vanish
```

docviews reports the change, its author, its date and its scope (run, paragraph
mark, row) — **255 revision containers** in the corpus, 121 of them in paragraph
marks or row properties that no Python text reader reports. Comments keep their
anchors: **999 comments, 1,189 anchors**, where python-docx 1.2.0 returns
comment text but no reference range (its maintainer has said the range is "not
currently planned").

Reading the same file, the earlier measurement on the largest corpus document
(8,152 paragraphs; benchmark dated 2026-10-17):

| | time | chars captured | peak RSS |
|---|---:|---:|---:|
| python-docx | 0.45 s | 821,329 | 93 MB |
| docviews | **0.44 s** | **823,292** | **63 MB** |

Same speed, two-thirds of the memory, more of the document. `doc.stream()` can
walk the body with `iterparse` and hold one paragraph at a time when memory
matters more than speed.

**Write microbenchmark (`tools/bench_compat_write.py`, run 2026-10-30):**
Python 3.13.16, python-docx 1.2.0; five-run medians. Each timed case builds N
paragraphs and a populated 20×4 table, then saves to an in-memory stream. The
saved package is reopened with both libraries, and every paragraph and table
cell is checked against the expected text.

| paragraphs | library | build | save | build + save | package |
|---:|---|---:|---:|---:|---:|
| 500 | python-docx | 0.132 s | **0.010 s** | 0.142 s | 38,502 B |
| 500 | docviews.compat | **0.014 s** | 0.017 s | **0.031 s** | 4,931 B |
| 5,000 | python-docx | 0.558 s | **0.014 s** | 0.571 s | 50,894 B |
| 5,000 | docviews.compat | **0.321 s** | 0.117 s | **0.438 s** | 17,253 B |

In this workload docviews is faster end-to-end at both sizes, while its 5,000-
paragraph save is still slower by itself. This is a controlled microbenchmark,
not a universal performance claim; the libraries use different package
starting points. The earlier append optimization also removed an O(n²)
paragraph-index walk (11.9 s down to 0.47 s at 5,000 appends in its original
benchmark), with a regression test.

## 3. Writing: what a save costs

Two measurements, same corpus, same loop.

**Open → save unchanged, then diff the packages** (`compare_roundtrip.py`):

| | docviews | python-docx |
|---|---:|---:|
| files opened | **831 / 832** | 762 / 832 |
| files that keep every package entry | **831** | 523 |
| files that keep every part byte-identical | **831** | **7** |
| parts byte-identical | **13,912 / 13,912** | 5,398 / 12,042 |
| parts meaning-same | **13,912 / 13,912** | 10,922 / 12,042 |

The single docviews failure is a password-protected document — refused with a
sentence, not a traceback. python-docx's 70 failures are ISO/IEC 29500 **Strict**
packages it cannot open; docviews reads them.

**What a save does to the *content* you care about** (`compare_save_impact.py`,
832 files, text + revisions + comments + anchors + footnotes):

| | files saved | documents whose content changed | documents that lost package entries |
|---|---:|---:|---:|
| docviews | 831 | **0** | **0** |
| python-docx | 762 | 0 | **239** |

python-docx does not corrupt text — it re-serialises every part it parsed,
which drops what it did not parse (footnote, endnote, custom-XML and media
parts it never loaded). docviews copies unedited parts verbatim and only
rewrites what was edited.

The v0.3 authoring scope includes paragraphs, runs, fonts and flags, styles,
tables (widths, merges, row rules, nested), sections (margins, orientation,
start type, gutter, first/even headers), headers and footers, comments with
anchor ranges, core properties, and inline images. The test suite and interop
sweep (39 / 42; the three refusals are Strict packages python-docx cannot open)
cover representative round trips, not every combination. `tests/test_compat_parity.py`
now specifically checks `Run.add_picture()` in body, header, and footer parts,
including existing header relationships, repeated PNGs, correct relationship
targets, and save/reopen with python-docx. It also checks Python-docx keyword
calls for run breaks, picture sources, table indices, character styles, and
saving to a writable stream.

## 4. Size, dependencies, install

| | python-docx | docviews |
|---|---|---|
| wheel | 252,987 B | **about 144 KiB** |
| installed | 2.8 MB + lxml 12 MB | **0.9 MB** |
| dependencies | `lxml`, `typing_extensions` | **none** (stdlib `zipfile` + `xml.etree`) |
| XML layer | lxml (C, fast, exposes elements) | ElementTree (slower, no element API) |
| platforms | needs an lxml wheel | pure Python, any platform |

Zero dependencies is a deployment argument (Lambda layers, locked-down images,
Pyodide) and a fidelity argument: there is no XML library between the file and
you that can normalise, reorder or rename something.

## 5. Where python-docx is still the right answer

| if your code… | use |
|---|---|
| builds XML by hand: `OxmlElement("w:br")`, `qn("w:val")`, `.element`, `.xpath`, `doc.part` | **python-docx** — docviews does not expose the lxml tree, by design |
| is a dependency of another library that does `import docx` | **python-docx** today — docviews' compat door is `import docviews.compat as docx`; it does not install a `docx` module |
| needs the largest possible stack of blog posts, StackOverflow answers, LLM training data | **python-docx** — 95.4M downloads/month against a project with no public release at the 2026-10-31 snapshot |
| needs floating/anchored shapes, charts, fields, symbol glyphs, TOC objects, `.doc` input or rendering *authored* | **neither** — python-docx does not author these either; docviews reads and preserves them byte-for-byte |

## 6. Where docviews is the only option

| capability | python-docx | docviews |
|---|---|---|
| original / accepted / markup views of the same paragraph | no | `doc.text("original"/"accepted"/"markup")` |
| tracked-change list with author, date, scope | no | `doc.revisions()` (255 containers in the corpus) |
| comment anchor ranges | no (maintainer: not planned) | `comment.anchors`, 1,189 measured |
| footnotes, endnotes, text boxes in the text | no | yes (11,883 unreachable chars recovered) |
| ISO/IEC 29500 Strict packages | refuses to open | opens |
| byte-identical save | no (7 / 832) | yes (831 / 831) |
| save that cannot drop an unparsed part | no | yes |
| streaming read (`iterparse`, one paragraph at a time) | no | `doc.stream()` |
| audit a folder of `.docx` files, no code | no | `docviews audit .` CLI |
| 45 extra names on the same objects (`editable`, `revisions`, `in_textbox`, `markup_text`…) | — | yes |

## 7. Maturity, honestly

| | python-docx | docviews |
|---|---|---|
| first release | 2013 | 2026 (v0.1) |
| latest | 1.2.0, 2025-06-16 | 1.0.0 stable metadata selected 2026-11-01; unpublished |
| use in the wild | 95.4M downloads/month | no public user base recorded at that snapshot |
| documentation | book-length, plus a decade of Q&A | README, compatibility and migration guides, contributor guide, census, and comparison; source is public, and two hosted CI runs passed |
| test suite here | — | 239 tests, corpus-backed; full suite passed on CPython 3.9–3.15 |
| upstream activity | dormant since 2025-06; tracked-changes PRs #1534/#1538 open since Jan 2026, review stalled | active |

The last row is the reason this project exists: python-docx's reading gaps are
not bugs that will be fixed next month — issue #340 has been open since
December 2016, and the maintainer has said comment reference ranges are not
planned. A library that refuses to answer a question for nine years is a
library with a fixed ceiling.

## 8. Provenance

| claim | produced by | when |
|---|---|---|
| sampled object/module names, callable parameter shapes, enum parity | `compat_coverage.py` → `COMPAT.md` | 2026-10-31; 276/276 names, 78/78 shapes, 471/471 enum members |
| documented recipe transcript diff | `examples/docs_recipes.py` ×2 | 2026-10-31; diff empty |
| targeted compatibility regressions | `tests/test_compat_parity.py` + `tests/test_authoring_v03.py` | 2026-10-31, 32 passed |
| round-trip preservation | `compare_roundtrip.py` (831 files) | 2026-10-17 |
| save impact | `compare_save_impact.py` (832 files) | 2026-10-17 |
| reader comparison | `readers.py` → `readers.jsonl` (167 files) | 2026-10-14 |
| revisions, comments, anchors, unreachable chars | `census2.py` → `CENSUS.md` §6 | 2026-10-14 |
| time / peak RSS | `/tmp/bench_one.py`, same file, separate processes | 2026-10-17 |
| build/save microbenchmark | `tools/bench_compat_write.py` (five-run medians, outputs reopened by both libraries) | 2026-10-31 |
| downloads | PyPI API | 2026-10-14 |
| suite | `python -m pytest` (`.[dev]`, python-docx ≥1.2.0) | 2026-11-01; 239 passed on each tested Python 3.9–3.15 patch release (seven versions) |
| external public-DOCX smoke set | `real_world_eval/download.py` + `real_world_eval/analyze.py` → `real_world_eval/analysis.json` | 2026-10-29, 8 official DOCX files |

The suite run used `python-docx==1.2.0`. The host interpreter had 1.1.2 and
failed two interoperability tests because `Document.comments` was not available;
the latest run in the compatible environment passed all 239 tests. The `[dev]`
extra now requires `python-docx>=1.2.0`, so a normal development install selects
the needed API while the published runtime package remains dependency-free.

**Caveat that applies to all of it:** the 832-file census is fixture-heavy —
python-docx's, mammoth's, docx4j's and Open XML SDK's own test suites. It
measures *preservation* extremely well and *"every real Word file on earth"* not
at all. The eight-file public-sector smoke set in §10 is the first external
package-preservation evidence, but it is a small, UK-government-heavy convenience
sample with no Word-rendered text gold-standard. A larger and more diverse
real-world corpus remains the largest open question.

---

## 9. The rating

Ten dimensions, 1–10, each tied to a number already in this file. A score is not
a preference — where a score is judgement rather than measurement, the reason is
under the table.

| dimension | python-docx | docviews | on what basis |
|---|:---:|:---:|---|
| API compatibility with python-docx code | **10** | 8 | The full documented-public-API matrix records 578 in-scope behavior matches and 6 structural/internal exclusions; the separate bounded audit has 276/276 sampled object names, 78/78 sampled call shapes, and 471/471 runtime enum members. `docx.oxml`, `document.element`, and `doc.part` remain outside the chosen target |
| reading fidelity | 6 | **10** | 84.1 % of corpus numbers vs 100 % in the earlier fixture census; no views, no revisions, no comment anchors, no footnotes/endnotes/text boxes |
| save safety | 5 | **10** | Earlier 832-file run: 7 python-docx files keep every part byte-identical vs 831 docviews; 239 python-docx files lose package entries |
| authoring breadth | **9** | 8 | both author paragraphs→styles→tables→sections→images; python-docx still has more mature edge-case coverage and a larger regression history |
| speed | 8 | **9** | 2026-10-31 build+save microbenchmark: docviews 0.031 vs 0.142 s (500 paragraphs) and 0.438 vs 0.571 s (5,000), with python-docx faster at saving alone |
| dependencies / footprint | 6 | **10** | lxml + typing_extensions + 12 MB vs none; 252,987 B vs about 144 KiB per wheel |
| maturity / ecosystem | **10** | 3 | 95.4M downloads/month since 2013 vs stable 1.0.0 metadata that remains unpublished; the repo and issue tracker are new and no adoption data is recorded at the 2026-10-31 recheck |
| documentation / support | **9** | 6 | python-docx has a book-length manual and established Q&A; docviews has migration/contributor guides and measured compatibility docs, but its public issue tracker is new and has no support history or SLA |
| verification / tests | 6 | **9** | 239 tests passed on each tested CPython 3.9–3.15 patch release (seven versions), plus an 8-file public-DOCX round-trip smoke set; no independent rendered-text gold set |
| openness (Strict, plain XML, streaming) | 6 | **9** | ISO-Strict opens here and is refused there; no C XML layer; `iterparse` streaming |
| **unweighted mean** | **7.50** | **8.20** | |

Weighted for three real situations (weights sum to 100):

| who you are | python-docx | docviews | winner |
|---|:---:|:---:|---|
| **A. migrating an existing python-docx codebase** (API 35, maturity 20, save 15, fidelity 10, docs 10, speed 5, deps 5) | **8.45** | 7.45 | **python-docx — by 1.00** |
| **B. starting a new document pipeline** (fidelity 25, save 20, API 10, deps 10, speed 10, maturity 10, openness 10, docs 5) | 6.95 | **8.70** | **docviews — by 1.75** |
| **C. reading or auditing documents** (fidelity 40, openness 15, save 10, deps 10, speed 10, maturity 10, docs 5) | 6.65 | **8.85** | **docviews — by 2.20** |

**Why python-docx still wins profile A.** It is the self-compatibility baseline,
and the current docviews matrix is still incomplete. It also wins because a
migration is an *organisational* act: the library you already run has 95.4M
downloads a month, every error message you might hit is already on the
internet, and a junior developer has heard of it. That is worth 1.00 point and
it is worth being honest about. The counter is narrow: if your code uses the
tested documented subset, the migration guide and recipe diff reduce the work,
but neither erases the maturity gap or guarantees that an untested application
will switch cleanly.

**Why docviews wins profiles B and C.** These profiles weight fidelity and save
safety heavily, where docviews leads 10–6 and 10–5. Its advantages in
zero-dependency deployment, Strict-package support, streaming, and the measured
build+save workload outweigh its lower API-compatibility, documentation, and
ecosystem scores in those use cases. Profile C is **8.85 vs 6.65**. For reading,
python-docx's losses are structural: it has no concept of original/accepted
views, revisions, or comment anchors, and its maintainer has said the anchors
are not planned. This is a technical-capability comparison, not an adoption
claim.

**What would move the scores.** In order of size:

1. **Complete the documented-API matrix.** The latest 276/276 names and 78/78
   parameter shapes are still a selected sample; special methods, import paths,
   and behavior across all documented objects/states are not fully audited.
   Close that evidence gap before raising the API score above 8.
2. **Publish it, then earn adoption.** A verified maintainer identity, explicit
   version choice, authorized release, public source/issue tracker, and external
   CI are prerequisites. They could justify moving `maturity` above 3, but a
   PyPI listing alone would not establish a strong ecosystem; raise the score
   only after sustained external use, issue traffic, and releases.
3. **Scale and diversity of the real-world corpus.** The 8 public DOCX files in
   `real_world_eval/` add external package-preservation evidence, but all come
   through GOV.UK and most are HM Treasury forms. A few hundred documents from
   diverse real workflows—ideally with independently rendered references—could
   confirm or knock down the reading-fidelity score. User-private files remain
   out of scope unless the user explicitly asks to audit them.
4. **ISO-Strict writing.** Reading Strict is a point of differentiation; writing
   it is not implemented, and it is the largest unknown in `save safety` for
   that class of file.
5. **The `docx` module question.** The compat door is `import docviews.compat as
   docx`. Shipping a `docx` alias package (opt-in) would let libraries that
   `import docx` internally ride along, which no amount of coverage otherwise
   buys.

**One-line verdict.** On the unweighted rubric docviews scores **8.20** to
python-docx's **7.50**, and it leads for new pipelines (**8.70 vs 6.95**) and
reading/auditing (**8.85 vs 6.65**). python-docx still leads the migration
profile (**8.45 vs 7.45**) because ecosystem maturity and support history matter
there. The documented-public-API matrix is now complete within the agreed
boundary: 578 in-scope behaviors match and six structural/internal records are
excluded. This establishes API-contract parity with python-docx 1.2.0, not
adoption, an SLA, undocumented internals, or behavior outside the documented
contract. Stable `1.0.0` metadata is selected but unpublished; broader
independent validation and an established user base remain separate concerns.

---

## 10. First external public-DOCX smoke set (2026-10-29)

The fixture census above is still the main evidence base. To test whether its
package-preservation result carried over beyond documents made for library test
suites, `real_world_eval/` adds eight public-sector documents fetched from their
official GOV.UK asset URLs: six HM Treasury guidance/forms, one DWP consent form
(now withdrawn), and one Office of the Advocate General for Scotland document.
The four GOV.UK source pages' standard footer says all content is available
under the Open Government Licence v3.0, except where otherwise stated; no
contrary document-specific reuse terms were found on these pages. The licence
still has exclusions, including third-party rights. The source, licence and
terms URLs, direct download URL, HTTP response, size, and SHA-256 are in the CSV
and JSON manifests.

The 2026-10-29 run's `docviews 0.3.0` label records the version reported by
that run; it is historical metadata. Stable version `1.0.0` was selected on
2026-11-01 after documented-API parity was verified; it remains unpublished.

| unchanged open → save | docviews 0.3.0 | python-docx 1.2.0 |
|---|---:|---:|
| valid packages opened | **8 / 8** | **8 / 8** |
| package entries retained | **248 / 248** | 246 / 248 |
| unchanged common package parts (byte-identical) | **248 / 248** | 128 / 246 |
| documents omitting package entries | **0 / 8** | 1 / 8 |

The one python-docx omission was two unreferenced `[trash]/*.dat` ZIP entries
from the WGA CG-05 template. This is a package-preservation finding, not a claim
of reader-visible content loss. docviews retained all package entries byte for
byte in all eight files; python-docx changed the bytes of 118 common parts while
saving. The sample contains custom XML and media and is more realistic than a
minimal generated document, but it is small and heavily biased toward UK
Treasury forms. It is not a Word-rendered text gold-standard, and it does not
move the ratings above.

A GOV.UK page also mislabeled a PDF asset with a `.docx` URL. The downloader
checks ZIP signatures and rejects it; the details and checksum are in
[`real_world_eval/REPORT.md`](real_world_eval/REPORT.md). The restricted World
Bank candidate and the NZ Ombudsman document that requires permission are not in
the corpus. No user-private files were inspected. This external smoke set is
separate from `corpus/`, and does not change its 832-file counts.

Reproduce from the project root with `python3 real_world_eval/download.py
--force` and `PYTHONPATH=src python3 real_world_eval/analyze.py` (using
`python-docx==1.2.0` for the comparison). Detailed caveats and per-file results:
[`real_world_eval/REPORT.md`](real_world_eval/REPORT.md); raw metrics:
[`real_world_eval/analysis.json`](real_world_eval/analysis.json).
