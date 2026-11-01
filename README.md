# docviews

**Read `.docx` in original, accepted, and markup views across text-bearing parts; author common content; zero runtime dependencies.**

**Documented-API parity:** verified against python-docx 1.2.0 within the agreed
boundary: documented public behavior is differentially tested, while
undocumented lxml, OOXML, and OPC internals are excluded. The source-backed
584-row matrix records **578 tested behavior matches, 0 partial/not-assessed
in-scope rows, and 6 structural/internal scope exclusions**. All 471 runtime
enum members match; the documented-only `MSO_THEME_COLOR_INDEX.MIXED` is
implemented with the published value `-2`. The full 239-test suite passed on
each tested CPython patch release from 3.9.25 through 3.15.2. See
[`API_PARITY_MATRIX.md`](API_PARITY_MATRIX.md) for scope and evidence. The
stable release **1.0.0** is published on PyPI. Install it with
`python -m pip install docviews`.

```python
from docviews import open_docx

doc = open_docx("contract.docx")
doc.text("accepted")     # what Word shows with changes accepted
doc.text("original")     # what Word shows with changes rejected
doc.text("markup")       # everything the file contains
for c in doc.comments():
    c.anchor_text        # the exact characters under review
for r in doc.revisions():
    r.kind, r.author, r.scope       # scope: text | paragraph-mark | row | block
```

Writing is the same object:

```python
from docviews import Document, Inches

doc = Document.new()                      # or open_docx("existing.docx")
doc.add_heading("Q3 Report", 0)
para = doc.add_paragraph("Revenue grew ")
para.add_run("18%").bold = True
doc.add_table(rows=2, cols=2, text=[["Region", "Revenue"], ["EMEA", "1,204,332"]])
doc.add_picture("chart.png", width=Inches(5))     # media part + relationship, wired up
doc.sections[0].footer.add_paragraph("Confidential")
doc.save("q3.docx")                       # atomic: never a half-written file
```

Already written against python-docx? If your code imports the top-level
`docx` module and uses the supported documented object API, start with this
change:

```python
import docviews.compat as docx            # instead of: import docx
```

If it also imports `docx.shared` or `docx.enum.*`, change those import paths to
`docviews.compat.shared` and `docviews.compat.enum.*`. Documented proxy classes
also have mirrored modules under `docviews.compat`; the full path inventory is
in [`API_PARITY_MATRIX.md`](API_PARITY_MATRIX.md). The documented public API
matrix records 578 tested behavior matches and no partial or unassessed
in-scope behaviors; six structural/internal entries are excluded by the agreed
boundary. The documented-only `MIXED` enum value is implemented separately from
the upstream runtime omission. This is parity for the documented contract, not
for undocumented internals or malformed states outside that contract.

Then a CLI that tells you what your *current* reader is losing:

```console
$ docviews audit contracts/
file                                 verdict     revs  unreach  expand  why
------------------------------------------------------------------------------
tracked-changes.docx                 lossy          7        0      6x  7 paragraphs change under acceptance (+5 inserted, -68 deleted characters)
tracked-tables.docx                  lossy          0        0      6x  9 table rows marked deleted (no text reader reports this)
un-report.docx                       lossy          0     9629      6x  text no python-docx API returns: footnotes 9,629 chars
strict.docx                          lossy          0        0      4x  ISO/IEC 29500 Strict package (python-docx raises on these)
expansion-bomb.docx                  lossy          0        0     59x  expands 59x (28.5 MB part)

5 files: 5 lossy (100.0%), 0 unreadable, 66 numbers in unreachable text
```

---

## Why this exists

Many Python document pipelines rely on `python-docx` (95.4M downloads/month in a PyPI measurement on 2026-10-14). But `paragraph.text` is not always the full document. Measured over **832 public documents** (method and raw data in [`CENSUS.md`](https://github.com/Vikas-93257/docviews/blob/main/CENSUS.md)):

| what happens | how often | why |
|---|---|---|
| text returned matches **neither** Word view | 10 of 22 documents containing revisions | `.text` drops inserted text (a *rejected* behaviour) and deleted text (an *accepted* behaviour) at the same time |
| **147,000 characters** sit in parts no API can reach | 152 documents | footnotes, endnotes and text boxes have no reader |
| a deleted table row or paragraph mark goes unreported | 121 of 255 revision containers, 20 documents | the markup sits in `w:trPr` / `w:pPr/w:rPr`, not around runs — no Python text reader reports it |
| the file **cannot be opened at all** | 75 of 832 (`9.0 %`), of which 61 ISO-Strict | the Transitional relationship type is hardcoded |
| a 0.49 MB file costs **416 MB of RAM** | 23 documents expand >20× | every part is parsed eagerly and kept |

And saving through it silently rewrites the package — measured by opening all 832 documents, saving each one **unchanged**, and diffing the two packages:

| open → save, unchanged | python-docx 1.2.0 | docviews 1.0.0 (published 2026-11-01) |
|---|---:|---:|
| documents opened | 762 / 832 | **831 / 832** |
| documents that come back with the **same package entries** | 523 | **831** |
| documents with **every part byte-identical** | **7** | **831** |
| documents losing parts on save | **239** | **0** |
| parts byte-identical after save | 5,398 / 12,042 | **13,912 / 13,912** |

The parts python-docx drops are not decoration: 232 documents lose `footnotes.xml.rels` and `endnotes.xml.rels`, 231 lose `fontTable.xml.rels`, and directory entries vanish. The prose usually survives that; the wiring around it does not always, and nobody chose to lose it. **docviews' rule is simpler: a part the caller did not edit is copied out of the package byte-for-byte.**

The paragraph-level proof, from `docx-revisions`' own test fixture:

```
the document offers exactly two readings:
    accepted : "We prefer  color spelling."
    rejected : "We prefer colour  spelling."

python-docx : "We prefer   spelling."      ← a sentence Word cannot display
docviews    : the two readings above
```

## Install

The stable release is available on PyPI. Install the latest version with:

```console
python -m pip install docviews
```

The runtime package has no dependencies, and the documented-API test suite is
verified on CPython 3.9–3.15. The published stable package version is `1.0.0`.

## Reading

```python
doc.text("original")            # reject all changes
doc.text("accepted")            # accept all changes
doc.text("markup")              # show everything, deletions included

for p in doc.paragraphs():      # body, headers, footers, footnotes, endnotes, comments
    p.text, p.style.name, p.original_text, p.alignment, p.is_revised, p.part_name
    for run in p.runs:
        run.text, run.bold, run.italic, run.font.size, run.font.color
```

`doc.iter_paragraphs()` streams the same walk at flat memory; `doc.stream()` walks the body alone; `doc.parts()`, `doc.revisions()`, `doc.comments()`, `doc.numbers()` are described below.

### Revisions that are not runs

Word marks a **deleted table row** with `w:del` inside `w:trPr`, and a **deleted paragraph mark** with `w:del` inside `w:pPr/w:rPr` — the row disappears, two paragraphs merge, and not one Python text reader reports either:

```python
for r in doc.revisions():
    r.kind, r.author, r.date, r.scope      # scope: text | paragraph-mark | row | block
```

```
9 tracked changes in tracked-tables.docx
  delete     Watertight         2026-10-02T09:00  id=9201   scope=row
```

### Comments, with the characters they are attached to

python-docx 1.2.0 ships comment *text* and its release notes say the "reference" range "is not currently planned". A comment without its range is a note with no idea what it is about:

```python
for c in doc.comments():                  # anchors move with the view you ask for
    c.author, c.text, c.resolved, c.is_reply, c.reply_ids
    c.anchors[0].paragraph_index, c.anchors[0].start, c.anchors[0].end
    c.anchor_text                         # the exact characters under review
```

Across the census corpus: **999 comments in 89 documents, 1,189 anchors, every anchor verified against its own view text**.

## Writing

### New documents

```python
doc = Document.new()
doc.add_heading("Report", 0)                     # 0 = Title, 1..9 = Headings
doc.add_paragraph("Body.", style="Heading 2", align="center")
doc.add_list_item("bullet", kind="bullet")       # or kind="number", level=1
doc.add_page_break()
table = doc.add_table(rows=3, cols=2, text=[["Region", "Revenue"], ["EMEA", "1,204,332"]])
table.cell(0, 0).paragraphs[0].runs[0].bold = True
table.add_row()
doc.save("report.docx")
```

Formatting is settable per run — `run.bold`, `.italic`, `.underline`, `.size` (EMU, so `Pt(12)` works), `.font_name`, `.color = "FF0000"` — and readable back the same way.

### Editing documents you did not create

```python
doc = open_docx("contract.docx")
para = doc.paragraphs()[3]
para.set_text("replaced wording")                # keeps the paragraph's style
para.insert_after("a new clause")                # Word's Enter key
para.runs[0].bold = True
doc.add_paragraph("Added by legal.")
doc.save("contract-v2.docx")                     # everything else copied verbatim
```

The guarantee is not a promise, it is a measurement: open all 832 corpus documents, add a paragraph, save, and **every part except `word/document.xml` comes back byte-identical** — tracked-change markup, comments and their anchors, footnotes, endnotes, text boxes, images, custom XML, theme, numbering.

That includes the two things an editor usually breaks first:

* **Tracked changes stay tracked.** Adding text to a document under review does not accept or reject anything.
* **Comments keep their ranges.** A comment whose anchor sits in a paragraph you edited still points at the same characters, because the anchors are re-read from the saved package, not remembered.

### Pictures, headers, footers, merged cells, styles, sections

```python
doc = Document.new()

doc.add_picture("chart.png", width=Inches(5))     # aspect ratio follows
doc.inline_shapes                                 # [InlineShape('Picture 1', 5.00x... in, word/media/image1.png)]

doc.sections[0].header.add_paragraph("Confidential — draft")   # part created on the spot
doc.sections[0].footer.add_paragraph("Page 1 of 1")

table = doc.add_table(rows=3, cols=3, text=[["a", "b", "c"], ["d", "e", "f"], ["g", "h", "i"]])
table.cell(0, 0).merge(table.cell(0, 2))          # w:gridSpan on the first cell
table.cell(1, 0).merge(table.cell(2, 0))          # w:vMerge restart/continue
table.cell(1, 0).text                             # "d\ng" — the merge is followed, not blanked

doc.styles.add_style("Callout", "paragraph", base="Normal", bold=True, size=Pt(12))
doc.add_paragraph("Read this.", style="Callout")

doc.add_section()                                 # a second section
doc.core_properties.title = "Q3 Report"           # docProps/core.xml
doc.save("q3.docx")
```

Writing a comment — and, unlike python-docx, knowing what it points at:

```python
run = doc.add_paragraph("The total is 1,204,332 units.").runs[0]
comment = doc.add_comment([run], text="Check against the ledger.",
                          author="AM", initials="AM")
comment.anchor_text          # "1,204,332 units." — the range, not just the note
doc.comments()               # and it reads back with its anchors from the file
```

Page furniture beyond the default header, styles with their inheritance, and
`word/settings.xml`:

```python
section = doc.sections[0]
section.first_page_header.add_paragraph("Cover page only")   # sets w:titlePg
section.even_page_header.add_paragraph("Even pages")         # sets w:evenAndOddHeaders
section.orientation = WD_ORIENT.LANDSCAPE                    # w:orient + swapped sides
section.start_type = WD_SECTION_START.ODD_PAGE               # w:type
section.header.add_table(rows=1, cols=2, text=[["a", "b"]])  # tables in headers

style = doc.styles["Heading 1"]                  # lookup by name, id or default
style.base_style, style.next_paragraph_style    # the inheritance chain
style.quick_style, style.priority, style.hidden = True, 5, False
doc.styles.latent_styles.add_latent_style("Callout", ui_priority=10)
doc.settings.odd_and_even_pages_header_footer = True
```

Images are measured **from their own bytes** — PNG, JPEG, GIF, BMP, TIFF, with the DPI the file declares — so this library still has zero dependencies and will not insert a picture at a guessed size. SVG is refused with a message that says what to do instead.

Everything above is checked against python-docx: it opens the saved file and sees the same pictures, the same header text, the same merged cells and the same style. [`tests/test_images.py`](https://github.com/Vikas-93257/docviews/blob/main/tests/test_images.py) and [`tests/test_authoring.py`](https://github.com/Vikas-93257/docviews/blob/main/tests/test_authoring.py) are that check.

### Writing is honest about what it does not do

Block-level revision re-flow (a whole paragraph wrapped in `w:ins`) is reported but not yet applied to the views (0 instances in the 832-document corpus). Fields, symbol-font symbols, `.doc` input and rendering are not there. Floating/anchored images are read but not written — `add_picture` inserts inline pictures, the kind a report or a template needs. Nothing above is silently approximated.

## Replacing python-docx in your code

For the complete import map, documented API scope, and migration notes, see
[`MIGRATION.md`](https://github.com/Vikas-93257/docviews/blob/main/MIGRATION.md)
and [`API_PARITY_MATRIX.md`](https://github.com/Vikas-93257/docviews/blob/main/API_PARITY_MATRIX.md).
Contributor and test instructions are in
[`CONTRIBUTING.md`](https://github.com/Vikas-93257/docviews/blob/main/CONTRIBUTING.md).
The compatibility surface matches python-docx 1.2.0's documented public API;
imports must be redirected to `docviews.compat`, and undocumented internals are
outside the parity claim.

```python
import docviews.compat as docx          # replace `import docx`
from docviews.compat.enum import WD_ALIGN_PARAGRAPH  # update enum/shared paths too

doc = docx.Document()
doc.add_heading("Quarterly Report", 0)
p = doc.add_paragraph("Total: ")
p.add_run("1,204,332").bold = True
p.alignment = WD_ALIGN_PARAGRAPH.LEFT
t = doc.add_table(rows=3, cols=2)
t.style = "Table Grid"
t.cell(0, 0).text = "Region"
doc.sections[0].left_margin = docx.Inches(1.2)
doc.core_properties.author = "Finance"
doc.save("out.docx")
```

`docviews.compat` provides python-docx's shapes — `Document()`, `add_paragraph/add_heading/add_run/add_table/add_page_break`, `add_picture`, `paragraphs`, `tables`, `sections[0].header/.footer`, `core_properties`, `styles.add_style`, `inline_shapes`, `style.name`, `run.font.bold`, `Pt/Inches/Cm/RGBColor`, `WD_ALIGN_PARAGRAPH` — on top of this reader, so you keep the anchors, the views, the scopes and the byte-exact save. [`examples/switch_check.py`](https://github.com/Vikas-93257/docviews/blob/main/examples/switch_check.py) is the acceptance test: the same script run against python-docx and against the shim prints **identical output** (run and diffed locally; see [`CENSUS.md`](https://github.com/Vikas-93257/docviews/blob/main/CENSUS.md) §6.3).

Two deliberate file-handling departures:

* `Document(path)` opens ISO-Strict packages (61 of the census files), which python-docx cannot open at all.
* `save()` copies unedited parts verbatim instead of re-serialising every part it parsed.

All 471 enum members present in python-docx 1.2.0 match in value, `str()`,
`repr()`, and `xml_value`. The published enum page also lists
`MSO_THEME_COLOR_INDEX.MIXED`, which the pinned runtime omits; docviews includes
it with the Microsoft Office API value `-2` and records the docs/runtime
conflict in [`COMPAT.md`](https://github.com/Vikas-93257/docviews/blob/main/COMPAT.md).

[`COMPAT.md`](https://github.com/Vikas-93257/docviews/blob/main/COMPAT.md) is the measured answer to "what is still missing": every public
attribute python-docx 1.2.0 puts on the objects it hands you is answered, and
the **471 enum members and 30 enum classes** python-docx ships are importable
with the same values (`WD_LINE_SPACING.DOUBLE == 2`) — from
`docviews.compat.enum`, and from same-named submodules
(`docviews.compat.enum.text`, `.style`, `.section`, `.table`, `.shape`, `.dml`),
so the *import line* survives the swap, not just the name.
What is *not* claimed is the XML layer: `docx.oxml` (`OxmlElement`, `qn`,
`element.xpath`), `doc.element` and the OPC part classes. Code that reaches into
the lxml tree is not migrated by changing an import — that code is rewritten
against `docviews.views`, and [`COMPAT.md`](https://github.com/Vikas-93257/docviews/blob/main/COMPAT.md) says so instead of pretending
otherwise.

Still not authorable anywhere in this library: floating/anchored shapes, charts,
fields and symbol-font glyphs, tables of contents as objects. They are read and
preserved byte-for-byte.

## Comparison snapshot

[`COMPARE.md`](https://github.com/Vikas-93257/docviews/blob/main/COMPARE.md) is the evidence-linked comparison, including methodology, caveats,
and a rated scorecard (support/adoption snapshot updated 2026-10-31). On that
rubric, docviews scores **8.20 vs 7.50** unweighted. python-docx still leads for
migrating an existing codebase (**8.45 vs 7.45**); docviews leads for a new
document pipeline (**8.70 vs 6.95**) and for reading/auditing (**8.85 vs 6.65**).
Those ecosystem scores do not substitute for the separate API matrix.

| measurement | python-docx 1.2.0 | docviews 1.0.0 (published 2026-11-01) |
|---|---|---|
| downloads / month | 95.4M (checked 2026-10-14) | published 2026-11-01; adoption data is not yet meaningful |
| documented API parity | baseline | 578 in-scope behavior matches; 0 partial/not-assessed; 6 structural/internal scope exclusions; details in `API_PARITY_MATRIX.md` |
| bounded compatibility audit | baseline | 276/276 sampled object names, 78/78 sampled callable parameter shapes, 471/471 runtime enum members |
| unchanged open → save | 762 / 832 open; 7 keep every part | 831 / 832 open; 13,912 / 13,912 parts byte-identical |
| tricky-document reading | 14,640 numbers lost (15.9%) | reference reader; 0 lost by construction |
| runtime dependencies | lxml + typing_extensions | none |
| wheel size | 252,987 B | about 144 KiB (local stable build) |

A separate eight-document public-sector smoke set sits outside the 832-file
fixture corpus: docviews preserved all 248/248 package entries; python-docx
omitted two unreferenced entries in one template. This is package-preservation
evidence, not a text-rendering gold standard or a broad real-world sample; see
[`real_world_eval/REPORT.md`](https://github.com/Vikas-93257/docviews/blob/main/real_world_eval/REPORT.md).

## Compatibility boundary

[`API_PARITY_MATRIX.md`](https://github.com/Vikas-93257/docviews/blob/main/API_PARITY_MATRIX.md)
is the source-backed documented-public-API inventory and differential record;
`COMPAT.md` is an additional bounded object-sample audit, not the full matrix.
The matrix records **578 tested behavior matches, 0 partial/not-assessed in-scope
entries, and 6 structural/internal exclusions**. It also reports complete
availability, call-shape, writability, enum, and import-path checks within the
agreed documented scope. Undocumented `docx.oxml`, lxml, and OPC internals and
malformed states outside the published contract are not covered.

The full 239-test suite passed on **each tested CPython patch release**:
3.9.25, 3.10.22, 3.11.18, 3.12.16, 3.13.16, 3.14.10, and 3.15.2. The twelve
selected recipes from python-docx's own documentation produce matching
save-and-reopen transcripts. Wheel and sdist builds were separately clean-
installed and smoke-tested; see `RELEASE_READINESS.md` for artifact details.
For the import boundary and migration caveats, see [`MIGRATION.md`](https://github.com/Vikas-93257/docviews/blob/main/MIGRATION.md),
[`COMPAT.md`](https://github.com/Vikas-93257/docviews/blob/main/COMPAT.md), and [`COMPARE.md`](https://github.com/Vikas-93257/docviews/blob/main/COMPARE.md).

## Guardrails, not silence

A `.docx` is a zip, and a zip can lie about its size. `docviews` refuses parts that exceed explicit limits instead of exhausting memory, and every part it could not read is reported — never silently dropped:

```python
from docviews import Limits
doc = open_docx("upload.docx", limits=Limits(max_part_bytes=8_000_000))
doc.warnings     # ["skipped word/footnotes.xml: expands 900x (limit 400x)"]
```

`doc.stream()` walks the body with `iterparse` and holds one paragraph at a time. On the worst file in the census:

| reader | peak RSS | paragraphs | time |
|---|---:|---:|---:|
| python-docx 1.2.0 | **416 MB** | 4,735 | 2.7 s |
| docviews (tree) | 346 MB | 4,735 | 2.5 s |
| **docviews (stream)** | **15 MB** | 4,735 | 2.4 s |

## Command line

```
docviews audit PATH [--json] [--ci] [--compare python-docx|mammoth|docx2python]
docviews text FILE --view accepted|original|markup
docviews revisions FILE [--json]
docviews comments FILE [--roots-only] [--threads] [--json]
docviews parts FILE
```

`--ci` turns the audit into a build gate:

```console
$ docviews audit contracts/ --ci --fail-on revisions,unreachable
--ci: 3 file(s) failed [revisions, unreachable]        # exit code 2
```

## What it does not do yet

1. **block-level revisions**: a whole paragraph wrapped in `w:ins`/`w:del` is *reported* (`scope="block"`) but its text does not move between views yet (0 instances in the 832-document corpus, so it is measured, not guessed)
2. **fields and symbol-font symbols** — a `PAGE` field shows its cached result, and a `w:sym` character is not resolved to a glyph
3. **`.doc` (binary) input** — the OLE container is not parsed; convert to `.docx` first
4. **rendering** — page geometry, pagination and layout; this library is about the content of a document, not its pixels
5. **floating and anchored shapes, charts, tables of contents as objects** — they are read and preserved byte-for-byte, but there is no API to author them

## Honesty rules this library follows

1. **Nothing is dropped silently.** Anything unread lands in `warnings`.
2. **Views are exact functions of the file** — no "best effort" text.
3. **Saving is not a loss.** Parts nobody edited are copied byte-for-byte, and that is asserted over the corpus, not asserted in prose.
4. **Claims are measured.** [`COMPAT.md`](https://github.com/Vikas-93257/docviews/blob/main/COMPAT.md) is generated, not typed; [`CENSUS.md`](https://github.com/Vikas-93257/docviews/blob/main/CENSUS.md) is the method; [`tests/test_corpus.py`](https://github.com/Vikas-93257/docviews/blob/main/tests/test_corpus.py), [`tests/test_comments.py`](https://github.com/Vikas-93257/docviews/blob/main/tests/test_comments.py) and [`tests/test_write.py`](https://github.com/Vikas-93257/docviews/blob/main/tests/test_write.py) re-run it and fail if a number moves without the census moving too. Revisions and comments are additionally cross-checked against hand-written XML counters ([`verify_revisions.py`](https://github.com/Vikas-93257/docviews/blob/main/verify_revisions.py)).

## Development

The runtime package has no dependencies. Install the optional test tools with
`pip install -e ".[dev]"`; the test extra pins the minimum compatible
`python-docx` version (1.2.0) used to verify interoperability.

The census-backed tests require a local `corpus/` directory. Its third-party
fixture files are gitignored until their redistribution terms are reviewed; if
it is absent, those marked tests skip. To recreate it, collect the source test
suites and follow the commands in [`CENSUS.md`](https://github.com/Vikas-93257/docviews/blob/main/CENSUS.md) using [`build_corpus.py`](https://github.com/Vikas-93257/docviews/blob/main/build_corpus.py).

```console
python -m pytest tests/test_views.py tests/test_comments.py tests/test_write.py tests/test_cli.py
python -m pytest tests/test_corpus.py                                            # census regression
python roundtrip_probe.py                                                        # save-every-file check
python compare_roundtrip.py                                                      # head-to-head vs python-docx
python compat_coverage.py                                                        # python-docx API coverage -> COMPAT.md
python authoring_impact.py                                                       # what each authoring call touches
python examples/switch_check.py --shim                                           # the compat acceptance test
```

## Project links

- [Source repository](https://github.com/Vikas-93257/docviews)
- [Issue tracker](https://github.com/Vikas-93257/docviews/issues)
- [Changelog](https://github.com/Vikas-93257/docviews/blob/main/CHANGELOG.md)
- [License](https://github.com/Vikas-93257/docviews/blob/main/LICENSE)

MIT licensed.
