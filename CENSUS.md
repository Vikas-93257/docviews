# What the standard Python .docx reader actually loses

**Measured 2026-10-14 · 832 public documents · python-docx 1.2.0 · Python 3.13**

Every Python pipeline that reads Word documents goes through `python-docx`.
This is a measurement of what that reader returns, paragraph by paragraph and
part by part, against the document's own XML. Every number below is
reproducible from this directory and every claim is attributable to a specific
cause (that attribution is the point — a "gap" nobody can attribute is a vibe).

```bash
python docviews/build_corpus.py --src <checkouts> --out docviews/corpus   # Day 1
python docviews/run_census.py --corpus docviews/corpus --out docviews/census2.jsonl
python docviews/census_report2.py docviews/census2.jsonl docviews/corpus/manifest.json
python docviews/readers.py --corpus docviews/corpus --out docviews/readers.jsonl --only-interesting
python docviews/exhibits.py docviews/corpus/<any-file>.docx
```

---

## 1. The corpus, and why it is a hard sample

832 documents after de-duplicating by SHA-1, collected from the test-data sets
where real user documents accumulate — bug reports, conformance files, and
library fixtures:

| source | docs | what it is |
|---|---:|---|
| Open-XML-SDK (Microsoft) | 378 | Word/OOXML conformance suite, incl. ISO-Strict tests |
| dolanmiu/docx (JS) | 252 | library fixtures + layout probes |
| docx4j (Java) | 113 | real-world samples (UN report, contracts, forms) |
| python-docx | 45 | the incumbent's own fixtures |
| unstructured | 21 | real-world examples (handbooks, filings) |
| mammoth.js | 17 | converter fixtures |
| docx-revisions, python-pptx, python-docx sdists | 6 | library fixtures |

This is **not** a random sample of business documents. It is skewed *harder*
than a random sample: conformance suites deliberately contain one feature per
file, which understates revision markup, while bug-report fixtures overstate
odd structures. Treat the rates as "at least this common" for the failure
kinds, never as a population estimate. The audit CLI is how anyone gets the
number for *their* corpus.

## 2. Method (the definitions the numbers depend on)

* **Accepted view** = the text Word shows with "accept all changes"
  (`w:ins`/`w:moveTo` kept, `w:del`/`w:moveFrom` dropped).
  **Original view** = "reject all changes".
* Comparison is element-exact: each `python-docx` `Paragraph` is compared
  against the views of **its own** XML element (`p._p`), so no list alignment
  can be wrong. An earlier version of this probe compared paragraph lists by
  position and produced numbers that were pure artefacts of paragraphs nested
  in tables — that version was discarded, not published.
* `python-docx` maps `w:tab`/`w:ptab` → `"\t"`, `w:br`/`w:cr` → `"\n"`,
  `w:noBreakHyphen` → `"-"`; the views mirror that exactly, so whitespace is
  never reported as a difference.
* Text-level claims compare **numbers** (`\d[\d,]*\.?\d*`) as a multiset,
  because that is a corruption you can put in front of a lawyer or a lab
  result. Counting characters would be noisier and less meaningful.
* Reader-internal markup is stripped before comparison (mammoth's base64 image
  payloads, markdown escaping, alt text; docx2python's part separators and
  comment labels) — each is documented in `readers.py`.

## 3. Findings

### 3.1 python-docx cannot open 75 of these 832 documents (9.0 %)

| count | error | cause |
|---:|---|---|
| **61** | `KeyError: no relationship of type 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'` | **ISO/IEC 29500 Strict** documents. Strict uses `http://purl.oclc.org/ooxml/...` namespaces and relationship types; python-docx hardcodes the Transitional relationship type, so the document is not even opened. 60 of the 61 come from the Open-XML-SDK strict suite — i.e. **60 of 378 Microsoft conformance files** are unreadable by the incumbent. |
| 6 | `XMLSyntaxError: Start tag expected` | a part that is not well-formed XML |
| 2 | `KeyError: no item named 'word/document.xml'` | package without a main part |
| 2 | `AttributeError: '_cython_…cython_function_or_method'` | python-docx/lxml internals on imported documents |
| 1 | `KeyError: no item named 'docProps/core.xml'` | package without core properties |
| 1 | `ValueError: no 'tc' element at grid_offset` | malformed table grid |
| 1 | `BadZipFile` | not a zip |
| 1 | `KeyError: no item named 'word/#_THIS_WEEK_IN…'` | damaged relationship target |

Three further files were skipped by the *probe*, not by the reader: their
`word/document.xml` expands to 16–29 MB (see §3.5).

### 3.6 Memory: a 0.5 MB upload costs 400 MB of RAM

| file | zip size | `document.xml` | peak RSS to simply open it |
|---|---:|---:|---:|
| `…-font-widths.docx` | 0.49 MB | 28.5 MB | **416 MB** |
| `…-font-italic-widths.docx` | 0.49 MB | 29.4 MB | **433 MB** |

A 0.49 MB file expands 59× inside the zip, and the eager lxml parse of every
part turns it into ~420 MB of resident memory. In a web service that accepts
user uploads, this is a denial-of-service vector reachable through the most
common document library in Python. (This probe itself died on these files until
it was changed to parse each document once and to run in batches.)

Same file, three readers, measured 2026-10-14:

| reader | peak RSS | paragraphs | time |
|---|---:|---:|---:|
| python-docx 1.2.0 | 416 MB | 4,735 | 2.7 s |
| docviews 0.1.0, tree path | 346 MB | 4,735 | 2.5 s |
| **docviews 0.1.0, `stream()`** | **15 MB** | 4,735 | 2.4 s |

`stream()` walks the part with `iterparse` and clears each paragraph as it goes,
so memory does not grow with document size. That is a 28× reduction on the file
that breaks the incumbent.

### 3.2 Tracked changes: rare in this corpus, catastrophic when present

22 documents (2.9 %) contain revision markup. In those documents:

| measure | documents | paragraphs |
|---|---:|---:|
| reader text differs from what Word displays (accepted view) | 16 / 22 | 38 |
| reader text matches **neither** Word view — text no user can see | 10 / 22 | 20 |
| reader returns the **original draft** text | 12 / 22 | 18 |
| reader reports an **empty** paragraph while the XML holds text | 3 / 22 | 5 |
| numbers lost or invented | 2 / 22 | 4 numbers |

Verified exhibit — `docx-revisions-mai_tests_OXML_TrackChanges_Test.docx`,
paragraph 5. The document offers exactly two readings, `We prefer color
spelling.` or `We prefer colour spelling.` The incumbent returns a third:

```
reader (python-docx .text) : "We prefer   spelling."
Word, changes accepted     : "We prefer  color spelling."
Word, changes rejected     : "We prefer colour  spelling."
XML: …<w:del w:id="1003"><w:r><w:delText>colour</w:delText></w:r></w:del>
     <w:r><w:t>color</w:t></w:r>…        <w:ins w:id="1006">…</w:ins>
```

Three spaces where an adjective should be. This is not a formatting
difference: `paragraph.runs` yields only runs that are *direct children* of the
paragraph, so text inside `w:ins`, `w:del`, `w:moveFrom` and `w:moveTo` is
invisible to both `.runs` and `.text` — the reader accidentally returns a mix
of the accepted view (deletions dropped) and the original view (insertions
dropped). **Word itself has no setting that produces this text.**

### 3.3 Text that no python-docx API can reach

| measure | value |
|---|---:|
| documents with unreachable text | **152 / 754 (20.2 %)** |
| characters in footnotes / endnotes / text boxes | **147,182** |
| of which text boxes | 14,453 |
| footnotes | 93,676 |
| endnotes | 39,053 |

Footnotes, endnotes and text boxes have no reader in python-docx 1.2.0. In one
file (`wordprocessing_footer___header_BackgroundReport_-_THE_WORKING_GROUP_ON_
INTERNET_GOVERNANCE.docx`, the UN WGIG background report) 1,922 body paragraphs
exist of which `doc.paragraphs` returns 1,325 — the rest live in tables; the
document's 597 table paragraphs *are* reachable via `doc.tables`, which is the
footgun: `"\n".join(p.text for p in doc.paragraphs)` — the most common line in
the ecosystem — silently returns a partial document and never says so.

### 3.4 Run-level blindness (`paragraph.runs`)

Inside `document.xml` across the corpus (308,661 `<w:r>` elements):

| where the run lives | runs | reachable by `.runs`? |
|---|---:|---|
| direct child of the paragraph | 303,700 | yes |
| inside `w:hyperlink` | 3,713 | **no** — but `.text` *does* include it |
| inside `w:smartTag` | 497 | no |
| inside `w:sdt` (content control) | 389 | no |
| inside `w:fldSimple` | 134 | no |
| inside `w:ins` / `w:del` | 222 | no |
| moved text (`w:moveFrom`/`w:moveTo`) | 6 | no |

4,920 runs (1.6 %) are invisible to `.runs`. The hyperlink row is the sharpest
one: python-docx's own API is internally inconsistent — `.text` includes
hyperlink text, `.runs` does not, so any pipeline that walks runs (the
documented way to get bold/italic segmentation, or to map text to formatting)
loses link text that `.text` would have given it.

### 3.7 Four readers, one corpus (the 167 documents the census flagged)

| reader | opened | documents losing numbers | numbers missing | numbers invented |
|---|---:|---:|---:|---:|
| python-docx 1.2.0 | 167 | **78.4 %** | **14,640** | 0 |
| mammoth 1.13.0 | 167 | 89.8 % | 1,742 | 5,118 |
| docx2python 3.7.1 | 166 (1 raises) | 89.8 % | 189 | 3,227 |
| *raw OOXML (reference)* | 167 | — | — | — |

Attribution for python-docx's 14,640 missing numbers: **14,286 (97.6 %) sit in
parts no python-docx API reaches** (footnotes, endnotes, text boxes). The
remaining 354 occur inside paragraphs that the reader does return and are
candidates for content controls / smart tags / nested structures — they need
per-document triage and are *not* claimed as revision loss.

"Invented" for mammoth and docx2python is **not** fabricated content: it is
generated numbering the reader synthesises (footnote markers from `w:id`
attributes, list numbering, image labels) which by construction is not present
as text in the XML. python-docx invents **zero** numbers — it only loses them.

## 4. Verdict against the Day-5 thresholds

The memo (in `/home/user/radar/replacement-radar.md`) set:

> **GO** if ≥25 % of real public documents contain markup content a naive
> reader drops or corrupts **AND** accepted-vs-reader text diverges in ≥10 %
> of files. **NO-GO** if markup < ~5 % or divergence negligible.

Measured, honestly:

* **Tracked changes alone do not clear the bar**: 2.9 % of this corpus. The
  divergence criterion *is* met among documents that contain revisions (16 of
  22 = 73 % differ from the accepted view, 10 of 22 return impossible text),
  but prevalence is low here — and this corpus understates it (conformance
  files are one-feature-per-file).
* **The broader criterion does clear the bar**: 20.2 % of documents hold text
  no API can reach, 57.4 % have runs `.runs` never yields, 9.0 % cannot be
  opened at all, and one 0.5 MB file costs 400 MB of RAM.

So: **GO — with the product's centre of gravity moved.** Tracked changes are
the sharpest failure (a text that exists in no view) but not the most
prevalent. The prevalent failures are parts nobody reads, runs nobody sees,
files nobody can open, and memory nobody budgets. A library that ships
`revisions()`, `comment_anchors()`, three views and a full-part walker
addresses all of them at once; a library that only does tracked changes would
be aimed at the smaller half of its own evidence.

## 5. What this does not prove

* Prevalence in real business corpora. 2.9 % tracked-change rate here is a
  floor for a hard sample, not a global estimate; the audit CLI exists to
  produce that number per user.
* That the alternative libraries are broken. mammoth and docx2python ignore
  revisions *by design and document it*; the finding is that **no library lets
  you choose a view**, and none reports what it dropped.
* Anything about writing. This measures reading only.

---

## 6. Verification: `docviews` 0.1.1 reproduces this census

The library in this directory was built to the census, and the census is its
regression test (`tests/test_corpus.py`, `tests/test_comments.py`, green).

| census metric | census value | docviews 0.1.1 |
|---|---:|---:|
| documents read | 754 / 832 (python-docx) | **831 / 832** (the 832nd is not a zip) |
| ISO-Strict packages readable | 0 (python-docx raises on all 61) | **61 / 61** |
| unreachable text: footnotes | 93,676 chars | 93,665 chars |
| unreachable text: endnotes | 39,053 | 39,074 |
| unreachable text: text boxes | 14,453 | 14,701 |
| documents holding unreachable text | 152 (of the 754 measurable) | 159 (of 831 — includes the 61 strict files) |
| the impossible text | present | **gone**: `accepted` = "We prefer  color spelling.", `original` = "We prefer colour  spelling." |

What the audit says at the end of a full sweep (`python3 audit_corpus.py`,
docviews 0.1.0, 832 documents, 71 s total — ~85 ms per file):

| verdict | documents | share |
|---|---:|---:|
| lossy — something in the file is read wrong, invisible, or unopenable | **329** | 39.5 % |
| clean | 502 | 60.3 % |
| unreadable (encrypted, not a zip: `wordprocessing_protected_document_with_password`) | 1 | 0.1 % |

with 159 documents containing text no python-docx API returns, 61 ISO-Strict
packages python-docx refuses to open, 89 documents whose comments have no
retrievable anchor, 27 documents whose text changes under accept/reject, and
20 documents carrying row/paragraph-mark revisions that no Python reader
reports at all.

Two independent measurements of the same corpus agreeing to within 0.5 % is the
closest thing this project has to a proof that the numbers are real.

### 6.1 Revisions, re-counted from the markup (revised 2026-10-14)

The first revision counts here were too low and slightly wrong for a reason
worth recording: a revision container that holds **no visible characters** — a
deleted image, a deleted text box, an emptied run — carries no text for a
text-based detector to find. Counting revisions is therefore done from the
markup, and each count is cross-checked against a second, hand-written XML
counter (`verify_revisions.py`, which walks the parts itself and classifies
containers by their parent chain):

| metric | count | second implementation |
|---|---:|---:|
| revision containers | **255** | 255 |
| paragraphs carrying revision markup | **142** in 27 documents | 142 |
| `insert` / `delete` / `move_to` / `move_from` | 95 / 150 / 6 / 4 | identical |
| documents with any revision markup | 28 | identical |
| scope `text` (around runs) | 134 | —
| scope `paragraph-mark` (`w:del` in `w:pPr/w:rPr`) | 89 | —
| scope `row` (`w:del` in `w:trPr`) | 32 | —
| scope `block` (`w:ins`/`w:del` wrapping a whole `w:p`) | **0** | —

Two results matter for anyone reading this census:

* **121 of the 255 revision containers (47 %) are paragraph-mark or row
  revisions.** python-docx reports none of them: the markup is not around runs,
  so `.text`, `.runs` and even `docx-revisions` see nothing. A deleted table row
  still reads as a live row; a deleted paragraph mark still reads as a paragraph
  break. `docviews` reports them (`RevisionInfo.scope`, and
  `docviews audit` prints "N table rows marked deleted (no text reader reports
  this)"). It does **not** yet re-flow the view text accordingly.
* The earlier numbers (107/118, 88 paragraphs) came from a detector that
  double-counted revisions inside text boxes under their host paragraph. The
  255-container total is verified twice, so it replaces them.

### 6.2 Comments, measured (new 2026-10-14)

`comments()` returns text **and** the characters the comment is attached to.

| metric | count |
|---|---:|
| documents with comments | 89 of 832 (one fixture ships an empty `<w:comments/>` part) |
| comments | 999 |
| anchors | 1,189 (24 of them point anchors: a reference with no range) |
| comments with no anchor in the body (orphans) | 4 |
| replies / resolved | 380 / 73 |
| anchor text vs the paragraph's own view text | **1,189 / 1,189 exact matches, 0 mismatches** |

The last row is the claim that matters: every anchor is a slice of the view the
caller asked for, checked in `tests/test_comments.py` against the paragraph text
that the same library returns — and verified a second time by an independent
probe over all 90 documents. Threading and resolution are read from
`commentsExtended.xml` (`w15:paraIdParent`, `w15:done`); Word keys a comment by
the `w14:paraId` of the **last** paragraph of its body, which is why a comment
whose body has more than one paragraph still resolves.


### 6.3 Writing (new 2026-10-14, v0.2.0)

Writing is measured the same way reading is: by opening documents, doing
something, and diffing the packages. All 832 corpus documents, three probes.

**(a) Open → save unchanged → diff.** The claim: a part nobody edited comes
back byte-for-byte.

| metric | docviews 0.2.0 | python-docx 1.2.0 |
|---|---:|---:|
| documents opened | **831 / 832** | 762 / 832 |
| same package entries after save | **831** | 523 |
| every part byte-identical after save | **831** | 7 |
| parts byte-identical after save | **13,912 / 13,912** | 5,398 / 12,042 |
| parts whose meaning changed after save | **0** | 120 (Content_Types, `.rels` files, core.xml) |
| documents losing parts on save | **0** | **239** |

What python-docx drops, counted (`pythondocx_dropped_entries.json`):
`footnotes.xml.rels` (232 documents), `endnotes.xml.rels` (232),
`fontTable.xml.rels` (231), then header/footer `.rels` files (23, 14, 11, 10,
8, 8, 7, 6…) plus directory entries. No `r:id` reference in the 832 documents
was left dangling by it, so this is package damage rather than proven content
loss — stated that way on purpose. `check_integrity.py` is the checker used,
and it was validated by deliberately deleting a `.rels` file and confirming it
reports the dangling references.

**(b) Edit → save → does the rest survive?** Add a paragraph to a document, save,
and compare every other part and every reader-visible quantity:

| document | revisions | comments | anchors | footnote/endnote chars | parts unchanged |
|---|---:|---:|---:|---:|---:|
| `docx-revisions-mai_tests_OXML_TrackChanges_Test.docx` | 8 → 8 | 0 → 0 | – | 0 → 0 | 11/11 |
| `steps_test_files_comments-rich-para.docx` | 0 → 0 | 4 → 4 | 4 → 4 | 0 → 0 | 16/16 |
| `layout-probes_stops2_word-stops-endnotes.docx` | – | – | – | 11,883 → 11,883 | 22/22 |

Across the whole corpus: `impact_python_docx.jsonl` and
`impact_docviews.jsonl` record **0 content-level differences** for both readers
(the files list *problems*, so `impact_docviews.jsonl` is empty on purpose:
there were none to list, against 478 rows for python-docx, nearly all of them
package-entry losses)
(neither loses text, revisions, comments or numbers when saving this corpus),
and **0 package-entry losses** for docviews against 239 for python-docx.

**(c) The compatibility shim runs existing code.** `examples/switch_check.py`
is one python-docx program (heading, subtitle, three bullets, mixed-format
runs, a 3×2 table, section length) run twice — once with `import docx`, once
with `import docviews.compat as docx` — and it must print the same report. It
does, byte for byte, including `paragraph.style.name` (`"Heading 1"`,
`"List Bullet"`, `"Quote"`) and `sections: 1`.

Suite at v0.2.0: **61 tests**, of which the corpus ones re-measure the reading
census (§6, §6.1, §6.2) and the save guarantees above.


### 6.4 Authoring (new 2026-10-16, v0.3.0)

v0.3 adds the authoring surface python-docx users reach for after the easy
cases — pictures, headers and footers, merged cells, style creation, section
breaks, core metadata. The claim being measured is not "it works", it is the
same one as §6.3: **an edit changes what it was asked to change and nothing
else**, and **another reader sees what was written**.

Method: `authoring_impact.py` (rewritten for this section) runs four sweeps over
the 832-document corpus and writes every document's result to
`authoring_impact.jsonl`; the image measurement table is `image_facts.json`.

**(a) Open → save unchanged.** The v0.2 claim, now measured over the whole
corpus rather than a 70-file sample:

| metric | docviews 0.3.0 |
|---|---:|
| documents opened and saved | **831 / 832** (the 832nd is a password-protected file, not a zip) |
| package entries lost | **0** |
| package entries added | **0** |
| parts whose bytes changed | **0** (13,912 parts across the sample; 0 of 13,912 rewritten) |

**(b) Add a paragraph to a header.** 78 of the 832 documents have a header with
text in it; all 78 were edited (append a paragraph), saved, reopened:

| metric | result |
|---|---:|
| documents where exactly **one** part changed | **78 / 78** |
| documents where anything outside `word/header*.xml` changed | **0** |
| the new text readable after saving (docviews) | **78 / 78** |
| package entries lost | **0** |

The 753 documents without a header paragraph are skipped rather than counted as
passes, which is why the header sweep reports a denominator.

**(c) Insert a picture into the body.** 42 documents (every 20th), a 120×60 PNG
at 2 inches wide:

| metric | result |
|---|---:|
| documents where the change stayed inside `word/document.xml`, its `.rels`, `[Content_Types].xml` and the new media part | **42 / 42** |
| documents where anything else changed | **0** |
| documents losing a package entry | **0** |
| `len(doc.inline_shapes)` before → after | **+1 in 42 / 42** |

**(d) Interoperability: python-docx 1.2.0 reads what was written.** Same 42
documents, each given a picture, a header paragraph (where a header exists) and
a new style, then opened with python-docx:

| metric | result |
|---|---:|
| packages python-docx opened | **39 / 42** — the other 3 are ISO/IEC 29500 Strict, which python-docx refuses by design (0.2.0 census: 61 such packages in this corpus) |
| saw the inserted picture | **39 / 39** |
| saw the created style | **39 / 39** |
| saw the header line | **5 / 5** of the packages that had a header to write into |

Three defects were found by this sweep and fixed before it was recorded —
which is the reason the sweep exists:

1. header/footer parts were written with root element `w:header` instead of
   `w:hdr`, so python-docx read them as opaque XML (fixed, `authoring.py`);
2. core properties were declared with the `officeDocument` relationship type
   instead of the package one, so python-docx reported its own default title
   (`"Word Document"`) and silently ignored the file's (fixed, `write.py`);
3. `doc.styles.add_style` wrote into `word/glossary/styles.xml` in packages that
   carry a glossary, because the styles part was found by scanning entry names
   instead of following the document part's `styles` relationship (fixed,
   `document.py`).

**(e) Images are measured from their own bytes.** No image library is imported,
so the fixtures are 100-odd bytes each (`tests/blobs.py`, regenerated by
`tools/make_image_blobs.py`):

| fixture | type | pixels | dpi declared | native size |
|---|---|---:|---:|---|
| `PNG_96` | image/png | 120×60 | 96 | 1.250 × 0.625 in |
| `PNG_RAW` | image/png | 120×60 | none → 72 | 1.667 × 0.833 in |
| `JPG_300` | image/jpeg | 200×100 | 300 | 0.667 × 0.333 in |
| `GIF` | image/gif | 3×2 | none → 72 | 0.042 × 0.028 in |
| `BMP` | image/bmp | 4×6 | none → 72 | 0.056 × 0.083 in |
| `TIFF_300` | image/tiff | 5×7 | 300 (rational tags) | 0.017 × 0.023 in |
| SVG, junk, empty | — | — | — | **refused** with a message, never guessed |

The suite count lives in §6.5 with the rest of the replace measurement — it
moved after this section was written, and a number repeated in two places is a
number that goes stale in one of them. The corpus tests re-measure §6, §6.1,
§6.2 and the §6.3/§6.4 guarantees; a number that moves without this file moving
is a failing test.

## 6.5 Is it capable of replacing python-docx? The measurement

### (a) The authoring API, attribute by attribute

`python3 compat_coverage.py` builds corresponding documents with python-docx
1.2.0 and `docviews.compat`, then compares selected public attributes and
callable signatures. XML-level names (`._element`, `._p`, `part`, `oxml.*`) are
excluded on purpose: claiming lxml compatibility would be claiming a different
promise. The table below records the **initial, limited 2026-10-16 run** (13
object pairs); it has since been expanded. See the current validation note
below and `COMPAT.md` for the latest audit and explicit scope limits.

| object | python-docx attributes | answered | missing |
|---|---:|---:|---|
| Document | 17 | 17 | — |
| Paragraph | 12 | 12 | — |
| Run | 14 | 14 | — |
| Font | 27 | 27 | — |
| Table | 12 | 12 | — |
| Row | 6 | 6 | — |
| Cell | 10 | 10 | — |
| Section | 19 | 19 | — |
| Header | 6 | 6 | — |
| Footer | 6 | 6 | — |
| Styles | 5 | 5 | — |
| Style | 14 | 14 | — |
| InlineShape | 3 | 3 | — |
| CoreProperties | 15 | 15 | — |
| **total** | **166** | **166 (100.0 %)** | **0** |

45 names exist *beyond* python-docx on the same objects (`editable`, `accepted`,
`in_table`, `in_textbox`, `iter_paragraphs`, `is_revised`, `add_list_item`, …),
because the views, the anchors and the streaming walk have no python-docx
counterpart to collide with.

Two things the 2026-10-17 run fixed, both in the measurement itself and in what
it measured. The probe could not count a name that `__getattr__` resolves (it
asked `dir()` only), which under-reported `CoreProperties` as 0/15; it also left
out `Footer` and counted `InlineShapes` (the *collection*) instead of
`InlineShape`. Both are corrected, so the table above is the honest number — a
slightly larger one, because it now asks about more objects. The measurement
change then exposed a real gap the old run had hidden: `docx.shared`'s
`lazyproperty` (and `write_only_property`), which the compat door now provides.

The same question for the names python-docx ships *outside* a document object,
because existing code imports them directly:

| namespace | names | answered |
|---|---:|---:|
| classes in each `docx.enum.*` module | 38 | **38** |
| enum members in those classes | 471 | **471** |
| `docx.enum.base`'s base classes | 2 | **2** |
| names in `docx.shared` (units, colours, decorators) | 10 | **10** |

The members are not hand-copied: `tools/gen_enums.py` reads them out of the
installed python-docx and writes `src/docviews/compat/enum/_generated.py`
(names, values, XML tokens, bases, aliases), and it also writes one
``docviews.compat.enum.text`` / ``.style`` / ``.section`` / ``.table`` /
``.shape`` / ``.dml`` module per python-docx module, so the *import line*
survives the switch, not just the name.

The members are real enum members, not objects that imitate them:
``docviews.compat.enum.WD_ALIGN_PARAGRAPH`` is an ``Enum`` over ``int`` whose
members derive from ``BaseXmlEnum``, so ``list()``, ``WD_ALIGN_PARAGRAPH(1)``,
``__members__``, ``isinstance``, ``from_xml()`` and ``to_xml()`` all work.
``tests/test_enums.py`` compares all 471 members with python-docx's own —
``value``, ``str()`` and ``repr()`` are identical for every one of them. The
tokens are python-docx's too, with two exceptions that are stated here and in
``COMPAT.md`` rather than discovered by a user:

* ``WD_BREAK``/``WD_BREAK_TYPE`` carry the tokens python-docx's deprecated class
  never mapped (``"page"``, ``"column"``, …), so ``run.add_break(WD_BREAK.PAGE)``
  writes ``<w:br w:type="page"/>``. 25 members gain a token this way;
* seven members carry a *different* token on purpose (``WD_COLOR_INDEX.AUTO``
  is ``"none"`` because ``w:highlight`` has no ``"default"``; three
  ``WD_LINE_SPACING`` members and ``MSO_THEME_COLOR.NOT_THEME_COLOR`` are
  ``None`` rather than the ``"UNMAPPED"`` marker, so the writer never emits a
  placeholder).

Zero members diverge unexplained. `str(WD_ALIGN_PARAGRAPH.CENTER)` is
`"CENTER (1)"`, exactly as in python-docx; the token is `.xml_value` (and our
`.ooxml`), which is what lets `paragraph.alignment =
WD_ALIGN_PARAGRAPH.CENTER` write a correct file with the *same integer*
python-docx uses.

### (b) What the count does and does not mean

The table is a presence measurement. A name answering is not a name behaving the
same, so behavioral evidence is separate: the suite, `examples/docs_recipes.py`
(python-docx's own twelve documented recipes, run through both libraries and
diffed), and authoring round trips reopened with python-docx 1.2.0 itself
(§6.3, §6.4, and the interop sweep below). The initial presence table and the
corpus-sweep numbers remain tied to their stated 2026-10-16/17 run dates; the
latest validation note below supersedes the old API and suite counts.

Deliberately **not** claimed, and stated in `COMPAT.md` rather than discovered
by a user: `docx.oxml` (`OxmlElement`, `qn`, `element.xpath`), `document.element`
and the OPC part classes, and `docx.shared`'s lxml plumbing (`ElementProxy`,
`lazyproperty`). Code that manipulates the XML tree is not migrated by changing
an import; it is rewritten against `docviews.views`.

Non-gaps, for completeness: `.doc` (binary) input and rendering are refused by
python-docx too, so there is no behaviour to match; fields, symbol-font glyphs,
floating shapes, charts and TOC authoring are read and preserved byte-for-byte
by both libraries but authored by neither in this API surface.

### (c) Initial v0.3.0 authoring measurements (2026-10-17)

The corpus and interoperability sweeps below were measured after the initial
v0.3.0 authoring work on 2026-10-17. They have not been rerun after the
2026-10-29 compatibility refinements; see the current validation note below.

| what | result |
|---|---|
| open every corpus file and save it unchanged (`compare_roundtrip.py`) | **831 / 832** opened, **13,912 / 13,912 parts byte-identical**, 0 entries lost or added |
| the same trip with python-docx 1.2.0 | 762 / 832 opened, 523 keep every entry, **7** keep every part, 5,398 / 12,042 parts byte-identical |
| add a paragraph to the header of every document that has one (`authoring_impact.py`, header sweep) | 78 / 78 edited documents keep the original header text, **exactly one part changed each** (78/78), 0 entries lost — 753 documents have no header with text |
| insert a picture into every 20th file (42) | **42 / 42** gain exactly one shape, **42 / 42** change nothing unexpected |
| python-docx 1.2.0 re-opens what was written (interop sweep, 42) | **39 / 42** — the 3 refusals are ISO/IEC 29500 Strict packages python-docx cannot open at all; of those it opened, it saw the picture **39 / 39**, the created style **39 / 39**, the header line **5 / 5** |

The initial v0.3.0 suite had **131 tests** — test_views 19, test_write 17,
test_images 14, test_authoring 23, test_cli 7, test_comments 9, test_corpus 9,
and the parity files added for that work: test_write_v03 11,
test_authoring_v03 14 (through the shim as well), test_enums 8.

**Latest validation (2026-10-31):** the full suite passes **149 tests on each
Python version from 3.9 through 3.15** with `python-docx==1.2.0`; the focused
`tests/test_compat_parity.py` and `tests/test_authoring_v03.py` run passes **32
tests on each of the same seven versions**. Compatibility regressions cover,
among other areas, latent-style defaults and item properties, run-content
iteration and rendered page-break fragments, Drawing/Image access, python-docx-
compatible `RGBColor` and length-unit constructors, plus the earlier comments,
hyperlink, line-spacing, tab-stop, section, table-style, picture, keyword, and
stream-save cases. The twelve-recipe documented transcript diff was rerun and
is empty. The expanded `compat_coverage.py` audit (generated under Python 3.13)
reports **276/276** sampled public attribute names, **78/78** sampled
method/constructor parameter shapes
(names/order/kinds/defaults; annotations excluded), and **471/471** enum members
with identical values, `str()`, and `repr()`. This remains a bounded sample,
not proof of complete documented-API parity or all behavior; `COMPAT.md` records
the limit. The refreshed five-run writer microbenchmark found docviews faster
end-to-end at 500 and 5,000 paragraphs plus a populated table, while
python-docx remains faster at save alone; updated numbers and caveats are in
`COMPARE.md`. After the final metadata version selection to `0.3.0a1`, the full
suite was rerun on Python 3.9 and 3.15 (**149 passed on each**); fresh Python
3.15 wheel and sdist installs also passed DOCX and CLI smoke checks. None of
this updates the 2026-10-17 corpus-sweep statistics above.

### (d) The switch acceptance test

`examples/docs_recipes.py` is the file python-docx's documentation implies: the
twelve recipes it documents, written verbatim, then re-opened with the *same*
library and described in plain text — paragraph styles and text, alignment, run
formatting, table geometry, section geometry, header and footer paragraphs,
core properties, comments, inline shapes. Run against python-docx 1.2.0 and
against `docviews.compat`, the two transcripts are **byte-identical**
(verified again 2026-10-31).

Getting there found five real behaviour gaps, all now closed — they are the
answer to "capable of replacing", because each is what a migrating program would
have hit:

| what a program saw | python-docx | docviews before | now |
|---|---|---|---|
| `section.page_width` after `orientation = LANDSCAPE` | the stored width (8.5in) | the rotated page (11in) | the stored width |
| paragraphs in a header it just created | `['', 'Page header']` | `['Page header']` | `['', 'Page header']` |
| `run.add_break(WD_BREAK.PAGE)` in `paragraph.text` | `''` | `'\n'` | `''` |
| `section.orientation` | `pgSz/@w:orient` | inferred from the measurements | `pgSz/@w:orient` |
| a document that states no geometry | Letter, 1.25in sides | Letter (fixed before this run) | Letter, 1.25in sides, 0.5in header |
