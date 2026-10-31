# Initial public real-world DOCX smoke evaluation

**Run date:** 2026-10-29  
**Purpose:** test package opening and unchanged-save preservation on documents not sourced from software test suites.  
**Status:** exploratory smoke evaluation; not a representative corpus or a fidelity gold-standard.

## Result

Eight DOCX files were downloaded from their publishers' official GOV.UK asset URLs. The four GOV.UK source pages' standard footer says all content is available under the Open Government Licence v3.0, except where otherwise stated; no contrary document-specific reuse terms were found on these pages. The licence still has exclusions, including third-party rights. The manifest links to the source pages, GOV.UK terms, and the OGL itself, and records each direct URL, HTTP response, file size, and SHA-256.

The `docviews 0.3.0` table label is the package version reported on the 2026-10-29 run date. It is historical metadata; the current alpha candidate is `0.3.0a1`.

| Measure | docviews 0.3.0 | python-docx 1.2.0 |
|---|---:|---:|
| Valid DOCX packages opened | **8 / 8** | **8 / 8** |
| Open warnings/errors | **0** | **0** |
| Package entries retained on unchanged save | **248 / 248** | **246 / 248** |
| Parts byte-identical after unchanged save | **248 / 248** | 128 / 246 |
| Documents with lost package entries | **0 / 8** | 1 / 8 |

`docviews` opened each file and serialized an unchanged in-memory document; every one of the **248 package entries** had identical bytes before and after. `python-docx` opened all eight, but its round-trip of `hmt_wga_2024_25_cg05.docx` omitted two unreferenced ZIP entries, `[trash]/0000.dat` and `[trash]/0001.dat` (1,247 and 4,118 bytes). The other six entries in the package remained present, but their bytes and many other package parts were rewritten as usual by `python-docx`.

Across the eight files, python-docx changed 118 of 248 common package-part byte streams during save. That is a package-level observation, **not evidence that 118 pieces of reader-visible content changed**. In the observed CG-05 case, the two omitted entries are not referenced in the package's XML or relationship files; no claim of visible-content loss is made.

The source set totals **2,771,590 compressed bytes** and 248 ZIP entries. It contains 32 media entries and 60 `customXml/` entries; the five current-year WGA templates contain all 60 custom-XML entries. This gives the evaluation real packages with custom XML, media, headers/footers and note parts—not just minimal generated documents.

## Documents

| File | Publisher / type | Size | ZIP entries | customXml entries | docviews unchanged save | python-docx omitted entries |
|---|---|---:|---:|---:|---|---|
| `hmt_wga_annex_written_guidance_2024.docx` | HM Treasury, long guidance | 2,288,575 B | 37 | 0 | exact, 37/37 | none |
| `hmt_wga_2024_25_cg02.docx` | HM Treasury, form | 60,000 B | 25 | 9 | exact, 25/25 | none |
| `hmt_wga_2024_25_cg03.docx` | HM Treasury, form | 73,657 B | 35 | 12 | exact, 35/35 | none |
| `hmt_wga_2024_25_cg04_central.docx` | HM Treasury, checklist | 99,983 B | 36 | 12 | exact, 36/36 | none |
| `hmt_wga_2024_25_cg04_local.docx` | HM Treasury, checklist | 86,515 B | 39 | 15 | exact, 39/39 | none |
| `hmt_wga_2024_25_cg05.docx` | HM Treasury, audit form | 74,956 B | 37 | 12 | exact, 37/37 | `[trash]/0000.dat`, `[trash]/0001.dat` |
| `dwp_atw_withdrawn_annex_b.docx` | DWP, consent form (withdrawn) | 34,597 B | 19 | 0 | exact, 19/19 | none |
| `oag_style_guidance_2016.docx` | Office of the Advocate General for Scotland, prose | 53,307 B | 20 | 0 | exact, 20/20 | none |

The detailed per-file part differences, package inventory, character counts, warnings, and measured durations are in `analysis.json`. The timings are one-run diagnostic data only and were **not** used to adjust the project's speed score.

## Scope and cautions

- This is a convenience sample: **six HM Treasury documents**, one DWP form, and one Office of the Advocate General for Scotland document, all obtained through GOV.UK. It is highly UK-public-sector-specific and not statistically representative.
- No user-private document was accessed. This does not validate private business documents, documents from other jurisdictions, or every DOCX feature.
- There is no independently rendered Word/LibreOffice text gold-standard for this set. Package preservation is directly measured; text-extraction counts from the two libraries use different scope (docviews reports all of its text-bearing parts, while the comparison reports python-docx body/table text) and are **not** treated as a correctness score.
- The external smoke result does not replace the 832-file fixture census. It adds independent-source package evidence, but the project's broader reading-fidelity claims still need a larger, more diverse real-world evaluation.
- Reuse terms are recorded per source page and remain subject to third-party exceptions, especially for logos, emblems, photographs, and embedded assets. Keep the attribution and licence note with any use of these files.

## Acquisition checks

A GOV.UK page advertised the Devolution factsheet at a URL ending in `.docx`, but the official asset returned HTTP 200 with `Content-Type: application/pdf`, `%PDF-1.6` signature, 237,526 bytes, and SHA-256 `f2834609984f31cd9bd32d812cc65d962aa5b221d1de468bcfdc53af6bef72e9`. It was excluded as a PDF, not counted as a DOCX. The downloader rejects non-ZIP data even if the URL ends in `.docx`.

The New Zealand Ombudsman's annual report was not included: its verified copyright terms allow personal-use copying with attribution and require contacting the office for other use. The World Bank candidate was excluded because permission was required. Neither appears in the manifest or corpus.

## Reproduce

From `/home/user/docviews` (use Python with `python-docx==1.2.0` for the comparison):

```bash
python3 real_world_eval/download.py --force
PYTHONPATH=src python3 real_world_eval/analyze.py
```

The downloader validates the ZIP signature, records SHA-256 values, and writes both JSON and CSV manifests. The evaluator checks the input hash, opens/saves each file with both libraries, and compares package entry names and uncompressed part bytes. It does not change the main `corpus/` directory.
