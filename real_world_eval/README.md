# Initial real-world DOCX smoke set

This is a deliberately small, openly sourced **external smoke set**, added after
benchmarking the `docviews`/`python-docx` comparison exposed that the original
832-file corpus is entirely drawn from software test suites. It is not a
statistically random sample, is not a substitute for the 832-file fixture census,
and does **not** validate Word-visible text against a trusted renderer.

## What it tests

Eight authentic public-sector DOCX files fetched from official GOV.UK asset URLs:

- one 2.3 MB HM Treasury guidance document with many paragraphs and tables;
- five current HM Treasury operational forms/checklists (CG-02/03/04/05);
- one Department for Work and Pensions consent form (listed as withdrawn);
- one Office of the Advocate General for Scotland document released through FOI.

Each of the four GOV.UK source pages' standard footer says all content is
available under the Open Government Licence v3.0, except where otherwise stated;
no contrary document-specific reuse terms were found on these pages. The licence
still has exclusions, including third-party rights. The manifest links to the
licence, GOV.UK terms, and each source page, and records the direct URL, HTTP
response, file size, and SHA-256 in `manifest.csv` and `manifest.json`. The
corpus is for local testing only. Do not redistribute embedded logos, emblems,
or other third-party assets beyond their applicable licence.

## Initial result (2026-10-29)

All eight files opened in both libraries. On unchanged save, docviews retained all
248/248 package entries byte-identically. python-docx 1.2.0 retained 246/248
entries and 128/246 common package parts byte-identically; its CG-05 round-trip
omitted two unreferenced `[trash]/*.dat` entries. This is package-preservation
evidence only—not proof of reader-visible content loss or of text fidelity. Full
results and limitations are in [`REPORT.md`](REPORT.md); per-file data are in
`analysis.json`.

## Reproduce

```bash
python3 real_world_eval/download.py
PYTHONPATH=src python3 real_world_eval/analyze.py
```

For an apples-to-apples comparison with the published head-to-head, use
`python-docx==1.2.0`. The analysis records each library's own open/save time,
entry and part counts, whether an unchanged save is byte-identical per part,
and the parts/entries that differ. `analysis.json` is machine-readable.

**Interpretation limit:** preservation is a package-level observation. It does
not prove text fidelity. `docviews` reports all text it sees, and `python-docx`
reports public body/table text, but only an independent Word/LibreOffice-rendered
comparison can establish whether a specific string is visible to a reader. This
smoke set does not modify the main fixture corpus or its published counts.

## Acquisition note

A separate GOV.UK publication page labelled a factsheet as DOCX, but the linked
asset returned HTTP 200 with `Content-Type: application/pdf` and `%PDF-1.6`
magic bytes. It was excluded rather than treating a mislabelled PDF as a DOCX:

- page: https://www.gov.uk/government/publications/devolution-guidance-for-civil-servants/devolution-guidance-for-civil-servants
- URL: https://assets.publishing.service.gov.uk/media/6915fe2a493305b49ce6e793/Devolution_factsheet_2025.docx
- observed length: 237,526 bytes; SHA-256: `f2834609984f31cd9bd32d812cc65d962aa5b221d1de468bcfdc53af6bef72e9`

This is a source-publishing discrepancy, not a `docviews` or `python-docx`
result. The downloader rejects non-ZIP content under a `.docx` URL.
