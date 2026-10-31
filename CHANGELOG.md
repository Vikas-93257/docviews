# Changelog

## 0.3.0a1 — Alpha release candidate (unpublished at the 2026-10-31 check)

This candidate combines the v0.3 authoring work with the compatibility and
release-readiness refinements completed through 2026-10-31:

- A `docviews.compat` bridge for a sampled subset of python-docx 1.2.0's
documented object and enum APIs. The latest audit records **276/276** sampled
public attribute names, **78/78** sampled callable parameter shapes on Python
3.13, and **471/471** enum members. These bounded counts are not full documented-
API behavioral parity and do not include the lxml/OOXML/OPC internals.
- Compatibility refinements for latent styles, run-content iteration and
rendered page breaks, Drawing/Image access, `RGBColor`, and python-docx-shaped
length-unit constructors and signatures.
- Authoring for paragraphs, runs, styles, tables, sections, headers/footers,
comments, core properties, and inline pictures, with unedited package parts
preserved byte-for-byte.
- `Run.add_picture()` support in the main document, header, and footer parts;
relationship IDs are reserved per source part and repeated images get distinct
media names. Keyword calls and stream-save behavior have focused regressions.
- Reading and auditing for original, accepted, and markup views, revisions,
comments and anchors, and text-bearing package parts.
- The twelve documented python-docx recipe transcripts match after save/reopen.
The full suite passed **149 tests on each Python version from 3.9 through
3.15** with python-docx 1.2.0 as the optional interoperability baseline. After
selecting metadata version `0.3.0a1`, the full suite was rerun on Python 3.9 and
3.15 (**149 passed on each**); fresh Python 3.15 wheel and sdist installs also
passed DOCX and CLI smoke checks.
- The five-run write benchmark shows docviews faster build-plus-save for the
measured 500- and 5,000-paragraph workloads; python-docx remains faster at
save-only for the 5,000-paragraph case. See `COMPARE.md` for timings and limits.
- An external smoke set of eight public-sector DOCX files from GOV.UK pages
whose footers state the Open Government Licence v3.0 (subject to its
third-party-rights exclusions) is documented separately from the fixture
corpus in `real_world_eval/REPORT.md`.

This remains an Alpha candidate, not a general drop-in replacement or a stable
1.0 release. Full documented-public-API parity is still the selected engineering
target, and the candidate has not been published as of the date above. See
`README.md`, `COMPARE.md`, `COMPAT.md`, and `MIGRATION.md` for usage,
measurements, compatibility scope, and migration limitations.
