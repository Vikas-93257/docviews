# Changelog

## 1.0.0 — Stable release metadata selected (2026-11-01; not published)

The first stable version is selected after completing the documented-public-API
parity assessment against python-docx 1.2.0. The package metadata now uses
version `1.0.0` and the `Production/Stable` classifier. This is a local release
candidate only: no PyPI upload has been made, and publication requires explicit
maintainer approval.

- The source-backed API matrix contains **584 documented records**: **578
  in-scope behavior matches**, **0 partial or unassessed in-scope records**, and
  **6 structural/internal exclusions** under the agreed boundary. Availability,
  callable shape, writability, enum, and import-path checks are recorded in
  `API_PARITY_MATRIX.md`.
- All **471 runtime enum members** match python-docx 1.2.0. The documented-only
  `MSO_THEME_COLOR_INDEX.MIXED` is implemented with Microsoft's published value
  `-2`; it has no XML token mapping because the upstream runtime omits it.
- The full **239-test suite** passed on CPython 3.9.25, 3.10.22, 3.11.18,
  3.12.16, 3.13.16, 3.14.10, and 3.15.2. Each run reported one upstream
  python-docx style-ID deprecation warning.
- The twelve selected python-docx documentation recipes produce matching
  save-and-reopen transcripts.
- The Sphinx 9.2.0 reference-docs build succeeded with one explicitly accepted
  upstream docutils warning at `research/python-docx-v1.2.0/docs/user/comments.rst:134`.
  The pinned upstream source was not changed.
- The wheel and source distribution were rebuilt after the metadata change,
  passed strict Twine checks, installed in clean Python 3.15.2 environments,
  and passed DOCX create/save/reopen and CLI smoke tests. Local artifacts and
  checksums are in `release_artifacts/1.0.0/`.

The parity claim is limited to the documented public API of python-docx 1.2.0.
Undocumented lxml/OOXML/OPC internals and malformed states outside the
published contract remain out of scope. No PyPI upload is authorized by this
readiness work.

---

## 0.3.0a1 — Historical unpublished alpha checkpoint (2026-10-31)

`0.3.0a1` was an internal checkpoint used before the documented-public-API
parity work was complete. It was never published and is superseded by the
`1.0.0` stable metadata selection. Its temporary local archives were removed
after the final stable artifacts were validated.
