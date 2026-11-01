# Release record and documentation status — 2026-11-01

## Published 1.0.0

- **PyPI:** [docviews 1.0.0](https://pypi.org/project/docviews/1.0.0/)
- **Source tag:** `v1.0.0` on commit `10f6da4ee1942e55712d60e771db085d5d33c0b9`
- **Hosted tag CI:** `tests #5` passed on `v1.0.0`; seven Python-version jobs
  and package validation completed successfully.
- **Release workflow:** `Publish docviews to PyPI #2` succeeded from `main` on
  commit `10f6da4`; all seven release-test jobs, the build/validation job, and
  the protected PyPI publish job passed. The `pypi` environment approval was
  recorded in the run.
- The earlier `Publish docviews to PyPI #1` was a no-upload preflight: tests and
  build passed; its publish job was intentionally skipped.
- **Original 1.0.0 artifacts:** wheel 147,615 bytes, SHA-256
  `71d225b8e4052a7b96e9db018866a0bb29aaad76e508ee17c350d90616d69357`; sdist
  181,911 bytes, SHA-256
  `c8401726076865b31bad3cce8d821defc140653da4ba88aa702a8ed7c10405a5`. These
  remain recorded in `release_artifacts/1.0.0/` and must not be replaced.

The documented-public-API parity target against python-docx 1.2.0 remains
bounded to the public documented contract; undocumented lxml/OOXML/OPC internals
remain excluded. The Windows CPython 3.14 suite passed 239 tests with one
upstream warning (163.42s); the Linux CPython 3.13.16 suite passed 239 tests
with one upstream warning (153.14s).

## Repository documentation correction

The README and related release/migration/comparison documentation are corrected
to reflect that 1.0.0 is published and `pip install docviews` works. This is a
repository-only documentation update: it does not change package code, version
metadata, release tags, or the existing PyPI files.

The README included pre-publication wording when 1.0.0 was uploaded. As a
result, the long description currently displayed on PyPI still contains that
stale text even after the GitHub README is corrected. PyPI keeps the metadata
captured for an existing release; a GitHub documentation commit does not change
the 1.0.0 PyPI description. No new release or upload is part of this update.

## Pushing the documentation update from Windows

1. Apply the documentation-only correction overlay to `C:\Test_ml\docviews`.
   It must not change `pyproject.toml`, `src/docviews/__init__.py`,
   `.github/workflows/publish.yml`, or `release_artifacts/1.0.0/`.
2. Run `git diff --check` and review `git status --short`. The changes should be
   documentation/evidence files only.
3. Commit and push the documentation changes to `main`. Do not create a new
   release tag or dispatch the PyPI workflow for this update.
4. Confirm the hosted `tests` workflow passes on the documentation commit.

The user pushes from Windows; no Git remote or authentication is available in
this workspace. No private user documents were inspected.
