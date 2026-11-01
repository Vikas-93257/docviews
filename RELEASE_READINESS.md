# Release readiness — 2026-11-01

## Release decision

The agreed target—parity with the **documented public API of python-docx 1.2.0**,
excluding undocumented lxml/OOXML/OPC internals—is assessed as complete. The
source-backed matrix records **578 tested in-scope behavior matches, 0 partial
or not-assessed in-scope records, and 6 structural/internal scope exclusions**.
The enum policy, Python-version range, docs warning, and package artifacts have
all been reviewed.

After that parity assessment, stable release metadata was selected as **1.0.0**
with the `Development Status :: 5 - Production/Stable` classifier. This is a
local metadata and artifact decision only: **no PyPI upload has been made or
authorized**. The temporary `0.3.0a1` validation archives were removed after
the stable 1.0.0 artifacts passed, to avoid confusion.

## Six readiness items

1. **Enum discrepancies — successfully completed.** The 471 runtime members
   match python-docx 1.2.0. The documentation-only
   `MSO_THEME_COLOR_INDEX.MIXED` is implemented with Microsoft's published
   value `-2`, separately from the upstream runtime omission; its XML value is
   unmapped.
2. **Partial behavior assessments — successfully completed.** All 578
   in-scope documented behavior records have differential evidence and are
   `tested_match`; there are no partial or unassessed in-scope records. Six
   structural/internal records are excluded under the agreed boundary.
3. **Python 3.9–3.15 support — prior matrix completed; post-fix matrix pending.**
   The pre-Windows-diagnostic source passed **239 tests on each** of CPython
   3.9.25, 3.10.22, 3.11.18, 3.12.16, 3.13.16, 3.14.10, and 3.15.2, with one
   upstream style-ID deprecation warning per run. After the stream-order fix,
   the full suite passed on CPython 3.13.16 (**239 passed**, one upstream
   warning) and on Windows CPython 3.14 (**239 passed**, one upstream warning,
   163.42s). Hosted CI must still confirm the final code on all seven versions.
4. **Docs-build warning — successfully completed by explicit acceptance.** A
   Sphinx 9.2.0 build of the pinned reference docs succeeded with one
   docutils warning at `research/python-docx-v1.2.0/docs/user/comments.rst:134`:
   `Unknown target name: "container"`. The source markup `_block-item container_`
   is an upstream documentation typo. It does not affect the runtime or API
   inventory. The reference checkout remains unchanged; the temporary build
   copy used a modern Sphinx-compatible `intersphinx_mapping`.
5. **Final artifact build and validation — successfully completed.** The wheel
   and sdist were rebuilt for 1.0.0 on Python 3.15.2 after source, docs, and
   metadata changes. Both passed `python -m twine check --strict`; each was
   clean-installed in its own Python 3.15.2 environment and passed metadata,
   DOCX create/save/reopen, `docviews --version`, and `docviews --help` checks.
   The sdist has 84 entries, includes the required package data/docs/licenses,
   and excludes tests. Checksums are in `release_artifacts/1.0.0/SHA256SUMS.txt`.
6. **Stable-release metadata — successfully completed after parity.** The
   selected version is `1.0.0`, the development-status classifier is
   `Production/Stable`, and package author metadata is `Vikas Maurya` without
   an invented email. This does **not** authorize publication.

## Final local artifacts

- `release_artifacts/1.0.0/docviews-1.0.0-py3-none-any.whl` — **147,615 bytes**
  SHA-256: `71d225b8e4052a7b96e9db018866a0bb29aaad76e508ee17c350d90616d69357`
- `release_artifacts/1.0.0/docviews-1.0.0.tar.gz` — **181,911 bytes**
  SHA-256: `c8401726076865b31bad3cce8d821defc140653da4ba88aa702a8ed7c10405a5`
- `release_artifacts/1.0.0/SHA256SUMS.txt` — checksum manifest

These files are local validation artifacts only. No superseded alpha archives
remain in the workspace.

## Prepared operational workflow

A non-automatic release workflow is staged locally at
`.github/workflows/publish.yml`. It is dispatched from `main` and only enables
publication when the release-tag input is `v1.0.0`, the version input is
`1.0.0`, and the typed confirmation is `PUBLISH docviews 1.0.0`. It runs the
full seven-version test matrix, rebuilds and checks the distributions, and
publishes only through a protected GitHub Environment named `pypi` using OIDC.
A missing/wrong confirmation skips the publish job. Both workflow files pass
`actionlint` 1.7.12, and the CI cache key uses `pyproject.toml`. A Windows
CPython 3.14/3.13 run exposed four POSIX-only test paths, a test that compared
shape counts from different libraries, and a real ordering mismatch between
tree and stream reads for text boxes. The paths and assertion were corrected;
the stream reader now emits an anchor paragraph before its text-box paragraphs,
and corpus sampling is normalized across operating systems. After these fixes,
a disposable CPython 3.13.16 checkout passed **239 tests** with one upstream
style-ID warning; wheel/sdist build, strict Twine checks, clean installs, DOCX
round-trip, and CLI smoke checks passed. The final source was then applied on
Windows and its full CPython 3.14 suite passed (**239 tests**, one upstream
warning, 163.42s). Hosted CI must still confirm the final code on all seven
Python versions. No GitHub-hosted run, OIDC exchange, or PyPI publish has occurred.

`RELEASE_RUNBOOK.md` records the manual setup steps. No PyPI token is stored in
the workspace, and no workflow dispatch, Git push, or upload has been made.

## Remaining before any public release

1. Review and push the local source, docs, CI cache fix, both workflow files,
   and runbook to `https://github.com/Vikas-93257/docviews` from Windows. The
   shared workspace
   has no Git metadata or configured remote/authentication; no unsupported push
   or repository-creation attempt was made.
2. Wait for hosted CI on the pushed revision and confirm all seven Python jobs
   and the package-validation job pass.
3. Protect `main` against unreviewed workflow changes; configure the GitHub
   `pypi` environment with a required reviewer and deployments from `main`; add
   a tag ruleset preventing release tags such as `v1.0.0` from being moved or
   deleted. Configure a PyPI Trusted Publisher for owner `Vikas-93257`,
   repository `docviews`, workflow `.github/workflows/publish.yml`, environment
   `pypi`. PyPI's pending-publisher setup does **not** reserve the package name.
4. The PyPI JSON endpoint returned `Not Found` for `docviews` on 2026-11-01.
   This is point-in-time evidence only; recheck immediately before release.
5. Obtain explicit final approval. Only then create/push tag `v1.0.0`, manually
   dispatch the workflow from `main` with tag/version inputs and the exact
   confirmation, review the build/test results, and approve the protected
   environment. This readiness work and workflow preparation are **not**
   authorization to upload.

No user-private documents were inspected.

## Evidence files

- `API_PARITY_MATRIX.md` — source-backed public-API inventory and differential
  results.
- `docs/api-parity/python_docx_1_2_0_comparison.json` — 584 comparisons and
  the 2026-10-31 comparison date.
- `docs/api-parity/python_docx_1_2_0_behavior.json` — tested behavior evidence,
  Python-version results, docs-warning review, and distribution validation.
- `docs/api-parity/python_docx_1_2_0_inventory.json` — official v1.2.0 source
  inventory.
- `.github/workflows/publish.yml` — manual, tag-guarded Trusted Publishing
  workflow; not pushed or dispatched.
- `.github/workflows/tests.yml` — seven-version test matrix and package check;
  cache now keys from `pyproject.toml`.
- `RELEASE_RUNBOOK.md` — account setup and maintainer steps for the first
  public release.
