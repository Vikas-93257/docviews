# First PyPI release runbook — docviews 1.0.0

**Status:** prepared locally; not published. Both workflow files pass
`actionlint` 1.7.12. After fixing Windows portability/test issues and a
text-box stream-order bug, the release checks passed in a throwaway CPython
3.13.16 checkout: 239 tests passed, wheel/sdist built, strict Twine checks
passed, and clean-install/DOCX/CLI smoke checks passed. The full Windows
CPython 3.14 suite also passed (239 tests, one upstream warning, 163.42s).
GitHub-hosted CI and OIDC have not been exercised; the workflow has not been
pushed or dispatched. Remaining steps require repository/PyPI account access
and maintainer approval.

## Prepared release controls

- Stable version: `1.0.0` (`Production/Stable`).
- Locally validated artifacts: `release_artifacts/1.0.0/`; verify with
  `sha256sum -c release_artifacts/1.0.0/SHA256SUMS.txt`.
- Manual workflow: `.github/workflows/publish.yml`.
- Dispatch the workflow from `main`. It only enables publication when its tag
  input is `v1.0.0`, version input is `1.0.0`, and confirmation is exactly
  `PUBLISH docviews 1.0.0`.
- The publish job requires the GitHub Environment named `pypi` and uses PyPI
  Trusted Publishing via GitHub OIDC; it does not use a stored PyPI token.
- The workflow rebuilds from the tag, runs the Python 3.9–3.15 test matrix,
  checks metadata, builds wheel/sdist, runs strict Twine validation and a wheel
  smoke test, and only then enables the protected publish job.
- If the confirmation input is absent/wrong, the publish job is skipped. A
  human must still approve the protected `pypi` environment when configured.

## Remaining actions requiring maintainer access

1. **Review and push the local changes from Windows.** The shared workspace has
   no Git metadata or remote/authentication, so it has not pushed anything.
   Push the stable metadata, updated docs, the CI cache fix, both workflow
   files, and this runbook to `https://github.com/Vikas-93257/docviews`.
2. **Wait for hosted CI on the pushed revision.** Confirm all seven Python jobs
   and the package-validation job pass before creating a release tag.
3. **Protect the GitHub repository and environment.** Protect `main` against
   unreviewed workflow changes. Create the `pypi` environment, require
   maintainer review, and allow deployments from `main` (the manual dispatch
   runs on that branch). Separately add a tag ruleset for release tags such as
   `v*` so `v1.0.0` cannot be moved or deleted after review.
4. **Configure a PyPI Trusted Publisher.** In PyPI account publishing settings
   (or a pending publisher for a not-yet-created project), use:
   - project name: `docviews`
   - owner: `Vikas-93257`
   - repository: `docviews`
   - workflow file: `.github/workflows/publish.yml`
   - environment: `pypi`

   PyPI's pending-publisher mechanism does **not** reserve a project name; a
   different user could register it before the first upload. See the [PyPI
   Trusted Publishers guide](https://docs.pypi.org/trusted-publishers/creating-a-project-through-oidc/).
5. **Recheck the name immediately before publishing.** The PyPI JSON endpoint
   returned `Not Found` for `docviews` on 2026-11-01. This is point-in-time
   evidence only, not a reservation; the workflow also performs a fail-closed
   check during the build.
6. **Obtain explicit final release approval.** Only after review, hosted CI,
   environment protection, and publisher setup, create/push tag `v1.0.0`. Then
   manually dispatch the workflow from `main`, passing tag `v1.0.0`, version
   `1.0.0`, and the exact confirmation text. Approve the protected environment
   only after reviewing the run's test/build results. This runbook and workflow
   preparation do not authorize an upload.
7. After a successful upload, verify the PyPI project page and install
   `docviews==1.0.0` in a fresh environment. Record the resulting public
   artifact hashes and release URL.

**No PyPI upload or Git push has been performed.** The local artifacts are
validation outputs, not a published release.
