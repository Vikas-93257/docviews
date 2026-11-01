# Public release record — docviews 1.0.0

`docviews==1.0.0` was published to PyPI through protected Trusted Publishing
from GitHub Actions on commit `10f6da4` on 2026-11-01. The original tag is
`v1.0.0`; do not move or recreate it. The `tests #5` tag run and release
workflow `Publish docviews to PyPI #2` passed. The release workflow ran the
seven-version test matrix, built and validated wheel/sdist distributions, and
completed the protected PyPI upload.

The project is installable with:

```console
python -m pip install docviews
```

## Repository documentation correction

The repository README and related guides are corrected so GitHub visitors
see accurate installation and publication status. This is a documentation-only
GitHub commit: it does not change the published 1.0.0 package, its version, or
its tag.

The original 1.0.0 PyPI long description still contains pre-publication wording
from the README at upload time. A GitHub documentation commit does not refresh
an existing PyPI release's description. Correcting PyPI itself would require a
separate future release and explicit maintainer approval; that is not part of
this documentation update.

## Windows maintainer steps

1. Apply the docs-only overlay to `C:\Test_ml\docviews`.
2. Review the changed paths and run `git diff --check`. Confirm no changes to
   `pyproject.toml`, `src/docviews/__init__.py`, the publish workflow, or the
   original `release_artifacts/1.0.0/` files.
3. Commit and push the documentation changes on `main`.
4. Wait for the hosted `tests` workflow on the docs commit to pass.

Do not create a new tag, dispatch the PyPI publish workflow, or upload a new
package for this repository-only documentation change.
