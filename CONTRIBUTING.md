# Contributing to docviews

Thanks for helping improve docviews. It is an Alpha OOXML reader/writer, so
changes should favor explicit behavior, package preservation, and regression
tests over silent approximations.

## Development setup

Use Python 3.9 or newer, then install the optional test dependencies:

```bash
python -m pip install -e ".[dev]"
python -m pytest -o addopts='' -q
```

The interoperability baseline is python-docx 1.2.0. The runtime package has
no third-party dependencies; python-docx is used only by development tests.
The full suite passed locally on Python 3.9–3.15. The first hosted run on the
public repository passed all seven Python jobs and package validation. A
follow-up workflow update to remove action-runtime and runner-image notices is
prepared; run CI again after pushing that update.

## Build and validate distributions

Install the packaging frontend and validator, then build both the wheel and the
source distribution from a clean checkout:

```bash
python -m pip install build twine
rm -rf build dist
python -m build
python -m twine check --strict dist/*
```

The CI workflow also installs the built wheel into an isolated target and
checks that it can create, save, and reopen a DOCX and start the `docviews` CLI.
Never upload a candidate during local validation; release approval and the
configured publisher identity are separate gates.

## Changes and tests

- Add a focused regression test for each behavior change. For compatibility
  fixes, exercise the python-docx call form and reopen saved output with
  python-docx 1.2.0.
- For package edits, check both the content and the relevant relationships and
  content-type entries. Preserve untouched parts byte-for-byte.
- Avoid adding personal, confidential, or unlicensed customer documents. Use
  small synthetic fixtures where possible; otherwise document source and reuse
  terms before proposing a fixture.
- Keep unsupported OOXML explicit. Do not claim a feature is authored merely
  because the reader can preserve its XML.
- Run the documented recipe diff when changing the compatibility layer:

  ```bash
  python examples/docs_recipes.py > /tmp/python-docx.txt
  PYTHONPATH=src python examples/docs_recipes.py --shim > /tmp/docviews.txt
  diff -u /tmp/python-docx.txt /tmp/docviews.txt
  ```

- Reproduce the in-memory build/save comparison after performance changes:

  ```bash
  PYTHONPATH=src python tools/bench_compat_write.py
  ```

## Pull requests and support

A useful change includes a clear problem statement, a minimal reproduction, a
test, and any user-facing limitations. Avoid promising general compatibility
from a single passing example; update `COMPAT.md`, `MIGRATION.md`, `COMPARE.md`,
or the changelog when scope, evidence, or known limitations change.

The public repository is [Vikas-93257/docviews](https://github.com/Vikas-93257/docviews),
and its [issue tracker](https://github.com/Vikas-93257/docviews/issues) is the
public channel for non-sensitive reports. The source is now on `main`; the first
hosted run passed, and a follow-up workflow cleanup is pending a push and rerun.
There is no support SLA. Do not post private documents or security details in
public issues.
