# Migrating from python-docx

`docviews.compat` is a bridge for a **tested subset** of python-docx's documented
API. It is useful for a controlled pilot; it is not a drop-in `docx` package
and does not guarantee identical behavior for every application. The project
is Alpha. Read this guide and test your own files before switching a production
workflow.

## Import changes

If your application only imports `docx` and uses the documented object API:

```diff
- import docx
+ import docviews.compat as docx
```

Update direct imports too:

```diff
- from docx.shared import Inches, Pt
- from docx.enum.text import WD_ALIGN_PARAGRAPH
+ from docviews.compat.shared import Inches, Pt
+ from docviews.compat.enum.text import WD_ALIGN_PARAGRAPH
```

`docviews.compat` does **not** install a top-level `docx` module. A third-party
package that imports `docx` internally will continue to use python-docx.

## Common calls

The compatibility layer supports document creation, paragraphs/runs, common
formatting, tables/cells, sections, headers/footers, comments, inline pictures,
and core properties. For example:

```python
import io
import docviews.compat as docx
from docviews.compat.shared import Inches
from docviews.compat.enum.text import WD_BREAK

report = docx.Document()
report.add_heading("Quarterly report", 0)
paragraph = report.add_paragraph("Revenue: ")
run = paragraph.add_run("1,204,332", style="Emphasis")
run.bold = True
run.add_break(break_type=WD_BREAK.PAGE)
report.add_picture(image_path_or_stream=io.BytesIO(png_bytes),
                   width=Inches(2))

output = io.BytesIO()
report.save(path_or_stream=output)  # returns None, as python-docx does
```

Recent parity work also covers character-style run creation, `table.cell(row_idx,
col_idx)` and row/column cell keywords, and picture insertion from a run in the
body, a header, or a footer. The current regressions live in
`tests/test_compat_parity.py`.

## Behavior to check deliberately

- Run your app's tests against real, representative documents. Include headers,
  footers, tables, comments, revisions, pictures, and any unusual templates it
  processes.
- Compare generated documents after reopening them with the application that
  consumes them (at least python-docx 1.2.0; ideally Word or LibreOffice too).
- Check any code that uses enum identity, error text, XML edge cases, or
  undocumented methods. `COMPAT.md` records the current bounded audit:
  **276/276** sampled public names, **78/78** sampled callable parameter shapes
  on Python 3.13, and **471/471** enum members. These are inventory and selected
  signature checks, not full documented-API behavioral equivalence.
- For a low-risk rollout, keep python-docx available, pilot docviews on a
  non-critical path, and retain an easy rollback to the original imports.

## Not covered by the compatibility bridge

Changing imports will not migrate code that depends on:

- `docx.oxml`, `OxmlElement`, `qn`, lxml elements, XPath, or the `docx.opc`
  package/part internals;
- attributes such as `document.element` or `document.part`;
- unsupported authoring such as charts, fields, floating/anchored shapes, or
  table-of-contents objects;
- `.doc` input or document rendering.

For those workflows, keep python-docx or select a tool designed for the
required feature. docviews can read and preserve some unsupported structures
without authoring or rendering them; preservation is not the same as editing
support.

## Release and support status

At the 2026-10-31 release-readiness check, the selected Alpha candidate is
version `0.3.0a1`; no PyPI release has been made. The public GitHub repository
is [Vikas-93257/docviews](https://github.com/Vikas-93257/docviews), and source
is on `main`. The [issue tracker](https://github.com/Vikas-93257/docviews/issues)
is open. Hosted Actions run #1 passed all seven Python jobs and the
package-validation job; run #2 passed after updating to Node 24-compatible
actions and pinning Ubuntu 24.04. The full suite passed **149 tests on each
Python version from 3.9 through 3.15** in the local matrix; after selecting
`0.3.0a1`, it was rerun on Python 3.9 and 3.15. The full documented-API
behavior matrix remains unfinished, and there is no production SLA or support
commitment.
