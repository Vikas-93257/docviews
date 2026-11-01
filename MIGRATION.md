# Migrating from python-docx

`docviews.compat` implements python-docx 1.2.0's **documented public API** within
the agreed boundary; the source-backed parity matrix records 578 tested
behavior matches and six structural/internal scope exclusions. It is not a
top-level `docx` package: redirect imports to `docviews.compat`, and code that
depends on undocumented lxml/OOXML/OPC internals remains out of scope. Stable
release metadata `1.0.0` is selected but has not been published. Read this guide
and test your application-specific workflows before switching production use.

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
- The completed documented-public-API comparison is in
  `API_PARITY_MATRIX.md`: 578 tested behavior matches, no partial or
  unassessed in-scope rows, and six structural/internal exclusions. `COMPAT.md`
  remains a separate bounded object-sample audit (276/276 sampled names,
  78/78 sampled call shapes, and 471/471 runtime enum members).
- For a production migration, still run your application's integration suite,
  validate its actual DOCX inputs/outputs, and retain a rollback path. API
  parity does not prove suitability for workflows that depend on private
  internals or capabilities outside python-docx's documented API.

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

Stable version `1.0.0` and the `Production/Stable` classifier are selected after
completion of the documented-public-API parity assessment. The version has not
been uploaded to PyPI; release artifacts are local and publication requires
explicit maintainer approval. The repository URL is
[github.com/Vikas-93257/docviews](https://github.com/Vikas-93257/docviews), and
the [issue tracker](https://github.com/Vikas-93257/docviews/issues) is the
public support channel. No production SLA or commercial support commitment is
made. See `RELEASE_READINESS.md` for the remaining release approval steps and
artifact checksums.
