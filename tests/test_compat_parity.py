"""Regression tests for python-docx-shaped values and mutation behavior."""
from __future__ import annotations
import importlib
import inspect
import io
import os
import subprocess
import sys
import zipfile
from datetime import datetime, timedelta, timezone
from pathlib import Path
from xml.etree import ElementTree as ET
import docx
import pytest
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
import docviews.compat as compat
from docviews.compat.enum import MSO_COLOR_TYPE, MSO_THEME_COLOR, WD_ALIGN_PARAGRAPH, WD_BREAK, WD_CELL_VERTICAL_ALIGNMENT, WD_COLOR_INDEX, WD_INLINE_SHAPE, WD_LINE_SPACING, WD_ROW_HEIGHT_RULE, WD_SECTION_START, WD_STYLE_TYPE, WD_TAB_ALIGNMENT, WD_TAB_LEADER, WD_TABLE_ALIGNMENT, WD_TABLE_DIRECTION, WD_UNDERLINE
from docviews.compat.shared import Inches, Pt, RGBColor
from conftest import make_docx
import blobs
CORE_PROPERTY_FIELDS = ('title', 'author', 'subject', 'comments', 'category', 'keywords', 'last_modified_by', 'revision', 'content_status', 'identifier', 'language', 'version', 'created', 'modified', 'last_printed')
CORE_PROPERTY_STRING_FIELDS = ('title', 'author', 'subject', 'comments', 'category', 'keywords', 'last_modified_by', 'content_status', 'identifier', 'language', 'version')
CORE_PROPERTY_DATE_FIELDS = ('created', 'modified', 'last_printed')
DOCUMENTED_PROXY_IMPORT_PATHS = ('docviews.compat.dml.color.ColorFormat', 'docviews.compat.opc.coreprops.CoreProperties', 'docviews.compat.section.Section', 'docviews.compat.section.Sections', 'docviews.compat.section._Footer', 'docviews.compat.section._Header', 'docviews.compat.settings.Settings', 'docviews.compat.shape.InlineShape', 'docviews.compat.shape.InlineShapes', 'docviews.compat.styles.latent.LatentStyles', 'docviews.compat.styles.latent._LatentStyle', 'docviews.compat.styles.style.BaseStyle', 'docviews.compat.styles.style.CharacterStyle', 'docviews.compat.styles.style.ParagraphStyle', 'docviews.compat.styles.style._NumberingStyle', 'docviews.compat.styles.style._TableStyle', 'docviews.compat.styles.styles.Styles', 'docviews.compat.table.Table', 'docviews.compat.table._Cell', 'docviews.compat.table._Column', 'docviews.compat.table._Columns', 'docviews.compat.table._Row', 'docviews.compat.table._Rows', 'docviews.compat.text.font.Font', 'docviews.compat.text.hyperlink.Hyperlink', 'docviews.compat.text.pagebreak.RenderedPageBreak', 'docviews.compat.text.paragraph.Paragraph', 'docviews.compat.text.parfmt.ParagraphFormat', 'docviews.compat.text.run.Font', 'docviews.compat.text.run.Run', 'docviews.compat.text.tabstops.TabStop', 'docviews.compat.text.tabstops.TabStops')

def _core_property_values(properties):
    return {name: getattr(properties, name) for name in CORE_PROPERTY_FIELDS}

def _capture_exception(operation):
    try:
        operation()
        return
    except Exception as exc:
        return (type(exc), str(exc))

def _patch_first_table_row_grid_edges(package: bytes, before: int, after: int) -> bytes:
    word_ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    namespaces = {'w': word_ns}
    output = io.BytesIO()
    ET.register_namespace('w', word_ns)
    with zipfile.ZipFile(io.BytesIO(package), 'r') as source:
        with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as target:
            for member in source.infolist():
                data = source.read(member.filename)
                if member.filename == 'word/document.xml':
                    root = ET.fromstring(data)
                    table = root.find('.//w:tbl', namespaces)
                    row = table.find('w:tr', namespaces)
                    cells = [child for child in row if child.tag == f'{{{word_ns}}}tc']
                    py_assert2 = len(cells)
                    py_assert5 = 2
                    py_assert4 = py_assert2 >= py_assert5
                    if not py_assert4:
                        py_format7 = pytest_ar._call_reprcompare(('>=',), (py_assert4,), ('%(py3)s\n{%(py3)s = %(py0)s(%(py1)s)\n} >= %(py6)s',), (py_assert2, py_assert5)) % {'py0': pytest_ar._saferepr(len) if 'len' in py_builtins.locals() or pytest_ar._should_repr_global_name(len) else 'len', 'py1': pytest_ar._saferepr(cells) if 'cells' in py_builtins.locals() or pytest_ar._should_repr_global_name(cells) else 'cells', 'py3': pytest_ar._saferepr(py_assert2), 'py6': pytest_ar._saferepr(py_assert5)}
                        py_format9 = 'assert %(py8)s' % {'py8': py_format7}
                        raise AssertionError(pytest_ar._format_explanation(py_format9))
                    py_assert2 = py_assert4 = py_assert5 = None
                    row.remove(cells[0])
                    row.remove(cells[-1])
                    for child in list(row):
                        if child.tag == f'{{{word_ns}}}trPr':
                            row.remove(child)
                    properties = ET.Element(f'{{{word_ns}}}trPr')
                    grid_before = ET.SubElement(properties, f'{{{word_ns}}}gridBefore')
                    grid_before.set(f'{{{word_ns}}}val', str(before))
                    grid_after = ET.SubElement(properties, f'{{{word_ns}}}gridAfter')
                    grid_after.set(f'{{{word_ns}}}val', str(after))
                    row.insert(0, properties)
                    data = ET.tostring(root, encoding='utf-8', xml_declaration=True)
                target.writestr(member, data)
                continue
    return output.getvalue()

def _P(text: str) -> str:
    return f'<w:p><w:r><w:t>{text}</w:t></w:r></w:p>'

def test_all_official_documentation_recipe_transcripts_match_after_reopen():
    """Run the documented examples through both implementations and diff output."""
    root = Path(__file__).resolve().parents[1]
    example = root / 'examples' / 'docs_recipes.py'
    reference_source = Path(docx.__file__).resolve().parent.parent
    candidate_source = root / 'src'
    env = os.environ.copy()
    paths = [str(reference_source), str(candidate_source)]
    if env.get('PYTHONPATH'):
        paths.append(env['PYTHONPATH'])
    env['PYTHONPATH'] = os.pathsep.join(paths)

    def transcript(*args):
        return subprocess.run([sys.executable, str(example), *args], cwd=root, env=env, check=True, capture_output=True, text=True).stdout
    reference = transcript()
    candidate = transcript('--shim')
    py_assert1 = reference == candidate
    assert not not py_assert1
    py_assert1 = None
    py_assert1 = reference.count
    py_assert3 = '### r'
    py_assert5 = py_assert1(py_assert3)
    py_assert8 = 12
    py_assert7 = py_assert5 == py_assert8
    assert not not py_assert7
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert8 = None

def test_documented_proxy_classes_are_importable_from_mirrored_modules():
    for path in DOCUMENTED_PROXY_IMPORT_PATHS:
        parts = path.split('.')
        resolved = None
        for split_at in range(len(parts), 0, -1):
            try:
                resolved = importlib.import_module('.'.join(parts[:split_at]))
                for component in parts[split_at:]:
                    resolved = getattr(resolved, component)
            except (ImportError, AttributeError):
                if split_at == 1:
                    raise
                continue
            break
        py_assert3 = isinstance(resolved, type)
        assert not not py_assert3
        py_assert3 = None
        continue

def test_sampled_public_inventory_and_callable_shapes_have_no_gaps():
    import compat_coverage as audit
    rows = audit.measure()
    py_assert1 = (not row['missing'] for row in rows)
    py_assert3 = all(py_assert1)
    assert not not py_assert3
    py_assert1 = py_assert3 = None
    checked, compared, matched, mismatches, unavailable = audit.signature_audit()
    py_assert4 = compared + unavailable
    py_assert1 = checked == py_assert4
    assert not not py_assert1
    py_assert1 = py_assert4 = None
    py_assert1 = compared == matched
    assert not not py_assert1
    py_assert1 = None
    py_assert2 = []
    py_assert1 = mismatches == py_assert2
    assert not not py_assert1
    py_assert1 = py_assert2 = None

def test_paragraph_text_style_and_alignment_are_read_write_compatible(tmp_path):
    reference = docx.Document()
    candidate = compat.Document()
    ref_paragraph = reference.add_paragraph('before')
    our_paragraph = candidate.add_paragraph('before')
    ref_paragraph.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.RIGHT
    our_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    py_assert1 = ref_paragraph.alignment
    py_assert5 = docx.enum
    py_assert7 = py_assert5.text
    py_assert9 = py_assert7.WD_ALIGN_PARAGRAPH
    py_assert11 = py_assert9.RIGHT
    py_assert3 = py_assert1 is py_assert11
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = py_assert11 = None
    py_assert1 = our_paragraph.alignment
    py_assert5 = WD_ALIGN_PARAGRAPH.RIGHT
    py_assert3 = py_assert1 is py_assert5
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert5 = None
    py_assert1 = ref_paragraph.paragraph_format
    py_assert3 = py_assert1.alignment
    py_assert7 = docx.enum
    py_assert9 = py_assert7.text
    py_assert11 = py_assert9.WD_ALIGN_PARAGRAPH
    py_assert13 = py_assert11.RIGHT
    py_assert5 = py_assert3 is py_assert13
    assert not not py_assert5
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = py_assert11 = py_assert13 = None
    py_assert1 = our_paragraph.paragraph_format
    py_assert3 = py_assert1.alignment
    py_assert7 = WD_ALIGN_PARAGRAPH.RIGHT
    py_assert5 = py_assert3 is py_assert7
    assert not not py_assert5
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = None
    ref_paragraph.style = reference.styles['Heading 1']
    our_paragraph.style = candidate.styles['Heading 1']
    ref_paragraph.text = 'after'
    our_paragraph.text = 'after'
    py_assert1 = ref_paragraph.text
    py_assert6 = our_paragraph.text
    py_assert3 = py_assert1 == py_assert6
    py_assert8 = 'after'
    py_assert4 = py_assert6 == py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = ref_paragraph.style
    py_assert3 = py_assert1.name
    py_assert8 = our_paragraph.style
    py_assert10 = py_assert8.name
    py_assert5 = py_assert3 == py_assert10
    py_assert12 = 'Heading 1'
    py_assert6 = py_assert10 == py_assert12
    assert not (not py_assert5 or not py_assert6)
    py_assert1 = py_assert3 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
    py_assert1 = ref_paragraph.style
    py_assert3 = py_assert1.style_id
    py_assert7 = our_paragraph.style
    py_assert9 = py_assert7.style_id
    py_assert5 = py_assert3 == py_assert9
    assert not not py_assert5
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = None
    ref_path = tmp_path / 'reference-paragraph.docx'
    our_path = tmp_path / 'candidate-paragraph.docx'
    reference.save(ref_path)
    candidate.save(our_path)
    reopened_reference = docx.Document(str(ref_path)).paragraphs[0]
    reopened_candidate = compat.Document(our_path.read_bytes()).paragraphs[0]
    py_assert1 = reopened_reference.text
    py_assert6 = reopened_candidate.text
    py_assert3 = py_assert1 == py_assert6
    py_assert8 = 'after'
    py_assert4 = py_assert6 == py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = reopened_reference.alignment
    py_assert5 = docx.enum
    py_assert7 = py_assert5.text
    py_assert9 = py_assert7.WD_ALIGN_PARAGRAPH
    py_assert11 = py_assert9.RIGHT
    py_assert3 = py_assert1 is py_assert11
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = py_assert11 = None
    py_assert1 = reopened_candidate.alignment
    py_assert5 = WD_ALIGN_PARAGRAPH.RIGHT
    py_assert3 = py_assert1 is py_assert5
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert5 = None
    py_assert1 = reopened_reference.style
    py_assert3 = py_assert1.name
    py_assert8 = reopened_candidate.style
    py_assert10 = py_assert8.name
    py_assert5 = py_assert3 == py_assert10
    py_assert12 = 'Heading 1'
    py_assert6 = py_assert10 == py_assert12
    assert not (not py_assert5 or not py_assert6)
    py_assert1 = py_assert3 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
    py_assert1 = reopened_reference.style
    py_assert3 = py_assert1.style_id
    py_assert7 = reopened_candidate.style
    py_assert9 = py_assert7.style_id
    py_assert5 = py_assert3 == py_assert9
    assert not not py_assert5
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = None

def test_quickstart_list_style_recipe_matches_python_docx(tmp_path):
    reference = docx.Document()
    candidate = compat.Document()
    ref_direct = reference.add_paragraph('Lorem ipsum dolor sit amet.', style='ListBullet')
    our_direct = candidate.add_paragraph('Lorem ipsum dolor sit amet.', style='ListBullet')
    ref_after = reference.add_paragraph('A second list item.')
    our_after = candidate.add_paragraph('A second list item.')
    ref_after.style = 'List Bullet'
    our_after.style = 'List Bullet'
    for ref_paragraph, our_paragraph in ((ref_direct, our_direct), (ref_after, our_after)):
        py_assert1 = ref_paragraph.text
        py_assert5 = our_paragraph.text
        py_assert3 = py_assert1 == py_assert5
        assert not not py_assert3
        py_assert1 = py_assert3 = py_assert5 = None
        py_assert1 = ref_paragraph.style
        py_assert3 = py_assert1.name
        py_assert8 = our_paragraph.style
        py_assert10 = py_assert8.name
        py_assert5 = py_assert3 == py_assert10
        py_assert12 = 'List Bullet'
        py_assert6 = py_assert10 == py_assert12
        assert not (not py_assert5 or not py_assert6)
        py_assert1 = py_assert3 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
        py_assert1 = ref_paragraph.style
        py_assert3 = py_assert1.style_id
        py_assert8 = our_paragraph.style
        py_assert10 = py_assert8.style_id
        py_assert5 = py_assert3 == py_assert10
        py_assert12 = 'ListBullet'
        py_assert6 = py_assert10 == py_assert12
        assert not (not py_assert5 or not py_assert6)
        py_assert1 = py_assert3 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
        continue
    ref_path = tmp_path / 'reference-quickstart-list.docx'
    our_path = tmp_path / 'candidate-quickstart-list.docx'
    reference.save(ref_path)
    candidate.save(our_path)
    reopened_reference = docx.Document(str(ref_path))
    reopened_candidate = compat.Document(our_path.read_bytes())
    py_assert0 = [paragraph.text for paragraph in reopened_candidate.paragraphs]
    py_assert3 = [paragraph.text for paragraph in reopened_reference.paragraphs]
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None
    py_assert0 = [paragraph.style.name for paragraph in reopened_candidate.paragraphs]
    py_assert4 = [paragraph.style.name for paragraph in reopened_reference.paragraphs]
    py_assert2 = py_assert0 == py_assert4
    py_assert6 = ['List Bullet', 'List Bullet']
    py_assert3 = py_assert4 == py_assert6
    assert not (not py_assert2 or not py_assert3)
    py_assert0 = py_assert2 = py_assert3 = py_assert4 = py_assert6 = None

def test_document_factory_signature_matches_python_docx():
    py_assert1 = inspect.signature
    py_assert4 = compat.Document
    py_assert6 = py_assert1(py_assert4)
    py_assert10 = inspect.signature
    py_assert13 = docx.Document
    py_assert15 = py_assert10(py_assert13)
    py_assert8 = py_assert6 == py_assert15
    assert not not py_assert8
    py_assert1 = py_assert4 = py_assert6 = py_assert8 = py_assert10 = py_assert13 = py_assert15 = None

def test_document_add_paragraph_arguments_match_and_round_trip(tmp_path):
    reference = docx.Document()
    candidate = compat.Document()
    pairs = [(reference.add_paragraph(), candidate.add_paragraph()), (reference.add_paragraph('First line\nSecond line\tcontinued', style='Heading 2'), candidate.add_paragraph('First line\nSecond line\tcontinued', style='Heading 2')), (reference.add_paragraph('Styled paragraph', style=reference.styles['Subtitle']), candidate.add_paragraph('Styled paragraph', style=candidate.styles['Subtitle']))]
    for ref_paragraph, our_paragraph in pairs:
        py_assert1 = ref_paragraph.text
        py_assert5 = our_paragraph.text
        py_assert3 = py_assert1 == py_assert5
        assert not not py_assert3
        py_assert1 = py_assert3 = py_assert5 = None
        py_assert1 = ref_paragraph.style
        py_assert3 = py_assert1.name
        py_assert7 = our_paragraph.style
        py_assert9 = py_assert7.name
        py_assert5 = py_assert3 == py_assert9
        assert not not py_assert5
        py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = None
        py_assert1 = ref_paragraph.style
        py_assert3 = py_assert1.style_id
        py_assert7 = our_paragraph.style
        py_assert9 = py_assert7.style_id
        py_assert5 = py_assert3 == py_assert9
        assert not not py_assert5
        py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = None
        continue
    ref_path = tmp_path / 'reference-add-paragraph.docx'
    our_path = tmp_path / 'candidate-add-paragraph.docx'
    reference.save(ref_path)
    candidate.save(our_path)
    reopened_reference = docx.Document(str(ref_path)).paragraphs
    reopened_candidate = compat.Document(our_path.read_bytes()).paragraphs
    py_assert0 = [p.text for p in reopened_reference]
    py_assert3 = [p.text for p in reopened_candidate]
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None
    py_assert0 = [p.style.style_id for p in reopened_reference]
    py_assert3 = [p.style.style_id for p in reopened_candidate]
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None

def test_document_paragraphs_and_inner_content_use_python_docx_body_boundary():
    body = _P('body') + '<w:sdt><w:sdtContent>' + _P('inside content control') + '</w:sdtContent></w:sdt>'
    source = make_docx(body)
    ours = compat.Document(source)
    theirs = docx.Document(io.BytesIO(source))
    py_assert0 = [p.text for p in ours.paragraphs]
    py_assert3 = [p.text for p in theirs.paragraphs]
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None
    py_assert0 = [p.text for p in ours.paragraphs]
    py_assert3 = ['body']
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None
    py_assert0 = 'inside content control'
    py_assert4 = ours.text
    py_assert6 = py_assert4()
    py_assert2 = py_assert0 in py_assert6
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert4 = py_assert6 = None
    py_assert0 = [type(item).__name__ for item in ours.iter_inner_content()]
    py_assert3 = ['_CompatParagraph']
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None

def test_document_save_matches_python_docx_for_paths_and_streams(tmp_path):
    reference = docx.Document()
    candidate = compat.Document()
    reference.add_heading('Saved heading', level=2)
    candidate.add_heading('Saved heading', level=2)
    reference.add_table(rows=1, cols=1).cell(0, 0).text = 'saved cell'
    candidate.add_table(rows=1, cols=1).cell(0, 0).text = 'saved cell'

    def assert_loaded_matches(ref_doc, our_doc):
        py_assert0 = [(p.text, p.style.style_id) for p in ref_doc.paragraphs]
        py_assert3 = [(p.text, p.style.style_id) for p in our_doc.paragraphs]
        py_assert2 = py_assert0 == py_assert3
        assert not not py_assert2
        py_assert0 = py_assert2 = py_assert3 = None
        py_assert0 = [table.cell(0, 0).text for table in ref_doc.tables]
        py_assert3 = [table.cell(0, 0).text for table in our_doc.tables]
        py_assert2 = py_assert0 == py_assert3
        assert not not py_assert2
        py_assert0 = py_assert2 = py_assert3 = None
    ref_path = tmp_path / 'reference-save.docx'
    our_path = tmp_path / 'candidate-save.docx'
    py_assert1 = reference.save
    py_assert4 = py_assert1(ref_path)
    py_assert7 = None
    py_assert6 = py_assert4 is py_assert7
    assert not not py_assert6
    py_assert1 = py_assert4 = py_assert6 = py_assert7 = None
    py_assert1 = candidate.save
    py_assert4 = py_assert1(our_path)
    py_assert7 = None
    py_assert6 = py_assert4 is py_assert7
    assert not not py_assert6
    py_assert1 = py_assert4 = py_assert6 = py_assert7 = None
    assert_loaded_matches(docx.Document(str(ref_path)), compat.Document(our_path.read_bytes()))
    ref_stream, our_stream = (io.BytesIO(), io.BytesIO())
    py_assert1 = reference.save
    py_assert4 = py_assert1(ref_stream)
    py_assert7 = None
    py_assert6 = py_assert4 is py_assert7
    assert not not py_assert6
    py_assert1 = py_assert4 = py_assert6 = py_assert7 = None
    py_assert1 = candidate.save
    py_assert4 = py_assert1(our_stream)
    py_assert7 = None
    py_assert6 = py_assert4 is py_assert7
    assert not not py_assert6
    py_assert1 = py_assert4 = py_assert6 = py_assert7 = None
    ref_stream.seek(0)
    our_stream.seek(0)
    assert_loaded_matches(docx.Document(ref_stream), compat.Document(our_stream.read()))

def test_document_add_picture_dimensions_match_python_docx(tmp_path):
    reference = docx.Document()
    candidate = compat.Document()
    dimensions = ({}, {'width': docx.shared.Inches(1)}, {'height': docx.shared.Inches(0.75)}, {'width': docx.shared.Inches(0.8), 'height': docx.shared.Inches(0.6)})

    def add_all(document, length_type):
        shapes = []
        for options in dimensions:
            converted = {name: length_type(value.inches) for name, value in options.items()}
            shapes.append(document.add_picture(io.BytesIO(blobs.PNG_96), **converted))
        return shapes
    ref_shapes = add_all(reference, docx.shared.Inches)
    our_shapes = add_all(candidate, Inches)

    def snapshot(shape):
        return (shape.type.name, int(shape.width), int(shape.height))
    py_assert0 = [snapshot(shape) for shape in ref_shapes]
    py_assert3 = [snapshot(shape) for shape in our_shapes]
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None
    py_assert2 = reference.inline_shapes
    py_assert4 = len(py_assert2)
    py_assert10 = candidate.inline_shapes
    py_assert12 = len(py_assert10)
    py_assert6 = py_assert4 == py_assert12
    py_assert14 = 4
    py_assert7 = py_assert12 == py_assert14
    assert not (not py_assert6 or not py_assert7)
    py_assert2 = py_assert4 = py_assert6 = py_assert7 = py_assert10 = py_assert12 = py_assert14 = None
    py_assert0 = [snapshot(shape) for shape in reference.inline_shapes]
    py_assert3 = [snapshot(shape) for shape in candidate.inline_shapes]
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None
    ref_shapes[1].width = docx.shared.Inches(1.5)
    our_shapes[1].width = Inches(1.5)
    py_assert1 = ref_shapes[1]
    py_assert3 = snapshot(py_assert1)
    py_assert7 = our_shapes[1]
    py_assert9 = snapshot(py_assert7)
    py_assert5 = py_assert3 == py_assert9
    assert not not py_assert5
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = None
    ref_path = tmp_path / 'reference-add-picture.docx'
    our_path = tmp_path / 'candidate-add-picture.docx'
    reference.save(ref_path)
    candidate.save(our_path)
    reopened_reference = docx.Document(str(ref_path))
    reopened_candidate = compat.Document(our_path.read_bytes())
    py_assert0 = [snapshot(shape) for shape in reopened_reference.inline_shapes]
    py_assert3 = [snapshot(shape) for shape in reopened_candidate.inline_shapes]
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None

def test_document_heading_and_page_break_match_python_docx(tmp_path):
    reference = docx.Document()
    candidate = compat.Document()
    expected = []
    ref_default = reference.add_heading('default heading')
    our_default = candidate.add_heading('default heading')
    expected.append((ref_default.text, ref_default.style.name, ref_default.style.style_id))
    py_assert0 = (our_default.text, our_default.style.name, our_default.style.style_id)
    py_assert3 = expected[-1]
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None
    for level in range(10):
        ref_heading = reference.add_heading(f'heading {level}', level)
        our_heading = candidate.add_heading(f'heading {level}', level)
        snapshot = (ref_heading.text, ref_heading.style.name, ref_heading.style.style_id)
        py_assert0 = (our_heading.text, our_heading.style.name, our_heading.style.style_id)
        py_assert2 = py_assert0 == snapshot
        assert not not py_assert2
        py_assert0 = py_assert2 = None
        expected.append(snapshot)
        continue
    for level in (-1, 10):
        ref_error = _capture_exception(lambda level=level: reference.add_heading('invalid', level))
        our_error = _capture_exception(lambda level=level: candidate.add_heading('invalid', level))
        py_assert1 = ref_error == our_error
        assert not not py_assert1
        py_assert1 = None
        continue
    ref_break = reference.add_page_break()
    our_break = candidate.add_page_break()
    py_assert1 = ref_break.text
    py_assert6 = our_break.text
    py_assert3 = py_assert1 == py_assert6
    py_assert8 = ''
    py_assert4 = py_assert6 == py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = ref_break.style
    py_assert3 = py_assert1.name
    py_assert7 = our_break.style
    py_assert9 = py_assert7.name
    py_assert5 = py_assert3 == py_assert9
    assert not not py_assert5
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = None
    py_assert2 = ref_break.runs
    py_assert4 = len(py_assert2)
    py_assert10 = our_break.runs
    py_assert12 = len(py_assert10)
    py_assert6 = py_assert4 == py_assert12
    py_assert14 = 1
    py_assert7 = py_assert12 == py_assert14
    assert not (not py_assert6 or not py_assert7)
    py_assert2 = py_assert4 = py_assert6 = py_assert7 = py_assert10 = py_assert12 = py_assert14 = None
    py_assert0 = ref_break.runs[0]
    py_assert2 = py_assert0.text
    py_assert6 = our_break.runs[0]
    py_assert8 = py_assert6.text
    py_assert4 = py_assert2 == py_assert8
    py_assert10 = ''
    py_assert5 = py_assert8 == py_assert10
    assert not (not py_assert4 or not py_assert5)
    py_assert0 = py_assert2 = py_assert4 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = None

    def page_break_count(path):
        with zipfile.ZipFile(path) as package:
            root = ET.fromstring(package.read('word/document.xml'))
        return sum((1 for element in root.iter() if element.tag.endswith('}br') if element.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type') == 'page'))
    ref_path = tmp_path / 'reference-heading-page-break.docx'
    our_path = tmp_path / 'candidate-heading-page-break.docx'
    reference.save(ref_path)
    candidate.save(our_path)
    py_assert2 = page_break_count(ref_path)
    py_assert8 = page_break_count(our_path)
    py_assert4 = py_assert2 == py_assert8
    py_assert10 = 1
    py_assert5 = py_assert8 == py_assert10
    assert not (not py_assert4 or not py_assert5)
    py_assert2 = py_assert4 = py_assert5 = py_assert8 = py_assert10 = None
    reopened_reference = docx.Document(str(ref_path))
    reopened_candidate = compat.Document(our_path.read_bytes())
    ref_paragraphs = [(p.text, p.style.name, p.style.style_id) for p in reopened_reference.paragraphs]
    our_paragraphs = [(p.text, p.style.name, p.style.style_id) for p in reopened_candidate.paragraphs]
    py_assert1 = ref_paragraphs == our_paragraphs
    assert not not py_assert1
    py_assert1 = None
    py_assert0 = ref_paragraphs[:-1]
    py_assert2 = py_assert0 == expected
    assert not not py_assert2
    py_assert0 = py_assert2 = None
    py_assert0 = ref_paragraphs[-1][0]
    py_assert3 = ''
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None
    py_assert2 = page_break_count(ref_path)
    py_assert8 = page_break_count(our_path)
    py_assert4 = py_assert2 == py_assert8
    py_assert10 = 1
    py_assert5 = py_assert8 == py_assert10
    assert not (not py_assert4 or not py_assert5)
    py_assert2 = py_assert4 = py_assert5 = py_assert8 = py_assert10 = None

def test_document_inner_content_yields_paragraphs_and_tables_in_order(tmp_path):
    reference = docx.Document()
    candidate = compat.Document()
    for document in (reference, candidate):
        document.add_paragraph('before')
        cell = document.add_table(rows=1, cols=1).cell(0, 0)
        cell.text = 'cell'
        nested = cell.add_table(rows=1, cols=2)
        nested.cell(0, 0).text = 'nested table'
        nested.cell(0, 1).text = 'second nested cell'
        document.add_paragraph('after')

    def summarize(items):
        result = []
        for item in items:
            if hasattr(item, 'runs'):
                result.append(('paragraph', item.text))
            else:
                result.append(('table', tuple((tuple((cell.text for cell in row.cells)) for row in item.rows))))
        return result
    py_assert2 = reference.iter_inner_content
    py_assert4 = py_assert2()
    py_assert6 = summarize(py_assert4)
    py_assert11 = candidate.iter_inner_content
    py_assert13 = py_assert11()
    py_assert15 = summarize(py_assert13)
    py_assert8 = py_assert6 == py_assert15
    assert not not py_assert8
    py_assert2 = py_assert4 = py_assert6 = py_assert8 = py_assert11 = py_assert13 = py_assert15 = None
    py_assert2 = reference.tables
    py_assert4 = len(py_assert2)
    py_assert10 = candidate.tables
    py_assert12 = len(py_assert10)
    py_assert6 = py_assert4 == py_assert12
    py_assert14 = 1
    py_assert7 = py_assert12 == py_assert14
    assert not (not py_assert6 or not py_assert7)
    py_assert2 = py_assert4 = py_assert6 = py_assert7 = py_assert10 = py_assert12 = py_assert14 = None
    py_assert0 = [table.cell(0, 0).text for table in reference.tables]
    py_assert4 = [table.cell(0, 0).text for table in candidate.tables]
    py_assert2 = py_assert0 == py_assert4
    py_assert6 = ['cell\n']
    py_assert3 = py_assert4 == py_assert6
    assert not (not py_assert2 or not py_assert3)
    py_assert0 = py_assert2 = py_assert3 = py_assert4 = py_assert6 = None

    def nested_table_snapshot(cell_or_table):
        table = cell_or_table if hasattr(cell_or_table, 'rows') else cell_or_table.tables[0]
        return (len(table.rows), len(table.columns), [int(column.width) if column.width is not None else None for column in table.columns], [[nested_cell.text for nested_cell in row.cells] for row in table.rows])
    py_assert0 = [nested_table_snapshot(table.cell(0, 0)) for table in reference.tables]
    py_assert4 = [nested_table_snapshot(table.cell(0, 0)) for table in candidate.tables]
    py_assert2 = py_assert0 == py_assert4
    py_assert6 = [(1, 2, [int(reference.tables[0].cell(0, 0).width) // 2] * 2, [['nested table', 'second nested cell']])]
    py_assert3 = py_assert4 == py_assert6
    assert not (not py_assert2 or not py_assert3)
    py_assert0 = py_assert2 = py_assert3 = py_assert4 = py_assert6 = None
    ref_cell, our_cell = (reference.tables[0].cell(0, 0), candidate.tables[0].cell(0, 0))
    py_assert0 = [paragraph.text for paragraph in ref_cell.paragraphs]
    py_assert4 = [paragraph.text for paragraph in our_cell.paragraphs]
    py_assert2 = py_assert0 == py_assert4
    py_assert6 = ['cell', '']
    py_assert3 = py_assert4 == py_assert6
    assert not (not py_assert2 or not py_assert3)
    py_assert0 = py_assert2 = py_assert3 = py_assert4 = py_assert6 = None

    def summarize_cell(cell):
        return [
            ('paragraph', item.text)
            if hasattr(item, 'runs')
            else ('table',) + nested_table_snapshot(item)
            for item in cell.iter_inner_content()
        ]
    expected_inner_content = [('paragraph', 'cell'), ('table',) + nested_table_snapshot(ref_cell.tables[0]), ('paragraph', '')]
    py_assert2 = summarize_cell(ref_cell)
    py_assert8 = summarize_cell(our_cell)
    py_assert4 = py_assert2 == py_assert8
    py_assert5 = py_assert8 == expected_inner_content
    assert not (not py_assert4 or not py_assert5)
    py_assert2 = py_assert4 = py_assert5 = py_assert8 = None
    py_assert0 = [p.text for p in reference.paragraphs]
    py_assert4 = [p.text for p in candidate.paragraphs]
    py_assert2 = py_assert0 == py_assert4
    py_assert6 = ['before', 'after']
    py_assert3 = py_assert4 == py_assert6
    assert not (not py_assert2 or not py_assert3)
    py_assert0 = py_assert2 = py_assert3 = py_assert4 = py_assert6 = None
    ref_path = tmp_path / 'reference-document-content.docx'
    our_path = tmp_path / 'candidate-document-content.docx'
    reference.save(ref_path)
    candidate.save(our_path)
    reopened_reference = docx.Document(str(ref_path))
    reopened_candidate = compat.Document(our_path.read_bytes())
    py_assert2 = reopened_reference.iter_inner_content
    py_assert4 = py_assert2()
    py_assert6 = summarize(py_assert4)
    py_assert11 = reopened_candidate.iter_inner_content
    py_assert13 = py_assert11()
    py_assert15 = summarize(py_assert13)
    py_assert8 = py_assert6 == py_assert15
    assert not not py_assert8
    py_assert2 = py_assert4 = py_assert6 = py_assert8 = py_assert11 = py_assert13 = py_assert15 = None
    py_assert2 = reopened_reference.tables
    py_assert4 = len(py_assert2)
    py_assert10 = reopened_candidate.tables
    py_assert12 = len(py_assert10)
    py_assert6 = py_assert4 == py_assert12
    py_assert14 = 1
    py_assert7 = py_assert12 == py_assert14
    assert not (not py_assert6 or not py_assert7)
    py_assert2 = py_assert4 = py_assert6 = py_assert7 = py_assert10 = py_assert12 = py_assert14 = None
    py_assert0 = [table.cell(0, 0).text for table in reopened_reference.tables]
    py_assert4 = [table.cell(0, 0).text for table in reopened_candidate.tables]
    py_assert2 = py_assert0 == py_assert4
    py_assert6 = ['cell\n']
    py_assert3 = py_assert4 == py_assert6
    assert not (not py_assert2 or not py_assert3)
    py_assert0 = py_assert2 = py_assert3 = py_assert4 = py_assert6 = None
    py_assert0 = [table.cell(0, 0).tables[0].cell(0, 0).text for table in reopened_reference.tables]
    py_assert4 = [table.cell(0, 0).tables[0].cell(0, 0).text for table in reopened_candidate.tables]
    py_assert2 = py_assert0 == py_assert4
    py_assert6 = ['nested table']
    py_assert3 = py_assert4 == py_assert6
    assert not (not py_assert2 or not py_assert3)
    py_assert0 = py_assert2 = py_assert3 = py_assert4 = py_assert6 = None
    ref_reopened_cell = reopened_reference.tables[0].cell(0, 0)
    our_reopened_cell = reopened_candidate.tables[0].cell(0, 0)
    py_assert2 = summarize_cell(ref_reopened_cell)
    py_assert7 = summarize_cell(our_reopened_cell)
    py_assert4 = py_assert2 == py_assert7
    assert not not py_assert4
    py_assert2 = py_assert4 = py_assert7 = None

def test_default_table_style_matches_python_docx():
    ours = compat.Document().add_table(rows=1, cols=1)
    theirs = docx.Document().add_table(rows=1, cols=1)
    py_assert1 = ours.style
    py_assert3 = py_assert1.name
    py_assert8 = theirs.style
    py_assert10 = py_assert8.name
    py_assert5 = py_assert3 == py_assert10
    py_assert12 = 'Normal Table'
    py_assert6 = py_assert10 == py_assert12
    assert not (not py_assert5 or not py_assert6)
    py_assert1 = py_assert3 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None

def test_document_add_table_style_arguments_match_python_docx(tmp_path):
    reference = docx.Document()
    candidate = compat.Document()
    ref_styles = (None, 'Table Grid', reference.styles['Table Grid'])
    our_styles = (None, 'Table Grid', candidate.styles['Table Grid'])
    expected = []
    for index, (ref_style, our_style) in enumerate(zip(ref_styles, our_styles)):
        ref_table = reference.add_table(rows=2, cols=2, style=ref_style)
        our_table = candidate.add_table(rows=2, cols=2, style=our_style)
        ref_table.cell(0, 0).text = f'table {index}'
        our_table.cell(0, 0).text = f'table {index}'
        ref_snapshot = (len(ref_table.rows), len(ref_table.columns), ref_table.style.style_id, ref_table.style.name)
        our_snapshot = (len(our_table.rows), len(our_table.columns), our_table.style.style_id, our_table.style.name)
        py_assert1 = ref_snapshot == our_snapshot
        assert not not py_assert1
        py_assert1 = None
        expected.append(ref_snapshot)
        continue
    py_assert2 = len(expected)
    py_assert5 = 3
    py_assert4 = py_assert2 == py_assert5
    assert not not py_assert4
    py_assert2 = py_assert4 = py_assert5 = None
    ref_path = tmp_path / 'reference-add-table.docx'
    our_path = tmp_path / 'candidate-add-table.docx'
    reference.save(ref_path)
    candidate.save(our_path)
    reopened_reference = docx.Document(str(ref_path))
    reopened_candidate = compat.Document(our_path.read_bytes())
    py_assert0 = [(len(table.rows), len(table.columns), table.style.style_id, table.style.name) for table in reopened_reference.tables]
    py_assert4 = [(len(table.rows), len(table.columns), table.style.style_id, table.style.name) for table in reopened_candidate.tables]
    py_assert2 = py_assert0 == py_assert4
    py_assert3 = py_assert4 == expected
    assert not (not py_assert2 or not py_assert3)
    py_assert0 = py_assert2 = py_assert3 = py_assert4 = None

def test_table_collections_and_mutators_match_python_docx(tmp_path):
    reference = docx.Document()
    candidate = compat.Document()
    ref_table = reference.add_table(rows=2, cols=2)
    our_table = candidate.add_table(rows=2, cols=2)
    ref_table.alignment = docx.enum.table.WD_TABLE_ALIGNMENT.CENTER
    our_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    ref_table.style = reference.styles['Table Grid']
    our_table.style = candidate.styles['Table Grid']
    for row_index in range(2):
        for column_index in range(2):
            text = f'{row_index}:{column_index}'
            ref_table.cell(row_index, column_index).text = text
            our_table.cell(row_index, column_index).text = text
    ref_cell, our_cell = (ref_table.cell(0, 0), our_table.cell(0, 0))
    ref_cell.vertical_alignment = docx.enum.table.WD_CELL_VERTICAL_ALIGNMENT.CENTER
    our_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    ref_cell.add_paragraph('second paragraph')
    our_cell.add_paragraph('second paragraph')
    ref_new_row = ref_table.add_row()
    our_new_row = our_table.add_row()
    ref_new_row.cells[0].text = 'added row'
    our_new_row.cells[0].text = 'added row'
    ref_new_row.height = docx.shared.Inches(0.4)
    our_new_row.height = Inches(0.4)
    ref_new_row.height_rule = docx.enum.table.WD_ROW_HEIGHT_RULE.EXACTLY
    our_new_row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
    ref_new_column = ref_table.add_column(docx.shared.Inches(1.5))
    our_new_column = our_table.add_column(Inches(1.5))
    for ref_cell, our_cell in zip(ref_new_column.cells, our_new_column.cells):
        ref_cell.text = our_cell.text = 'added column'

    def table_snapshot(table):
        return {'alignment': table.alignment.name if table.alignment is not None else None, 'style': (table.style.style_id, table.style.name), 'rows': [[cell.text for cell in row.cells] for row in table.rows], 'row_cells': [cell.text for cell in table.row_cells(1)], 'column_cells': [cell.text for cell in table.column_cells(2)], 'row_indexed': [cell.text for cell in table.rows[1].cells], 'rows_sliced': [[cell.text for cell in row.cells] for row in table.rows[:2]], 'column_indexed': [cell.text for cell in table.columns[2].cells], 'first_cell_paragraphs': [p.text for p in table.cell(0, 0).paragraphs], 'first_cell_vertical_alignment': table.cell(0, 0).vertical_alignment.name if table.cell(0, 0).vertical_alignment is not None else None, 'row_metrics': [(int(row.height) if row.height is not None else None, row.height_rule.name if row.height_rule is not None else None) for row in table.rows], 'column_widths': [int(column.width) if column.width is not None else None for column in table.columns], 'dimensions': (len(table.rows), len(table.columns))}
    for ref_target, our_target in ((ref_table.rows.table, our_table.rows.table), (ref_table.columns.table, our_table.columns.table), (ref_table.rows[0].table, our_table.rows[0].table), (ref_table.columns[0].table, our_table.columns[0].table)):
        py_assert1 = []
        py_assert3 = ref_target is ref_table
        py_assert0 = py_assert3
        if py_assert3:
            py_assert9 = our_target is our_table
            py_assert0 = py_assert9
        assert not not py_assert0
        py_assert0 = py_assert1 = py_assert3 = py_assert9 = None
        continue
    py_assert2 = table_snapshot(ref_table)
    py_assert7 = table_snapshot(our_table)
    py_assert4 = py_assert2 == py_assert7
    assert not not py_assert4
    py_assert2 = py_assert4 = py_assert7 = None
    py_assert2 = ref_table.rows
    py_assert4 = len(py_assert2)
    py_assert10 = our_table.rows
    py_assert12 = len(py_assert10)
    py_assert6 = py_assert4 == py_assert12
    py_assert14 = 3
    py_assert7 = py_assert12 == py_assert14
    assert not (not py_assert6 or not py_assert7)
    py_assert2 = py_assert4 = py_assert6 = py_assert7 = py_assert10 = py_assert12 = py_assert14 = None
    py_assert2 = ref_table.columns
    py_assert4 = len(py_assert2)
    py_assert10 = our_table.columns
    py_assert12 = len(py_assert10)
    py_assert6 = py_assert4 == py_assert12
    py_assert14 = 3
    py_assert7 = py_assert12 == py_assert14
    assert not (not py_assert6 or not py_assert7)
    py_assert2 = py_assert4 = py_assert6 = py_assert7 = py_assert10 = py_assert12 = py_assert14 = None
    ref_path = tmp_path / 'reference-table-collections.docx'
    our_path = tmp_path / 'candidate-table-collections.docx'
    reference.save(ref_path)
    candidate.save(our_path)
    reopened_reference = docx.Document(str(ref_path)).tables[0]
    reopened_candidate = compat.Document(our_path.read_bytes()).tables[0]
    py_assert2 = table_snapshot(reopened_reference)
    py_assert7 = table_snapshot(reopened_candidate)
    py_assert4 = py_assert2 == py_assert7
    assert not not py_assert4
    py_assert2 = py_assert4 = py_assert7 = None

def test_document_add_table_uses_last_section_usable_width(tmp_path):
    reference = docx.Document()
    candidate = compat.Document()
    ref_section = reference.add_section(docx.enum.section.WD_SECTION_START.CONTINUOUS)
    our_section = candidate.add_section(WD_SECTION_START.CONTINUOUS)
    ref_section.page_width = docx.shared.Inches(10)
    ref_section.left_margin = docx.shared.Inches(0.75)
    ref_section.right_margin = docx.shared.Inches(0.5)
    our_section.page_width = Inches(10)
    our_section.left_margin = Inches(0.75)
    our_section.right_margin = Inches(0.5)
    ref_table = reference.add_table(rows=1, cols=3)
    our_table = candidate.add_table(rows=1, cols=3)

    def widths(table):
        return ([int(column.width) for column in table.columns], [int(cell.width) for cell in table.rows[0].cells])
    py_assert2 = widths(ref_table)
    py_assert7 = widths(our_table)
    py_assert4 = py_assert2 == py_assert7
    assert not not py_assert4
    py_assert2 = py_assert4 = py_assert7 = None
    ref_path = tmp_path / 'reference-last-section-table.docx'
    our_path = tmp_path / 'candidate-last-section-table.docx'
    reference.save(ref_path)
    candidate.save(our_path)
    py_assert1 = docx.Document(ref_path).tables[0]
    py_assert3 = widths(py_assert1)
    py_assert7 = compat.Document(our_path).tables[0]
    py_assert9 = widths(py_assert7)
    py_assert5 = py_assert3 == py_assert9
    assert not not py_assert5
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = None

def test_table_autofit_and_direction_match_python_docx(tmp_path):
    reference = docx.Document()
    candidate = compat.Document()
    ref_table = reference.add_table(rows=1, cols=1)
    our_table = candidate.add_table(rows=1, cols=1)
    py_assert1 = ref_table.autofit
    py_assert6 = our_table.autofit
    py_assert3 = py_assert1 is py_assert6
    py_assert8 = True
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    ref_table.autofit = False
    our_table.autofit = False
    py_assert1 = ref_table.autofit
    py_assert6 = our_table.autofit
    py_assert3 = py_assert1 is py_assert6
    py_assert8 = False
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    ref_table.autofit = True
    our_table.autofit = True
    py_assert1 = ref_table.autofit
    py_assert6 = our_table.autofit
    py_assert3 = py_assert1 is py_assert6
    py_assert8 = True
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = ref_table.table_direction
    py_assert6 = our_table.table_direction
    py_assert3 = py_assert1 is py_assert6
    py_assert8 = None
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    for ref_direction, our_direction in ((docx.enum.table.WD_TABLE_DIRECTION.LTR, WD_TABLE_DIRECTION.LTR), (docx.enum.table.WD_TABLE_DIRECTION.RTL, WD_TABLE_DIRECTION.RTL)):
        ref_table.table_direction = ref_direction
        our_table.table_direction = our_direction
        expected_direction = bool(ref_direction)
        py_assert1 = ref_table.table_direction
        py_assert3 = py_assert1 is expected_direction
        assert not not py_assert3
        py_assert1 = py_assert3 = None
        py_assert1 = our_table.table_direction
        py_assert3 = py_assert1 is expected_direction
        assert not not py_assert3
        py_assert1 = py_assert3 = None
        continue
    ref_table.table_direction = None
    our_table.table_direction = None
    py_assert1 = ref_table.table_direction
    py_assert6 = our_table.table_direction
    py_assert3 = py_assert1 is py_assert6
    py_assert8 = None
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    ref_path = tmp_path / 'reference-table-layout.docx'
    our_path = tmp_path / 'candidate-table-layout.docx'
    ref_table.autofit = False
    our_table.autofit = False
    ref_table.table_direction = docx.enum.table.WD_TABLE_DIRECTION.RTL
    our_table.table_direction = WD_TABLE_DIRECTION.RTL
    reference.save(ref_path)
    candidate.save(our_path)
    reopened_reference = docx.Document(str(ref_path)).tables[0]
    reopened_candidate = compat.Document(our_path.read_bytes()).tables[0]
    py_assert1 = reopened_reference.autofit
    py_assert6 = reopened_candidate.autofit
    py_assert3 = py_assert1 is py_assert6
    py_assert8 = False
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = reopened_reference.table_direction
    py_assert4 = True
    py_assert3 = py_assert1 is py_assert4
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert4 = None
    py_assert1 = reopened_candidate.table_direction
    py_assert4 = True
    py_assert3 = py_assert1 is py_assert4
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert4 = None

def test_core_properties_fields_defaults_setters_and_round_trip_match(tmp_path):
    reference = docx.Document()
    candidate = compat.Document()
    ref_props = reference.core_properties
    our_props = candidate.core_properties
    py_assert2 = _core_property_values(ref_props)
    py_assert7 = _core_property_values(our_props)
    py_assert4 = py_assert2 == py_assert7
    assert not not py_assert4
    py_assert2 = py_assert4 = py_assert7 = None
    for name in CORE_PROPERTY_FIELDS:
        py_assert4 = getattr(ref_props, name)
        py_assert6 = type(py_assert4)
        py_assert13 = getattr(our_props, name)
        py_assert15 = type(py_assert13)
        py_assert8 = py_assert6 is py_assert15
        assert not not py_assert8
        py_assert4 = py_assert6 = py_assert8 = py_assert13 = py_assert15 = None
        continue
    values = {'title': 'Parity title', 'author': 'Vikas Maurya', 'subject': 'Core property test', 'comments': 'Round-trip metadata', 'category': 'Engineering', 'keywords': 'docx, parity', 'last_modified_by': 'Vikas', 'revision': 7, 'content_status': 'Draft', 'identifier': 'urn:docviews:parity', 'language': 'en-IN', 'version': '3', 'created': datetime(2024, 5, 6, 7, 8, 9, tzinfo=timezone(timedelta(hours=5, minutes=30))), 'modified': datetime(2024, 5, 7, 8, 9, 10), 'last_printed': datetime(2024, 5, 8, 9, 10, 11)}
    for name, value in values.items():
        setattr(ref_props, name, value)
        setattr(our_props, name, value)
    py_assert2 = _core_property_values(ref_props)
    py_assert7 = _core_property_values(our_props)
    py_assert4 = py_assert2 == py_assert7
    assert not not py_assert4
    py_assert2 = py_assert4 = py_assert7 = None
    for name in CORE_PROPERTY_FIELDS:
        py_assert4 = getattr(ref_props, name)
        py_assert6 = type(py_assert4)
        py_assert13 = getattr(our_props, name)
        py_assert15 = type(py_assert13)
        py_assert8 = py_assert6 is py_assert15
        assert not not py_assert8
        py_assert4 = py_assert6 = py_assert8 = py_assert13 = py_assert15 = None
        continue
    ref_path = tmp_path / 'reference-core-properties.docx'
    our_path = tmp_path / 'candidate-core-properties.docx'
    reference.save(ref_path)
    candidate.save(our_path)
    reopened_reference = docx.Document(str(ref_path)).core_properties
    reopened_candidate = compat.Document(our_path.read_bytes()).core_properties
    py_assert2 = _core_property_values(reopened_reference)
    py_assert7 = _core_property_values(reopened_candidate)
    py_assert4 = py_assert2 == py_assert7
    assert not not py_assert4
    py_assert2 = py_assert4 = py_assert7 = None
    for name in CORE_PROPERTY_FIELDS:
        py_assert4 = getattr(reopened_reference, name)
        py_assert6 = type(py_assert4)
        py_assert13 = getattr(reopened_candidate, name)
        py_assert15 = type(py_assert13)
        py_assert8 = py_assert6 is py_assert15
        assert not not py_assert8
        py_assert4 = py_assert6 = py_assert8 = py_assert13 = py_assert15 = None
        continue

def test_core_properties_parse_w3cdtf_offsets_and_partial_dates_like_upstream():
    template = io.BytesIO()
    docx.Document().save(template)
    original = template.getvalue()
    output = io.BytesIO()
    old_date = b'2013-12-23T23:15:00Z'
    with zipfile.ZipFile(io.BytesIO(original), 'r') as source:
        with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as target:
            for member in source.infolist():
                data = source.read(member.filename)
                if member.filename == 'docProps/core.xml':
                    data = data.replace(old_date, b'2024-02-03T10:15:00+05:30', 1)
                    data = data.replace(old_date, b'2019-04', 1)
                target.writestr(member, data)
    package = output.getvalue()
    reference = docx.Document(io.BytesIO(package)).core_properties
    candidate = compat.Document(package).core_properties
    for name in ('created', 'modified'):
        py_assert3 = getattr(reference, name)
        py_assert9 = getattr(candidate, name)
        py_assert5 = py_assert3 == py_assert9
        assert not not py_assert5
        py_assert3 = py_assert5 = py_assert9 = None
        py_assert4 = getattr(reference, name)
        py_assert6 = type(py_assert4)
        py_assert13 = getattr(candidate, name)
        py_assert15 = type(py_assert13)
        py_assert8 = py_assert6 is py_assert15
        assert not not py_assert8
        py_assert4 = py_assert6 = py_assert8 = py_assert13 = py_assert15 = None
        continue
    py_assert1 = reference.created
    py_assert5 = 2024
    py_assert7 = 2
    py_assert9 = 3
    py_assert11 = 4
    py_assert13 = 45
    py_assert16 = timezone.utc
    py_assert18 = datetime(py_assert5, py_assert7, py_assert9, py_assert11, py_assert13, tzinfo=py_assert16)
    py_assert3 = py_assert1 == py_assert18
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = py_assert11 = py_assert13 = py_assert16 = py_assert18 = None
    py_assert1 = reference.modified
    py_assert5 = 2019
    py_assert7 = 4
    py_assert9 = 1
    py_assert11 = 0
    py_assert13 = 0
    py_assert16 = timezone.utc
    py_assert18 = datetime(py_assert5, py_assert7, py_assert9, py_assert11, py_assert13, tzinfo=py_assert16)
    py_assert3 = py_assert1 == py_assert18
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = py_assert11 = py_assert13 = py_assert16 = py_assert18 = None

def test_core_properties_setter_validation_matches_python_docx():
    reference = docx.Document().core_properties
    candidate = compat.Document().core_properties
    for name in CORE_PROPERTY_STRING_FIELDS:
        for value in (None, 123, 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'):
            ref_error = _capture_exception(lambda name=name, value=value: setattr(reference, name, value))
            our_error = _capture_exception(lambda name=name, value=value: setattr(candidate, name, value))
            py_assert1 = ref_error == our_error
            assert not not py_assert1
            py_assert1 = None
            if not len(str(value)) <= 255:
                continue
            py_assert3 = getattr(reference, name)
            py_assert9 = getattr(candidate, name)
            py_assert5 = py_assert3 == py_assert9
            assert not not py_assert5
            py_assert3 = py_assert5 = py_assert9 = None
            continue
        continue
    for name in CORE_PROPERTY_DATE_FIELDS:
        ref_error = _capture_exception(lambda name=name: setattr(reference, name, None))
        our_error = _capture_exception(lambda name=name: setattr(candidate, name, None))
        py_assert1 = ref_error == our_error
        assert not not py_assert1
        py_assert1 = None
        continue
    for value in (0, -1, 1.5, '2'):
        ref_error = _capture_exception(lambda value=value: setattr(reference, 'revision', value))
        our_error = _capture_exception(lambda value=value: setattr(candidate, 'revision', value))
        py_assert1 = ref_error == our_error
        assert not not py_assert1
        py_assert1 = None
        continue

def test_table_cell_row_and_column_dynamic_writability_match_and_round_trip(tmp_path):
    reference = docx.Document()
    candidate = compat.Document()
    ref_table = reference.add_table(rows=2, cols=2)
    our_table = candidate.add_table(rows=2, cols=2)

    def observed(table):
        cell = table.cell(0, 0)
        row = table.rows[0]
        column = table.columns[1]
        return {'text': cell.text, 'grid_span': cell.grid_span, 'cell_width': int(cell.width) if cell.width is not None else None, 'column_width': int(column.width) if column.width is not None else None, 'row_height': int(row.height) if row.height is not None else None, 'grid_cols_before': row.grid_cols_before, 'grid_cols_after': row.grid_cols_after, 'autofit': table.autofit, 'table_direction': table.table_direction, 'merged_grid_span': table.cell(1, 0).grid_span}
    ref_cell, our_cell = (ref_table.cell(0, 0), our_table.cell(0, 0))
    py_assert1 = ref_cell.grid_span
    py_assert6 = our_cell.grid_span
    py_assert3 = py_assert1 == py_assert6
    py_assert8 = 1
    py_assert4 = py_assert6 == py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    ref_grid_error = _capture_exception(lambda: setattr(ref_cell, 'grid_span', 3))
    our_grid_error = _capture_exception(lambda: setattr(our_cell, 'grid_span', 3))
    py_assert0 = ref_grid_error[0]
    py_assert4 = our_grid_error[0]
    py_assert2 = py_assert0 is py_assert4
    py_assert3 = py_assert4 is AttributeError
    assert not (not py_assert2 or not py_assert3)
    py_assert0 = py_assert2 = py_assert3 = py_assert4 = None
    for name in ('grid_cols_before', 'grid_cols_after'):
        ref_row, our_row = (ref_table.rows[0], our_table.rows[0])
        ref_error = _capture_exception(lambda name=name: setattr(ref_row, name, 1))
        our_error = _capture_exception(lambda name=name: setattr(our_row, name, 1))
        py_assert0 = ref_error[0]
        py_assert4 = our_error[0]
        py_assert2 = py_assert0 is py_assert4
        py_assert3 = py_assert4 is AttributeError
        assert not (not py_assert2 or not py_assert3)
        py_assert0 = py_assert2 = py_assert3 = py_assert4 = None
        py_assert3 = getattr(ref_row, name)
        py_assert10 = getattr(our_row, name)
        py_assert5 = py_assert3 == py_assert10
        py_assert12 = 0
        py_assert6 = py_assert10 == py_assert12
        assert not (not py_assert5 or not py_assert6)
        py_assert3 = py_assert5 = py_assert6 = py_assert10 = py_assert12 = None
        continue
    ref_cell.text = 'Cell text'
    our_cell.text = 'Cell text'
    ref_cell.width = Inches(1.25)
    our_cell.width = Inches(1.25)
    ref_table.columns[1].width = Inches(2.5)
    our_table.columns[1].width = Inches(2.5)
    ref_table.rows[0].height = Inches(0.4)
    our_table.rows[0].height = Inches(0.4)
    ref_table.autofit = our_table.autofit = False
    ref_table.table_direction = docx.enum.table.WD_TABLE_DIRECTION.RTL
    our_table.table_direction = WD_TABLE_DIRECTION.RTL
    ref_table.cell(1, 0).merge(ref_table.cell(1, 1))
    our_table.cell(1, 0).merge(our_table.cell(1, 1))
    py_assert2 = observed(ref_table)
    py_assert7 = observed(our_table)
    py_assert4 = py_assert2 == py_assert7
    assert not not py_assert4
    py_assert2 = py_assert4 = py_assert7 = None
    ref_path = tmp_path / 'reference-table-writability.docx'
    our_path = tmp_path / 'candidate-table-writability.docx'
    reference.save(ref_path)
    candidate.save(our_path)
    reopened_reference = docx.Document(str(ref_path)).tables[0]
    reopened_candidate = compat.Document(our_path.read_bytes()).tables[0]
    py_assert2 = observed(reopened_reference)
    py_assert7 = observed(reopened_candidate)
    py_assert4 = py_assert2 == py_assert7
    assert not not py_assert4
    py_assert2 = py_assert4 = py_assert7 = None

def test_table_row_grid_before_and_after_match_irregular_docx_rows():
    document = docx.Document()
    document.add_table(rows=1, cols=4)
    source = io.BytesIO()
    document.save(source)
    package = _patch_first_table_row_grid_edges(source.getvalue(), before=1, after=1)
    reference = docx.Document(io.BytesIO(package)).tables[0].rows[0]
    candidate = compat.Document(package).tables[0].rows[0]
    py_assert1 = reference.grid_cols_before
    py_assert6 = candidate.grid_cols_before
    py_assert3 = py_assert1 == py_assert6
    py_assert8 = 1
    py_assert4 = py_assert6 == py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = reference.grid_cols_after
    py_assert6 = candidate.grid_cols_after
    py_assert3 = py_assert1 == py_assert6
    py_assert8 = 1
    py_assert4 = py_assert6 == py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert2 = reference.cells
    py_assert4 = len(py_assert2)
    py_assert10 = candidate.cells
    py_assert12 = len(py_assert10)
    py_assert6 = py_assert4 == py_assert12
    py_assert14 = 2
    py_assert7 = py_assert12 == py_assert14
    assert not (not py_assert6 or not py_assert7)
    py_assert2 = py_assert4 = py_assert6 = py_assert7 = py_assert10 = py_assert12 = py_assert14 = None
    py_assert0 = [cell.text for cell in reference.cells]
    py_assert3 = [cell.text for cell in candidate.cells]
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None

def test_style_base_and_next_paragraph_relationships_match_and_round_trip(tmp_path):
    reference = docx.Document()
    candidate = compat.Document()
    ref_base = reference.styles.add_style('Parity Base', docx.enum.style.WD_STYLE_TYPE.PARAGRAPH)
    ref_derived = reference.styles.add_style('Parity Derived', docx.enum.style.WD_STYLE_TYPE.PARAGRAPH)
    our_base = candidate.styles.add_style('Parity Base', WD_STYLE_TYPE.PARAGRAPH)
    our_derived = candidate.styles.add_style('Parity Derived', WD_STYLE_TYPE.PARAGRAPH)
    py_assert1 = ref_derived.base_style
    py_assert6 = our_derived.base_style
    py_assert3 = py_assert1 is py_assert6
    py_assert8 = None
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = ref_derived.next_paragraph_style
    py_assert3 = py_assert1 is ref_derived
    assert not not py_assert3
    py_assert1 = py_assert3 = None
    py_assert1 = our_derived.next_paragraph_style
    py_assert3 = py_assert1 is our_derived
    assert not not py_assert3
    py_assert1 = py_assert3 = None
    ref_derived.base_style = ref_base
    ref_derived.next_paragraph_style = ref_base
    our_derived.base_style = our_base
    our_derived.next_paragraph_style = our_base
    py_assert1 = ref_derived.base_style
    py_assert3 = py_assert1.name
    py_assert8 = our_derived.base_style
    py_assert10 = py_assert8.name
    py_assert5 = py_assert3 == py_assert10
    py_assert12 = 'Parity Base'
    py_assert6 = py_assert10 == py_assert12
    assert not (not py_assert5 or not py_assert6)
    py_assert1 = py_assert3 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
    py_assert1 = ref_derived.next_paragraph_style
    py_assert3 = py_assert1.name
    py_assert8 = our_derived.next_paragraph_style
    py_assert10 = py_assert8.name
    py_assert5 = py_assert3 == py_assert10
    py_assert12 = 'Parity Base'
    py_assert6 = py_assert10 == py_assert12
    assert not (not py_assert5 or not py_assert6)
    py_assert1 = py_assert3 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
    linked_path = tmp_path / 'style-links.docx'
    reference.save(linked_path)
    reopened_reference = docx.Document(str(linked_path)).styles['Parity Derived']
    candidate.save(linked_path.with_name('candidate-style-links.docx'))
    reopened_candidate = compat.Document(linked_path.with_name('candidate-style-links.docx')).styles['Parity Derived']
    py_assert1 = reopened_reference.base_style
    py_assert3 = py_assert1.name
    py_assert8 = reopened_candidate.base_style
    py_assert10 = py_assert8.name
    py_assert5 = py_assert3 == py_assert10
    py_assert12 = 'Parity Base'
    py_assert6 = py_assert10 == py_assert12
    assert not (not py_assert5 or not py_assert6)
    py_assert1 = py_assert3 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
    py_assert1 = reopened_reference.next_paragraph_style
    py_assert3 = py_assert1.name
    py_assert8 = reopened_candidate.next_paragraph_style
    py_assert10 = py_assert8.name
    py_assert5 = py_assert3 == py_assert10
    py_assert12 = 'Parity Base'
    py_assert6 = py_assert10 == py_assert12
    assert not (not py_assert5 or not py_assert6)
    py_assert1 = py_assert3 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
    ref_derived.base_style = None
    ref_derived.next_paragraph_style = None
    our_derived.base_style = None
    our_derived.next_paragraph_style = None
    py_assert1 = ref_derived.base_style
    py_assert6 = our_derived.base_style
    py_assert3 = py_assert1 is py_assert6
    py_assert8 = None
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = ref_derived.next_paragraph_style
    py_assert3 = py_assert1 is ref_derived
    assert not not py_assert3
    py_assert1 = py_assert3 = None
    py_assert1 = our_derived.next_paragraph_style
    py_assert3 = py_assert1 is our_derived
    assert not not py_assert3
    py_assert1 = py_assert3 = None
    ref_derived.next_paragraph_style = ref_derived
    our_derived.next_paragraph_style = our_derived
    py_assert1 = ref_derived.next_paragraph_style
    py_assert3 = py_assert1 is ref_derived
    assert not not py_assert3
    py_assert1 = py_assert3 = None
    py_assert1 = our_derived.next_paragraph_style
    py_assert3 = py_assert1 is our_derived
    assert not not py_assert3
    py_assert1 = py_assert3 = None
    unset_path = tmp_path / 'style-links-unset.docx'
    reference.save(unset_path)
    candidate.save(unset_path.with_name('candidate-style-links-unset.docx'))
    reopened_reference = docx.Document(str(unset_path)).styles['Parity Derived']
    reopened_candidate = compat.Document(unset_path.with_name('candidate-style-links-unset.docx')).styles['Parity Derived']
    py_assert1 = reopened_reference.base_style
    py_assert6 = reopened_candidate.base_style
    py_assert3 = py_assert1 is py_assert6
    py_assert8 = None
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = reopened_reference.next_paragraph_style
    py_assert3 = py_assert1 is reopened_reference
    assert not not py_assert3
    py_assert1 = py_assert3 = None
    py_assert1 = reopened_candidate.next_paragraph_style
    py_assert3 = py_assert1 is reopened_candidate
    assert not not py_assert3
    py_assert1 = py_assert3 = None
@pytest.mark.parametrize(('kind', 'reference_style_type', 'candidate_style_type'), [('Character', docx.enum.style.WD_STYLE_TYPE.CHARACTER, WD_STYLE_TYPE.CHARACTER), ('Table', docx.enum.style.WD_STYLE_TYPE.TABLE, WD_STYLE_TYPE.TABLE)])

def test_character_and_table_style_relationships_match(kind, reference_style_type, candidate_style_type, tmp_path):
    reference = docx.Document()
    candidate = compat.Document()
    ref_base = reference.styles.add_style(f'{kind} Base', reference_style_type)
    ref_derived = reference.styles.add_style(f'{kind} Derived', reference_style_type)
    our_base = candidate.styles.add_style(f'{kind} Base', candidate_style_type)
    our_derived = candidate.styles.add_style(f'{kind} Derived', candidate_style_type)
    py_assert1 = ref_derived.base_style
    py_assert6 = our_derived.base_style
    py_assert3 = py_assert1 is py_assert6
    py_assert8 = None
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    ref_derived.base_style = ref_base
    our_derived.base_style = our_base
    py_assert1 = ref_derived.base_style
    py_assert3 = py_assert1.name
    py_assert8 = our_derived.base_style
    py_assert10 = py_assert8.name
    py_assert5 = py_assert3 == py_assert10
    py_assert12 = f'{kind} Base'
    py_assert6 = py_assert10 == py_assert12
    assert not (not py_assert5 or not py_assert6)
    py_assert1 = py_assert3 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
    base_link_ref = tmp_path / f'reference-{kind.lower()}-base-link.docx'
    base_link_ours = tmp_path / f'candidate-{kind.lower()}-base-link.docx'
    reference.save(base_link_ref)
    candidate.save(base_link_ours)
    py_assert0 = docx.Document(str(base_link_ref)).styles[f'{kind} Derived']
    py_assert2 = py_assert0.base_style
    py_assert4 = py_assert2.name
    py_assert7 = f'{kind} Base'
    py_assert6 = py_assert4 == py_assert7
    assert not not py_assert6
    py_assert0 = py_assert2 = py_assert4 = py_assert6 = py_assert7 = None
    py_assert0 = compat.Document(base_link_ours.read_bytes()).styles[f'{kind} Derived']
    py_assert2 = py_assert0.base_style
    py_assert4 = py_assert2.name
    py_assert7 = f'{kind} Base'
    py_assert6 = py_assert4 == py_assert7
    assert not not py_assert6
    py_assert0 = py_assert2 = py_assert4 = py_assert6 = py_assert7 = None
    ref_derived.base_style = None
    our_derived.base_style = None
    py_assert1 = ref_derived.base_style
    py_assert6 = our_derived.base_style
    py_assert3 = py_assert1 is py_assert6
    py_assert8 = None
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    if kind == 'Table':
        ref_next = reference.styles.add_style('Table Next Paragraph', docx.enum.style.WD_STYLE_TYPE.PARAGRAPH)
        our_next = candidate.styles.add_style('Table Next Paragraph', WD_STYLE_TYPE.PARAGRAPH)
        py_assert1 = ref_derived.next_paragraph_style
        py_assert3 = py_assert1 is ref_derived
        assert not not py_assert3
        py_assert1 = py_assert3 = None
        py_assert1 = our_derived.next_paragraph_style
        py_assert3 = py_assert1 is our_derived
        assert not not py_assert3
        py_assert1 = py_assert3 = None
        ref_derived.next_paragraph_style = ref_next
        our_derived.next_paragraph_style = our_next
        py_assert1 = ref_derived.next_paragraph_style
        py_assert3 = py_assert1.name
        py_assert8 = our_derived.next_paragraph_style
        py_assert10 = py_assert8.name
        py_assert5 = py_assert3 == py_assert10
        py_assert12 = 'Table Next Paragraph'
        py_assert6 = py_assert10 == py_assert12
        assert not (not py_assert5 or not py_assert6)
        py_assert1 = py_assert3 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
        next_link_ref = tmp_path / 'reference-table-next-link.docx'
        next_link_ours = tmp_path / 'candidate-table-next-link.docx'
        reference.save(next_link_ref)
        candidate.save(next_link_ours)
        py_assert0 = docx.Document(str(next_link_ref)).styles['Table Derived']
        py_assert2 = py_assert0.next_paragraph_style
        py_assert4 = py_assert2.name
        py_assert7 = 'Table Next Paragraph'
        py_assert6 = py_assert4 == py_assert7
        assert not not py_assert6
        py_assert0 = py_assert2 = py_assert4 = py_assert6 = py_assert7 = None
        py_assert0 = compat.Document(next_link_ours.read_bytes()).styles['Table Derived']
        py_assert2 = py_assert0.next_paragraph_style
        py_assert4 = py_assert2.name
        py_assert7 = 'Table Next Paragraph'
        py_assert6 = py_assert4 == py_assert7
        assert not not py_assert6
        py_assert0 = py_assert2 = py_assert4 = py_assert6 = py_assert7 = None
        ref_derived.next_paragraph_style = None
        our_derived.next_paragraph_style = None
        py_assert1 = ref_derived.next_paragraph_style
        py_assert3 = py_assert1 is ref_derived
        assert not not py_assert3
        py_assert1 = py_assert3 = None
        py_assert1 = our_derived.next_paragraph_style
        py_assert3 = py_assert1 is our_derived
        assert not not py_assert3
        py_assert1 = py_assert3 = None
    ref_path = tmp_path / f'reference-{kind.lower()}-style.docx'
    our_path = tmp_path / f'candidate-{kind.lower()}-style.docx'
    reference.save(ref_path)
    candidate.save(our_path)
    reopened_reference = docx.Document(str(ref_path)).styles[f'{kind} Derived']
    reopened_candidate = compat.Document(our_path.read_bytes()).styles[f'{kind} Derived']
    py_assert1 = reopened_reference.base_style
    py_assert6 = reopened_candidate.base_style
    py_assert3 = py_assert1 is py_assert6
    py_assert8 = None
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    if kind == 'Table':
        py_assert1 = reopened_reference.next_paragraph_style
        py_assert3 = py_assert1 is reopened_reference
        assert not not py_assert3
        py_assert1 = py_assert3 = None
        py_assert1 = reopened_candidate.next_paragraph_style
        py_assert3 = py_assert1 is reopened_candidate
        assert not not py_assert3
        py_assert1 = py_assert3 = None

def test_documented_style_proxy_types_members_and_defaults_match(tmp_path):
    from docviews.compat.styles.style import BaseStyle, CharacterStyle, ParagraphStyle, _NumberingStyle, _TableStyle
    reference = docx.Document()
    candidate = compat.Document()
    style_types = (('Parity Character', docx.enum.style.WD_STYLE_TYPE.CHARACTER, WD_STYLE_TYPE.CHARACTER, 'CharacterStyle'), ('Parity Paragraph', docx.enum.style.WD_STYLE_TYPE.PARAGRAPH, WD_STYLE_TYPE.PARAGRAPH, 'ParagraphStyle'), ('Parity Table', docx.enum.style.WD_STYLE_TYPE.TABLE, WD_STYLE_TYPE.TABLE, '_TableStyle'), ('Parity List', docx.enum.style.WD_STYLE_TYPE.LIST, WD_STYLE_TYPE.LIST, '_NumberingStyle'))
    pairs = [(reference.styles.add_style(name, ref_type), candidate.styles.add_style(name, our_type), class_name) for name, ref_type, our_type, class_name in style_types]
    pairs.extend(((reference.styles['Normal'], candidate.styles['Normal'], 'ParagraphStyle'), (reference.styles['Default Paragraph Font'], candidate.styles['Default Paragraph Font'], 'CharacterStyle'), (reference.styles['Table Grid'], candidate.styles['Table Grid'], '_TableStyle')))
    common_members = ('builtin', 'delete', 'hidden', 'locked', 'name', 'priority', 'quick_style', 'type', 'unhide_when_used')
    for ref_style, our_style, expected_class in pairs:
        py_assert2 = type(ref_style)
        py_assert4 = py_assert2.__name__
        py_assert10 = type(our_style)
        py_assert12 = py_assert10.__name__
        py_assert6 = py_assert4 == py_assert12
        py_assert7 = py_assert12 == expected_class
        assert not (not py_assert6 or not py_assert7)
        py_assert2 = py_assert4 = py_assert6 = py_assert7 = py_assert10 = py_assert12 = None
        py_assert3 = isinstance(our_style, BaseStyle)
        assert not not py_assert3
        py_assert3 = None
        for member in common_members:
            py_assert3 = hasattr(ref_style, member)
            py_assert9 = hasattr(our_style, member)
            py_assert5 = py_assert3 is py_assert9
            assert not not py_assert5
            py_assert3 = py_assert5 = py_assert9 = None
            continue
        py_assert1 = ref_style.name
        py_assert5 = our_style.name
        py_assert3 = py_assert1 == py_assert5
        assert not not py_assert3
        py_assert1 = py_assert3 = py_assert5 = None
        py_assert1 = ref_style.style_id
        py_assert5 = our_style.style_id
        py_assert3 = py_assert1 == py_assert5
        assert not not py_assert3
        py_assert1 = py_assert3 = py_assert5 = None
        py_assert1 = ref_style.type
        py_assert3 = py_assert1.name
        py_assert7 = our_style.type
        py_assert9 = py_assert7.name
        py_assert5 = py_assert3 == py_assert9
        assert not not py_assert5
        py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = None
        if expected_class == 'CharacterStyle':
            py_assert3 = isinstance(our_style, CharacterStyle)
            assert not not py_assert3
            py_assert3 = None
            py_assert2 = 'font'
            py_assert4 = hasattr(ref_style, py_assert2)
            py_assert10 = 'font'
            py_assert12 = hasattr(our_style, py_assert10)
            py_assert6 = py_assert4 is py_assert12
            py_assert14 = True
            py_assert7 = py_assert12 is py_assert14
            assert not (not py_assert6 or not py_assert7)
            py_assert2 = py_assert4 = py_assert6 = py_assert7 = py_assert10 = py_assert12 = py_assert14 = None
            py_assert2 = 'base_style'
            py_assert4 = hasattr(ref_style, py_assert2)
            py_assert10 = 'base_style'
            py_assert12 = hasattr(our_style, py_assert10)
            py_assert6 = py_assert4 is py_assert12
            py_assert14 = True
            py_assert7 = py_assert12 is py_assert14
            assert not (not py_assert6 or not py_assert7)
            py_assert2 = py_assert4 = py_assert6 = py_assert7 = py_assert10 = py_assert12 = py_assert14 = None
            py_assert2 = 'paragraph_format'
            py_assert4 = hasattr(ref_style, py_assert2)
            py_assert10 = 'paragraph_format'
            py_assert12 = hasattr(our_style, py_assert10)
            py_assert6 = py_assert4 is py_assert12
            py_assert14 = False
            py_assert7 = py_assert12 is py_assert14
            assert not (not py_assert6 or not py_assert7)
            py_assert2 = py_assert4 = py_assert6 = py_assert7 = py_assert10 = py_assert12 = py_assert14 = None
            continue
        if expected_class == 'ParagraphStyle':
            py_assert3 = isinstance(our_style, ParagraphStyle)
            assert not not py_assert3
            py_assert3 = None
            py_assert2 = 'paragraph_format'
            py_assert4 = hasattr(ref_style, py_assert2)
            py_assert10 = 'paragraph_format'
            py_assert12 = hasattr(our_style, py_assert10)
            py_assert6 = py_assert4 is py_assert12
            py_assert14 = True
            py_assert7 = py_assert12 is py_assert14
            assert not (not py_assert6 or not py_assert7)
            py_assert2 = py_assert4 = py_assert6 = py_assert7 = py_assert10 = py_assert12 = py_assert14 = None
            py_assert2 = 'next_paragraph_style'
            py_assert4 = hasattr(ref_style, py_assert2)
            py_assert10 = 'next_paragraph_style'
            py_assert12 = hasattr(our_style, py_assert10)
            py_assert6 = py_assert4 is py_assert12
            py_assert14 = True
            py_assert7 = py_assert12 is py_assert14
            assert not (not py_assert6 or not py_assert7)
            py_assert2 = py_assert4 = py_assert6 = py_assert7 = py_assert10 = py_assert12 = py_assert14 = None
            continue
        if expected_class == '_TableStyle':
            py_assert3 = isinstance(our_style, _TableStyle)
            assert not not py_assert3
            py_assert3 = None
            py_assert2 = 'font'
            py_assert4 = hasattr(ref_style, py_assert2)
            py_assert10 = 'font'
            py_assert12 = hasattr(our_style, py_assert10)
            py_assert6 = py_assert4 is py_assert12
            py_assert14 = True
            py_assert7 = py_assert12 is py_assert14
            assert not (not py_assert6 or not py_assert7)
            py_assert2 = py_assert4 = py_assert6 = py_assert7 = py_assert10 = py_assert12 = py_assert14 = None
            py_assert2 = 'paragraph_format'
            py_assert4 = hasattr(ref_style, py_assert2)
            py_assert10 = 'paragraph_format'
            py_assert12 = hasattr(our_style, py_assert10)
            py_assert6 = py_assert4 is py_assert12
            py_assert14 = True
            py_assert7 = py_assert12 is py_assert14
            assert not (not py_assert6 or not py_assert7)
            py_assert2 = py_assert4 = py_assert6 = py_assert7 = py_assert10 = py_assert12 = py_assert14 = None
            continue
        py_assert3 = isinstance(our_style, _NumberingStyle)
        assert not not py_assert3
        py_assert3 = None
        py_assert2 = 'base_style'
        py_assert4 = hasattr(ref_style, py_assert2)
        py_assert10 = 'base_style'
        py_assert12 = hasattr(our_style, py_assert10)
        py_assert6 = py_assert4 is py_assert12
        py_assert14 = False
        py_assert7 = py_assert12 is py_assert14
        assert not (not py_assert6 or not py_assert7)
        py_assert2 = py_assert4 = py_assert6 = py_assert7 = py_assert10 = py_assert12 = py_assert14 = None
        py_assert2 = 'font'
        py_assert4 = hasattr(ref_style, py_assert2)
        py_assert10 = 'font'
        py_assert12 = hasattr(our_style, py_assert10)
        py_assert6 = py_assert4 is py_assert12
        py_assert14 = False
        py_assert7 = py_assert12 is py_assert14
        assert not (not py_assert6 or not py_assert7)
        py_assert2 = py_assert4 = py_assert6 = py_assert7 = py_assert10 = py_assert12 = py_assert14 = None
        py_assert2 = 'paragraph_format'
        py_assert4 = hasattr(ref_style, py_assert2)
        py_assert10 = 'paragraph_format'
        py_assert12 = hasattr(our_style, py_assert10)
        py_assert6 = py_assert4 is py_assert12
        py_assert14 = False
        py_assert7 = py_assert12 is py_assert14
        assert not (not py_assert6 or not py_assert7)
        py_assert2 = py_assert4 = py_assert6 = py_assert7 = py_assert10 = py_assert12 = py_assert14 = None
        py_assert2 = 'next_paragraph_style'
        py_assert4 = hasattr(ref_style, py_assert2)
        py_assert10 = 'next_paragraph_style'
        py_assert12 = hasattr(our_style, py_assert10)
        py_assert6 = py_assert4 is py_assert12
        py_assert14 = False
        py_assert7 = py_assert12 is py_assert14
        assert not (not py_assert6 or not py_assert7)
        py_assert2 = py_assert4 = py_assert6 = py_assert7 = py_assert10 = py_assert12 = py_assert14 = None
        continue
    ref_character, our_character, _ = pairs[0]
    ref_paragraph, our_paragraph, _ = pairs[1]
    ref_table, our_table, _ = pairs[2]
    style_flags = ('hidden', 'locked', 'quick_style', 'unhide_when_used')
    for ref_style, our_style, _ in pairs[:4]:
        py_assert1 = (getattr(ref_style, flag) is getattr(our_style, flag) is False for flag in style_flags)
        py_assert3 = all(py_assert1)
        assert not not py_assert3
        py_assert1 = py_assert3 = None
        py_assert1 = ref_style.priority
        py_assert6 = our_style.priority
        py_assert3 = py_assert1 is py_assert6
        py_assert8 = None
        py_assert4 = py_assert6 is py_assert8
        assert not (not py_assert3 or not py_assert4)
        py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
        ref_style.name = our_style.name = ref_style.name + ' Renamed'
        for flag in style_flags:
            setattr(ref_style, flag, True)
            setattr(our_style, flag, True)
        ref_style.priority = our_style.priority = 37
        for flag in style_flags:
            setattr(ref_style, flag, False)
            setattr(our_style, flag, False)
            py_assert3 = getattr(ref_style, flag)
            py_assert10 = getattr(our_style, flag)
            py_assert5 = py_assert3 is py_assert10
            py_assert12 = False
            py_assert6 = py_assert10 is py_assert12
            assert not (not py_assert5 or not py_assert6)
            py_assert3 = py_assert5 = py_assert6 = py_assert10 = py_assert12 = None
            setattr(ref_style, flag, True)
            setattr(our_style, flag, True)
            continue
        ref_style.priority = our_style.priority = None
        py_assert1 = ref_style.priority
        py_assert6 = our_style.priority
        py_assert3 = py_assert1 is py_assert6
        py_assert8 = None
        py_assert4 = py_assert6 is py_assert8
        assert not (not py_assert3 or not py_assert4)
        py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
        ref_style.priority = our_style.priority = 37
        continue
    ref_character.font.name = 'Arial'
    our_character.font.name = 'Arial'
    ref_character.font.size = docx.shared.Pt(13)
    our_character.font.size = Pt(13)
    ref_paragraph.font.name = 'Verdana'
    our_paragraph.font.name = 'Verdana'
    ref_table.font.name = 'Calibri'
    our_table.font.name = 'Calibri'
    ref_table.font.size = docx.shared.Pt(11)
    our_table.font.size = Pt(11)
    ref_paragraph.next_paragraph_style = reference.styles['Normal']
    our_paragraph.next_paragraph_style = candidate.styles['Normal']
    ref_paragraph.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    our_paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ref_table.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.RIGHT
    our_table.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    def style_snapshot(style):
        base_style = getattr(style, 'base_style', None)
        next_style = getattr(style, 'next_paragraph_style', None)
        return (type(style).__name__, style.name, style.style_id, style.type.name, style.builtin, style.hidden, style.locked, style.priority, style.quick_style, style.unhide_when_used, type(base_style).__name__ if base_style is not None else None, getattr(base_style, 'style_id', None), type(next_style).__name__ if next_style is not None else None, getattr(next_style, 'style_id', None))
    for ref_style, our_style, _ in pairs:
        py_assert2 = style_snapshot(ref_style)
        py_assert7 = style_snapshot(our_style)
        py_assert4 = py_assert2 == py_assert7
        assert not not py_assert4
        py_assert2 = py_assert4 = py_assert7 = None
        continue
    py_assert0 = (ref_character.font.name, ref_character.font.size)
    py_assert3 = (our_character.font.name, our_character.font.size)
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None
    py_assert1 = ref_paragraph.font
    py_assert3 = py_assert1.name
    py_assert8 = our_paragraph.font
    py_assert10 = py_assert8.name
    py_assert5 = py_assert3 == py_assert10
    py_assert12 = 'Verdana'
    py_assert6 = py_assert10 == py_assert12
    assert not (not py_assert5 or not py_assert6)
    py_assert1 = py_assert3 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
    py_assert0 = (ref_table.font.name, ref_table.font.size)
    py_assert3 = (our_table.font.name, our_table.font.size)
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None
    py_assert1 = ref_paragraph.next_paragraph_style
    py_assert3 = py_assert1.style_id
    py_assert8 = our_paragraph.next_paragraph_style
    py_assert10 = py_assert8.style_id
    py_assert5 = py_assert3 == py_assert10
    py_assert12 = 'Normal'
    py_assert6 = py_assert10 == py_assert12
    assert not (not py_assert5 or not py_assert6)
    py_assert1 = py_assert3 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
    py_assert1 = ref_paragraph.paragraph_format
    py_assert3 = py_assert1.alignment
    py_assert5 = py_assert3.name
    py_assert10 = our_paragraph.paragraph_format
    py_assert12 = py_assert10.alignment
    py_assert14 = py_assert12.name
    py_assert7 = py_assert5 == py_assert14
    py_assert16 = 'CENTER'
    py_assert8 = py_assert14 == py_assert16
    assert not (not py_assert7 or not py_assert8)
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert8 = py_assert10 = py_assert12 = py_assert14 = py_assert16 = None
    py_assert1 = ref_table.paragraph_format
    py_assert3 = py_assert1.alignment
    py_assert5 = py_assert3.name
    py_assert10 = our_table.paragraph_format
    py_assert12 = py_assert10.alignment
    py_assert14 = py_assert12.name
    py_assert7 = py_assert5 == py_assert14
    py_assert16 = 'RIGHT'
    py_assert8 = py_assert14 == py_assert16
    assert not (not py_assert7 or not py_assert8)
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert8 = py_assert10 = py_assert12 = py_assert14 = py_assert16 = None
    default_types = ((docx.enum.style.WD_STYLE_TYPE.PARAGRAPH, WD_STYLE_TYPE.PARAGRAPH), (docx.enum.style.WD_STYLE_TYPE.CHARACTER, WD_STYLE_TYPE.CHARACTER), (docx.enum.style.WD_STYLE_TYPE.TABLE, WD_STYLE_TYPE.TABLE), (docx.enum.style.WD_STYLE_TYPE.LIST, WD_STYLE_TYPE.LIST))

    def defaults_snapshot(document, types):
        result = []
        for style_type in types:
            style = document.styles.default(style_type)
            result.append(None if style is None else (type(style).__name__, style.name, style.style_id, style.type.name))
        return result
    py_assert2 = [pair[0] for pair in default_types]
    py_assert4 = defaults_snapshot(reference, py_assert2)
    py_assert9 = [pair[1] for pair in default_types]
    py_assert11 = defaults_snapshot(candidate, py_assert9)
    py_assert6 = py_assert4 == py_assert11
    assert not not py_assert6
    py_assert2 = py_assert4 = py_assert6 = py_assert9 = py_assert11 = None
    ref_list_default = reference.styles.default(docx.enum.style.WD_STYLE_TYPE.LIST)
    our_list_default = candidate.styles.default(WD_STYLE_TYPE.LIST)
    py_assert2 = type(ref_list_default)
    py_assert4 = py_assert2.__name__
    py_assert10 = type(our_list_default)
    py_assert12 = py_assert10.__name__
    py_assert6 = py_assert4 == py_assert12
    py_assert14 = '_NumberingStyle'
    py_assert7 = py_assert12 == py_assert14
    assert not (not py_assert6 or not py_assert7)
    py_assert2 = py_assert4 = py_assert6 = py_assert7 = py_assert10 = py_assert12 = py_assert14 = None
    py_assert3 = isinstance(our_list_default, _NumberingStyle)
    assert not not py_assert3
    py_assert3 = None
    for member in ('font', 'paragraph_format', 'next_paragraph_style'):
        py_assert3 = hasattr(ref_list_default, member)
        py_assert10 = hasattr(our_list_default, member)
        py_assert5 = py_assert3 is py_assert10
        py_assert12 = False
        py_assert6 = py_assert10 is py_assert12
        assert not (not py_assert5 or not py_assert6)
        py_assert3 = py_assert5 = py_assert6 = py_assert10 = py_assert12 = None
        continue
    deletion_types = (('Character', docx.enum.style.WD_STYLE_TYPE.CHARACTER, WD_STYLE_TYPE.CHARACTER, 'CharacterStyle'), ('Paragraph', docx.enum.style.WD_STYLE_TYPE.PARAGRAPH, WD_STYLE_TYPE.PARAGRAPH, 'ParagraphStyle'), ('Table', docx.enum.style.WD_STYLE_TYPE.TABLE, WD_STYLE_TYPE.TABLE, '_TableStyle'), ('List', docx.enum.style.WD_STYLE_TYPE.LIST, WD_STYLE_TYPE.LIST, '_NumberingStyle'))
    for label, ref_type, our_type, class_name in deletion_types:
        name = f'Parity Deleted {label} Style'
        ref_deleted = reference.styles.add_style(name, ref_type)
        our_deleted = candidate.styles.add_style(name, our_type)
        py_assert2 = type(ref_deleted)
        py_assert4 = py_assert2.__name__
        py_assert10 = type(our_deleted)
        py_assert12 = py_assert10.__name__
        py_assert6 = py_assert4 == py_assert12
        py_assert7 = py_assert12 == class_name
        assert not (not py_assert6 or not py_assert7)
        py_assert2 = py_assert4 = py_assert6 = py_assert7 = py_assert10 = py_assert12 = None
        ref_deleted.delete()
        our_deleted.delete()
        with pytest.raises(KeyError):
            reference.styles[name]
        with pytest.raises(KeyError):
            candidate.styles[name]
        continue
    ref_path = tmp_path / 'reference-style-proxies.docx'
    our_path = tmp_path / 'candidate-style-proxies.docx'
    reference.save(ref_path)
    candidate.save(our_path)
    reopened_reference = docx.Document(str(ref_path))
    reopened_candidate = compat.Document(our_path.read_bytes())
    for name, _, _, _ in style_types:
        ref_style = reopened_reference.styles[name + ' Renamed']
        our_style = reopened_candidate.styles[name + ' Renamed']
        py_assert2 = style_snapshot(ref_style)
        py_assert7 = style_snapshot(our_style)
        py_assert4 = py_assert2 == py_assert7
        assert not not py_assert4
        py_assert2 = py_assert4 = py_assert7 = None
        if hasattr(ref_style, 'font'):
            py_assert0 = (ref_style.font.name, ref_style.font.size)
            py_assert3 = (our_style.font.name, our_style.font.size)
            py_assert2 = py_assert0 == py_assert3
            assert not not py_assert2
            py_assert0 = py_assert2 = py_assert3 = None
        if hasattr(ref_style, 'paragraph_format'):
            py_assert1 = ref_style.paragraph_format
            py_assert3 = py_assert1.alignment
            py_assert5 = py_assert3.name
            py_assert9 = our_style.paragraph_format
            py_assert11 = py_assert9.alignment
            py_assert13 = py_assert11.name
            py_assert7 = py_assert5 == py_assert13
            assert not not py_assert7
            py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = py_assert11 = py_assert13 = None
        if not hasattr(ref_style, 'next_paragraph_style'):
            continue
        py_assert1 = ref_style.next_paragraph_style
        py_assert3 = py_assert1.style_id
        py_assert7 = our_style.next_paragraph_style
        py_assert9 = py_assert7.style_id
        py_assert5 = py_assert3 == py_assert9
        assert not not py_assert5
        py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = None
        continue
    py_assert2 = [pair[0] for pair in default_types]
    py_assert4 = defaults_snapshot(reopened_reference, py_assert2)
    py_assert9 = [pair[1] for pair in default_types]
    py_assert11 = defaults_snapshot(reopened_candidate, py_assert9)
    py_assert6 = py_assert4 == py_assert11
    assert not not py_assert6
    py_assert2 = py_assert4 = py_assert6 = py_assert9 = py_assert11 = None

def test_styles_add_style_builtin_and_duplicate_behavior_matches_python_docx():
    reference = docx.Document()
    candidate = compat.Document()
    style_types = (('Character', docx.enum.style.WD_STYLE_TYPE.CHARACTER, WD_STYLE_TYPE.CHARACTER, 'CharacterStyle'), ('Paragraph', docx.enum.style.WD_STYLE_TYPE.PARAGRAPH, WD_STYLE_TYPE.PARAGRAPH, 'ParagraphStyle'), ('Table', docx.enum.style.WD_STYLE_TYPE.TABLE, WD_STYLE_TYPE.TABLE, '_TableStyle'), ('List', docx.enum.style.WD_STYLE_TYPE.LIST, WD_STYLE_TYPE.LIST, '_NumberingStyle'))
    for label, ref_type, our_type, proxy_name in style_types:
        name = f'Parity Builtin {label}'
        ref_style = reference.styles.add_style(name, ref_type, builtin=True)
        our_style = candidate.styles.add_style(name, our_type, builtin=True)
        py_assert2 = type(ref_style)
        py_assert4 = py_assert2.__name__
        py_assert10 = type(our_style)
        py_assert12 = py_assert10.__name__
        py_assert6 = py_assert4 == py_assert12
        py_assert7 = py_assert12 == proxy_name
        assert not (not py_assert6 or not py_assert7)
        py_assert2 = py_assert4 = py_assert6 = py_assert7 = py_assert10 = py_assert12 = None
        py_assert1 = ref_style.builtin
        py_assert6 = our_style.builtin
        py_assert3 = py_assert1 is py_assert6
        py_assert8 = True
        py_assert4 = py_assert6 is py_assert8
        assert not (not py_assert3 or not py_assert4)
        py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
        py_assert0 = reference.styles[name]
        py_assert2 = py_assert0.style_id
        py_assert5 = candidate.styles[name]
        py_assert7 = py_assert5.style_id
        py_assert4 = py_assert2 == py_assert7
        assert not not py_assert4
        py_assert0 = py_assert2 = py_assert4 = py_assert5 = py_assert7 = None
        with pytest.raises(ValueError):
            reference.styles.add_style(name, ref_type)
        with pytest.raises(ValueError):
            candidate.styles.add_style(name, our_type)
        continue

def test_styles_default_returns_none_without_a_default_definition():
    source = compat.Document().to_bytes()
    rewritten = io.BytesIO()
    word_ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    with zipfile.ZipFile(io.BytesIO(source)) as package:
        with zipfile.ZipFile(rewritten, 'w', zipfile.ZIP_DEFLATED) as output:
            for info in package.infolist():
                data = package.read(info.filename)
                if info.filename == 'word/styles.xml':
                    root = ET.fromstring(data)
                    for style in root.findall(f'{{{word_ns}}}style'):
                        if style.get(qn('w:type')) == 'numbering':
                            style.attrib.pop(qn('w:default'), None)
                    data = ET.tostring(root, encoding='utf-8', xml_declaration=True)
                output.writestr(info, data)
    raw = rewritten.getvalue()
    reference = docx.Document(io.BytesIO(raw))
    candidate = compat.Document(raw)
    py_assert1 = reference.styles
    py_assert3 = py_assert1.default
    py_assert6 = docx.enum
    py_assert8 = py_assert6.style
    py_assert10 = py_assert8.WD_STYLE_TYPE
    py_assert12 = py_assert10.LIST
    py_assert14 = py_assert3(py_assert12)
    py_assert17 = None
    py_assert16 = py_assert14 is py_assert17
    assert not not py_assert16
    py_assert1 = py_assert3 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = py_assert14 = py_assert16 = py_assert17 = None
    py_assert1 = candidate.styles
    py_assert3 = py_assert1.default
    py_assert6 = WD_STYLE_TYPE.LIST
    py_assert8 = py_assert3(py_assert6)
    py_assert11 = None
    py_assert10 = py_assert8 is py_assert11
    assert not not py_assert10
    py_assert1 = py_assert3 = py_assert6 = py_assert8 = py_assert10 = py_assert11 = None

def test_complete_default_style_collection_matches_python_docx():
    reference = docx.Document()
    candidate = compat.Document()

    def style_snapshot(style):
        next_style = getattr(style, 'next_paragraph_style', None)
        base_style = getattr(style, 'base_style', None)
        return (type(style).__name__, style.style_id, style.name, style.type.name, style.builtin, style.hidden, style.locked, style.priority, style.quick_style, style.unhide_when_used, type(base_style).__name__ if base_style is not None else None, getattr(base_style, 'style_id', None), type(next_style).__name__ if next_style is not None else None, getattr(next_style, 'style_id', None))
    reference_styles = list(reference.styles)
    candidate_styles = list(candidate.styles)
    py_assert2 = len(reference_styles)
    py_assert8 = len(candidate_styles)
    py_assert4 = py_assert2 == py_assert8
    py_assert10 = 164
    py_assert5 = py_assert8 == py_assert10
    assert not (not py_assert4 or not py_assert5)
    py_assert2 = py_assert4 = py_assert5 = py_assert8 = py_assert10 = None
    py_assert0 = [style_snapshot(style) for style in reference_styles]
    py_assert3 = [style_snapshot(style) for style in candidate_styles]
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None
    for ref_style, our_style in zip(reference_styles, candidate_styles):
        py_assert0 = reference.styles[ref_style.name]
        py_assert2 = py_assert0.style_id
        py_assert6 = ref_style.style_id
        py_assert4 = py_assert2 == py_assert6
        assert not not py_assert4
        py_assert0 = py_assert2 = py_assert4 = py_assert6 = None
        py_assert0 = candidate.styles[ref_style.name]
        py_assert2 = py_assert0.style_id
        py_assert6 = ref_style.style_id
        py_assert4 = py_assert2 == py_assert6
        assert not not py_assert4
        py_assert0 = py_assert2 = py_assert4 = py_assert6 = None
        continue

def test_latent_styles_match_python_docx_public_properties(tmp_path):
    document = compat.Document()
    latent = document.styles.latent_styles
    latent.default_priority = 43
    latent.default_to_hidden = False
    latent.default_to_locked = True
    latent.default_to_quick_style = True
    latent.default_to_unhide_when_used = False
    latent.load_count = 276
    entry = latent.add_latent_style('Coverage Latent')
    entry.hidden = True
    entry.locked = False
    entry.priority = 17
    entry.quick_style = True
    entry.unhide_when_used = None
    alias = latent.add_latent_style('Heading 1')
    py_assert1 = alias.name
    py_assert4 = 'Heading 1'
    py_assert3 = py_assert1 == py_assert4
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert4 = None
    py_assert0 = latent['Coverage Latent']
    py_assert2 = py_assert0.name
    py_assert5 = 'Coverage Latent'
    py_assert4 = py_assert2 == py_assert5
    assert not not py_assert4
    py_assert0 = py_assert2 = py_assert4 = py_assert5 = None
    py_assert2 = len(latent)
    py_assert5 = 2
    py_assert4 = py_assert2 >= py_assert5
    assert not not py_assert4
    py_assert2 = py_assert4 = py_assert5 = None
    path = tmp_path / 'latent-styles.docx'
    document.save(path)
    reference = docx.Document(str(path)).styles.latent_styles
    py_assert1 = reference.default_priority
    py_assert6 = latent.default_priority
    py_assert3 = py_assert1 == py_assert6
    py_assert8 = 43
    py_assert4 = py_assert6 == py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = reference.default_to_hidden
    py_assert6 = latent.default_to_hidden
    py_assert3 = py_assert1 == py_assert6
    py_assert8 = False
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = reference.default_to_locked
    py_assert6 = latent.default_to_locked
    py_assert3 = py_assert1 == py_assert6
    py_assert8 = True
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = reference.default_to_quick_style
    py_assert6 = latent.default_to_quick_style
    py_assert3 = py_assert1 == py_assert6
    py_assert8 = True
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = reference.default_to_unhide_when_used
    py_assert6 = latent.default_to_unhide_when_used
    py_assert3 = py_assert1 == py_assert6
    py_assert8 = False
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = reference.load_count
    py_assert6 = latent.load_count
    py_assert3 = py_assert1 == py_assert6
    py_assert8 = 276
    py_assert4 = py_assert6 == py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    reference_entry = reference['Coverage Latent']
    py_assert1 = reference_entry.name
    py_assert5 = entry.name
    py_assert3 = py_assert1 == py_assert5
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert5 = None
    py_assert1 = reference_entry.hidden
    py_assert6 = entry.hidden
    py_assert3 = py_assert1 == py_assert6
    py_assert8 = True
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = reference_entry.locked
    py_assert6 = entry.locked
    py_assert3 = py_assert1 == py_assert6
    py_assert8 = False
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = reference_entry.priority
    py_assert6 = entry.priority
    py_assert3 = py_assert1 == py_assert6
    py_assert8 = 17
    py_assert4 = py_assert6 == py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = reference_entry.quick_style
    py_assert6 = entry.quick_style
    py_assert3 = py_assert1 == py_assert6
    py_assert8 = True
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = reference_entry.unhide_when_used
    py_assert6 = entry.unhide_when_used
    py_assert3 = py_assert1 == py_assert6
    py_assert8 = None
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    entry.delete()
    with pytest.raises(KeyError):
        latent['Coverage Latent']
    with pytest.raises(AttributeError):
        _ = entry.name

def test_list_style_uses_numbering_style_proxy_like_python_docx(tmp_path):
    reference = docx.Document()
    candidate = compat.Document()
    ref_style = reference.styles.add_style('Parity List Style', docx.enum.style.WD_STYLE_TYPE.LIST)
    our_style = candidate.styles.add_style('Parity List Style', WD_STYLE_TYPE.LIST)
    py_assert2 = type(ref_style)
    py_assert4 = py_assert2.__name__
    py_assert10 = type(our_style)
    py_assert12 = py_assert10.__name__
    py_assert6 = py_assert4 == py_assert12
    py_assert14 = '_NumberingStyle'
    py_assert7 = py_assert12 == py_assert14
    assert not (not py_assert6 or not py_assert7)
    py_assert2 = py_assert4 = py_assert6 = py_assert7 = py_assert10 = py_assert12 = py_assert14 = None
    py_assert1 = ref_style.type
    py_assert5 = docx.enum
    py_assert7 = py_assert5.style
    py_assert9 = py_assert7.WD_STYLE_TYPE
    py_assert11 = py_assert9.LIST
    py_assert3 = py_assert1 is py_assert11
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = py_assert11 = None
    py_assert1 = our_style.type
    py_assert5 = WD_STYLE_TYPE.LIST
    py_assert3 = py_assert1 is py_assert5
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert5 = None
    py_assert1 = ref_style.name
    py_assert6 = our_style.name
    py_assert3 = py_assert1 == py_assert6
    py_assert8 = 'Parity List Style'
    py_assert4 = py_assert6 == py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = ref_style.builtin
    py_assert6 = our_style.builtin
    py_assert3 = py_assert1 is py_assert6
    py_assert8 = False
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert2 = 'next_paragraph_style'
    py_assert4 = hasattr(ref_style, py_assert2)
    py_assert7 = False
    py_assert6 = py_assert4 is py_assert7
    assert not not py_assert6
    py_assert2 = py_assert4 = py_assert6 = py_assert7 = None
    py_assert2 = 'next_paragraph_style'
    py_assert4 = hasattr(our_style, py_assert2)
    py_assert7 = False
    py_assert6 = py_assert4 is py_assert7
    assert not not py_assert6
    py_assert2 = py_assert4 = py_assert6 = py_assert7 = None
    py_assert2 = 'paragraph_format'
    py_assert4 = hasattr(ref_style, py_assert2)
    py_assert7 = False
    py_assert6 = py_assert4 is py_assert7
    assert not not py_assert6
    py_assert2 = py_assert4 = py_assert6 = py_assert7 = None
    py_assert2 = 'paragraph_format'
    py_assert4 = hasattr(our_style, py_assert2)
    py_assert7 = False
    py_assert6 = py_assert4 is py_assert7
    assert not not py_assert6
    py_assert2 = py_assert4 = py_assert6 = py_assert7 = None
    py_assert1 = reference.styles['Parity List Style']
    py_assert3 = type(py_assert1)
    py_assert5 = py_assert3.__name__
    py_assert10 = candidate.styles['Parity List Style']
    py_assert12 = type(py_assert10)
    py_assert14 = py_assert12.__name__
    py_assert7 = py_assert5 == py_assert14
    py_assert16 = '_NumberingStyle'
    py_assert8 = py_assert14 == py_assert16
    assert not (not py_assert7 or not py_assert8)
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert8 = py_assert10 = py_assert12 = py_assert14 = py_assert16 = None
    py_assert2 = candidate.styles
    py_assert4 = py_assert2.get_by_id
    py_assert7 = our_style.style_id
    py_assert10 = WD_STYLE_TYPE.LIST
    py_assert12 = py_assert4(py_assert7, py_assert10)
    py_assert14 = type(py_assert12)
    py_assert16 = py_assert14.__name__
    py_assert19 = '_NumberingStyle'
    py_assert18 = py_assert16 == py_assert19
    assert not not py_assert18
    py_assert2 = py_assert4 = py_assert7 = py_assert10 = py_assert12 = py_assert14 = py_assert16 = py_assert18 = py_assert19 = None
    ref_path = tmp_path / 'reference-list-style.docx'
    our_path = tmp_path / 'candidate-list-style.docx'
    reference.save(ref_path)
    candidate.save(our_path)
    reopened_reference = docx.Document(str(ref_path)).styles['Parity List Style']
    reopened_candidate = compat.Document(our_path.read_bytes()).styles['Parity List Style']
    py_assert2 = type(reopened_reference)
    py_assert4 = py_assert2.__name__
    py_assert10 = type(reopened_candidate)
    py_assert12 = py_assert10.__name__
    py_assert6 = py_assert4 == py_assert12
    py_assert14 = '_NumberingStyle'
    py_assert7 = py_assert12 == py_assert14
    assert not (not py_assert6 or not py_assert7)
    py_assert2 = py_assert4 = py_assert6 = py_assert7 = py_assert10 = py_assert12 = py_assert14 = None
    py_assert1 = reopened_reference.type
    py_assert5 = docx.enum
    py_assert7 = py_assert5.style
    py_assert9 = py_assert7.WD_STYLE_TYPE
    py_assert11 = py_assert9.LIST
    py_assert3 = py_assert1 is py_assert11
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = py_assert11 = None
    py_assert1 = reopened_candidate.type
    py_assert5 = WD_STYLE_TYPE.LIST
    py_assert3 = py_assert1 is py_assert5
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert5 = None

def test_table_cell_paragraphs_and_enum_properties_round_trip(tmp_path):
    document = compat.Document()
    table = document.add_table(rows=2, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = document.styles['Table Grid']
    cell = table.cell(row_idx=0, col_idx=0)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    paragraph = cell.paragraphs[0]
    paragraph.text = 'cell value'
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    table.rows[0].height = Inches(0.4)
    table.rows[0].height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
    py_assert1 = table.alignment
    py_assert5 = WD_ALIGN_PARAGRAPH.CENTER
    py_assert3 = py_assert1 is py_assert5
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert5 = None
    py_assert1 = cell.vertical_alignment
    py_assert5 = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    py_assert3 = py_assert1 is py_assert5
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert5 = None
    py_assert1 = paragraph.alignment
    py_assert5 = WD_ALIGN_PARAGRAPH.CENTER
    py_assert3 = py_assert1 is py_assert5
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert5 = None
    py_assert0 = table.rows[0]
    py_assert2 = py_assert0.height_rule
    py_assert6 = WD_ROW_HEIGHT_RULE.EXACTLY
    py_assert4 = py_assert2 is py_assert6
    assert not not py_assert4
    py_assert0 = py_assert2 = py_assert4 = py_assert6 = None
    py_assert0 = table.row_cells(row_idx=0)[0]
    py_assert2 = py_assert0.text
    py_assert5 = 'cell value'
    py_assert4 = py_assert2 == py_assert5
    assert not not py_assert4
    py_assert0 = py_assert2 = py_assert4 = py_assert5 = None
    py_assert0 = table.column_cells(column_idx=0)[0]
    py_assert2 = py_assert0.text
    py_assert5 = 'cell value'
    py_assert4 = py_assert2 == py_assert5
    assert not not py_assert4
    py_assert0 = py_assert2 = py_assert4 = py_assert5 = None
    py_assert1 = table.rows
    py_assert3 = py_assert1.table
    py_assert5 = py_assert3 is table
    assert not not py_assert5
    py_assert1 = py_assert3 = py_assert5 = None
    py_assert1 = table.columns
    py_assert3 = py_assert1.table
    py_assert5 = py_assert3 is table
    assert not not py_assert5
    py_assert1 = py_assert3 = py_assert5 = None
    py_assert0 = table.rows[0]
    py_assert2 = py_assert0.table
    py_assert4 = py_assert2 is table
    assert not not py_assert4
    py_assert0 = py_assert2 = py_assert4 = None
    py_assert0 = table.columns[0]
    py_assert2 = py_assert0.table
    py_assert4 = py_assert2 is table
    assert not not py_assert4
    py_assert0 = py_assert2 = py_assert4 = None
    path = tmp_path / 'table.docx'
    document.save(path)
    reopened = docx.Document(str(path))
    py_assert0 = reopened.tables[0]
    py_assert2 = py_assert0.alignment
    py_assert6 = docx.enum
    py_assert8 = py_assert6.text
    py_assert10 = py_assert8.WD_ALIGN_PARAGRAPH
    py_assert12 = py_assert10.CENTER
    py_assert4 = py_assert2 is py_assert12
    assert not not py_assert4
    py_assert0 = py_assert2 = py_assert4 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
    py_assert0 = reopened.tables[0]
    py_assert2 = py_assert0.cell
    py_assert4 = 0
    py_assert6 = 0
    py_assert8 = py_assert2(py_assert4, py_assert6)
    py_assert10 = py_assert8.vertical_alignment
    py_assert14 = docx.enum
    py_assert16 = py_assert14.table
    py_assert18 = py_assert16.WD_CELL_VERTICAL_ALIGNMENT
    py_assert20 = py_assert18.CENTER
    py_assert12 = py_assert10 is py_assert20
    assert not not py_assert12
    py_assert0 = py_assert2 = py_assert4 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = py_assert14 = py_assert16 = py_assert18 = py_assert20 = None
    py_assert0 = reopened.tables[0].cell(0, 0).paragraphs[0]
    py_assert2 = py_assert0.alignment
    py_assert6 = docx.enum
    py_assert8 = py_assert6.text
    py_assert10 = py_assert8.WD_ALIGN_PARAGRAPH
    py_assert12 = py_assert10.CENTER
    py_assert4 = py_assert2 is py_assert12
    assert not not py_assert4
    py_assert0 = py_assert2 = py_assert4 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
    py_assert0 = reopened.tables[0].rows[0]
    py_assert2 = py_assert0.height_rule
    py_assert6 = docx.enum
    py_assert8 = py_assert6.table
    py_assert10 = py_assert8.WD_ROW_HEIGHT_RULE
    py_assert12 = py_assert10.EXACTLY
    py_assert4 = py_assert2 is py_assert12
    assert not not py_assert4
    py_assert0 = py_assert2 = py_assert4 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
    py_assert0 = reopened.tables[0]
    py_assert2 = py_assert0.cell
    py_assert4 = 0
    py_assert6 = 0
    py_assert8 = py_assert2(py_assert4, py_assert6)
    py_assert10 = py_assert8.text
    py_assert13 = 'cell value'
    py_assert12 = py_assert10 == py_assert13
    assert not not py_assert12
    py_assert0 = py_assert2 = py_assert4 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = py_assert13 = None

def test_shared_unit_constructors_match_python_docx():
    cases = ((docx.shared.Inches, compat.Inches, 1.25), (docx.shared.Cm, compat.Cm, 3.5), (docx.shared.Mm, compat.Mm, 12.75), (docx.shared.Pt, compat.Pt, 13.5), (docx.shared.Emu, compat.Emu, 123456), (docx.shared.Twips, compat.Twips, 1.5))
    for reference_type, compat_type, value in cases:
        reference = reference_type(value)
        actual = compat_type(value)
        py_assert2 = int(actual)
        py_assert7 = int(reference)
        py_assert4 = py_assert2 == py_assert7
        assert not not py_assert4
        py_assert2 = py_assert4 = py_assert7 = None
        py_assert3 = compat.Length
        py_assert5 = isinstance(actual, py_assert3)
        assert not not py_assert5
        py_assert3 = py_assert5 = None
        py_assert3 = isinstance(actual, compat_type)
        assert not not py_assert3
        py_assert3 = None
        for unit in ('emu', 'inches', 'cm', 'mm', 'pt', 'twips'):
            py_assert3 = getattr(actual, unit)
            py_assert9 = getattr(reference, unit)
            py_assert5 = py_assert3 == py_assert9
            assert not not py_assert5
            py_assert3 = py_assert5 = py_assert9 = None
            continue
        continue
    reference_length = docx.shared.Length(123456)
    candidate_length = compat.Length(123456)
    py_assert2 = int(reference_length)
    py_assert7 = int(candidate_length)
    py_assert4 = py_assert2 == py_assert7
    assert not not py_assert4
    py_assert2 = py_assert4 = py_assert7 = None
    for unit in ('emu', 'inches', 'cm', 'mm', 'pt', 'twips'):
        py_assert3 = getattr(candidate_length, unit)
        py_assert9 = getattr(reference_length, unit)
        py_assert5 = py_assert3 == py_assert9
        assert not not py_assert5
        py_assert3 = py_assert5 = py_assert9 = None
        continue

def test_font_color_object_matches_python_docx_defaults_and_writes():
    reference_font = docx.Document().add_paragraph('colored').runs[0].font
    reference_color = reference_font.color
    document = compat.Document()
    run = document.add_paragraph('colored').runs[0]
    ref_assignment_error = _capture_exception(lambda: setattr(reference_font, 'color', None))
    our_assignment_error = _capture_exception(lambda: setattr(run.font, 'color', None))
    py_assert0 = ref_assignment_error[0]
    py_assert4 = our_assignment_error[0]
    py_assert2 = py_assert0 is py_assert4
    py_assert3 = py_assert4 is AttributeError
    assert not (not py_assert2 or not py_assert3)
    py_assert0 = py_assert2 = py_assert3 = py_assert4 = None
    color = run.font.color
    py_assert2 = None
    py_assert1 = color is not py_assert2
    assert not not py_assert1
    py_assert1 = py_assert2 = None
    py_assert1 = reference_color.rgb
    py_assert6 = color.rgb
    py_assert3 = py_assert1 is py_assert6
    py_assert8 = None
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = reference_color.theme_color
    py_assert6 = color.theme_color
    py_assert3 = py_assert1 is py_assert6
    py_assert8 = None
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = reference_color.type
    py_assert6 = color.type
    py_assert3 = py_assert1 is py_assert6
    py_assert8 = None
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    reference_color.rgb = docx.shared.RGBColor(12, 34, 56)
    color.rgb = RGBColor(12, 34, 56)
    py_assert1 = reference_color.rgb
    py_assert6 = color.rgb
    py_assert3 = py_assert1 == py_assert6
    py_assert9 = 12
    py_assert11 = 34
    py_assert13 = 56
    py_assert15 = RGBColor(py_assert9, py_assert11, py_assert13)
    py_assert4 = py_assert6 == py_assert15
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert9 = py_assert11 = py_assert13 = py_assert15 = None
    py_assert2 = color.rgb
    py_assert5 = compat.RGBColor
    py_assert7 = isinstance(py_assert2, py_assert5)
    assert not not py_assert7
    py_assert2 = py_assert5 = py_assert7 = None
    py_assert1 = color.rgb
    py_assert5 = docx.shared
    py_assert7 = py_assert5.RGBColor
    py_assert9 = 12
    py_assert11 = 34
    py_assert13 = 56
    py_assert15 = py_assert7(py_assert9, py_assert11, py_assert13)
    py_assert3 = py_assert1 == py_assert15
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = py_assert11 = py_assert13 = py_assert15 = None
    py_assert1 = RGBColor.from_string
    py_assert3 = '0C2238'
    py_assert5 = py_assert1(py_assert3)
    py_assert9 = docx.shared
    py_assert11 = py_assert9.RGBColor
    py_assert13 = py_assert11.from_string
    py_assert15 = '0C2238'
    py_assert17 = py_assert13(py_assert15)
    py_assert7 = py_assert5 == py_assert17
    assert not not py_assert7
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = py_assert11 = py_assert13 = py_assert15 = py_assert17 = None
    py_assert2 = RGBColor.from_string
    py_assert4 = '0C2238'
    py_assert6 = py_assert2(py_assert4)
    py_assert8 = str(py_assert6)
    py_assert11 = '0C2238'
    py_assert10 = py_assert8 == py_assert11
    assert not not py_assert10
    py_assert2 = py_assert4 = py_assert6 = py_assert8 = py_assert10 = py_assert11 = None
    with pytest.raises(TypeError, match='RGBColor\\(\\) takes three integer values'):
        RGBColor(12.5, 34, 56)
    py_assert1 = color.type
    py_assert5 = MSO_COLOR_TYPE.RGB
    py_assert3 = py_assert1 is py_assert5
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert5 = None
    style_font = document.styles['Heading 1'].font
    style_font.color.rgb = RGBColor(1, 2, 3)
    py_assert2 = style_font.color
    py_assert4 = py_assert2.rgb
    py_assert7 = compat.RGBColor
    py_assert9 = isinstance(py_assert4, py_assert7)
    assert not not py_assert9
    py_assert2 = py_assert4 = py_assert7 = py_assert9 = None
    reference_color.theme_color = docx.enum.dml.MSO_THEME_COLOR_INDEX.ACCENT_1
    color.theme_color = MSO_THEME_COLOR.ACCENT_1
    py_assert1 = reference_color.theme_color
    py_assert3 = py_assert1.name
    py_assert8 = color.theme_color
    py_assert10 = py_assert8.name
    py_assert5 = py_assert3 == py_assert10
    py_assert12 = 'ACCENT_1'
    py_assert6 = py_assert10 == py_assert12
    assert not (not py_assert5 or not py_assert6)
    py_assert1 = py_assert3 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
    py_assert1 = reference_color.type
    py_assert3 = py_assert1.name
    py_assert8 = color.type
    py_assert10 = py_assert8.name
    py_assert5 = py_assert3 == py_assert10
    py_assert12 = 'THEME'
    py_assert6 = py_assert10 == py_assert12
    assert not (not py_assert5 or not py_assert6)
    py_assert1 = py_assert3 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
    py_assert1 = color.theme_color
    py_assert5 = MSO_THEME_COLOR.ACCENT_1
    py_assert3 = py_assert1 is py_assert5
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert5 = None
    py_assert1 = color.type
    py_assert5 = MSO_COLOR_TYPE.THEME
    py_assert3 = py_assert1 is py_assert5
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert5 = None
    reference_color.rgb = None
    color.rgb = None
    py_assert1 = reference_color.rgb
    py_assert6 = color.rgb
    py_assert3 = py_assert1 is py_assert6
    py_assert8 = None
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = reference_color.theme_color
    py_assert6 = color.theme_color
    py_assert3 = py_assert1 is py_assert6
    py_assert8 = None
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = reference_color.type
    py_assert6 = color.type
    py_assert3 = py_assert1 is py_assert6
    py_assert8 = None
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None

def test_documented_font_properties_match_python_docx_and_round_trip(tmp_path):
    reference = docx.Document()
    candidate = compat.Document()
    ref_font = reference.add_paragraph().add_run('font').font
    our_font = candidate.add_paragraph().add_run('font').font
    flag_properties = ('all_caps', 'bold', 'complex_script', 'cs_bold', 'cs_italic', 'double_strike', 'emboss', 'hidden', 'imprint', 'italic', 'math', 'no_proof', 'outline', 'rtl', 'shadow', 'small_caps', 'snap_to_grid', 'spec_vanish', 'strike', 'web_hidden')
    properties = flag_properties + ('name', 'size', 'highlight_color', 'subscript', 'superscript', 'underline')

    def normalize(value):
        if value is None or isinstance(value, bool):
            return value
        if hasattr(value, 'name'):
            return value.name
        if isinstance(value, int):
            return int(value)
        return value

    def snapshot(font):
        return {name: normalize(getattr(font, name)) for name in properties}

    def assert_fonts_match():
        py_assert2 = snapshot(ref_font)
        py_assert7 = snapshot(our_font)
        py_assert4 = py_assert2 == py_assert7
        assert not not py_assert4
        py_assert2 = py_assert4 = py_assert7 = None
    assert_fonts_match()
    for name in flag_properties:
        for value in (True, False, None):
            setattr(ref_font, name, value)
            setattr(our_font, name, value)
            assert_fonts_match()
    for name in flag_properties:
        value = name in {'rtl', 'bold', 'italic', 'all_caps', 'small_caps'}
        setattr(ref_font, name, value)
        setattr(our_font, name, value)
    ref_font.name = our_font.name = 'Aptos'
    assert_fonts_match()
    ref_font.name = our_font.name = None
    assert_fonts_match()
    ref_font.name = our_font.name = 'Arial'
    ref_font.size = docx.shared.Pt(14)
    our_font.size = Pt(14)
    assert_fonts_match()
    ref_font.size = our_font.size = None
    assert_fonts_match()
    ref_font.size = docx.shared.Pt(12)
    our_font.size = Pt(12)
    for name in ('YELLOW', 'RED'):
        ref_font.highlight_color = getattr(docx.enum.text.WD_COLOR_INDEX, name)
        our_font.highlight_color = getattr(WD_COLOR_INDEX, name)
        assert_fonts_match()
    ref_font.highlight_color = our_font.highlight_color = None
    assert_fonts_match()
    for name in ('SINGLE', 'DOUBLE', 'WAVY', 'DOTTED'):
        ref_value = getattr(docx.enum.text.WD_UNDERLINE, name)
        our_value = getattr(WD_UNDERLINE, name)
        ref_font.underline = ref_value
        our_font.underline = our_value
        assert_fonts_match()
    for value in (True, False, None):
        ref_font.underline = our_font.underline = value
        assert_fonts_match()
    ref_font.subscript = our_font.subscript = True
    assert_fonts_match()
    ref_font.superscript = our_font.superscript = True
    assert_fonts_match()
    ref_font.subscript = our_font.subscript = None
    assert_fonts_match()
    ref_font.superscript = our_font.superscript = False
    assert_fonts_match()
    ref_font.superscript = our_font.superscript = None
    assert_fonts_match()
    ref_font.superscript = our_font.superscript = True
    ref_font.highlight_color = docx.enum.text.WD_COLOR_INDEX.YELLOW
    our_font.highlight_color = WD_COLOR_INDEX.YELLOW
    ref_font.underline = docx.enum.text.WD_UNDERLINE.DOUBLE
    our_font.underline = WD_UNDERLINE.DOUBLE
    assert_fonts_match()
    ref_path = tmp_path / 'reference-font-properties.docx'
    our_path = tmp_path / 'candidate-font-properties.docx'
    reference.save(ref_path)
    candidate.save(our_path)
    reopened_reference = docx.Document(str(ref_path)).paragraphs[0].runs[0].font
    reopened_candidate = compat.Document(our_path.read_bytes()).paragraphs[0].runs[0].font
    ref_font, our_font = (reopened_reference, reopened_candidate)
    assert_fonts_match()

def test_pictures_get_unique_parts_mutable_dimensions_and_story_rels(tmp_path):
    document = compat.Document()
    first = document.add_picture(blobs.PNG_96, width=Inches(1))
    second = document.add_picture(blobs.PNG_96, width=Inches(1))
    first.width = Inches(2)
    body_run = document.add_paragraph().add_run()
    third = body_run.add_picture(blobs.PNG_96, width=Inches(0.75))
    section = document.sections[0]
    header_run = section.header.add_paragraph('header image').add_run()
    header_shape = header_run.add_picture(blobs.PNG_96, width=Inches(0.5))
    header_shape_2 = header_run.add_picture(blobs.PNG_96, width=Inches(0.4))
    footer_run = section.footer.add_paragraph('footer image').add_run()
    footer_shape = footer_run.add_picture(blobs.PNG_96, width=Inches(0.25))
    for run in (body_run, header_run, footer_run):
        drawing = next((item for item in run.iter_inner_content() if type(item).__name__ == 'Drawing'))
        py_assert1 = drawing.has_picture
        py_assert4 = True
        py_assert3 = py_assert1 is py_assert4
        assert not not py_assert3
        py_assert1 = py_assert3 = py_assert4 = None
        py_assert1 = drawing.image
        py_assert3 = py_assert1.blob
        py_assert7 = blobs.PNG_96
        py_assert5 = py_assert3 == py_assert7
        assert not not py_assert5
        py_assert1 = py_assert3 = py_assert5 = py_assert7 = None
        continue
    py_assert0 = [first.name, second.name, third.name, header_shape.name, header_shape_2.name, footer_shape.name]
    py_assert3 = ['Picture 1', 'Picture 2', 'Picture 3', 'Picture 4', 'Picture 5', 'Picture 6']
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None
    py_assert1 = (shape.type is WD_INLINE_SHAPE.PICTURE for shape in (first, second, third, header_shape, header_shape_2, footer_shape))
    py_assert3 = all(py_assert1)
    assert not not py_assert3
    py_assert1 = py_assert3 = None
    py_assert1 = first.width
    py_assert5 = 2
    py_assert7 = Inches(py_assert5)
    py_assert3 = py_assert1 == py_assert7
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = None
    py_assert2 = document.inline_shapes
    py_assert4 = len(py_assert2)
    py_assert7 = 3
    py_assert6 = py_assert4 == py_assert7
    assert not not py_assert6
    py_assert2 = py_assert4 = py_assert6 = py_assert7 = None
    all_shapes = document._inline_shapes()
    py_assert2 = len(all_shapes)
    py_assert5 = 6
    py_assert4 = py_assert2 == py_assert5
    assert not not py_assert4
    py_assert2 = py_assert4 = py_assert5 = None
    py_assert0 = {shape.media for shape in all_shapes}
    py_assert3 = {f'word/media/image{i}.png' for i in range(1, 7)}
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None
    path = tmp_path / 'pictures.docx'
    document.save(path)
    reopened = docx.Document(str(path))
    py_assert2 = reopened.inline_shapes
    py_assert4 = len(py_assert2)
    py_assert7 = 3
    py_assert6 = py_assert4 == py_assert7
    assert not not py_assert6
    py_assert2 = py_assert4 = py_assert6 = py_assert7 = None
    py_assert0 = [shape.width for shape in reopened.inline_shapes]
    py_assert3 = [Inches(2), Inches(1), Inches(0.75)]
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None
    with zipfile.ZipFile(path) as package:
        names = set(package.namelist())
        py_assert0 = {f'word/media/image{i}.png' for i in range(1, 7)}
        py_assert2 = py_assert0 <= names
        if not py_assert2:
            py_format4 = pytest_ar._call_reprcompare(('<=',), (py_assert2,), ('%(py1)s <= %(py3)s',), (py_assert0, names)) % {'py1': pytest_ar._saferepr(py_assert0), 'py3': pytest_ar._saferepr(names) if 'names' in py_builtins.locals() or pytest_ar._should_repr_global_name(names) else 'names'}
            py_format6 = 'assert %(py5)s' % {'py5': py_format4}
            raise AssertionError(pytest_ar._format_explanation(py_format6))
        py_assert0 = py_assert2 = None
        header_rels = package.read('word/_rels/header1.xml.rels')
        footer_rels = package.read('word/_rels/footer1.xml.rels')
        py_assert0 = b'Id="rId1"'
        py_assert2 = py_assert0 in header_rels
        if not py_assert2:
            py_format4 = pytest_ar._call_reprcompare(('in',), (py_assert2,), ('%(py1)s in %(py3)s',), (py_assert0, header_rels)) % {'py1': pytest_ar._saferepr(py_assert0), 'py3': pytest_ar._saferepr(header_rels) if 'header_rels' in py_builtins.locals() or pytest_ar._should_repr_global_name(header_rels) else 'header_rels'}
            py_format6 = 'assert %(py5)s' % {'py5': py_format4}
            raise AssertionError(pytest_ar._format_explanation(py_format6))
        py_assert0 = py_assert2 = None
        py_assert0 = b'Target="media/image4.png"'
        py_assert2 = py_assert0 in header_rels
        if not py_assert2:
            py_format4 = pytest_ar._call_reprcompare(('in',), (py_assert2,), ('%(py1)s in %(py3)s',), (py_assert0, header_rels)) % {'py1': pytest_ar._saferepr(py_assert0), 'py3': pytest_ar._saferepr(header_rels) if 'header_rels' in py_builtins.locals() or pytest_ar._should_repr_global_name(header_rels) else 'header_rels'}
            py_format6 = 'assert %(py5)s' % {'py5': py_format4}
            raise AssertionError(pytest_ar._format_explanation(py_format6))
        py_assert0 = py_assert2 = None
        py_assert0 = b'Id="rId2"'
        py_assert2 = py_assert0 in header_rels
        if not py_assert2:
            py_format4 = pytest_ar._call_reprcompare(('in',), (py_assert2,), ('%(py1)s in %(py3)s',), (py_assert0, header_rels)) % {'py1': pytest_ar._saferepr(py_assert0), 'py3': pytest_ar._saferepr(header_rels) if 'header_rels' in py_builtins.locals() or pytest_ar._should_repr_global_name(header_rels) else 'header_rels'}
            py_format6 = 'assert %(py5)s' % {'py5': py_format4}
            raise AssertionError(pytest_ar._format_explanation(py_format6))
        py_assert0 = py_assert2 = None
        py_assert0 = b'Target="media/image5.png"'
        py_assert2 = py_assert0 in header_rels
        if not py_assert2:
            py_format4 = pytest_ar._call_reprcompare(('in',), (py_assert2,), ('%(py1)s in %(py3)s',), (py_assert0, header_rels)) % {'py1': pytest_ar._saferepr(py_assert0), 'py3': pytest_ar._saferepr(header_rels) if 'header_rels' in py_builtins.locals() or pytest_ar._should_repr_global_name(header_rels) else 'header_rels'}
            py_format6 = 'assert %(py5)s' % {'py5': py_format4}
            raise AssertionError(pytest_ar._format_explanation(py_format6))
        py_assert0 = py_assert2 = None
        py_assert0 = b'Id="rId1"'
        py_assert2 = py_assert0 in footer_rels
        if not py_assert2:
            py_format4 = pytest_ar._call_reprcompare(('in',), (py_assert2,), ('%(py1)s in %(py3)s',), (py_assert0, footer_rels)) % {'py1': pytest_ar._saferepr(py_assert0), 'py3': pytest_ar._saferepr(footer_rels) if 'footer_rels' in py_builtins.locals() or pytest_ar._should_repr_global_name(footer_rels) else 'footer_rels'}
            py_format6 = 'assert %(py5)s' % {'py5': py_format4}
            raise AssertionError(pytest_ar._format_explanation(py_format6))
        py_assert0 = py_assert2 = None
        py_assert0 = b'Target="media/image6.png"'
        py_assert2 = py_assert0 in footer_rels
        if not py_assert2:
            py_format4 = pytest_ar._call_reprcompare(('in',), (py_assert2,), ('%(py1)s in %(py3)s',), (py_assert0, footer_rels)) % {'py1': pytest_ar._saferepr(py_assert0), 'py3': pytest_ar._saferepr(footer_rels) if 'footer_rels' in py_builtins.locals() or pytest_ar._should_repr_global_name(footer_rels) else 'footer_rels'}
            py_format6 = 'assert %(py5)s' % {'py5': py_format4}
            raise AssertionError(pytest_ar._format_explanation(py_format6))
        py_assert0 = py_assert2 = None
        header_xml = package.read('word/header1.xml')
        footer_xml = package.read('word/footer1.xml')
        py_assert1 = []
        py_assert2 = b'<wp:inline'
        py_assert0, py_assert4 = (py_assert4, py_assert2 in header_xml)
        if py_assert4:
            py_assert9 = b'rId2'
            py_assert11 = py_assert9 in header_xml
            py_assert0 = py_assert11
        if not py_assert0:
            py_format6 = pytest_ar._call_reprcompare(('in',), (py_assert4,), ('%(py3)s in %(py5)s',), (py_assert2, header_xml)) % {'py3': pytest_ar._saferepr(py_assert2), 'py5': pytest_ar._saferepr(header_xml) if 'header_xml' in py_builtins.locals() or pytest_ar._should_repr_global_name(header_xml) else 'header_xml'}
            py_format8 = '%(py7)s' % {'py7': py_format6}
            py_assert1.append(py_format8)
            if py_assert4:
                py_format13 = pytest_ar._call_reprcompare(('in',), (py_assert11,), ('%(py10)s in %(py12)s',), (py_assert9, header_xml)) % {'py10': pytest_ar._saferepr(py_assert9), 'py12': pytest_ar._saferepr(header_xml) if 'header_xml' in py_builtins.locals() or pytest_ar._should_repr_global_name(header_xml) else 'header_xml'}
                py_format15 = '%(py14)s' % {'py14': py_format13}
                py_assert1.append(py_format15)
            py_format16 = pytest_ar._format_boolop(py_assert1, 0) % {}
            py_format18 = 'assert %(py17)s' % {'py17': py_format16}
            raise AssertionError(pytest_ar._format_explanation(py_format18))
        py_assert0 = py_assert1 = py_assert2 = py_assert4 = py_assert9 = py_assert11 = None
        py_assert1 = []
        py_assert2 = b'<wp:inline'
        py_assert0, py_assert4 = (py_assert4, py_assert2 in footer_xml)
        if py_assert4:
            py_assert9 = b'rId1'
            py_assert11 = py_assert9 in footer_xml
            py_assert0 = py_assert11
        if not py_assert0:
            py_format6 = pytest_ar._call_reprcompare(('in',), (py_assert4,), ('%(py3)s in %(py5)s',), (py_assert2, footer_xml)) % {'py3': pytest_ar._saferepr(py_assert2), 'py5': pytest_ar._saferepr(footer_xml) if 'footer_xml' in py_builtins.locals() or pytest_ar._should_repr_global_name(footer_xml) else 'footer_xml'}
            py_format8 = '%(py7)s' % {'py7': py_format6}
            py_assert1.append(py_format8)
            if py_assert4:
                py_format13 = pytest_ar._call_reprcompare(('in',), (py_assert11,), ('%(py10)s in %(py12)s',), (py_assert9, footer_xml)) % {'py10': pytest_ar._saferepr(py_assert9), 'py12': pytest_ar._saferepr(footer_xml) if 'footer_xml' in py_builtins.locals() or pytest_ar._should_repr_global_name(footer_xml) else 'footer_xml'}
                py_format15 = '%(py14)s' % {'py14': py_format13}
                py_assert1.append(py_format15)
            py_format16 = pytest_ar._format_boolop(py_assert1, 0) % {}
            py_format18 = 'assert %(py17)s' % {'py17': py_format16}
            raise AssertionError(pytest_ar._format_explanation(py_format18))
        py_assert0 = py_assert1 = py_assert2 = py_assert4 = py_assert9 = py_assert11 = None
    py_assert1 = reopened.sections[0].header.paragraphs[-1].runs[-1]
    py_assert3 = py_assert1._r
    py_assert5 = py_assert3.xpath
    py_assert7 = 'w:drawing/wp:inline'
    py_assert9 = py_assert5(py_assert7)
    py_assert11 = len(py_assert9)
    py_assert14 = 2
    py_assert13 = py_assert11 == py_assert14
    assert not not py_assert13
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = py_assert11 = py_assert13 = py_assert14 = None
    py_assert1 = reopened.sections[0].footer.paragraphs[-1].runs[-1]
    py_assert3 = py_assert1._r
    py_assert5 = py_assert3.xpath
    py_assert7 = 'w:drawing/wp:inline'
    py_assert9 = py_assert5(py_assert7)
    py_assert11 = len(py_assert9)
    py_assert14 = 1
    py_assert13 = py_assert11 == py_assert14
    assert not not py_assert13
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = py_assert11 = py_assert13 = py_assert14 = None
    py_assert2 = compat.Document
    py_assert5 = path.read_bytes
    py_assert7 = py_assert5()
    py_assert9 = py_assert2(py_assert7)
    py_assert11 = py_assert9.inline_shapes
    py_assert13 = len(py_assert11)
    py_assert16 = 3
    py_assert15 = py_assert13 == py_assert16
    assert not not py_assert15
    py_assert2 = py_assert5 = py_assert7 = py_assert9 = py_assert11 = py_assert13 = py_assert15 = py_assert16 = None

def test_run_picture_adds_relationship_to_existing_header_part(tmp_path):
    source = Path(__file__).resolve().parents[1] / 'corpus' / 'resources_AlteredParts_header-simple-plus-image.docx'
    if not source.exists():
        pytest.skip('the optional repository corpus is not present')
    document = compat.Document(source.read_bytes())
    header = document.sections[0].header
    paragraph = header.paragraphs[0]
    shape = paragraph.add_run().add_picture(blobs.PNG_96, width=Inches(0.5))
    py_assert1 = shape.name
    py_assert4 = 'Picture 2'
    py_assert3 = py_assert1 == py_assert4
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert4 = None
    all_shapes = document._inline_shapes()
    py_assert2 = len(all_shapes)
    py_assert5 = 2
    py_assert4 = py_assert2 == py_assert5
    assert not not py_assert4
    py_assert2 = py_assert4 = py_assert5 = None
    py_assert0 = {item.media for item in all_shapes}
    py_assert3 = {'word/media/image1.png', 'word/media/image2.png'}
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None
    path = tmp_path / 'existing-header-picture.docx'
    document.save(path)
    reopened = docx.Document(str(path))
    py_assert1 = (len(run._r.xpath('w:drawing/wp:inline')) for run in reopened.sections[0].header.paragraphs[0].runs)
    py_assert3 = sum(py_assert1)
    py_assert6 = 2
    py_assert5 = py_assert3 == py_assert6
    assert not not py_assert5
    py_assert1 = py_assert3 = py_assert5 = py_assert6 = None
    with zipfile.ZipFile(path) as package:
        rels = package.read('word/_rels/header2.xml.rels')
        py_assert1 = []
        py_assert2 = b'Id="rId1"'
        py_assert0, py_assert4 = (py_assert4, py_assert2 in rels)
        if py_assert4:
            py_assert9 = b'Target="media/image1.png"'
            py_assert11 = py_assert9 in rels
            py_assert0 = py_assert11
        if not py_assert0:
            py_format6 = pytest_ar._call_reprcompare(('in',), (py_assert4,), ('%(py3)s in %(py5)s',), (py_assert2, rels)) % {'py3': pytest_ar._saferepr(py_assert2), 'py5': pytest_ar._saferepr(rels) if 'rels' in py_builtins.locals() or pytest_ar._should_repr_global_name(rels) else 'rels'}
            py_format8 = '%(py7)s' % {'py7': py_format6}
            py_assert1.append(py_format8)
            if py_assert4:
                py_format13 = pytest_ar._call_reprcompare(('in',), (py_assert11,), ('%(py10)s in %(py12)s',), (py_assert9, rels)) % {'py10': pytest_ar._saferepr(py_assert9), 'py12': pytest_ar._saferepr(rels) if 'rels' in py_builtins.locals() or pytest_ar._should_repr_global_name(rels) else 'rels'}
                py_format15 = '%(py14)s' % {'py14': py_format13}
                py_assert1.append(py_format15)
            py_format16 = pytest_ar._format_boolop(py_assert1, 0) % {}
            py_format18 = 'assert %(py17)s' % {'py17': py_format16}
            raise AssertionError(pytest_ar._format_explanation(py_format18))
        py_assert0 = py_assert1 = py_assert2 = py_assert4 = py_assert9 = py_assert11 = None
        py_assert1 = []
        py_assert2 = b'Id="rId2"'
        py_assert0, py_assert4 = (py_assert4, py_assert2 in rels)
        if py_assert4:
            py_assert9 = b'Target="media/image2.png"'
            py_assert11 = py_assert9 in rels
            py_assert0 = py_assert11
        if not py_assert0:
            py_format6 = pytest_ar._call_reprcompare(('in',), (py_assert4,), ('%(py3)s in %(py5)s',), (py_assert2, rels)) % {'py3': pytest_ar._saferepr(py_assert2), 'py5': pytest_ar._saferepr(rels) if 'rels' in py_builtins.locals() or pytest_ar._should_repr_global_name(rels) else 'rels'}
            py_format8 = '%(py7)s' % {'py7': py_format6}
            py_assert1.append(py_format8)
            if py_assert4:
                py_format13 = pytest_ar._call_reprcompare(('in',), (py_assert11,), ('%(py10)s in %(py12)s',), (py_assert9, rels)) % {'py10': pytest_ar._saferepr(py_assert9), 'py12': pytest_ar._saferepr(rels) if 'rels' in py_builtins.locals() or pytest_ar._should_repr_global_name(rels) else 'rels'}
                py_format15 = '%(py14)s' % {'py14': py_format13}
                py_assert1.append(py_format15)
            py_format16 = pytest_ar._format_boolop(py_assert1, 0) % {}
            py_format18 = 'assert %(py17)s' % {'py17': py_format16}
            raise AssertionError(pytest_ar._format_explanation(py_format18))
        py_assert0 = py_assert1 = py_assert2 = py_assert4 = py_assert9 = py_assert11 = None
        py_assert1 = package.read
        py_assert3 = 'word/header2.xml'
        py_assert5 = py_assert1(py_assert3)
        py_assert7 = py_assert5.count
        py_assert9 = b'<wp:inline'
        py_assert11 = py_assert7(py_assert9)
        py_assert14 = 2
        py_assert13 = py_assert11 == py_assert14
        if not py_assert13:
            py_format16 = pytest_ar._call_reprcompare(('==',), (py_assert13,), ('%(py12)s\n{%(py12)s = %(py8)s\n{%(py8)s = %(py6)s\n{%(py6)s = %(py2)s\n{%(py2)s = %(py0)s.read\n}(%(py4)s)\n}.count\n}(%(py10)s)\n} == %(py15)s',), (py_assert11, py_assert14)) % {'py0': pytest_ar._saferepr(package) if 'package' in py_builtins.locals() or pytest_ar._should_repr_global_name(package) else 'package', 'py2': pytest_ar._saferepr(py_assert1), 'py4': pytest_ar._saferepr(py_assert3), 'py6': pytest_ar._saferepr(py_assert5), 'py8': pytest_ar._saferepr(py_assert7), 'py10': pytest_ar._saferepr(py_assert9), 'py12': pytest_ar._saferepr(py_assert11), 'py15': pytest_ar._saferepr(py_assert14)}
            py_format18 = 'assert %(py17)s' % {'py17': py_format16}
            raise AssertionError(pytest_ar._format_explanation(py_format18))
        py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = py_assert11 = py_assert13 = py_assert14 = None

def test_python_docx_keywords_character_styles_and_stream_save():
    document = compat.Document()
    character_style = document.styles.add_style('Inline Label', WD_STYLE_TYPE.CHARACTER)
    builtin_style = document.styles.add_style('Synthetic Builtin', WD_STYLE_TYPE.PARAGRAPH, True)
    py_assert1 = builtin_style.builtin
    py_assert4 = True
    py_assert3 = py_assert1 is py_assert4
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert4 = None
    paragraph = document.add_paragraph()
    run = paragraph.add_run(text=None, style=character_style)
    run.add_break(break_type=WD_BREAK.PAGE)
    run.add_picture(image_path_or_stream=io.BytesIO(blobs.PNG_96), width=Inches(0.5))
    document.add_picture(image_path_or_stream=io.BytesIO(blobs.PNG_96), width=Inches(1))
    py_assert1 = run.style
    py_assert3 = py_assert1.name
    py_assert6 = 'Inline Label'
    py_assert5 = py_assert3 == py_assert6
    assert not not py_assert5
    py_assert1 = py_assert3 = py_assert5 = py_assert6 = None
    py_assert1 = document.styles
    py_assert3 = py_assert1.get_by_id
    py_assert6 = character_style.style_id
    py_assert9 = WD_STYLE_TYPE.CHARACTER
    py_assert11 = py_assert3(py_assert6, style_type=py_assert9)
    py_assert13 = py_assert11.name
    py_assert16 = 'Inline Label'
    py_assert15 = py_assert13 == py_assert16
    assert not not py_assert15
    py_assert1 = py_assert3 = py_assert6 = py_assert9 = py_assert11 = py_assert13 = py_assert15 = py_assert16 = None
    py_assert1 = document.styles
    py_assert3 = py_assert1.get_style_id
    py_assert7 = WD_STYLE_TYPE.CHARACTER
    py_assert9 = py_assert3(style_or_name=character_style, style_type=py_assert7)
    py_assert13 = character_style.style_id
    py_assert11 = py_assert9 == py_assert13
    assert not not py_assert11
    py_assert1 = py_assert3 = py_assert7 = py_assert9 = py_assert11 = py_assert13 = None
    stream = io.BytesIO()
    py_assert1 = document.save
    py_assert4 = py_assert1(path_or_stream=stream)
    py_assert7 = None
    py_assert6 = py_assert4 is py_assert7
    assert not not py_assert6
    py_assert1 = py_assert4 = py_assert6 = py_assert7 = None
    stream.seek(0)
    reopened = docx.Document(stream)
    py_assert2 = reopened.inline_shapes
    py_assert4 = len(py_assert2)
    py_assert7 = 2
    py_assert6 = py_assert4 == py_assert7
    assert not not py_assert6
    py_assert2 = py_assert4 = py_assert6 = py_assert7 = None
    py_assert0 = [shape.width for shape in reopened.inline_shapes]
    py_assert3 = [Inches(0.5), Inches(1)]
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None
    saved_run = reopened.paragraphs[0].runs[0]
    py_assert1 = saved_run.style
    py_assert3 = py_assert1.name
    py_assert6 = 'Inline Label'
    py_assert5 = py_assert3 == py_assert6
    assert not not py_assert5
    py_assert1 = py_assert3 = py_assert5 = py_assert6 = None
    py_assert0 = saved_run._r.br_lst[0]
    py_assert2 = py_assert0.type
    py_assert5 = 'page'
    py_assert4 = py_assert2 == py_assert5
    assert not not py_assert4
    py_assert0 = py_assert2 = py_assert4 = py_assert5 = None

def test_header_footer_block_content_matches_python_docx_and_round_trip(tmp_path):
    reference = docx.Document()
    candidate = compat.Document()
    for kind in ('header', 'footer'):
        ref_story = getattr(reference.sections[0], kind)
        our_story = getattr(candidate.sections[0], kind)
        ref_story.paragraphs[0].text = f'{kind} opening'
        our_story.paragraphs[0].text = f'{kind} opening'
        ref_story.add_paragraph(f'{kind} added paragraph')
        our_story.add_paragraph(f'{kind} added paragraph')
        ref_table = ref_story.add_table(rows=1, cols=2, width=docx.shared.Inches(3))
        our_table = our_story.add_table(rows=1, cols=2, width=Inches(3))
        ref_table.cell(0, 0).text = f'{kind} left'
        our_table.cell(0, 0).text = f'{kind} left'
        ref_table.cell(0, 1).text = f'{kind} right'
        our_table.cell(0, 1).text = f'{kind} right'
        ref_story.add_paragraph(f'{kind} styled paragraph', style='Subtitle')
        our_story.add_paragraph(f'{kind} styled paragraph', style='Subtitle')
        ref_second_table = ref_story.add_table(rows=2, cols=1, width=docx.shared.Inches(2))
        our_second_table = our_story.add_table(rows=2, cols=1, width=Inches(2))
        ref_second_table.style = 'Table Grid'
        our_second_table.style = 'Table Grid'
        for row_index in range(2):
            text = f'{kind} second table row {row_index}'
            ref_second_table.cell(row_index, 0).text = text
            our_second_table.cell(row_index, 0).text = text
        continue

    def story_snapshot(story):
        def table_snapshot(table):
            return (
                len(table.rows),
                len(table.columns),
                [int(column.width) if column.width is not None else None for column in table.columns],
                [[cell.text for cell in row.cells] for row in table.rows],
            )

        return {
            'paragraphs': [(paragraph.text, paragraph.style.name) for paragraph in story.paragraphs],
            'tables': [table_snapshot(table) for table in story.tables],
            'inner_content': [
                ('paragraph', item.text, item.style.name)
                if hasattr(item, 'runs')
                else ('table', table_snapshot(item))
                for item in story.iter_inner_content()
            ],
        }
    for kind in ('header', 'footer'):
        py_assert2 = reference.sections[0]
        py_assert5 = getattr(py_assert2, kind)
        py_assert7 = story_snapshot(py_assert5)
        py_assert12 = candidate.sections[0]
        py_assert15 = getattr(py_assert12, kind)
        py_assert17 = story_snapshot(py_assert15)
        py_assert9 = py_assert7 == py_assert17
        assert not not py_assert9
        py_assert2 = py_assert5 = py_assert7 = py_assert9 = py_assert12 = py_assert15 = py_assert17 = None
        continue
    ref_path = tmp_path / 'reference-header-footer-blocks.docx'
    our_path = tmp_path / 'candidate-header-footer-blocks.docx'
    reference.save(ref_path)
    candidate.save(our_path)
    reopened_reference = docx.Document(str(ref_path))
    reopened_candidate = compat.Document(our_path.read_bytes())
    for kind in ('header', 'footer'):
        py_assert2 = reopened_reference.sections[0]
        py_assert5 = getattr(py_assert2, kind)
        py_assert7 = story_snapshot(py_assert5)
        py_assert12 = reopened_candidate.sections[0]
        py_assert15 = getattr(py_assert12, kind)
        py_assert17 = story_snapshot(py_assert15)
        py_assert9 = py_assert7 == py_assert17
        assert not not py_assert9
        py_assert2 = py_assert5 = py_assert7 = py_assert9 = py_assert12 = py_assert15 = py_assert17 = None
        continue
@pytest.mark.parametrize('kind', ['header', 'footer'])

def test_header_footer_linking_setter_and_inherited_content_match(tmp_path, kind):
    reference = docx.Document()
    candidate = compat.Document()
    ref_first = getattr(reference.sections[0], kind)
    our_first = getattr(candidate.sections[0], kind)
    py_assert1 = ref_first.is_linked_to_previous
    py_assert4 = True
    py_assert3 = py_assert1 is py_assert4
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert4 = None
    py_assert1 = our_first.is_linked_to_previous
    py_assert4 = True
    py_assert3 = py_assert1 is py_assert4
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert4 = None
    ref_first.paragraphs[0].text = 'shared definition'
    our_first.paragraphs[0].text = 'shared definition'
    py_assert1 = ref_first.is_linked_to_previous
    py_assert4 = False
    py_assert3 = py_assert1 is py_assert4
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert4 = None
    py_assert1 = our_first.is_linked_to_previous
    py_assert4 = False
    py_assert3 = py_assert1 is py_assert4
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert4 = None
    ref_second = reference.add_section(docx.enum.section.WD_SECTION_START.NEW_PAGE)
    our_second = candidate.add_section(WD_SECTION_START.NEW_PAGE)
    ref_second_part = getattr(ref_second, kind)
    our_second_part = getattr(our_second, kind)
    py_assert1 = ref_second_part.is_linked_to_previous
    py_assert4 = True
    py_assert3 = py_assert1 is py_assert4
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert4 = None
    py_assert1 = our_second_part.is_linked_to_previous
    py_assert4 = True
    py_assert3 = py_assert1 is py_assert4
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert4 = None
    py_assert0 = ref_second_part.paragraphs[0]
    py_assert2 = py_assert0.text
    py_assert6 = our_second_part.paragraphs[0]
    py_assert8 = py_assert6.text
    py_assert4 = py_assert2 == py_assert8
    py_assert10 = 'shared definition'
    py_assert5 = py_assert8 == py_assert10
    assert not (not py_assert4 or not py_assert5)
    py_assert0 = py_assert2 = py_assert4 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = None
    py_assert1 = ref_second_part.is_linked_to_previous
    py_assert4 = True
    py_assert3 = py_assert1 is py_assert4
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert4 = None
    py_assert1 = our_second_part.is_linked_to_previous
    py_assert4 = True
    py_assert3 = py_assert1 is py_assert4
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert4 = None
    ref_second_part.is_linked_to_previous = False
    our_second_part.is_linked_to_previous = False
    py_assert1 = ref_second_part.is_linked_to_previous
    py_assert4 = False
    py_assert3 = py_assert1 is py_assert4
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert4 = None
    py_assert1 = our_second_part.is_linked_to_previous
    py_assert4 = False
    py_assert3 = py_assert1 is py_assert4
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert4 = None
    py_assert2 = ref_second_part.paragraphs
    py_assert4 = len(py_assert2)
    py_assert10 = our_second_part.paragraphs
    py_assert12 = len(py_assert10)
    py_assert6 = py_assert4 == py_assert12
    py_assert14 = 1
    py_assert7 = py_assert12 == py_assert14
    assert not (not py_assert6 or not py_assert7)
    py_assert2 = py_assert4 = py_assert6 = py_assert7 = py_assert10 = py_assert12 = py_assert14 = None
    py_assert0 = ref_second_part.paragraphs[0]
    py_assert2 = py_assert0.text
    py_assert6 = our_second_part.paragraphs[0]
    py_assert8 = py_assert6.text
    py_assert4 = py_assert2 == py_assert8
    py_assert10 = ''
    py_assert5 = py_assert8 == py_assert10
    assert not (not py_assert4 or not py_assert5)
    py_assert0 = py_assert2 = py_assert4 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = None
    ref_second_part.paragraphs[0].text = 'section-specific'
    our_second_part.paragraphs[0].text = 'section-specific'
    ref_false_path = tmp_path / f'reference-{kind}-unlinked.docx'
    our_false_path = tmp_path / f'candidate-{kind}-unlinked.docx'
    reference.save(ref_false_path)
    candidate.save(our_false_path)
    reference = docx.Document(str(ref_false_path))
    candidate = compat.Document(our_false_path.read_bytes())
    ref_second_part = getattr(reference.sections[1], kind)
    our_second_part = getattr(candidate.sections[1], kind)
    py_assert1 = ref_second_part.is_linked_to_previous
    py_assert4 = False
    py_assert3 = py_assert1 is py_assert4
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert4 = None
    py_assert1 = our_second_part.is_linked_to_previous
    py_assert4 = False
    py_assert3 = py_assert1 is py_assert4
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert4 = None
    py_assert0 = ref_second_part.paragraphs[0]
    py_assert2 = py_assert0.text
    py_assert6 = our_second_part.paragraphs[0]
    py_assert8 = py_assert6.text
    py_assert4 = py_assert2 == py_assert8
    py_assert10 = 'section-specific'
    py_assert5 = py_assert8 == py_assert10
    assert not (not py_assert4 or not py_assert5)
    py_assert0 = py_assert2 = py_assert4 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = None
    ref_second_part.is_linked_to_previous = True
    our_second_part.is_linked_to_previous = True
    py_assert1 = ref_second_part.is_linked_to_previous
    py_assert4 = True
    py_assert3 = py_assert1 is py_assert4
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert4 = None
    py_assert1 = our_second_part.is_linked_to_previous
    py_assert4 = True
    py_assert3 = py_assert1 is py_assert4
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert4 = None
    py_assert0 = ref_second_part.paragraphs[0]
    py_assert2 = py_assert0.text
    py_assert6 = our_second_part.paragraphs[0]
    py_assert8 = py_assert6.text
    py_assert4 = py_assert2 == py_assert8
    py_assert10 = 'shared definition'
    py_assert5 = py_assert8 == py_assert10
    assert not (not py_assert4 or not py_assert5)
    py_assert0 = py_assert2 = py_assert4 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = None
    py_assert1 = ref_second_part.is_linked_to_previous
    py_assert4 = True
    py_assert3 = py_assert1 is py_assert4
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert4 = None
    py_assert1 = our_second_part.is_linked_to_previous
    py_assert4 = True
    py_assert3 = py_assert1 is py_assert4
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert4 = None
    ref_linked_path = tmp_path / f'reference-{kind}-linked.docx'
    our_linked_path = tmp_path / f'candidate-{kind}-linked.docx'
    reference.save(ref_linked_path)
    candidate.save(our_linked_path)
    reopened_reference = docx.Document(str(ref_linked_path))
    reopened_candidate = compat.Document(our_linked_path.read_bytes())
    ref_reopened = getattr(reopened_reference.sections[1], kind)
    our_reopened = getattr(reopened_candidate.sections[1], kind)
    py_assert1 = ref_reopened.is_linked_to_previous
    py_assert4 = True
    py_assert3 = py_assert1 is py_assert4
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert4 = None
    py_assert1 = our_reopened.is_linked_to_previous
    py_assert4 = True
    py_assert3 = py_assert1 is py_assert4
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert4 = None
    py_assert0 = ref_reopened.paragraphs[0]
    py_assert2 = py_assert0.text
    py_assert6 = our_reopened.paragraphs[0]
    py_assert8 = py_assert6.text
    py_assert4 = py_assert2 == py_assert8
    py_assert10 = 'shared definition'
    py_assert5 = py_assert8 == py_assert10
    assert not (not py_assert4 or not py_assert5)
    py_assert0 = py_assert2 = py_assert4 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = None
    relationship_type = f'/{kind}'
    counts = []
    for path in (ref_linked_path, our_linked_path):
        with zipfile.ZipFile(path) as package:
            relationships = ET.fromstring(package.read('word/_rels/document.xml.rels'))
            counts.append(sum(((rel.get('Type') or '').endswith(relationship_type) for rel in relationships)))
    py_assert0 = counts[0]
    py_assert4 = counts[1]
    py_assert2 = py_assert0 == py_assert4
    py_assert6 = 1
    py_assert3 = py_assert4 == py_assert6
    assert not (not py_assert2 or not py_assert3)
    py_assert0 = py_assert2 = py_assert3 = py_assert4 = py_assert6 = None

def test_sections_sequence_matches_python_docx_indexing_and_slicing():
    reference = docx.Document()
    candidate = compat.Document()
    start_types = (WD_SECTION_START.EVEN_PAGE, WD_SECTION_START.CONTINUOUS)
    for start_type in start_types:
        reference.add_section(start_type)
        candidate.add_section(start_type)

    def summarize(section):
        return (section.start_type.name, int(section.page_width))
    ref_sections = reference.sections
    our_sections = candidate.sections
    py_assert2 = len(ref_sections)
    py_assert8 = len(our_sections)
    py_assert4 = py_assert2 == py_assert8
    py_assert10 = 3
    py_assert5 = py_assert8 == py_assert10
    assert not (not py_assert4 or not py_assert5)
    py_assert2 = py_assert4 = py_assert5 = py_assert8 = py_assert10 = None
    py_assert0 = [summarize(section) for section in ref_sections]
    py_assert3 = [summarize(section) for section in our_sections]
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None
    for index in (0, 1, 2, -1, -2, -3):
        py_assert1 = ref_sections[index]
        py_assert3 = summarize(py_assert1)
        py_assert7 = our_sections[index]
        py_assert9 = summarize(py_assert7)
        py_assert5 = py_assert3 == py_assert9
        assert not not py_assert5
        py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = None
        continue
    for selection in (slice(None), slice(1, None), slice(None, None, -1), slice(0, 3, 2)):
        py_assert0 = [summarize(section) for section in ref_sections[selection]]
        py_assert3 = [summarize(section) for section in our_sections[selection]]
        py_assert2 = py_assert0 == py_assert3
        assert not not py_assert2
        py_assert0 = py_assert2 = py_assert3 = None
    for index in (3, -4):
        with pytest.raises(IndexError) as ref_error:
            ref_sections[index]
        with pytest.raises(IndexError) as our_error:
            our_sections[index]
        py_assert2 = ref_error.value
        py_assert4 = str(py_assert2)
        py_assert9 = our_error.value
        py_assert11 = str(py_assert9)
        py_assert6 = py_assert4 == py_assert11
        assert not not py_assert6
        py_assert2 = py_assert4 = py_assert6 = py_assert9 = py_assert11 = None
        continue
@pytest.mark.parametrize('story_name', ('first_page_header', 'first_page_footer', 'even_page_header', 'even_page_footer'))

def test_special_section_header_footer_stories_match_python_docx(tmp_path, story_name):
    reference = docx.Document()
    candidate = compat.Document()
    ref_first_section = reference.sections[0]
    our_first_section = candidate.sections[0]

    def activation_state(document, section):
        if story_name.startswith('first_page_'):
            return section.different_first_page_header_footer
        return document.settings.odd_and_even_pages_header_footer
    py_assert3 = activation_state(reference, ref_first_section)
    py_assert6 = False
    py_assert5 = py_assert3 is py_assert6
    assert not not py_assert5
    py_assert3 = py_assert5 = py_assert6 = None
    py_assert3 = activation_state(candidate, our_first_section)
    py_assert6 = False
    py_assert5 = py_assert3 is py_assert6
    assert not not py_assert5
    py_assert3 = py_assert5 = py_assert6 = None
    ref_first = getattr(ref_first_section, story_name)
    our_first = getattr(our_first_section, story_name)
    py_assert1 = ref_first.is_linked_to_previous
    py_assert4 = True
    py_assert3 = py_assert1 is py_assert4
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert4 = None
    py_assert1 = our_first.is_linked_to_previous
    py_assert4 = True
    py_assert3 = py_assert1 is py_assert4
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert4 = None
    ref_first.paragraphs[0].text = 'shared special story'
    our_first.paragraphs[0].text = 'shared special story'
    py_assert3 = activation_state(reference, ref_first_section)
    py_assert6 = False
    py_assert5 = py_assert3 is py_assert6
    assert not not py_assert5
    py_assert3 = py_assert5 = py_assert6 = None
    py_assert3 = activation_state(candidate, our_first_section)
    py_assert6 = False
    py_assert5 = py_assert3 is py_assert6
    assert not not py_assert5
    py_assert3 = py_assert5 = py_assert6 = None
    ref_second_section = reference.add_section(docx.enum.section.WD_SECTION_START.NEW_PAGE)
    our_second_section = candidate.add_section(WD_SECTION_START.NEW_PAGE)
    ref_second = getattr(ref_second_section, story_name)
    our_second = getattr(our_second_section, story_name)
    py_assert1 = ref_second.is_linked_to_previous
    py_assert4 = True
    py_assert3 = py_assert1 is py_assert4
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert4 = None
    py_assert1 = our_second.is_linked_to_previous
    py_assert4 = True
    py_assert3 = py_assert1 is py_assert4
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert4 = None
    py_assert0 = ref_second.paragraphs[0]
    py_assert2 = py_assert0.text
    py_assert6 = our_second.paragraphs[0]
    py_assert8 = py_assert6.text
    py_assert4 = py_assert2 == py_assert8
    py_assert10 = 'shared special story'
    py_assert5 = py_assert8 == py_assert10
    assert not (not py_assert4 or not py_assert5)
    py_assert0 = py_assert2 = py_assert4 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = None
    ref_second.is_linked_to_previous = False
    our_second.is_linked_to_previous = False
    py_assert0 = ref_second.paragraphs[0]
    py_assert2 = py_assert0.text
    py_assert6 = our_second.paragraphs[0]
    py_assert8 = py_assert6.text
    py_assert4 = py_assert2 == py_assert8
    py_assert10 = ''
    py_assert5 = py_assert8 == py_assert10
    assert not (not py_assert4 or not py_assert5)
    py_assert0 = py_assert2 = py_assert4 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = None
    ref_second.paragraphs[0].text = 'section-specific story'
    our_second.paragraphs[0].text = 'section-specific story'
    ref_path = tmp_path / f'reference-{story_name}.docx'
    our_path = tmp_path / f'candidate-{story_name}.docx'
    reference.save(ref_path)
    candidate.save(our_path)
    reopened_reference = docx.Document(str(ref_path))
    reopened_candidate = compat.Document(our_path.read_bytes())
    ref_reopened = getattr(reopened_reference.sections[1], story_name)
    our_reopened = getattr(reopened_candidate.sections[1], story_name)
    py_assert1 = ref_reopened.is_linked_to_previous
    py_assert4 = False
    py_assert3 = py_assert1 is py_assert4
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert4 = None
    py_assert1 = our_reopened.is_linked_to_previous
    py_assert4 = False
    py_assert3 = py_assert1 is py_assert4
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert4 = None
    py_assert0 = ref_reopened.paragraphs[0]
    py_assert2 = py_assert0.text
    py_assert6 = our_reopened.paragraphs[0]
    py_assert8 = py_assert6.text
    py_assert4 = py_assert2 == py_assert8
    py_assert10 = 'section-specific story'
    py_assert5 = py_assert8 == py_assert10
    assert not (not py_assert4 or not py_assert5)
    py_assert0 = py_assert2 = py_assert4 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = None
    py_assert2 = reopened_reference.sections[0]
    py_assert4 = activation_state(reopened_reference, py_assert2)
    py_assert7 = False
    py_assert6 = py_assert4 is py_assert7
    assert not not py_assert6
    py_assert2 = py_assert4 = py_assert6 = py_assert7 = None
    py_assert2 = reopened_candidate.sections[0]
    py_assert4 = activation_state(reopened_candidate, py_assert2)
    py_assert7 = False
    py_assert6 = py_assert4 is py_assert7
    assert not not py_assert6
    py_assert2 = py_assert4 = py_assert6 = py_assert7 = None
    ref_reopened.is_linked_to_previous = True
    our_reopened.is_linked_to_previous = True
    py_assert0 = ref_reopened.paragraphs[0]
    py_assert2 = py_assert0.text
    py_assert6 = our_reopened.paragraphs[0]
    py_assert8 = py_assert6.text
    py_assert4 = py_assert2 == py_assert8
    py_assert10 = 'shared special story'
    py_assert5 = py_assert8 == py_assert10
    assert not (not py_assert4 or not py_assert5)
    py_assert0 = py_assert2 = py_assert4 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = None
    py_assert2 = reopened_reference.sections[0]
    py_assert4 = activation_state(reopened_reference, py_assert2)
    py_assert7 = False
    py_assert6 = py_assert4 is py_assert7
    assert not not py_assert6
    py_assert2 = py_assert4 = py_assert6 = py_assert7 = None
    py_assert2 = reopened_candidate.sections[0]
    py_assert4 = activation_state(reopened_candidate, py_assert2)
    py_assert7 = False
    py_assert6 = py_assert4 is py_assert7
    assert not not py_assert6
    py_assert2 = py_assert4 = py_assert6 = py_assert7 = None

def test_section_geometry_and_settings_match_python_docx_and_round_trip(tmp_path):
    reference = docx.Document()
    candidate = compat.Document()
    ref_section = reference.sections[0]
    our_section = candidate.sections[0]
    length_fields = ('page_width', 'page_height', 'left_margin', 'right_margin', 'top_margin', 'bottom_margin', 'header_distance', 'footer_distance', 'gutter')

    def snapshot(section):
        return {**{name: int(getattr(section, name)) if getattr(section, name) is not None else None for name in length_fields}, 'orientation': section.orientation.name, 'start_type': section.start_type.name, 'different_first_page_header_footer': section.different_first_page_header_footer}
    py_assert2 = snapshot(ref_section)
    py_assert7 = snapshot(our_section)
    py_assert4 = py_assert2 == py_assert7
    assert not not py_assert4
    py_assert2 = py_assert4 = py_assert7 = None
    values = {'page_width': Inches(9), 'page_height': Inches(12), 'left_margin': Inches(0.75), 'right_margin': Inches(0.8), 'top_margin': Inches(0.65), 'bottom_margin': Inches(0.7), 'header_distance': Inches(0.3), 'footer_distance': Inches(0.35), 'gutter': Inches(0.15)}
    for name, value in values.items():
        setattr(ref_section, name, value)
        setattr(our_section, name, value)
    ref_section.orientation = docx.enum.section.WD_ORIENTATION.LANDSCAPE
    our_section.orientation = compat.enum.WD_ORIENTATION.LANDSCAPE
    ref_section.start_type = docx.enum.section.WD_SECTION_START.ODD_PAGE
    our_section.start_type = WD_SECTION_START.ODD_PAGE
    ref_section.different_first_page_header_footer = True
    our_section.different_first_page_header_footer = True
    py_assert2 = snapshot(ref_section)
    py_assert7 = snapshot(our_section)
    py_assert4 = py_assert2 == py_assert7
    assert not not py_assert4
    py_assert2 = py_assert4 = py_assert7 = None
    for name in length_fields:
        setattr(ref_section, name, None)
        setattr(our_section, name, None)
    ref_section.start_type = None
    our_section.start_type = None
    ref_section.orientation = None
    our_section.orientation = None
    ref_section.different_first_page_header_footer = False
    our_section.different_first_page_header_footer = False
    py_assert2 = snapshot(ref_section)
    py_assert7 = snapshot(our_section)
    py_assert4 = py_assert2 == py_assert7
    assert not not py_assert4
    py_assert2 = py_assert4 = py_assert7 = None
    for enabled in (True, False):
        reference.settings.odd_and_even_pages_header_footer = enabled
        candidate.settings.odd_and_even_pages_header_footer = enabled
        py_assert1 = reference.settings
        py_assert3 = py_assert1.odd_and_even_pages_header_footer
        py_assert5 = py_assert3 is enabled
        assert not not py_assert5
        py_assert1 = py_assert3 = py_assert5 = None
        py_assert1 = candidate.settings
        py_assert3 = py_assert1.odd_and_even_pages_header_footer
        py_assert5 = py_assert3 is enabled
        assert not not py_assert5
        py_assert1 = py_assert3 = py_assert5 = None
        ref_settings_path = tmp_path / f'reference-even-odd-{enabled}.docx'
        our_settings_path = tmp_path / f'candidate-even-odd-{enabled}.docx'
        reference.save(ref_settings_path)
        candidate.save(our_settings_path)
        py_assert1 = docx.Document
        py_assert5 = str(ref_settings_path)
        py_assert7 = py_assert1(py_assert5)
        py_assert9 = py_assert7.settings
        py_assert11 = py_assert9.odd_and_even_pages_header_footer
        py_assert16 = compat.Document
        py_assert19 = our_settings_path.read_bytes
        py_assert21 = py_assert19()
        py_assert23 = py_assert16(py_assert21)
        py_assert25 = py_assert23.settings
        py_assert27 = py_assert25.odd_and_even_pages_header_footer
        py_assert13 = py_assert11 is py_assert27
        py_assert14 = py_assert27 is enabled
        assert not (not py_assert13 or not py_assert14)
        py_assert1 = py_assert5 = py_assert7 = py_assert9 = py_assert11 = py_assert13 = py_assert14 = py_assert16 = py_assert19 = py_assert21 = py_assert23 = py_assert25 = py_assert27 = None
        continue
    ref_path = tmp_path / 'reference-section-properties.docx'
    our_path = tmp_path / 'candidate-section-properties.docx'
    reference.save(ref_path)
    candidate.save(our_path)
    reopened_reference = docx.Document(str(ref_path))
    reopened_candidate = compat.Document(our_path.read_bytes())
    py_assert1 = reopened_reference.sections[0]
    py_assert3 = snapshot(py_assert1)
    py_assert7 = reopened_candidate.sections[0]
    py_assert9 = snapshot(py_assert7)
    py_assert5 = py_assert3 == py_assert9
    assert not not py_assert5
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = None
    py_assert1 = reopened_reference.settings
    py_assert3 = py_assert1.odd_and_even_pages_header_footer
    py_assert8 = reopened_candidate.settings
    py_assert10 = py_assert8.odd_and_even_pages_header_footer
    py_assert5 = py_assert3 is py_assert10
    py_assert12 = False
    py_assert6 = py_assert10 is py_assert12
    assert not (not py_assert5 or not py_assert6)
    py_assert1 = py_assert3 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None

def test_section_iter_inner_content_matches_python_docx_in_document_order(tmp_path):
    reference = docx.Document()
    candidate = compat.Document()
    for document in (reference, candidate):
        document.add_paragraph('before the break')
        table = document.add_table(rows=1, cols=1)
        table.cell(0, 0).text = 'inside table'
        document.add_paragraph('section-ending paragraph')
        document.add_section(WD_SECTION_START.CONTINUOUS)
        document.add_paragraph('after the break')
        document.add_table(rows=1, cols=1).cell(0, 0).text = 'second section'

    def summarize(items):
        result = []
        for item in items:
            if hasattr(item, 'runs'):
                result.append(('paragraph', item.text))
            else:
                result.append(('table', tuple((tuple((cell.text for cell in row.cells)) for row in item.rows))))
        return result
    py_assert2 = reference.sections
    py_assert4 = len(py_assert2)
    py_assert10 = candidate.sections
    py_assert12 = len(py_assert10)
    py_assert6 = py_assert4 == py_assert12
    py_assert14 = 2
    py_assert7 = py_assert12 == py_assert14
    assert not (not py_assert6 or not py_assert7)
    py_assert2 = py_assert4 = py_assert6 = py_assert7 = py_assert10 = py_assert12 = py_assert14 = None
    for ref_section, our_section in zip(reference.sections, candidate.sections):
        py_assert2 = ref_section.iter_inner_content
        py_assert4 = py_assert2()
        py_assert6 = summarize(py_assert4)
        py_assert11 = our_section.iter_inner_content
        py_assert13 = py_assert11()
        py_assert15 = summarize(py_assert13)
        py_assert8 = py_assert6 == py_assert15
        assert not not py_assert8
        py_assert2 = py_assert4 = py_assert6 = py_assert8 = py_assert11 = py_assert13 = py_assert15 = None
        continue
    ref_path = tmp_path / 'reference-section-content.docx'
    our_path = tmp_path / 'candidate-section-content.docx'
    reference.save(ref_path)
    candidate.save(our_path)
    reopened_reference = docx.Document(str(ref_path))
    reopened_candidate = compat.Document(our_path.read_bytes())
    for ref_section, our_section in zip(reopened_reference.sections, reopened_candidate.sections):
        py_assert2 = ref_section.iter_inner_content
        py_assert4 = py_assert2()
        py_assert6 = summarize(py_assert4)
        py_assert11 = our_section.iter_inner_content
        py_assert13 = py_assert11()
        py_assert15 = summarize(py_assert13)
        py_assert8 = py_assert6 == py_assert15
        assert not not py_assert8
        py_assert2 = py_assert4 = py_assert6 = py_assert8 = py_assert11 = py_assert13 = py_assert15 = None
        continue

def test_add_section_uses_python_docx_start_type_default_and_argument(tmp_path):
    document = compat.Document()
    py_assert1 = document.add_section
    py_assert3 = py_assert1()
    py_assert5 = py_assert3.start_type
    py_assert9 = WD_SECTION_START.NEW_PAGE
    py_assert7 = py_assert5 is py_assert9
    assert not not py_assert7
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = None
    py_assert1 = document.add_section
    py_assert4 = WD_SECTION_START.ODD_PAGE
    py_assert6 = py_assert1(py_assert4)
    py_assert8 = py_assert6.start_type
    py_assert12 = WD_SECTION_START.ODD_PAGE
    py_assert10 = py_assert8 is py_assert12
    assert not not py_assert10
    py_assert1 = py_assert4 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
    sections = document.sections
    reference = docx.Document()
    reference.add_section()
    reference.add_section(docx.enum.section.WD_SECTION_START.ODD_PAGE)
    reference_sections = reference.sections
    py_assert1 = sections.count
    py_assert3 = sections[0]
    py_assert5 = py_assert1(py_assert3)
    py_assert10 = reference_sections.count
    py_assert12 = reference_sections[0]
    py_assert14 = py_assert10(py_assert12)
    py_assert7 = py_assert5 == py_assert14
    py_assert16 = 0
    py_assert8 = py_assert14 == py_assert16
    assert not (not py_assert7 or not py_assert8)
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert8 = py_assert10 = py_assert12 = py_assert14 = py_assert16 = None
    with pytest.raises(ValueError):
        sections.index(sections[0])
    with pytest.raises(ValueError):
        reference_sections.index(reference_sections[0])
    py_assert1 = sections[0:2]
    py_assert3 = len(py_assert1)
    py_assert8 = reference_sections[0:2]
    py_assert10 = len(py_assert8)
    py_assert5 = py_assert3 == py_assert10
    py_assert12 = 2
    py_assert6 = py_assert10 == py_assert12
    assert not (not py_assert5 or not py_assert6)
    py_assert1 = py_assert3 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
    path = tmp_path / 'sections.docx'
    document.save(path)
    sections = docx.Document(str(path)).sections
    py_assert0 = sections[1]
    py_assert2 = py_assert0.start_type
    py_assert6 = docx.enum
    py_assert8 = py_assert6.section
    py_assert10 = py_assert8.WD_SECTION_START
    py_assert12 = py_assert10.NEW_PAGE
    py_assert4 = py_assert2 is py_assert12
    assert not not py_assert4
    py_assert0 = py_assert2 = py_assert4 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
    py_assert0 = sections[2]
    py_assert2 = py_assert0.start_type
    py_assert6 = docx.enum
    py_assert8 = py_assert6.section
    py_assert10 = py_assert8.WD_SECTION_START
    py_assert12 = py_assert10.ODD_PAGE
    py_assert4 = py_assert2 is py_assert12
    assert not not py_assert4
    py_assert0 = py_assert2 = py_assert4 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None

def test_hyperlink_objects_match_python_docx_public_properties():
    body = "<w:p><w:hyperlink w:anchor='section1'><w:r><w:t>linked</w:t><w:lastRenderedPageBreak/></w:r><w:r><w:rPr><w:b/></w:rPr><w:t> text</w:t></w:r></w:hyperlink></w:p>"
    source = make_docx(body)
    ours = compat.Document(source).paragraphs[0]
    theirs = docx.Document(io.BytesIO(source)).paragraphs[0]
    our_link, = ours.hyperlinks
    their_link, = theirs.hyperlinks
    py_assert1 = our_link.address
    py_assert6 = their_link.address
    py_assert3 = py_assert1 == py_assert6
    py_assert8 = ''
    py_assert4 = py_assert6 == py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = our_link.fragment
    py_assert6 = their_link.fragment
    py_assert3 = py_assert1 == py_assert6
    py_assert8 = 'section1'
    py_assert4 = py_assert6 == py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = our_link.url
    py_assert6 = their_link.url
    py_assert3 = py_assert1 == py_assert6
    py_assert8 = ''
    py_assert4 = py_assert6 == py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = our_link.text
    py_assert6 = their_link.text
    py_assert3 = py_assert1 == py_assert6
    py_assert8 = 'linked text'
    py_assert4 = py_assert6 == py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = our_link.contains_page_break
    py_assert6 = their_link.contains_page_break
    py_assert3 = py_assert1 == py_assert6
    py_assert8 = True
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert0 = [run.text for run in our_link.runs]
    py_assert4 = [run.text for run in their_link.runs]
    py_assert2 = py_assert0 == py_assert4
    py_assert6 = ['linked', ' text']
    py_assert3 = py_assert4 == py_assert6
    assert not (not py_assert2 or not py_assert3)
    py_assert0 = py_assert2 = py_assert3 = py_assert4 = py_assert6 = None
    py_assert1 = our_link.paragraph
    py_assert3 = py_assert1 is ours
    assert not not py_assert3
    py_assert1 = py_assert3 = None
    py_assert0 = list(ours.iter_inner_content())[0]
    py_assert2 = py_assert0.text
    py_assert6 = their_link.text
    py_assert4 = py_assert2 == py_assert6
    assert not not py_assert4
    py_assert0 = py_assert2 = py_assert4 = py_assert6 = None

def test_run_inner_content_and_rendered_break_fragments_match_python_docx():
    source = docx.Document()
    paragraph = source.add_paragraph()
    paragraph.add_run('before')
    rendered = paragraph.add_run()
    rendered.add_text('left')
    rendered._r.append(OxmlElement('w:lastRenderedPageBreak'))
    rendered.add_text('right')
    rendered.add_break(docx.enum.text.WD_BREAK.PAGE)
    rendered.add_text('tail')
    hard_only = paragraph.add_run()
    hard_only.add_break(docx.enum.text.WD_BREAK.PAGE)
    payload = io.BytesIO()
    source.save(payload)
    reference = docx.Document(io.BytesIO(payload.getvalue()))
    document = compat.Document(io.BytesIO(payload.getvalue()))
    expected_run = reference.paragraphs[0].runs[1]
    actual_run = document.paragraphs[0].runs[1]
    expected_items = list(expected_run.iter_inner_content())
    actual_items = list(actual_run.iter_inner_content())
    signature = lambda items: [item if isinstance(item, str) else type(item).__name__ for item in items]
    py_assert2 = signature(actual_items)
    py_assert7 = signature(expected_items)
    py_assert4 = py_assert2 == py_assert7
    assert not not py_assert4
    py_assert2 = py_assert4 = py_assert7 = None
    py_assert1 = actual_run.contains_page_break
    py_assert6 = expected_run.contains_page_break
    py_assert3 = py_assert1 is py_assert6
    py_assert8 = True
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert0 = document.paragraphs[0]
    py_assert2 = py_assert0.contains_page_break
    py_assert6 = reference.paragraphs[0]
    py_assert8 = py_assert6.contains_page_break
    py_assert4 = py_assert2 is py_assert8
    py_assert10 = True
    py_assert5 = py_assert8 is py_assert10
    assert not (not py_assert4 or not py_assert5)
    py_assert0 = py_assert2 = py_assert4 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = None
    expected_paragraph_breaks = reference.paragraphs[0].rendered_page_breaks
    actual_paragraph_breaks = document.paragraphs[0].rendered_page_breaks
    py_assert2 = len(expected_paragraph_breaks)
    py_assert8 = len(actual_paragraph_breaks)
    py_assert4 = py_assert2 == py_assert8
    py_assert10 = 1
    py_assert5 = py_assert8 == py_assert10
    assert not (not py_assert4 or not py_assert5)
    py_assert2 = py_assert4 = py_assert5 = py_assert8 = py_assert10 = None
    py_assert0 = actual_paragraph_breaks[0]
    py_assert2 = py_assert0.preceding_paragraph_fragment
    py_assert4 = py_assert2.text
    py_assert7 = expected_paragraph_breaks[0]
    py_assert9 = py_assert7.preceding_paragraph_fragment
    py_assert11 = py_assert9.text
    py_assert6 = py_assert4 == py_assert11
    assert not not py_assert6
    py_assert0 = py_assert2 = py_assert4 = py_assert6 = py_assert7 = py_assert9 = py_assert11 = None
    py_assert0 = actual_paragraph_breaks[0]
    py_assert2 = py_assert0.following_paragraph_fragment
    py_assert4 = py_assert2.text
    py_assert7 = expected_paragraph_breaks[0]
    py_assert9 = py_assert7.following_paragraph_fragment
    py_assert11 = py_assert9.text
    py_assert6 = py_assert4 == py_assert11
    assert not not py_assert6
    py_assert0 = py_assert2 = py_assert4 = py_assert6 = py_assert7 = py_assert9 = py_assert11 = None
    rendered_break = next((item for item in actual_items if type(item).__name__ == 'RenderedPageBreak'))
    expected_break = next((item for item in expected_items if type(item).__name__ == 'RenderedPageBreak'))
    py_assert1 = rendered_break.preceding_paragraph_fragment
    py_assert3 = py_assert1.text
    py_assert8 = expected_break.preceding_paragraph_fragment
    py_assert10 = py_assert8.text
    py_assert5 = py_assert3 == py_assert10
    py_assert12 = 'beforeleft'
    py_assert6 = py_assert10 == py_assert12
    assert not (not py_assert5 or not py_assert6)
    py_assert1 = py_assert3 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
    py_assert1 = rendered_break.following_paragraph_fragment
    py_assert3 = py_assert1.text
    py_assert8 = expected_break.following_paragraph_fragment
    py_assert10 = py_assert8.text
    py_assert5 = py_assert3 == py_assert10
    py_assert12 = 'righttail'
    py_assert6 = py_assert10 == py_assert12
    assert not (not py_assert5 or not py_assert6)
    py_assert1 = py_assert3 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
    actual_hard_only = document.paragraphs[0].runs[2]
    expected_hard_only = reference.paragraphs[0].runs[2]
    py_assert1 = actual_hard_only.contains_page_break
    py_assert6 = expected_hard_only.contains_page_break
    py_assert3 = py_assert1 is py_assert6
    py_assert8 = False
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    picture_source = docx.Document()
    picture_source.add_paragraph().add_run().add_picture(io.BytesIO(blobs.PNG_96))
    picture_payload = io.BytesIO()
    picture_source.save(picture_payload)
    picture_reference = docx.Document(io.BytesIO(picture_payload.getvalue()))
    picture_document = compat.Document(io.BytesIO(picture_payload.getvalue()))
    expected_drawing = list(picture_reference.paragraphs[0].runs[0].iter_inner_content())
    actual_drawing = list(picture_document.paragraphs[0].runs[0].iter_inner_content())
    py_assert0 = [type(item).__name__ for item in actual_drawing]
    py_assert4 = [type(item).__name__ for item in expected_drawing]
    py_assert2 = py_assert0 == py_assert4
    py_assert6 = ['Drawing']
    py_assert3 = py_assert4 == py_assert6
    assert not (not py_assert2 or not py_assert3)
    py_assert0 = py_assert2 = py_assert3 = py_assert4 = py_assert6 = None
    py_assert0 = actual_drawing[0]
    py_assert2 = py_assert0._drawing
    py_assert4 = py_assert2.tag
    py_assert8 = expected_drawing[0]
    py_assert10 = py_assert8._drawing
    py_assert12 = py_assert10.tag
    py_assert14 = str(py_assert12)
    py_assert6 = py_assert4 == py_assert14
    assert not not py_assert6
    py_assert0 = py_assert2 = py_assert4 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = py_assert14 = None
    py_assert0 = actual_drawing[0]
    py_assert2 = py_assert0.has_picture
    py_assert6 = expected_drawing[0]
    py_assert8 = py_assert6.has_picture
    py_assert4 = py_assert2 is py_assert8
    py_assert10 = True
    py_assert5 = py_assert8 is py_assert10
    assert not (not py_assert4 or not py_assert5)
    py_assert0 = py_assert2 = py_assert4 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = None
    actual_image = actual_drawing[0].image
    expected_image = expected_drawing[0].image
    for attribute in ('blob', 'content_type', 'ext', 'filename', 'px_width', 'px_height', 'horz_dpi', 'vert_dpi', 'width', 'height', 'sha1'):
        py_assert3 = getattr(actual_image, attribute)
        py_assert9 = getattr(expected_image, attribute)
        py_assert5 = py_assert3 == py_assert9
        assert not not py_assert5
        py_assert3 = py_assert5 = py_assert9 = None
        continue
    py_assert1 = actual_image.scaled_dimensions
    py_assert3 = 914400
    py_assert5 = py_assert1(width=py_assert3)
    py_assert9 = expected_image.scaled_dimensions
    py_assert11 = 914400
    py_assert13 = py_assert9(width=py_assert11)
    py_assert7 = py_assert5 == py_assert13
    assert not not py_assert7
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = py_assert11 = py_assert13 = None
    unsaved = compat.Document()
    unsaved.add_picture(io.BytesIO(blobs.PNG_96), width=914400)
    unsaved_drawing = next((item for item in unsaved.paragraphs[0].runs[0].iter_inner_content() if type(item).__name__ == 'Drawing'))
    py_assert1 = unsaved_drawing.image
    py_assert3 = py_assert1.blob
    py_assert7 = blobs.PNG_96
    py_assert5 = py_assert3 == py_assert7
    assert not not py_assert5
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = None

def test_hyperlink_rendered_break_fragments_match_python_docx():
    source = make_docx("<w:p><w:r><w:t>before</w:t></w:r><w:hyperlink w:anchor='fragment'><w:r><w:t>left</w:t><w:lastRenderedPageBreak/><w:t>right</w:t></w:r></w:hyperlink><w:r><w:t>after</w:t></w:r></w:p>")
    reference = docx.Document(io.BytesIO(source))
    document = compat.Document(source)
    expected_run = reference.paragraphs[0].hyperlinks[0].runs[0]
    actual_run = document.paragraphs[0].hyperlinks[0].runs[0]
    expected_break = next((item for item in expected_run.iter_inner_content() if type(item).__name__ == 'RenderedPageBreak'))
    actual_break = next((item for item in actual_run.iter_inner_content() if type(item).__name__ == 'RenderedPageBreak'))
    py_assert1 = actual_break.preceding_paragraph_fragment
    py_assert3 = py_assert1.text
    py_assert8 = expected_break.preceding_paragraph_fragment
    py_assert10 = py_assert8.text
    py_assert5 = py_assert3 == py_assert10
    py_assert12 = 'beforeleftright'
    py_assert6 = py_assert10 == py_assert12
    assert not (not py_assert5 or not py_assert6)
    py_assert1 = py_assert3 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
    py_assert1 = actual_break.following_paragraph_fragment
    py_assert3 = py_assert1.text
    py_assert8 = expected_break.following_paragraph_fragment
    py_assert10 = py_assert8.text
    py_assert5 = py_assert3 == py_assert10
    py_assert12 = 'after'
    py_assert6 = py_assert10 == py_assert12
    assert not (not py_assert5 or not py_assert6)
    py_assert1 = py_assert3 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None

def test_paragraph_and_run_mutation_methods_match_python_docx(tmp_path):
    reference = docx.Document()
    candidate = compat.Document()
    ref_paragraph = reference.add_paragraph(style='Heading 2')
    our_paragraph = candidate.add_paragraph(style='Heading 2')
    ref_paragraph.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    our_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ref_run = ref_paragraph.add_run('before')
    our_run = our_paragraph.add_run('before')
    ref_style = reference.styles.add_style('Run Mutation Style', docx.enum.style.WD_STYLE_TYPE.CHARACTER)
    our_style = candidate.styles.add_style('Run Mutation Style', WD_STYLE_TYPE.CHARACTER)
    ref_run.style = ref_style
    our_run.style = our_style
    ref_styled_run = ref_paragraph.add_run(text=' styled', style=ref_style)
    our_styled_run = our_paragraph.add_run(text=' styled', style=our_style)
    py_assert0 = (ref_styled_run.text, ref_styled_run.style.name)
    py_assert4 = (our_styled_run.text, our_styled_run.style.name)
    py_assert2 = py_assert0 == py_assert4
    py_assert6 = (' styled', 'Run Mutation Style')
    py_assert3 = py_assert4 == py_assert6
    assert not (not py_assert2 or not py_assert3)
    py_assert0 = py_assert2 = py_assert3 = py_assert4 = py_assert6 = None
    ref_empty_run = ref_paragraph.add_run()
    our_empty_run = our_paragraph.add_run()
    py_assert1 = ref_empty_run.text
    py_assert6 = our_empty_run.text
    py_assert3 = py_assert1 == py_assert6
    py_assert8 = ''
    py_assert4 = py_assert6 == py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    ref_run.bold = our_run.bold = True
    ref_run.italic = our_run.italic = True
    ref_run.underline = our_run.underline = True
    py_assert1 = ref_run.style
    py_assert3 = py_assert1.name
    py_assert8 = our_run.style
    py_assert10 = py_assert8.name
    py_assert5 = py_assert3 == py_assert10
    py_assert12 = 'Run Mutation Style'
    py_assert6 = py_assert10 == py_assert12
    assert not (not py_assert5 or not py_assert6)
    py_assert1 = py_assert3 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
    ref_text = ref_run.add_text(' added text ')
    our_text = our_run.add_text(' added text ')
    assert ref_text is not None and our_text is not None
    ref_run.add_tab()
    our_run.add_tab()
    ref_run.add_break()
    our_run.add_break()
    break_names = ('LINE', 'PAGE', 'COLUMN', 'LINE_CLEAR_LEFT', 'LINE_CLEAR_RIGHT', 'LINE_CLEAR_ALL')
    for name in break_names:
        ref_run.add_break(getattr(docx.enum.text.WD_BREAK, name))
        our_run.add_break(getattr(WD_BREAK, name))
    py_assert1 = ref_run.text
    py_assert5 = our_run.text
    py_assert3 = py_assert1 == py_assert5
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert5 = None
    py_assert1 = ref_run.bold
    py_assert6 = our_run.bold
    py_assert3 = py_assert1 is py_assert6
    py_assert8 = True
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    word_ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    def break_snapshot(run):
        element = getattr(run, '_r', getattr(run, '_element', None))
        return [(child.get(f'{{{word_ns}}}type'), child.get(f'{{{word_ns}}}clear')) for child in element if isinstance(child.tag, str) if child.tag.endswith('}br')]
    py_assert2 = break_snapshot(ref_run)
    py_assert7 = break_snapshot(our_run)
    py_assert4 = py_assert2 == py_assert7
    assert not not py_assert4
    py_assert2 = py_assert4 = py_assert7 = None
    py_assert1 = ref_run.clear
    py_assert3 = py_assert1()
    py_assert5 = py_assert3 is ref_run
    assert not not py_assert5
    py_assert1 = py_assert3 = py_assert5 = None
    py_assert1 = our_run.clear
    py_assert3 = py_assert1()
    py_assert5 = py_assert3 is our_run
    assert not not py_assert5
    py_assert1 = py_assert3 = py_assert5 = None
    py_assert1 = ref_run.text
    py_assert6 = our_run.text
    py_assert3 = py_assert1 == py_assert6
    py_assert8 = ''
    py_assert4 = py_assert6 == py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = ref_run.bold
    py_assert6 = our_run.bold
    py_assert3 = py_assert1 is py_assert6
    py_assert8 = True
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = ref_run.italic
    py_assert6 = our_run.italic
    py_assert3 = py_assert1 is py_assert6
    py_assert8 = True
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = ref_run.underline
    py_assert6 = our_run.underline
    py_assert3 = py_assert1 is py_assert6
    py_assert8 = True
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = ref_run.style
    py_assert3 = py_assert1.name
    py_assert8 = our_run.style
    py_assert10 = py_assert8.name
    py_assert5 = py_assert3 == py_assert10
    py_assert12 = 'Run Mutation Style'
    py_assert6 = py_assert10 == py_assert12
    assert not (not py_assert5 or not py_assert6)
    py_assert1 = py_assert3 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
    ref_paragraph.add_run('remove this')
    our_paragraph.add_run('remove this')
    py_assert1 = ref_paragraph.clear
    py_assert3 = py_assert1()
    py_assert5 = py_assert3 is ref_paragraph
    assert not not py_assert5
    py_assert1 = py_assert3 = py_assert5 = None
    py_assert1 = our_paragraph.clear
    py_assert3 = py_assert1()
    py_assert5 = py_assert3 is our_paragraph
    assert not not py_assert5
    py_assert1 = py_assert3 = py_assert5 = None
    py_assert1 = ref_paragraph.text
    py_assert6 = our_paragraph.text
    py_assert3 = py_assert1 == py_assert6
    py_assert8 = ''
    py_assert4 = py_assert6 == py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = ref_paragraph.style
    py_assert3 = py_assert1.style_id
    py_assert7 = our_paragraph.style
    py_assert9 = py_assert7.style_id
    py_assert5 = py_assert3 == py_assert9
    assert not not py_assert5
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = None
    py_assert1 = ref_paragraph.alignment
    py_assert5 = docx.enum
    py_assert7 = py_assert5.text
    py_assert9 = py_assert7.WD_ALIGN_PARAGRAPH
    py_assert11 = py_assert9.CENTER
    py_assert3 = py_assert1 is py_assert11
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = py_assert11 = None
    py_assert1 = our_paragraph.alignment
    py_assert5 = WD_ALIGN_PARAGRAPH.CENTER
    py_assert3 = py_assert1 is py_assert5
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert5 = None
    ref_anchor = reference.add_paragraph('following')
    our_anchor = candidate.add_paragraph('following')
    ref_inserted = ref_anchor.insert_paragraph_before('inserted', style='Subtitle')
    our_inserted = our_anchor.insert_paragraph_before('inserted', style='Subtitle')
    py_assert0 = (ref_inserted.text, ref_inserted.style.style_id)
    py_assert3 = (our_inserted.text, our_inserted.style.style_id)
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None
    ref_blank = ref_anchor.insert_paragraph_before()
    our_blank = our_anchor.insert_paragraph_before()
    py_assert0 = (ref_blank.text, ref_blank.style.style_id)
    py_assert3 = (our_blank.text, our_blank.style.style_id)
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None
    py_assert0 = [p.text for p in reference.paragraphs]
    py_assert3 = [p.text for p in candidate.paragraphs]
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None
    ref_path = tmp_path / 'reference-paragraph-run-mutators.docx'
    our_path = tmp_path / 'candidate-paragraph-run-mutators.docx'
    reference.save(ref_path)
    candidate.save(our_path)
    reopened_reference = docx.Document(str(ref_path))
    reopened_candidate = compat.Document(our_path.read_bytes())
    py_assert0 = [(p.text, p.style.style_id, p.alignment) for p in reopened_reference.paragraphs]
    py_assert3 = [(p.text, p.style.style_id, p.alignment) for p in reopened_candidate.paragraphs]
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None

def test_paragraph_spacing_and_tab_stops_match_python_docx(tmp_path):
    document = compat.Document()
    paragraph = document.add_paragraph('tabbed')
    paragraph_format = paragraph.paragraph_format
    paragraph_format.line_spacing_rule = WD_LINE_SPACING.DOUBLE
    py_assert1 = paragraph_format.line_spacing_rule
    py_assert5 = WD_LINE_SPACING.DOUBLE
    py_assert3 = py_assert1 is py_assert5
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert5 = None
    py_assert1 = paragraph_format.line_spacing
    py_assert4 = 2.0
    py_assert3 = py_assert1 == py_assert4
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert4 = None
    paragraph_format.line_spacing = 1
    py_assert1 = paragraph_format.line_spacing_rule
    py_assert5 = WD_LINE_SPACING.SINGLE
    py_assert3 = py_assert1 is py_assert5
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert5 = None
    py_assert1 = paragraph_format.line_spacing
    py_assert4 = 1.0
    py_assert3 = py_assert1 == py_assert4
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert4 = None
    paragraph_format.line_spacing = Pt(14)
    py_assert1 = paragraph_format.line_spacing_rule
    py_assert5 = WD_LINE_SPACING.EXACTLY
    py_assert3 = py_assert1 is py_assert5
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert5 = None
    py_assert1 = paragraph_format.line_spacing
    py_assert5 = 14
    py_assert7 = Pt(py_assert5)
    py_assert3 = py_assert1 == py_assert7
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = None
    for rule in (WD_LINE_SPACING.AT_LEAST, WD_LINE_SPACING.MULTIPLE, WD_LINE_SPACING.EXACTLY):
        paragraph_format.line_spacing_rule = rule
        py_assert1 = paragraph_format.line_spacing_rule
        py_assert3 = py_assert1 is rule
        assert not not py_assert3
        py_assert1 = py_assert3 = None
        continue
    paragraph_format.line_spacing = Pt(14)
    stops = paragraph_format.tab_stops
    right = stops.add_tab_stop(Inches(3), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
    stops.add_tab_stop(Inches(1))
    py_assert2 = len(stops)
    py_assert5 = 2
    py_assert4 = py_assert2 == py_assert5
    assert not not py_assert4
    py_assert2 = py_assert4 = py_assert5 = None
    py_assert0 = [stop.position for stop in stops]
    py_assert3 = [Inches(1), Inches(3)]
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None
    py_assert0 = stops[0]
    py_assert2 = py_assert0.alignment
    py_assert6 = WD_TAB_ALIGNMENT.LEFT
    py_assert4 = py_assert2 is py_assert6
    assert not not py_assert4
    py_assert0 = py_assert2 = py_assert4 = py_assert6 = None
    py_assert0 = stops[0]
    py_assert2 = py_assert0.leader
    py_assert6 = WD_TAB_LEADER.SPACES
    py_assert4 = py_assert2 is py_assert6
    assert not not py_assert4
    py_assert0 = py_assert2 = py_assert4 = py_assert6 = None
    py_assert1 = right.alignment
    py_assert5 = WD_TAB_ALIGNMENT.RIGHT
    py_assert3 = py_assert1 is py_assert5
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert5 = None
    py_assert1 = right.leader
    py_assert5 = WD_TAB_LEADER.DOTS
    py_assert3 = py_assert1 is py_assert5
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert5 = None
    right.position = Inches(0.5)
    right.leader = None
    py_assert0 = [stop.position for stop in stops]
    py_assert3 = [Inches(0.5), Inches(1)]
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None
    py_assert0 = stops[0]
    py_assert2 = py_assert0.alignment
    py_assert6 = WD_TAB_ALIGNMENT.RIGHT
    py_assert4 = py_assert2 is py_assert6
    assert not not py_assert4
    py_assert0 = py_assert2 = py_assert4 = py_assert6 = None
    py_assert1 = right.leader
    py_assert5 = WD_TAB_LEADER.SPACES
    py_assert3 = py_assert1 is py_assert5
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert5 = None
    path = tmp_path / 'paragraph-format.docx'
    document.save(path)
    reopened = docx.Document(str(path))
    saved_format = reopened.paragraphs[0].paragraph_format
    py_assert1 = saved_format.line_spacing_rule
    py_assert5 = docx.enum
    py_assert7 = py_assert5.text
    py_assert9 = py_assert7.WD_LINE_SPACING
    py_assert11 = py_assert9.EXACTLY
    py_assert3 = py_assert1 is py_assert11
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = py_assert11 = None
    py_assert1 = saved_format.line_spacing
    py_assert5 = 14
    py_assert7 = Pt(py_assert5)
    py_assert3 = py_assert1 == py_assert7
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = None
    saved_stops = saved_format.tab_stops
    py_assert0 = [stop.position for stop in saved_stops]
    py_assert3 = [Inches(0.5), Inches(1)]
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None
    py_assert0 = saved_stops[0]
    py_assert2 = py_assert0.alignment
    py_assert6 = docx.enum
    py_assert8 = py_assert6.text
    py_assert10 = py_assert8.WD_TAB_ALIGNMENT
    py_assert12 = py_assert10.RIGHT
    py_assert4 = py_assert2 is py_assert12
    assert not not py_assert4
    py_assert0 = py_assert2 = py_assert4 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
    py_assert0 = saved_stops[0]
    py_assert2 = py_assert0.leader
    py_assert6 = docx.enum
    py_assert8 = py_assert6.text
    py_assert10 = py_assert8.WD_TAB_LEADER
    py_assert12 = py_assert10.SPACES
    py_assert4 = py_assert2 is py_assert12
    assert not not py_assert4
    py_assert0 = py_assert2 = py_assert4 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
    stops.clear_all()
    py_assert2 = len(stops)
    py_assert5 = 0
    py_assert4 = py_assert2 == py_assert5
    assert not not py_assert4
    py_assert2 = py_assert4 = py_assert5 = None
    with pytest.raises(IndexError, match='TabStops object is empty'):
        stops[0]
    stops.add_tab_stop(Inches(1))
    del stops[0]
    py_assert2 = len(stops)
    py_assert5 = 0
    py_assert4 = py_assert2 == py_assert5
    assert not not py_assert4
    py_assert2 = py_assert4 = py_assert5 = None

def test_paragraph_format_properties_match_python_docx_and_round_trip(tmp_path):
    reference = docx.Document()
    candidate = compat.Document()
    ref_paragraph = reference.add_paragraph('formatted')
    our_paragraph = candidate.add_paragraph('formatted')
    ref_format = ref_paragraph.paragraph_format
    our_format = our_paragraph.paragraph_format

    def normalize(value):
        if value is None:
            return
        if hasattr(value, 'name'):
            return value.name
        if isinstance(value, int):
            return int(value)
        return value
    properties = ('alignment', 'first_line_indent', 'keep_together', 'keep_with_next', 'left_indent', 'line_spacing', 'line_spacing_rule', 'page_break_before', 'right_indent', 'space_after', 'space_before', 'widow_control')

    def assert_formats_match(ref_format, candidate_format):
        py_assert0 = {name: normalize(getattr(ref_format, name)) for name in properties}
        py_assert3 = {name: normalize(getattr(candidate_format, name)) for name in properties}
        py_assert2 = py_assert0 == py_assert3
        assert not not py_assert2
        py_assert0 = py_assert2 = py_assert3 = None
    assert_formats_match(ref_format, our_format)
    ref_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    our_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ref_format.first_line_indent = docx.shared.Inches(-0.25)
    our_format.first_line_indent = Inches(-0.25)
    ref_format.left_indent = docx.shared.Inches(-0.25)
    our_format.left_indent = Inches(-0.25)
    assert_formats_match(ref_format, our_format)
    ref_format.left_indent = docx.shared.Inches(0.5)
    our_format.left_indent = Inches(0.5)
    ref_format.right_indent = docx.shared.Inches(0.25)
    our_format.right_indent = Inches(0.25)
    ref_format.space_before = docx.shared.Pt(8)
    our_format.space_before = Pt(8)
    ref_format.space_after = docx.shared.Pt(6)
    our_format.space_after = Pt(6)
    ref_format.keep_together = True
    our_format.keep_together = True
    ref_format.keep_with_next = False
    our_format.keep_with_next = False
    ref_format.page_break_before = True
    our_format.page_break_before = True
    ref_format.widow_control = False
    our_format.widow_control = False
    assert_formats_match(ref_format, our_format)
    for ref_value, our_value in ((1, 1), (1.5, 1.5), (docx.shared.Pt(14), Pt(14)), (None, None)):
        ref_format.line_spacing = ref_value
        our_format.line_spacing = our_value
        assert_formats_match(ref_format, our_format)
    for name in ('SINGLE', 'ONE_POINT_FIVE', 'DOUBLE', 'AT_LEAST', 'MULTIPLE', 'EXACTLY'):
        ref_format.line_spacing_rule = getattr(docx.enum.text.WD_LINE_SPACING, name)
        our_format.line_spacing_rule = getattr(WD_LINE_SPACING, name)
        assert_formats_match(ref_format, our_format)
    for name in ('keep_together', 'keep_with_next', 'page_break_before', 'widow_control'):
        setattr(ref_format, name, None)
        setattr(our_format, name, None)
        assert_formats_match(ref_format, our_format)
        setattr(ref_format, name, False)
        setattr(our_format, name, False)
        assert_formats_match(ref_format, our_format)
    for name in ('alignment', 'first_line_indent', 'left_indent', 'right_indent', 'space_after', 'space_before'):
        setattr(ref_format, name, None)
        setattr(our_format, name, None)
        assert_formats_match(ref_format, our_format)
    ref_stops = ref_format.tab_stops
    our_stops = our_format.tab_stops
    ref_right = ref_stops.add_tab_stop(docx.shared.Inches(3), docx.enum.text.WD_TAB_ALIGNMENT.RIGHT, docx.enum.text.WD_TAB_LEADER.DOTS)
    our_right = our_stops.add_tab_stop(Inches(3), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
    ref_stops.add_tab_stop(docx.shared.Inches(1))
    our_stops.add_tab_stop(Inches(1))

    def stop_snapshot(stops):
        return [(int(stop.position), stop.alignment.name, stop.leader.name) for stop in stops]
    py_assert2 = stop_snapshot(ref_stops)
    py_assert7 = stop_snapshot(our_stops)
    py_assert4 = py_assert2 == py_assert7
    assert not not py_assert4
    py_assert2 = py_assert4 = py_assert7 = None
    ref_right.position = docx.shared.Inches(0.5)
    our_right.position = Inches(0.5)
    ref_right.leader = None
    our_right.leader = None
    py_assert2 = stop_snapshot(ref_stops)
    py_assert7 = stop_snapshot(our_stops)
    py_assert4 = py_assert2 == py_assert7
    assert not not py_assert4
    py_assert2 = py_assert4 = py_assert7 = None
    py_assert2 = stop_snapshot(ref_stops)
    py_assert5 = [(int(docx.shared.Inches(0.5)), 'RIGHT', 'SPACES'), (int(docx.shared.Inches(1)), 'LEFT', 'SPACES')]
    py_assert4 = py_assert2 == py_assert5
    assert not not py_assert4
    py_assert2 = py_assert4 = py_assert5 = None
    ref_path = tmp_path / 'reference-format-properties.docx'
    our_path = tmp_path / 'candidate-format-properties.docx'
    reference.save(ref_path)
    candidate.save(our_path)
    reopened_reference = docx.Document(str(ref_path)).paragraphs[0]
    reopened_candidate = compat.Document(our_path.read_bytes()).paragraphs[0]
    assert_formats_match(reopened_reference.paragraph_format, reopened_candidate.paragraph_format)
    py_assert2 = reopened_reference.paragraph_format
    py_assert4 = py_assert2.tab_stops
    py_assert6 = stop_snapshot(py_assert4)
    py_assert11 = reopened_candidate.paragraph_format
    py_assert13 = py_assert11.tab_stops
    py_assert15 = stop_snapshot(py_assert13)
    py_assert8 = py_assert6 == py_assert15
    assert not not py_assert8
    py_assert2 = py_assert4 = py_assert6 = py_assert8 = py_assert11 = py_assert13 = py_assert15 = None
    ref_stops.clear_all()
    our_stops.clear_all()
    py_assert2 = len(ref_stops)
    py_assert8 = len(our_stops)
    py_assert4 = py_assert2 == py_assert8
    py_assert10 = 0
    py_assert5 = py_assert8 == py_assert10
    assert not (not py_assert4 or not py_assert5)
    py_assert2 = py_assert4 = py_assert5 = py_assert8 = py_assert10 = None
    ref_stops.add_tab_stop(docx.shared.Inches(1))
    our_stops.add_tab_stop(Inches(1))
    del ref_stops[0], our_stops[0]
    py_assert2 = len(ref_stops)
    py_assert8 = len(our_stops)
    py_assert4 = py_assert2 == py_assert8
    py_assert10 = 0
    py_assert5 = py_assert8 == py_assert10
    assert not (not py_assert4 or not py_assert5)
    py_assert2 = py_assert4 = py_assert5 = py_assert8 = py_assert10 = None
    py_assert1 = lambda: ref_stops[0]
    py_assert3 = _capture_exception(py_assert1)
    py_assert7 = lambda: our_stops[0]
    py_assert9 = _capture_exception(py_assert7)
    py_assert5 = py_assert3 == py_assert9
    assert not not py_assert5
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = None

def test_document_add_comment_matches_python_docx_and_round_trip(tmp_path):
    reference = docx.Document()
    candidate = compat.Document()
    ref_paragraph = reference.add_paragraph()
    our_paragraph = candidate.add_paragraph()
    ref_runs = [ref_paragraph.add_run(text) for text in ('first', ' middle', ' last')]
    our_runs = [our_paragraph.add_run(text) for text in ('first', ' middle', ' last')]
    ref_comment = reference.add_comment([ref_runs[0], ref_runs[-1]], text='Please review', author='Reviewer', initials=None)
    our_comment = candidate.add_comment([our_runs[0], our_runs[-1]], text='Please review', author='Reviewer', initials=None)

    def comment_snapshot(comment):
        return (comment.comment_id, comment.author, comment.initials, comment.text, [paragraph.text for paragraph in comment.paragraphs])
    py_assert2 = comment_snapshot(ref_comment)
    py_assert7 = comment_snapshot(our_comment)
    py_assert4 = py_assert2 == py_assert7
    assert not not py_assert4
    py_assert2 = py_assert4 = py_assert7 = None
    py_assert1 = ref_comment.comment_id
    py_assert6 = our_comment.comment_id
    py_assert3 = py_assert1 == py_assert6
    py_assert8 = 0
    py_assert4 = py_assert6 == py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    ref_single = reference.add_comment(ref_runs[1], text='Single-run comment')
    our_single = candidate.add_comment(our_runs[1], text='Single-run comment')
    py_assert2 = comment_snapshot(ref_single)
    py_assert7 = comment_snapshot(our_single)
    py_assert4 = py_assert2 == py_assert7
    assert not not py_assert4
    py_assert2 = py_assert4 = py_assert7 = None
    py_assert1 = ref_single.comment_id
    py_assert6 = our_single.comment_id
    py_assert3 = py_assert1 == py_assert6
    py_assert8 = 1
    py_assert4 = py_assert6 == py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = ref_single.initials
    py_assert6 = our_single.initials
    py_assert3 = py_assert1 == py_assert6
    py_assert8 = ''
    py_assert4 = py_assert6 == py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None

    class RunSequence:

        def __init__(self, values):
            self.values = values

        def __getitem__(self, index):
            return self.values[index]

        def __len__(self):
            return len(self.values)
    ref_custom = reference.add_comment(RunSequence([ref_runs[0], ref_runs[1]]), text='Custom sequence')
    our_custom = candidate.add_comment(RunSequence([our_runs[0], our_runs[1]]), text='Custom sequence')
    py_assert2 = comment_snapshot(ref_custom)
    py_assert7 = comment_snapshot(our_custom)
    py_assert4 = py_assert2 == py_assert7
    assert not not py_assert4
    py_assert2 = py_assert4 = py_assert7 = None
    py_assert1 = ref_custom.comment_id
    py_assert6 = our_custom.comment_id
    py_assert3 = py_assert1 == py_assert6
    py_assert8 = 2
    py_assert4 = py_assert6 == py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    ref_second_paragraph = reference.add_paragraph()
    our_second_paragraph = candidate.add_paragraph()
    ref_second_runs = [ref_second_paragraph.add_run(text) for text in ('another', ' paragraph')]
    our_second_runs = [our_second_paragraph.add_run(text) for text in ('another', ' paragraph')]
    ref_cross = reference.add_comment([ref_runs[0], ref_second_runs[-1]], text='Cross-paragraph comment')
    our_cross = candidate.add_comment([our_runs[0], our_second_runs[-1]], text='Cross-paragraph comment')
    py_assert2 = comment_snapshot(ref_cross)
    py_assert7 = comment_snapshot(our_cross)
    py_assert4 = py_assert2 == py_assert7
    assert not not py_assert4
    py_assert2 = py_assert4 = py_assert7 = None
    py_assert1 = ref_cross.comment_id
    py_assert6 = our_cross.comment_id
    py_assert3 = py_assert1 == py_assert6
    py_assert8 = 3
    py_assert4 = py_assert6 == py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    for empty_sequence in ([], ()):
        ref_error = _capture_exception(lambda empty_sequence=empty_sequence: reference.add_comment(empty_sequence))
        our_error = _capture_exception(lambda empty_sequence=empty_sequence: candidate.add_comment(empty_sequence))
        py_assert1 = ref_error == our_error
        assert not not py_assert1
        py_assert1 = None
        continue

    def anchor_positions(path):
        with zipfile.ZipFile(path) as package:
            root = ET.fromstring(package.read('word/document.xml'))
        word_ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        body = root.find('.//{%s}body' % word_ns)
        events = []
        for paragraph_index, paragraph in enumerate(body.findall('{%s}p' % word_ns)):
            run_index = 0
            for child in paragraph:
                local = child.tag.rsplit('}', 1)[-1]
                if local in ('commentRangeStart', 'commentRangeEnd'):
                    events.append((paragraph_index, local, child.get('{%s}id' % word_ns), run_index))
                elif local == 'r':
                    run_index += 1
        return events
    ref_path = tmp_path / 'reference-document-comment.docx'
    our_path = tmp_path / 'candidate-document-comment.docx'
    reference.save(ref_path)
    candidate.save(our_path)
    py_assert2 = anchor_positions(ref_path)
    py_assert7 = anchor_positions(our_path)
    py_assert4 = py_assert2 == py_assert7
    assert not not py_assert4
    py_assert2 = py_assert4 = py_assert7 = None
    reopened_reference = docx.Document(str(ref_path))
    reopened_candidate = compat.Document(our_path.read_bytes())
    for comment_id in (0, 1, 2, 3):
        py_assert2 = reopened_reference.comments
        py_assert4 = py_assert2.get
        py_assert7 = py_assert4(comment_id)
        py_assert9 = comment_snapshot(py_assert7)
        py_assert14 = reopened_candidate.comments
        py_assert16 = py_assert14.get
        py_assert19 = py_assert16(comment_id)
        py_assert21 = comment_snapshot(py_assert19)
        py_assert11 = py_assert9 == py_assert21
        assert not not py_assert11
        py_assert2 = py_assert4 = py_assert7 = py_assert9 = py_assert11 = py_assert14 = py_assert16 = py_assert19 = py_assert21 = None
        continue
    py_assert2 = anchor_positions(ref_path)
    py_assert7 = anchor_positions(our_path)
    py_assert4 = py_assert2 == py_assert7
    assert not not py_assert4
    py_assert2 = py_assert4 = py_assert7 = None

def test_comment_range_offsets_inside_hyperlink_match_python_docx(tmp_path):
    source = make_docx("<w:p><w:r><w:t>before </w:t></w:r><w:hyperlink w:anchor='section1'><w:r><w:t>linked</w:t></w:r><w:r><w:t> text</w:t></w:r></w:hyperlink><w:r><w:t> after</w:t></w:r></w:p>")
    reference = docx.Document(io.BytesIO(source))
    candidate = compat.Document(source)
    ref_link = reference.paragraphs[0].hyperlinks[0]
    our_link = candidate.paragraphs[0].hyperlinks[0]
    py_assert0 = [run.text for run in ref_link.runs]
    py_assert4 = [run.text for run in our_link.runs]
    py_assert2 = py_assert0 == py_assert4
    py_assert6 = ['linked', ' text']
    py_assert3 = py_assert4 == py_assert6
    assert not (not py_assert2 or not py_assert3)
    py_assert0 = py_assert2 = py_assert3 = py_assert4 = py_assert6 = None
    ref_comment = reference.add_comment([ref_link.runs[0], ref_link.runs[1]], text='Review linked phrase')
    our_comment = candidate.add_comment([our_link.runs[0], our_link.runs[1]], text='Review linked phrase')
    py_assert1 = ref_comment.comment_id
    py_assert6 = our_comment.comment_id
    py_assert3 = py_assert1 == py_assert6
    py_assert8 = 0
    py_assert4 = py_assert6 == py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    record = candidate.comments.get(0)
    py_assert1 = record.anchor_text
    py_assert4 = 'linked text'
    py_assert3 = py_assert1 == py_assert4
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert4 = None
    py_assert0 = [(anchor.start, anchor.end, anchor.text) for anchor in record.anchors]
    py_assert3 = [(7, 18, 'linked text')]
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None

    def marker_positions(path):
        with zipfile.ZipFile(path) as package:
            root = ET.fromstring(package.read('word/document.xml'))
        word_ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        paragraph = root.find('.//{%s}p' % word_ns)
        events = []
        offset = 0

        def walk(parent):
            nonlocal offset
            for child in parent:
                local = child.tag.rsplit('}', 1)[-1]
                if local in ('commentRangeStart', 'commentRangeEnd'):
                    events.append((local, offset, child.get('{%s}id' % word_ns), parent.tag.rsplit('}', 1)[-1]))
                elif local == 't':
                    offset += len(child.text or '')
                else:
                    walk(child)
        walk(paragraph)
        return events
    ref_path = tmp_path / 'reference-hyperlink-comment.docx'
    our_path = tmp_path / 'candidate-hyperlink-comment.docx'
    reference.save(ref_path)
    candidate.save(our_path)
    expected = [('commentRangeStart', 7, '0', 'hyperlink'), ('commentRangeEnd', 18, '0', 'hyperlink')]
    py_assert2 = marker_positions(ref_path)
    py_assert8 = marker_positions(our_path)
    py_assert4 = py_assert2 == py_assert8
    py_assert5 = py_assert8 == expected
    assert not (not py_assert4 or not py_assert5)
    py_assert2 = py_assert4 = py_assert5 = py_assert8 = None
    reopened = compat.Document(our_path.read_bytes()).comments.get(0)
    py_assert1 = reopened.anchor_text
    py_assert4 = 'linked text'
    py_assert3 = py_assert1 == py_assert4
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert4 = None
    py_assert0 = [(anchor.start, anchor.end, anchor.text) for anchor in reopened.anchors]
    py_assert3 = [(7, 18, 'linked text')]
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None

def test_comments_collection_and_rich_comment_api_match_python_docx(tmp_path):
    reference = docx.Document()
    candidate = compat.Document()
    ref_run = reference.add_paragraph('review this').runs[0]
    our_run = candidate.add_paragraph('review this').runs[0]
    ref_comments = reference.comments
    our_comments = candidate.comments
    ref_comment = ref_comments.add_comment('First line\nSecond line', author='Reviewer', initials='RV')
    our_comment = our_comments.add_comment('First line\nSecond line', author='Reviewer', initials='RV')
    py_assert2 = len(ref_comments)
    py_assert8 = len(our_comments)
    py_assert4 = py_assert2 == py_assert8
    py_assert10 = 1
    py_assert5 = py_assert8 == py_assert10
    assert not (not py_assert4 or not py_assert5)
    py_assert2 = py_assert4 = py_assert5 = py_assert8 = py_assert10 = None
    py_assert1 = ref_comments.get
    py_assert3 = 999
    py_assert5 = py_assert1(py_assert3)
    py_assert10 = our_comments.get
    py_assert12 = 999
    py_assert14 = py_assert10(py_assert12)
    py_assert7 = py_assert5 is py_assert14
    py_assert16 = None
    py_assert8 = py_assert14 is py_assert16
    assert not (not py_assert7 or not py_assert8)
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert8 = py_assert10 = py_assert12 = py_assert14 = py_assert16 = None

    def comment_snapshot(comment):
        return (comment.comment_id, comment.author, comment.initials, comment.text, [paragraph.text for paragraph in comment.paragraphs], [(len(table.rows), len(table.columns), [[cell.text for cell in row.cells] for row in table.rows]) for table in comment.tables])
    py_assert2 = ref_comments.get
    py_assert4 = 0
    py_assert6 = py_assert2(py_assert4)
    py_assert8 = comment_snapshot(py_assert6)
    py_assert13 = our_comments.get
    py_assert15 = 0
    py_assert17 = py_assert13(py_assert15)
    py_assert19 = comment_snapshot(py_assert17)
    py_assert10 = py_assert8 == py_assert19
    assert not not py_assert10
    py_assert2 = py_assert4 = py_assert6 = py_assert8 = py_assert10 = py_assert13 = py_assert15 = py_assert17 = py_assert19 = None
    py_assert1 = ref_comment.comment_id
    py_assert6 = our_comment.comment_id
    py_assert3 = py_assert1 == py_assert6
    py_assert8 = 0
    py_assert4 = py_assert6 == py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = ref_comment.timestamp
    py_assert3 = py_assert1.tzinfo
    py_assert6 = None
    py_assert5 = py_assert3 is not py_assert6
    assert not not py_assert5
    py_assert1 = py_assert3 = py_assert5 = py_assert6 = None
    py_assert1 = our_comment.timestamp
    py_assert3 = py_assert1.tzinfo
    py_assert6 = None
    py_assert5 = py_assert3 is not py_assert6
    assert not not py_assert5
    py_assert1 = py_assert3 = py_assert5 = py_assert6 = None
    ref_run.mark_comment_range(ref_run, ref_comment.comment_id)
    our_run.mark_comment_range(our_run, our_comment.comment_id)
    py_assert1 = our_comments.get
    py_assert3 = 0
    py_assert5 = py_assert1(py_assert3)
    py_assert7 = py_assert5.anchor_text
    py_assert10 = 'review this'
    py_assert9 = py_assert7 == py_assert10
    assert not not py_assert9
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = py_assert10 = None
    ref_comment.initials = None
    our_comment.initials = None
    py_assert1 = ref_comment.initials
    py_assert6 = our_comment.initials
    py_assert3 = py_assert1 is py_assert6
    py_assert8 = None
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    ref_comment.paragraphs[0].add_run(' emphasized').bold = True
    our_comment.paragraphs[0].add_run(' emphasized').bold = True
    ref_comment.add_paragraph('A third paragraph')
    our_comment.add_paragraph('A third paragraph')
    ref_table = ref_comment.add_table(rows=1, cols=1, width=docx.shared.Inches(2))
    our_table = our_comment.add_table(rows=1, cols=1, width=Inches(2))
    ref_table.cell(0, 0).text = 'comment table'
    our_table.cell(0, 0).text = 'comment table'

    def inner_content_snapshot(comment):
        snapshot = []
        for item in comment.iter_inner_content():
            if hasattr(item, 'runs'):
                snapshot.append(('paragraph', item.text))
            else:
                snapshot.append(('table', len(item.rows), len(item.columns), [[cell.text for cell in row.cells] for row in item.rows]))
        return snapshot
    py_assert2 = inner_content_snapshot(ref_comment)
    py_assert7 = inner_content_snapshot(our_comment)
    py_assert4 = py_assert2 == py_assert7
    assert not not py_assert4
    py_assert2 = py_assert4 = py_assert7 = None
    py_assert2 = comment_snapshot(ref_comment)
    py_assert7 = comment_snapshot(our_comment)
    py_assert4 = py_assert2 == py_assert7
    assert not not py_assert4
    py_assert2 = py_assert4 = py_assert7 = None
    ref_path = tmp_path / 'reference-comments-collection.docx'
    our_path = tmp_path / 'candidate-comments-collection.docx'
    reference.save(ref_path)
    candidate.save(our_path)
    reopened_reference = docx.Document(str(ref_path))
    reopened_candidate = compat.Document(our_path.read_bytes())
    py_assert2 = reopened_reference.comments
    py_assert4 = len(py_assert2)
    py_assert10 = reopened_candidate.comments
    py_assert12 = len(py_assert10)
    py_assert6 = py_assert4 == py_assert12
    py_assert14 = 1
    py_assert7 = py_assert12 == py_assert14
    assert not (not py_assert6 or not py_assert7)
    py_assert2 = py_assert4 = py_assert6 = py_assert7 = py_assert10 = py_assert12 = py_assert14 = None
    py_assert2 = reopened_reference.comments
    py_assert4 = py_assert2.get
    py_assert6 = 0
    py_assert8 = py_assert4(py_assert6)
    py_assert10 = comment_snapshot(py_assert8)
    py_assert15 = reopened_candidate.comments
    py_assert17 = py_assert15.get
    py_assert19 = 0
    py_assert21 = py_assert17(py_assert19)
    py_assert23 = comment_snapshot(py_assert21)
    py_assert12 = py_assert10 == py_assert23
    assert not not py_assert12
    py_assert2 = py_assert4 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = py_assert15 = py_assert17 = py_assert19 = py_assert21 = py_assert23 = None
    py_assert0 = reopened_reference.paragraphs[0].runs[0]
    py_assert2 = py_assert0.text
    py_assert6 = reopened_candidate.paragraphs[0].runs[0]
    py_assert8 = py_assert6.text
    py_assert4 = py_assert2 == py_assert8
    py_assert10 = 'review this'
    py_assert5 = py_assert8 == py_assert10
    assert not (not py_assert4 or not py_assert5)
    py_assert0 = py_assert2 = py_assert4 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = None
    py_assert2 = reopened_reference.comments
    py_assert4 = py_assert2.get
    py_assert6 = 0
    py_assert8 = py_assert4(py_assert6)
    py_assert10 = py_assert8.tables
    py_assert12 = len(py_assert10)
    py_assert18 = reopened_candidate.comments
    py_assert20 = py_assert18.get
    py_assert22 = 0
    py_assert24 = py_assert20(py_assert22)
    py_assert26 = py_assert24.tables
    py_assert28 = len(py_assert26)
    py_assert14 = py_assert12 == py_assert28
    py_assert30 = 1
    py_assert15 = py_assert28 == py_assert30
    assert not (not py_assert14 or not py_assert15)
    py_assert2 = py_assert4 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = py_assert14 = py_assert15 = py_assert18 = py_assert20 = py_assert22 = py_assert24 = py_assert26 = py_assert28 = py_assert30 = None

def test_documented_text_writer_controls_match_python_docx_and_round_trip(tmp_path):
    reference = docx.Document()
    candidate = compat.Document()
    ref_empty = reference.add_paragraph()
    our_empty = candidate.add_paragraph()
    py_assert1 = ref_empty.text
    py_assert6 = our_empty.text
    py_assert3 = py_assert1 == py_assert6
    py_assert8 = ''
    py_assert4 = py_assert6 == py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = ref_empty.runs
    py_assert6 = our_empty.runs
    py_assert3 = py_assert1 == py_assert6
    py_assert8 = []
    py_assert4 = py_assert6 == py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    text_values = ('alpha\tbeta\ngamma\rdelta', ' leading text \n trailing text ', 'first\r\nsecond')
    for index, text in enumerate(text_values):
        ref_style = reference.styles['Heading 2']
        our_style = candidate.styles['Heading 2']
        ref_paragraph = reference.add_paragraph(text, style=ref_style)
        our_paragraph = candidate.add_paragraph(text, style=our_style)
        py_assert0 = (ref_paragraph.text, ref_paragraph.style.style_id)
        py_assert3 = (our_paragraph.text, our_paragraph.style.style_id)
        py_assert2 = py_assert0 == py_assert3
        assert not not py_assert2
        py_assert0 = py_assert2 = py_assert3 = None
        ref_paragraph.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.RIGHT
        our_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        ref_paragraph.runs[0].bold = True
        our_paragraph.runs[0].bold = True
        ref_paragraph.text = text
        our_paragraph.text = text
        py_assert1 = ref_paragraph.text
        py_assert6 = our_paragraph.text
        py_assert3 = py_assert1 == py_assert6
        py_assert9 = text.replace
        py_assert11 = '\r'
        py_assert13 = '\n'
        py_assert15 = py_assert9(py_assert11, py_assert13)
        py_assert4 = py_assert6 == py_assert15
        assert not (not py_assert3 or not py_assert4)
        py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert9 = py_assert11 = py_assert13 = py_assert15 = None
        py_assert2 = ref_paragraph.runs
        py_assert4 = len(py_assert2)
        py_assert10 = our_paragraph.runs
        py_assert12 = len(py_assert10)
        py_assert6 = py_assert4 == py_assert12
        py_assert14 = 1
        py_assert7 = py_assert12 == py_assert14
        assert not (not py_assert6 or not py_assert7)
        py_assert2 = py_assert4 = py_assert6 = py_assert7 = py_assert10 = py_assert12 = py_assert14 = None
        py_assert0 = ref_paragraph.runs[0]
        py_assert2 = py_assert0.bold
        py_assert6 = our_paragraph.runs[0]
        py_assert8 = py_assert6.bold
        py_assert4 = py_assert2 is py_assert8
        py_assert10 = None
        py_assert5 = py_assert8 is py_assert10
        assert not (not py_assert4 or not py_assert5)
        py_assert0 = py_assert2 = py_assert4 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = None
        py_assert1 = ref_paragraph.style
        py_assert3 = py_assert1.style_id
        py_assert7 = our_paragraph.style
        py_assert9 = py_assert7.style_id
        py_assert5 = py_assert3 == py_assert9
        assert not not py_assert5
        py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = None
        py_assert1 = ref_paragraph.alignment
        py_assert3 = py_assert1.name
        py_assert8 = our_paragraph.alignment
        py_assert10 = py_assert8.name
        py_assert5 = py_assert3 == py_assert10
        py_assert12 = 'RIGHT'
        py_assert6 = py_assert10 == py_assert12
        assert not (not py_assert5 or not py_assert6)
        py_assert1 = py_assert3 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
        ref_run = ref_paragraph.add_run(text)
        our_run = our_paragraph.add_run(text)
        ref_run.bold = our_run.bold = True
        ref_run.text = text
        our_run.text = text
        py_assert1 = ref_run.text
        py_assert6 = our_run.text
        py_assert3 = py_assert1 == py_assert6
        py_assert9 = text.replace
        py_assert11 = '\r'
        py_assert13 = '\n'
        py_assert15 = py_assert9(py_assert11, py_assert13)
        py_assert4 = py_assert6 == py_assert15
        assert not (not py_assert3 or not py_assert4)
        py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert9 = py_assert11 = py_assert13 = py_assert15 = None
        py_assert1 = ref_run.bold
        py_assert6 = our_run.bold
        py_assert3 = py_assert1 is py_assert6
        py_assert8 = True
        py_assert4 = py_assert6 is py_assert8
        assert not (not py_assert3 or not py_assert4)
        py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
        ref_anchor = reference.add_paragraph(f'anchor {index}')
        our_anchor = candidate.add_paragraph(f'anchor {index}')
        ref_inserted = ref_anchor.insert_paragraph_before(text, style=reference.styles['Title'])
        our_inserted = our_anchor.insert_paragraph_before(text, style=candidate.styles['Title'])
        py_assert0 = (ref_inserted.text, ref_inserted.style.style_id)
        py_assert3 = (our_inserted.text, our_inserted.style.style_id)
        py_assert2 = py_assert0 == py_assert3
        assert not not py_assert2
        py_assert0 = py_assert2 = py_assert3 = None
        py_assert0 = (ref_inserted.text, ref_inserted.style.style_id)
        py_assert3 = (text.replace('\r', '\n'), 'Title')
        py_assert2 = py_assert0 == py_assert3
        assert not not py_assert2
        py_assert0 = py_assert2 = py_assert3 = None
        continue
    ref_table = reference.add_table(rows=1, cols=1)
    our_table = candidate.add_table(rows=1, cols=1)
    ref_cell, our_cell = (ref_table.cell(0, 0), our_table.cell(0, 0))
    for text in text_values:
        ref_cell.text = text
        our_cell.text = text
        py_assert1 = ref_cell.text
        py_assert6 = our_cell.text
        py_assert3 = py_assert1 == py_assert6
        py_assert9 = text.replace
        py_assert11 = '\r'
        py_assert13 = '\n'
        py_assert15 = py_assert9(py_assert11, py_assert13)
        py_assert4 = py_assert6 == py_assert15
        assert not (not py_assert3 or not py_assert4)
        py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert9 = py_assert11 = py_assert13 = py_assert15 = None
        py_assert2 = ref_cell.paragraphs
        py_assert4 = len(py_assert2)
        py_assert10 = our_cell.paragraphs
        py_assert12 = len(py_assert10)
        py_assert6 = py_assert4 == py_assert12
        py_assert14 = 1
        py_assert7 = py_assert12 == py_assert14
        assert not (not py_assert6 or not py_assert7)
        py_assert2 = py_assert4 = py_assert6 = py_assert7 = py_assert10 = py_assert12 = py_assert14 = None
        ref_added = ref_cell.add_paragraph(text, style=reference.styles['Heading 3'])
        our_added = our_cell.add_paragraph(text, style=candidate.styles['Heading 3'])
        py_assert0 = (ref_added.text, ref_added.style.style_id)
        py_assert3 = (our_added.text, our_added.style.style_id)
        py_assert2 = py_assert0 == py_assert3
        assert not not py_assert2
        py_assert0 = py_assert2 = py_assert3 = None
        py_assert1 = ref_added.text
        py_assert5 = text.replace
        py_assert7 = '\r'
        py_assert9 = '\n'
        py_assert11 = py_assert5(py_assert7, py_assert9)
        py_assert3 = py_assert1 == py_assert11
        assert not not py_assert3
        py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = py_assert11 = None
        continue
    for kind in ('header', 'footer'):
        ref_story = getattr(reference.sections[0], kind)
        our_story = getattr(candidate.sections[0], kind)
        style_name = 'Header' if kind == 'header' else 'Footer'
        ref_added = ref_story.add_paragraph(text_values[0], style=reference.styles[style_name])
        our_added = our_story.add_paragraph(text_values[0], style=candidate.styles[style_name])
        py_assert0 = (ref_added.text, ref_added.style.style_id)
        py_assert3 = (our_added.text, our_added.style.style_id)
        py_assert2 = py_assert0 == py_assert3
        assert not not py_assert2
        py_assert0 = py_assert2 = py_assert3 = None
        py_assert1 = ref_added.text
        py_assert4 = text_values[0]
        py_assert6 = py_assert4.replace
        py_assert8 = '\r'
        py_assert10 = '\n'
        py_assert12 = py_assert6(py_assert8, py_assert10)
        py_assert3 = py_assert1 == py_assert12
        assert not not py_assert3
        py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
        continue

    def snapshot(document):
        return {'paragraphs': [(p.text, p.style.style_id, p.alignment.name if p.alignment is not None else None) for p in document.paragraphs], 'cell': (ref_cell.text if document is reference else our_cell.text, [(p.text, p.style.style_id) for p in (ref_cell.paragraphs if document is reference else our_cell.paragraphs)]), 'stories': {kind: [(p.text, p.style.style_id) for p in getattr(document.sections[0], kind).paragraphs] for kind in ('header', 'footer')}}
    py_assert2 = snapshot(reference)
    py_assert7 = snapshot(candidate)
    py_assert4 = py_assert2 == py_assert7
    assert not not py_assert4
    py_assert2 = py_assert4 = py_assert7 = None
    ref_path = tmp_path / 'reference-text-controls.docx'
    our_path = tmp_path / 'candidate-text-controls.docx'
    reference.save(ref_path)
    candidate.save(our_path)
    reopened_reference = docx.Document(str(ref_path))
    reopened_candidate = compat.Document(our_path.read_bytes())
    py_assert0 = [(p.text, p.style.style_id) for p in reopened_reference.paragraphs]
    py_assert3 = [(p.text, p.style.style_id) for p in reopened_candidate.paragraphs]
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None
    for kind in ('header', 'footer'):
        py_assert0 = [(p.text, p.style.style_id) for p in getattr(reopened_reference.sections[0], kind).paragraphs]
        py_assert3 = [(p.text, p.style.style_id) for p in getattr(reopened_candidate.sections[0], kind).paragraphs]
        py_assert2 = py_assert0 == py_assert3
        assert not not py_assert2
        py_assert0 = py_assert2 = py_assert3 = None
        continue
    reopened_ref_cell = reopened_reference.tables[0].cell(0, 0)
    reopened_our_cell = reopened_candidate.tables[0].cell(0, 0)
    py_assert1 = reopened_ref_cell.text
    py_assert5 = reopened_our_cell.text
    py_assert3 = py_assert1 == py_assert5
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert5 = None
    py_assert0 = [(p.text, p.style.style_id) for p in reopened_ref_cell.paragraphs]
    py_assert3 = [(p.text, p.style.style_id) for p in reopened_our_cell.paragraphs]
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None

def test_documented_style_assignment_types_and_lookup_errors_match():
    reference = docx.Document()
    candidate = compat.Document()

    def assert_errors_match(ref_operation, our_operation):
        py_assert2 = _capture_exception(ref_operation)
        py_assert7 = _capture_exception(our_operation)
        py_assert4 = py_assert2 == py_assert7
        assert not not py_assert4
        py_assert2 = py_assert4 = py_assert7 = None
    assert_errors_match(lambda: reference.add_paragraph('x', style='Missing Paragraph Style'), lambda: candidate.add_paragraph('x', style='Missing Paragraph Style'))
    assert_errors_match(lambda: setattr(reference.add_paragraph(), 'style', reference.styles['Default Paragraph Font']), lambda: setattr(candidate.add_paragraph(), 'style', candidate.styles['Default Paragraph Font']))
    assert_errors_match(lambda: setattr(reference.add_paragraph().add_run('x'), 'style', reference.styles['Normal']), lambda: setattr(candidate.add_paragraph().add_run('x'), 'style', candidate.styles['Normal']))
    assert_errors_match(lambda: setattr(reference.add_table(rows=1, cols=1), 'style', reference.styles['Normal']), lambda: setattr(candidate.add_table(rows=1, cols=1), 'style', candidate.styles['Normal']))
    assert_errors_match(lambda: reference.add_table(rows=1, cols=1, style='Missing Table Style'), lambda: candidate.add_table(rows=1, cols=1, style='Missing Table Style'))
    py_assert2 = reference.tables
    py_assert4 = len(py_assert2)
    py_assert9 = candidate.tables
    py_assert11 = len(py_assert9)
    py_assert6 = py_assert4 == py_assert11
    assert not not py_assert6
    py_assert2 = py_assert4 = py_assert6 = py_assert9 = py_assert11 = None
    ref_cell = reference.add_table(rows=1, cols=1).cell(0, 0)
    our_cell = candidate.add_table(rows=1, cols=1).cell(0, 0)
    assert_errors_match(lambda: ref_cell.add_paragraph('x', style=reference.styles['Table Grid']), lambda: our_cell.add_paragraph('x', style=candidate.styles['Table Grid']))
    assert_errors_match(lambda: reference.sections[0].header.add_paragraph('x', style=reference.styles['Default Paragraph Font']), lambda: candidate.sections[0].header.add_paragraph('x', style=candidate.styles['Default Paragraph Font']))
    ref_comment = reference.comments.add_comment()
    our_comment = candidate.comments.add_comment()
    assert_errors_match(lambda: ref_comment.add_paragraph('x', style=reference.styles['Default Paragraph Font']), lambda: our_comment.add_paragraph('x', style=candidate.styles['Default Paragraph Font']))
    ref_paragraph = reference.add_paragraph(style='Heading 1')
    our_paragraph = candidate.add_paragraph(style='Heading 1')
    py_assert1 = ref_paragraph.style
    py_assert3 = py_assert1.style_id
    py_assert8 = our_paragraph.style
    py_assert10 = py_assert8.style_id
    py_assert5 = py_assert3 == py_assert10
    py_assert12 = 'Heading1'
    py_assert6 = py_assert10 == py_assert12
    assert not (not py_assert5 or not py_assert6)
    py_assert1 = py_assert3 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
    ref_paragraph.style = reference.styles['Normal']
    our_paragraph.style = candidate.styles['Normal']
    py_assert1 = ref_paragraph.style
    py_assert3 = py_assert1.style_id
    py_assert8 = our_paragraph.style
    py_assert10 = py_assert8.style_id
    py_assert5 = py_assert3 == py_assert10
    py_assert12 = 'Normal'
    py_assert6 = py_assert10 == py_assert12
    assert not (not py_assert5 or not py_assert6)
    py_assert1 = py_assert3 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
    ref_paragraph.style = None
    our_paragraph.style = None
    py_assert1 = ref_paragraph.style
    py_assert3 = py_assert1.style_id
    py_assert8 = our_paragraph.style
    py_assert10 = py_assert8.style_id
    py_assert5 = py_assert3 == py_assert10
    py_assert12 = 'Normal'
    py_assert6 = py_assert10 == py_assert12
    assert not (not py_assert5 or not py_assert6)
    py_assert1 = py_assert3 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
    ref_run = ref_paragraph.add_run('styled')
    our_run = our_paragraph.add_run('styled')
    ref_run.style = reference.styles['Emphasis']
    our_run.style = candidate.styles['Emphasis']
    py_assert1 = ref_run.style
    py_assert3 = py_assert1.style_id
    py_assert8 = our_run.style
    py_assert10 = py_assert8.style_id
    py_assert5 = py_assert3 == py_assert10
    py_assert12 = 'Emphasis'
    py_assert6 = py_assert10 == py_assert12
    assert not (not py_assert5 or not py_assert6)
    py_assert1 = py_assert3 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
    ref_run.style = None
    our_run.style = None
    py_assert1 = ref_run.style
    py_assert3 = py_assert1.style_id
    py_assert8 = our_run.style
    py_assert10 = py_assert8.style_id
    py_assert5 = py_assert3 == py_assert10
    py_assert12 = 'DefaultParagraphFont'
    py_assert6 = py_assert10 == py_assert12
    assert not (not py_assert5 or not py_assert6)
    py_assert1 = py_assert3 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
    ref_table = reference.add_table(rows=1, cols=1, style='Table Grid')
    our_table = candidate.add_table(rows=1, cols=1, style='Table Grid')
    py_assert1 = ref_table.style
    py_assert3 = py_assert1.style_id
    py_assert8 = our_table.style
    py_assert10 = py_assert8.style_id
    py_assert5 = py_assert3 == py_assert10
    py_assert12 = 'TableGrid'
    py_assert6 = py_assert10 == py_assert12
    assert not (not py_assert5 or not py_assert6)
    py_assert1 = py_assert3 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
    ref_table.style = None
    our_table.style = None
    py_assert1 = ref_table.style
    py_assert3 = py_assert1.style_id
    py_assert7 = our_table.style
    py_assert9 = py_assert7.style_id
    py_assert5 = py_assert3 == py_assert9
    assert not not py_assert5
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = None

def test_paragraph_alignment_values_and_style_inheritance_match(tmp_path):
    reference = docx.Document()
    candidate = compat.Document()
    ref_paragraph = reference.add_paragraph('aligned')
    our_paragraph = candidate.add_paragraph('aligned')

    def alignment_name(value):
        if value is not None:
            return value.name
    for name in docx.enum.text.WD_ALIGN_PARAGRAPH.__members__:
        ref_value = getattr(docx.enum.text.WD_ALIGN_PARAGRAPH, name)
        our_value = getattr(WD_ALIGN_PARAGRAPH, name)
        ref_paragraph.alignment = ref_value
        our_paragraph.alignment = our_value
        py_assert2 = ref_paragraph.alignment
        py_assert4 = alignment_name(py_assert2)
        py_assert10 = our_paragraph.alignment
        py_assert12 = alignment_name(py_assert10)
        py_assert6 = py_assert4 == py_assert12
        py_assert7 = py_assert12 == name
        assert not (not py_assert6 or not py_assert7)
        py_assert2 = py_assert4 = py_assert6 = py_assert7 = py_assert10 = py_assert12 = None
        py_assert2 = ref_paragraph.paragraph_format
        py_assert4 = py_assert2.alignment
        py_assert6 = alignment_name(py_assert4)
        py_assert12 = our_paragraph.paragraph_format
        py_assert14 = py_assert12.alignment
        py_assert16 = alignment_name(py_assert14)
        py_assert8 = py_assert6 == py_assert16
        py_assert9 = py_assert16 == name
        assert not (not py_assert8 or not py_assert9)
        py_assert2 = py_assert4 = py_assert6 = py_assert8 = py_assert9 = py_assert12 = py_assert14 = py_assert16 = None
        continue
    ref_paragraph.alignment = our_paragraph.alignment = None
    py_assert1 = ref_paragraph.alignment
    py_assert6 = our_paragraph.alignment
    py_assert3 = py_assert1 is py_assert6
    py_assert8 = None
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = ref_paragraph.paragraph_format
    py_assert3 = py_assert1.alignment
    py_assert8 = our_paragraph.paragraph_format
    py_assert10 = py_assert8.alignment
    py_assert5 = py_assert3 is py_assert10
    py_assert12 = None
    py_assert6 = py_assert10 is py_assert12
    assert not (not py_assert5 or not py_assert6)
    py_assert1 = py_assert3 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
    ref_normal = reference.styles['Normal']
    our_normal = candidate.styles['Normal']
    ref_normal.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    our_normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    py_assert1 = ref_paragraph.alignment
    py_assert6 = our_paragraph.alignment
    py_assert3 = py_assert1 is py_assert6
    py_assert8 = None
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = ref_paragraph.paragraph_format
    py_assert3 = py_assert1.alignment
    py_assert8 = our_paragraph.paragraph_format
    py_assert10 = py_assert8.alignment
    py_assert5 = py_assert3 is py_assert10
    py_assert12 = None
    py_assert6 = py_assert10 is py_assert12
    assert not (not py_assert5 or not py_assert6)
    py_assert1 = py_assert3 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
    ref_path = tmp_path / 'reference-alignment-inheritance.docx'
    our_path = tmp_path / 'candidate-alignment-inheritance.docx'
    reference.save(ref_path)
    candidate.save(our_path)
    reopened_reference = docx.Document(str(ref_path)).paragraphs[0]
    reopened_candidate = compat.Document(our_path.read_bytes()).paragraphs[0]
    py_assert1 = reopened_reference.alignment
    py_assert6 = reopened_candidate.alignment
    py_assert3 = py_assert1 is py_assert6
    py_assert8 = None
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert1 = reopened_reference.paragraph_format
    py_assert3 = py_assert1.alignment
    py_assert8 = reopened_candidate.paragraph_format
    py_assert10 = py_assert8.alignment
    py_assert5 = py_assert3 is py_assert10
    py_assert12 = None
    py_assert6 = py_assert10 is py_assert12
    assert not (not py_assert5 or not py_assert6)
    py_assert1 = py_assert3 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = py_assert12 = None
    py_assert1 = reopened_reference.style
    py_assert3 = py_assert1.paragraph_format
    py_assert5 = py_assert3.alignment
    py_assert9 = reopened_candidate.style
    py_assert11 = py_assert9.paragraph_format
    py_assert13 = py_assert11.alignment
    py_assert7 = py_assert5 == py_assert13
    assert not not py_assert7
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = py_assert11 = py_assert13 = None

def test_document_and_run_picture_formats_sources_and_native_sizes_match(tmp_path):
    formats = (('png-dpi', 'png', blobs.PNG_96), ('png-default-dpi', 'png', blobs.PNG_RAW), ('jpeg', 'jpg', blobs.JPG_300), ('gif', 'gif', blobs.GIF), ('bmp', 'bmp', blobs.BMP), ('tiff', 'tiff', blobs.TIFF_300))
    reference = docx.Document()
    candidate = compat.Document()

    def shape_snapshot(shape):
        return (shape.type.name, int(shape.width), int(shape.height))
    paths = {}
    for index, (label, extension, data) in enumerate(formats):
        path = tmp_path / f'{label}.{extension}'
        path.write_bytes(data)
        paths[label] = path
        if index % 2 == 0:
            ref_shape = reference.add_picture(str(path))
            our_shape = candidate.add_picture(str(path))
        else:
            ref_stream = io.BytesIO(data)
            our_stream = io.BytesIO(data)
            ref_stream.seek(9)
            our_stream.seek(9)
            ref_shape = reference.add_picture(ref_stream)
            our_shape = candidate.add_picture(our_stream)
        py_assert2 = shape_snapshot(ref_shape)
        py_assert7 = shape_snapshot(our_shape)
        py_assert4 = py_assert2 == py_assert7
        assert not not py_assert4
        py_assert2 = py_assert4 = py_assert7 = None
        continue
    ref_run = reference.add_paragraph().add_run()
    our_run = candidate.add_paragraph().add_run()
    ref_run_shape = ref_run.add_picture(str(paths['bmp']))
    our_run_shape = our_run.add_picture(str(paths['bmp']))
    py_assert2 = shape_snapshot(ref_run_shape)
    py_assert7 = shape_snapshot(our_run_shape)
    py_assert4 = py_assert2 == py_assert7
    assert not not py_assert4
    py_assert2 = py_assert4 = py_assert7 = None
    ref_stream, our_stream = (io.BytesIO(blobs.TIFF_300), io.BytesIO(blobs.TIFF_300))
    ref_stream.seek(4)
    our_stream.seek(4)
    py_assert2 = ref_run.add_picture
    py_assert5 = py_assert2(ref_stream)
    py_assert7 = shape_snapshot(py_assert5)
    py_assert12 = our_run.add_picture
    py_assert15 = py_assert12(our_stream)
    py_assert17 = shape_snapshot(py_assert15)
    py_assert9 = py_assert7 == py_assert17
    assert not not py_assert9
    py_assert2 = py_assert5 = py_assert7 = py_assert9 = py_assert12 = py_assert15 = py_assert17 = None
    py_assert0 = [shape_snapshot(shape) for shape in reference.inline_shapes]
    py_assert3 = [shape_snapshot(shape) for shape in candidate.inline_shapes]
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None
    ref_path = tmp_path / 'reference-all-picture-formats.docx'
    our_path = tmp_path / 'candidate-all-picture-formats.docx'
    reference.save(ref_path)
    candidate.save(our_path)
    reopened_reference = docx.Document(str(ref_path))
    reopened_candidate = compat.Document(our_path.read_bytes())
    py_assert0 = [shape_snapshot(shape) for shape in reopened_reference.inline_shapes]
    py_assert3 = [shape_snapshot(shape) for shape in reopened_candidate.inline_shapes]
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None
    py_assert2 = reopened_reference.inline_shapes
    py_assert4 = len(py_assert2)
    py_assert10 = reopened_candidate.inline_shapes
    py_assert12 = len(py_assert10)
    py_assert6 = py_assert4 == py_assert12
    py_assert16 = len(formats)
    py_assert18 = 2
    py_assert20 = py_assert16 + py_assert18
    py_assert7 = py_assert12 == py_assert20
    assert not (not py_assert6 or not py_assert7)
    py_assert2 = py_assert4 = py_assert6 = py_assert7 = py_assert10 = py_assert12 = py_assert16 = py_assert18 = py_assert20 = None
    py_assert0 = [shape_snapshot(shape) for shape in reopened_reference.paragraphs[-1].runs[0].iter_inner_content() if hasattr(shape, 'type')]
    py_assert3 = [shape_snapshot(shape) for shape in reopened_candidate.paragraphs[-1].runs[0].iter_inner_content() if hasattr(shape, 'type')]
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None

def test_external_multiple_hyperlinks_match_python_docx_and_round_trip(tmp_path):
    reference = docx.Document()
    paragraph = reference.add_paragraph()
    paragraph.add_run('before ')

    def append_link(target, pieces, *, fragment=None, page_break=False):
        rel_id = paragraph.part.relate_to(target, RT.HYPERLINK, is_external=True)
        hyperlink = OxmlElement('w:hyperlink')
        hyperlink.set(qn('r:id'), rel_id)
        if fragment is not None:
            hyperlink.set(qn('w:anchor'), fragment)
        for index, text in enumerate(pieces):
            run = OxmlElement('w:r')
            if index:
                properties = OxmlElement('w:rPr')
                properties.append(OxmlElement('w:b'))
                run.append(properties)
            text_element = OxmlElement('w:t')
            if text.startswith(' ') or text.endswith(' '):
                text_element.set(qn('xml:space'), 'preserve')
            text_element.text = text
            run.append(text_element)
            if page_break and index == 0:
                run.append(OxmlElement('w:lastRenderedPageBreak'))
            hyperlink.append(run)
        paragraph._p.append(hyperlink)
    append_link('https://example.com/a?q=1#embedded', ['first', ' link'], page_break=True)
    paragraph.add_run(' between ')
    append_link('https://example.com/base', ['second', ' target'], fragment='bookmark-1')
    paragraph.add_run(' after')
    payload = io.BytesIO()
    reference.save(payload)
    source = payload.getvalue()
    expected_document = docx.Document(io.BytesIO(source))
    candidate = compat.Document(source)

    def link_snapshot(link):
        return (link.address, link.fragment, link.url, link.text, link.contains_page_break, [(run.text, run.bold, run.italic) for run in link.runs])
    ref_paragraph = expected_document.paragraphs[0]
    our_paragraph = candidate.paragraphs[0]
    py_assert0 = [link_snapshot(link) for link in ref_paragraph.hyperlinks]
    py_assert3 = [link_snapshot(link) for link in our_paragraph.hyperlinks]
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None
    py_assert0 = [link.url for link in our_paragraph.hyperlinks]
    py_assert3 = ['https://example.com/a?q=1#embedded', 'https://example.com/base#bookmark-1']
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None

    def inner_content_snapshot(paragraph):
        return [('link', item.address, item.fragment, item.url, item.text) if hasattr(item, 'url') else ('run', item.text) for item in paragraph.iter_inner_content()]
    py_assert2 = inner_content_snapshot(our_paragraph)
    py_assert7 = inner_content_snapshot(ref_paragraph)
    py_assert4 = py_assert2 == py_assert7
    assert not not py_assert4
    py_assert2 = py_assert4 = py_assert7 = None
    ref_path = tmp_path / 'reference-external-links.docx'
    our_path = tmp_path / 'candidate-external-links.docx'
    expected_document.save(ref_path)
    candidate.save(our_path)
    reopened_reference = docx.Document(str(ref_path)).paragraphs[0]
    reopened_candidate = compat.Document(our_path.read_bytes()).paragraphs[0]
    py_assert0 = [link_snapshot(link) for link in reopened_reference.hyperlinks]
    py_assert3 = [link_snapshot(link) for link in reopened_candidate.hyperlinks]
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None

def test_rendered_page_break_boundary_fragments_match_python_docx():
    cases = {'break-at-start': ('<w:p><w:r><w:lastRenderedPageBreak/><w:t>after</w:t></w:r></w:p>', [(None, 'after')]), 'break-at-end': ('<w:p><w:r><w:t>before</w:t><w:lastRenderedPageBreak/></w:r></w:p>', [('before', None)]), 'break-only': ('<w:p><w:r><w:lastRenderedPageBreak/></w:r></w:p>', [(None, None)]), 'multiple-breaks': ('<w:p><w:r><w:t>before</w:t><w:lastRenderedPageBreak/><w:t>middle</w:t><w:lastRenderedPageBreak/><w:t>after</w:t></w:r></w:p>', [('before', 'middleafter'), 'later-break-raises'])}

    def fragment_text(fragment):
        if fragment is None:
            return
        return fragment.text
    for body, expected in cases.values():
        source = make_docx(body)
        reference = docx.Document(io.BytesIO(source)).paragraphs[0]
        candidate = compat.Document(source).paragraphs[0]
        ref_breaks = reference.rendered_page_breaks
        our_breaks = candidate.rendered_page_breaks
        py_assert2 = len(ref_breaks)
        py_assert8 = len(our_breaks)
        py_assert4 = py_assert2 == py_assert8
        py_assert12 = len(expected)
        py_assert5 = py_assert8 == py_assert12
        assert not (not py_assert4 or not py_assert5)
        py_assert2 = py_assert4 = py_assert5 = py_assert8 = py_assert12 = None
        py_assert1 = reference.contains_page_break
        py_assert6 = candidate.contains_page_break
        py_assert3 = py_assert1 is py_assert6
        py_assert8 = True
        py_assert4 = py_assert6 is py_assert8
        assert not (not py_assert3 or not py_assert4)
        py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
        for ref_break, our_break, wanted in zip(ref_breaks, our_breaks, expected):
            if wanted == 'later-break-raises':
                for property_name in ('preceding_paragraph_fragment', 'following_paragraph_fragment'):
                    py_assert1 = lambda name=property_name: getattr(ref_break, name)
                    py_assert3 = _capture_exception(py_assert1)
                    py_assert7 = lambda name=property_name: getattr(our_break, name)
                    py_assert9 = _capture_exception(py_assert7)
                    py_assert5 = py_assert3 == py_assert9
                    assert not not py_assert5
                    py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = None
                    continue
                continue
            actual = (fragment_text(ref_break.preceding_paragraph_fragment), fragment_text(ref_break.following_paragraph_fragment))
            compatible = (fragment_text(our_break.preceding_paragraph_fragment), fragment_text(our_break.following_paragraph_fragment))
            py_assert1 = actual == compatible
            py_assert2 = compatible == wanted
            assert not (not py_assert1 or not py_assert2)
            py_assert1 = py_assert2 = None
            continue
        continue

def test_length_integer_arithmetic_and_rgb_color_channel_boundaries_match():
    ref_values = (docx.shared.Length(0), docx.shared.Length(1), docx.shared.Length(914400), docx.shared.Length(-635))
    our_values = (compat.Length(0), compat.Length(1), compat.Length(914400), compat.Length(-635))
    for ref_value, our_value in zip(ref_values, our_values):
        operations = (lambda value: value + 635, lambda value: value - 12700, lambda value: value * 3, lambda value: value // 635, lambda value: value / 3, lambda value: value + docx.shared.Length(25400))
        candidate_operations = (lambda value: value + 635, lambda value: value - 12700, lambda value: value * 3, lambda value: value // 635, lambda value: value / 3, lambda value: value + compat.Length(25400))
        for ref_operation, our_operation in zip(operations, candidate_operations):
            ref_result, our_result = (ref_operation(ref_value), our_operation(our_value))
            py_assert1 = ref_result == our_result
            assert not not py_assert1
            py_assert1 = None
            py_assert2 = type(ref_result)
            py_assert7 = type(our_result)
            py_assert4 = py_assert2 is py_assert7
            assert not not py_assert4
            py_assert2 = py_assert4 = py_assert7 = None
            continue
        continue
    for channels in ((0, 0, 0), (255, 255, 255), (0, 127, 255)):
        py_assert2 = docx.shared
        py_assert4 = py_assert2.RGBColor
        py_assert7 = py_assert4(*channels)
        py_assert9 = tuple(py_assert7)
        py_assert14 = compat.RGBColor
        py_assert17 = py_assert14(*channels)
        py_assert19 = tuple(py_assert17)
        py_assert11 = py_assert9 == py_assert19
        assert not not py_assert11
        py_assert2 = py_assert4 = py_assert7 = py_assert9 = py_assert11 = py_assert14 = py_assert17 = py_assert19 = None
        py_assert2 = docx.shared
        py_assert4 = py_assert2.RGBColor
        py_assert7 = py_assert4(*channels)
        py_assert9 = repr(py_assert7)
        py_assert14 = compat.RGBColor
        py_assert17 = py_assert14(*channels)
        py_assert19 = repr(py_assert17)
        py_assert11 = py_assert9 == py_assert19
        assert not not py_assert11
        py_assert2 = py_assert4 = py_assert7 = py_assert9 = py_assert11 = py_assert14 = py_assert17 = py_assert19 = None
        py_assert2 = docx.shared
        py_assert4 = py_assert2.RGBColor
        py_assert7 = py_assert4(*channels)
        py_assert9 = str(py_assert7)
        py_assert14 = compat.RGBColor
        py_assert17 = py_assert14(*channels)
        py_assert19 = str(py_assert17)
        py_assert11 = py_assert9 == py_assert19
        assert not not py_assert11
        py_assert2 = py_assert4 = py_assert7 = py_assert9 = py_assert11 = py_assert14 = py_assert17 = py_assert19 = None
        continue
    invalid_values = (-1, 256)
    for channel_index in range(3):
        for invalid in invalid_values:
            channels = [10, 20, 30]
            channels[channel_index] = invalid
            py_assert1 = lambda channels=tuple(channels): docx.shared.RGBColor(*channels)
            py_assert3 = _capture_exception(py_assert1)
            py_assert7 = lambda channels=tuple(channels): compat.RGBColor(*channels)
            py_assert9 = _capture_exception(py_assert7)
            py_assert5 = py_assert3 == py_assert9
            assert not not py_assert5
            py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = None
            continue
        continue
    for value in (1.5, '12', None):
        py_assert1 = lambda value=value: docx.shared.RGBColor(value, 2, 3)
        py_assert3 = _capture_exception(py_assert1)
        py_assert7 = lambda value=value: compat.RGBColor(value, 2, 3)
        py_assert9 = _capture_exception(py_assert7)
        py_assert5 = py_assert3 == py_assert9
        assert not not py_assert5
        py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = None
        continue

def test_comment_text_controls_and_timestamp_parsing_match_python_docx(tmp_path):
    reference = docx.Document()
    candidate = compat.Document()
    ref_comments = reference.comments
    our_comments = candidate.comments
    for text, author, initials in (('alpha\tbeta\rgamma\ndelta', 'Editor', 'ED'), ('\nleading paragraph', 'Reviewer', None), ('trailing paragraph\n', '', '')):
        ref_comment = ref_comments.add_comment(text, author, initials)
        our_comment = our_comments.add_comment(text, author, initials)

        def snapshot(comment):
            return (comment.author, comment.initials, comment.text, [(paragraph.text, [(run.text, run.bold, run.italic) for run in paragraph.runs]) for paragraph in comment.paragraphs])
        py_assert2 = snapshot(ref_comment)
        py_assert7 = snapshot(our_comment)
        py_assert4 = py_assert2 == py_assert7
        assert not not py_assert4
        py_assert2 = py_assert4 = py_assert7 = None
        py_assert1 = ref_comment.timestamp
        py_assert3 = py_assert1.tzinfo
        py_assert6 = None
        py_assert5 = py_assert3 is not py_assert6
        assert not not py_assert5
        py_assert1 = py_assert3 = py_assert5 = py_assert6 = None
        py_assert1 = our_comment.timestamp
        py_assert3 = py_assert1.tzinfo
        py_assert6 = None
        py_assert5 = py_assert3 is not py_assert6
        assert not not py_assert5
        py_assert1 = py_assert3 = py_assert5 = py_assert6 = None
        continue
    ref_comment = ref_comments.get(0)
    our_comment = our_comments.get(0)
    ref_added = ref_comment.add_paragraph('comment\tparagraph\rwith break', style=reference.styles['Heading 2'])
    our_added = our_comment.add_paragraph('comment\tparagraph\rwith break', style=candidate.styles['Heading 2'])
    py_assert0 = (ref_added.text, ref_added.style.style_id)
    py_assert3 = (our_added.text, our_added.style.style_id)
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None
    ref_path = tmp_path / 'reference-comment-controls.docx'
    our_path = tmp_path / 'candidate-comment-controls.docx'
    reference.save(ref_path)
    candidate.save(our_path)

    def patch_timestamp(package_bytes, timestamp):
        word_ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        date_key = f'{{{word_ns}}}date'
        output = io.BytesIO()
        with zipfile.ZipFile(io.BytesIO(package_bytes), 'r') as source:
            with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as target:
                for member in source.infolist():
                    data = source.read(member.filename)
                    if member.filename == 'word/comments.xml':
                        root = ET.fromstring(data)
                        comment = root.find(f'{{{word_ns}}}comment')
                        if timestamp is None:
                            comment.attrib.pop(date_key, None)
                        else:
                            comment.set(date_key, timestamp)
                        data = ET.tostring(root, encoding='utf-8', xml_declaration=True)
                    target.writestr(member, data)
        return output.getvalue()
    reference_bytes = ref_path.read_bytes()
    for value in ('2026-11-01T12:34:56Z', '2024-02-29T23:59:58+05:30', '2024-02-29T23:59:58', '2024-02-29T23:59:58.123456-04:00', None):
        fixture = patch_timestamp(reference_bytes, value)
        ref_timestamp = docx.Document(io.BytesIO(fixture)).comments.get(0).timestamp
        our_timestamp = compat.Document(fixture).comments.get(0).timestamp
        py_assert1 = ref_timestamp == our_timestamp
        assert not not py_assert1
        py_assert1 = None
        py_assert0 = ref_timestamp.tzinfo if ref_timestamp else None
        py_assert3 = our_timestamp.tzinfo if our_timestamp else None
        py_assert2 = py_assert0 == py_assert3
        assert not not py_assert2
        py_assert0 = py_assert2 = py_assert3 = None
        continue

def test_paragraph_runs_exclude_hyperlinks_and_revision_wrappers_like_python_docx():
    body = "<w:p><w:r><w:t>direct before</w:t></w:r><w:hyperlink w:anchor='internal'><w:r><w:t>linked</w:t><w:lastRenderedPageBreak/><w:t> across pages</w:t></w:r></w:hyperlink><w:ins w:id='1' w:author='Editor' w:date='2024-01-01T00:00:00Z'><w:r><w:t>inserted revision</w:t></w:r></w:ins><w:r><w:t>direct after</w:t><w:lastRenderedPageBreak/></w:r></w:p>"
    source = make_docx(body)
    reference = docx.Document(io.BytesIO(source)).paragraphs[0]
    candidate = compat.Document(source).paragraphs[0]
    py_assert0 = [run.text for run in reference.runs]
    py_assert4 = [run.text for run in candidate.runs]
    py_assert2 = py_assert0 == py_assert4
    py_assert6 = ['direct before', 'direct after']
    py_assert3 = py_assert4 == py_assert6
    assert not (not py_assert2 or not py_assert3)
    py_assert0 = py_assert2 = py_assert3 = py_assert4 = py_assert6 = None
    py_assert1 = reference.text
    py_assert6 = candidate.text
    py_assert3 = py_assert1 == py_assert6
    py_assert8 = 'direct beforelinked across pagesdirect after'
    py_assert4 = py_assert6 == py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None

    def inner_snapshot(paragraph):
        return [('link', item.text, item.address, item.fragment) if hasattr(item, 'url') else ('run', item.text) for item in paragraph.iter_inner_content()]
    py_assert2 = inner_snapshot(reference)
    py_assert7 = inner_snapshot(candidate)
    py_assert4 = py_assert2 == py_assert7
    assert not not py_assert4
    py_assert2 = py_assert4 = py_assert7 = None
    py_assert1 = reference.contains_page_break
    py_assert6 = candidate.contains_page_break
    py_assert3 = py_assert1 is py_assert6
    py_assert8 = True
    py_assert4 = py_assert6 is py_assert8
    assert not (not py_assert3 or not py_assert4)
    py_assert1 = py_assert3 = py_assert4 = py_assert6 = py_assert8 = None
    py_assert2 = reference.rendered_page_breaks
    py_assert4 = len(py_assert2)
    py_assert10 = candidate.rendered_page_breaks
    py_assert12 = len(py_assert10)
    py_assert6 = py_assert4 == py_assert12
    py_assert14 = 2
    py_assert7 = py_assert12 == py_assert14
    assert not (not py_assert6 or not py_assert7)
    py_assert2 = py_assert4 = py_assert6 = py_assert7 = py_assert10 = py_assert12 = py_assert14 = None
    ref_breaks = reference.rendered_page_breaks
    our_breaks = candidate.rendered_page_breaks
    py_assert0 = ref_breaks[0]
    py_assert2 = py_assert0.preceding_paragraph_fragment
    py_assert4 = py_assert2.text
    py_assert7 = our_breaks[0]
    py_assert9 = py_assert7.preceding_paragraph_fragment
    py_assert11 = py_assert9.text
    py_assert6 = py_assert4 == py_assert11
    assert not not py_assert6
    py_assert0 = py_assert2 = py_assert4 = py_assert6 = py_assert7 = py_assert9 = py_assert11 = None
    py_assert0 = ref_breaks[0]
    py_assert2 = py_assert0.following_paragraph_fragment
    py_assert4 = py_assert2.text
    py_assert7 = our_breaks[0]
    py_assert9 = py_assert7.following_paragraph_fragment
    py_assert11 = py_assert9.text
    py_assert6 = py_assert4 == py_assert11
    assert not not py_assert6
    py_assert0 = py_assert2 = py_assert4 = py_assert6 = py_assert7 = py_assert9 = py_assert11 = None
    for property_name in ('preceding_paragraph_fragment', 'following_paragraph_fragment'):
        py_assert1 = lambda name=property_name: getattr(ref_breaks[1], name)
        py_assert3 = _capture_exception(py_assert1)
        py_assert7 = lambda name=property_name: getattr(our_breaks[1], name)
        py_assert9 = _capture_exception(py_assert7)
        py_assert5 = py_assert3 == py_assert9
        assert not not py_assert5
        py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = None
        continue

def test_table_alignment_cell_enum_and_rectangular_merges_match(tmp_path):
    reference = docx.Document()
    candidate = compat.Document()
    ref_table = reference.add_table(rows=2, cols=2)
    our_table = candidate.add_table(rows=2, cols=2)
    for name in docx.enum.table.WD_TABLE_ALIGNMENT.__members__:
        ref_value = getattr(docx.enum.table.WD_TABLE_ALIGNMENT, name)
        our_value = getattr(WD_TABLE_ALIGNMENT, name)
        ref_table.alignment = ref_value
        our_table.alignment = our_value
        py_assert1 = ref_table.alignment
        py_assert3 = py_assert1.name
        py_assert8 = our_table.alignment
        py_assert10 = py_assert8.name
        py_assert5 = py_assert3 == py_assert10
        py_assert6 = py_assert10 == name
        assert not (not py_assert5 or not py_assert6)
        py_assert1 = py_assert3 = py_assert5 = py_assert6 = py_assert8 = py_assert10 = None
        continue
    for name in docx.enum.table.WD_CELL_VERTICAL_ALIGNMENT.__members__:
        ref_value = getattr(docx.enum.table.WD_CELL_VERTICAL_ALIGNMENT, name)
        our_value = getattr(WD_CELL_VERTICAL_ALIGNMENT, name)
        ref_table.cell(0, 0).vertical_alignment = ref_value
        our_table.cell(0, 0).vertical_alignment = our_value
        py_assert1 = ref_table.cell
        py_assert3 = 0
        py_assert5 = 0
        py_assert7 = py_assert1(py_assert3, py_assert5)
        py_assert9 = py_assert7.vertical_alignment
        py_assert11 = py_assert9.name
        py_assert16 = our_table.cell
        py_assert18 = 0
        py_assert20 = 0
        py_assert22 = py_assert16(py_assert18, py_assert20)
        py_assert24 = py_assert22.vertical_alignment
        py_assert26 = py_assert24.name
        py_assert13 = py_assert11 == py_assert26
        py_assert14 = py_assert26 == name
        assert not (not py_assert13 or not py_assert14)
        py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = py_assert11 = py_assert13 = py_assert14 = py_assert16 = py_assert18 = py_assert20 = py_assert22 = py_assert24 = py_assert26 = None
        continue
    ref_table.cell(0, 0).vertical_alignment = None
    our_table.cell(0, 0).vertical_alignment = None
    py_assert1 = ref_table.cell
    py_assert3 = 0
    py_assert5 = 0
    py_assert7 = py_assert1(py_assert3, py_assert5)
    py_assert9 = py_assert7.vertical_alignment
    py_assert14 = our_table.cell
    py_assert16 = 0
    py_assert18 = 0
    py_assert20 = py_assert14(py_assert16, py_assert18)
    py_assert22 = py_assert20.vertical_alignment
    py_assert11 = py_assert9 is py_assert22
    py_assert24 = None
    py_assert12 = py_assert22 is py_assert24
    assert not (not py_assert11 or not py_assert12)
    py_assert1 = py_assert3 = py_assert5 = py_assert7 = py_assert9 = py_assert11 = py_assert12 = py_assert14 = py_assert16 = py_assert18 = py_assert20 = py_assert22 = py_assert24 = None
    for row in range(2):
        for column in range(2):
            ref_cell = ref_table.cell(row, column)
            our_cell = our_table.cell(row, column)
            text = f'{row}:{column}'
            ref_cell.text = our_cell.text = text
            ref_cell.add_paragraph(f'second {text}')
            our_cell.add_paragraph(f'second {text}')
    ref_merged = ref_table.cell(0, 0).merge(ref_table.cell(1, 1))
    our_merged = our_table.cell(0, 0).merge(our_table.cell(1, 1))
    py_assert1 = ref_merged.text
    py_assert5 = our_merged.text
    py_assert3 = py_assert1 == py_assert5
    assert not not py_assert3
    py_assert1 = py_assert3 = py_assert5 = None
    py_assert0 = [p.text for p in ref_merged.paragraphs]
    py_assert3 = [p.text for p in our_merged.paragraphs]
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None
    py_assert0 = [[cell.text for cell in row.cells] for row in ref_table.rows]
    py_assert3 = [[cell.text for cell in row.cells] for row in our_table.rows]
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None
    ref_vertical = reference.add_table(rows=2, cols=1)
    our_vertical = candidate.add_table(rows=2, cols=1)
    ref_vertical.cell(0, 0).text = 'top'
    our_vertical.cell(0, 0).text = 'top'
    ref_vertical.cell(1, 0).text = 'bottom'
    our_vertical.cell(1, 0).text = 'bottom'
    ref_result = ref_vertical.cell(0, 0).merge(ref_vertical.cell(1, 0))
    our_result = our_vertical.cell(0, 0).merge(our_vertical.cell(1, 0))
    py_assert0 = (ref_result.text, [p.text for p in ref_result.paragraphs])
    py_assert3 = (our_result.text, [p.text for p in our_result.paragraphs])
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None
    ref_path = tmp_path / 'reference-merged-tables.docx'
    our_path = tmp_path / 'candidate-merged-tables.docx'
    reference.save(ref_path)
    candidate.save(our_path)
    reopened_reference = docx.Document(str(ref_path))
    reopened_candidate = compat.Document(our_path.read_bytes())
    py_assert0 = [[cell.text for row in table.rows for cell in row.cells] for table in reopened_reference.tables]
    py_assert3 = [[cell.text for row in table.rows for cell in row.cells] for table in reopened_candidate.tables]
    py_assert2 = py_assert0 == py_assert3
    assert not not py_assert2
    py_assert0 = py_assert2 = py_assert3 = None
