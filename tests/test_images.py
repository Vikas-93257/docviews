"""Pictures: detect the format, place it, and find it again.

The image files here are 100-odd bytes each and live in ``tests/blobs.py``, so
the suite needs no image library. Every expectation is checked on the saved
package *and* against python-docx, because a picture that only this library can
read is not a picture.
"""

from __future__ import annotations

import io
import zipfile

import pytest

from docviews import DocxError, open_docx
from docviews.document import Document
from docviews.images import ImageError, detect

import blobs
from conftest import CORPUS

docx = pytest.importorskip("docx")             # python-docx, for cross-checks

EMU_PER_INCH = 914400


# ------------------------------------------------------------------ detection
def test_png_size_and_dpi_come_from_the_file():
    info = detect(blobs.PNG_96)
    assert (info.content_type, info.extension) == ("image/png", "png")
    assert (info.width_px, info.height_px) == (120, 60)
    assert round(info.dpi_x) == 96
    # 120 px at 96 dpi = 1.25 in
    assert round(info.native_width / EMU_PER_INCH, 2) == 1.25


def test_a_png_without_a_physical_size_is_72_dpi():
    info = detect(blobs.PNG_RAW)
    assert round(info.dpi_x) == 72
    assert round(info.native_width / EMU_PER_INCH, 2) == round(120 / 72, 2)


def test_jpeg_and_tiff_carry_their_own_dpi():
    assert round(detect(blobs.JPG_300).dpi_x) == 300
    assert round(detect(blobs.TIFF_300).dpi_x) == 300


def test_gif_and_bmp_are_measured():
    assert (detect(blobs.GIF).width_px, detect(blobs.GIF).height_px) == (3, 2)
    assert (detect(blobs.BMP).width_px, detect(blobs.BMP).height_px) == (4, 6)


def test_one_measurement_keeps_the_aspect_ratio():
    info = detect(blobs.PNG_96)                     # 120x60
    width, height = info.scaled(width=2 * EMU_PER_INCH, height=None)
    assert width == 2 * EMU_PER_INCH
    assert height == 1 * EMU_PER_INCH


def test_an_unreadable_image_is_refused_not_guessed():
    for data in (blobs.SVG, blobs.JUNK, b""):
        with pytest.raises(ImageError):
            detect(data)


def test_the_door_accepts_paths_bytes_and_file_objects(tmp_path):
    path = tmp_path / "picture.png"
    path.write_bytes(blobs.PNG_96)
    for source in (path, str(path), blobs.PNG_96,
                   blobs.file_like("picture.png", blobs.PNG_96)):
        assert detect(source).extension == "png"


# --------------------------------------------------------------- add_picture
def test_a_picture_lands_in_the_package_with_its_own_parts():
    doc = Document.new()
    shape = doc.add_picture(blobs.PNG_96, width=2 * EMU_PER_INCH)
    assert shape.width == 2 * EMU_PER_INCH
    assert shape.height == 1 * EMU_PER_INCH          # aspect kept
    assert shape.media == "word/media/image1.png"
    with zipfile.ZipFile(io.BytesIO(doc.to_bytes())) as zf:
        assert zf.read("word/media/image1.png") == blobs.PNG_96
        types = zf.read("[Content_Types].xml").decode()
        assert 'Extension="png"' in types
        rels = zf.read("word/_rels/document.xml.rels").decode()
        assert 'Target="media/image1.png"' in rels


def test_python_docx_sees_the_picture_we_wrote(tmp_path):
    doc = Document.new()
    doc.add_picture(blobs.PNG_96, width=2 * EMU_PER_INCH)
    doc.add_picture(blobs.JPG_300, height=EMU_PER_INCH)
    path = tmp_path / "pictures.docx"
    doc.save(path)
    theirs = docx.Document(str(path))
    assert len(theirs.inline_shapes) == 2
    assert [round(s.width / EMU_PER_INCH, 2) for s in theirs.inline_shapes] == [2.0, 2.0]
    assert [round(s.height / EMU_PER_INCH, 2) for s in theirs.inline_shapes] == [1.0, 1.0]


def test_each_format_gets_a_part_name_word_would_use():
    doc = Document.new()
    doc.add_picture(blobs.PNG_96)
    doc.add_picture(blobs.JPG_300)
    doc.add_picture(blobs.GIF)
    with zipfile.ZipFile(io.BytesIO(doc.to_bytes())) as zf:
        names = zf.namelist()
    # numbering is per extension, which is what Word and python-docx both do
    assert "word/media/image1.png" in names
    assert "word/media/image1.jpeg" in names
    assert "word/media/image1.gif" in names


def test_inline_shapes_is_true_immediately_and_after_a_round_trip():
    doc = Document.new()
    assert doc.inline_shapes == []
    doc.add_picture(blobs.PNG_96, width=EMU_PER_INCH)
    assert len(doc.inline_shapes) == 1                # before saving anything
    reopened = open_docx(doc.to_bytes())
    assert len(reopened.inline_shapes) == 1
    assert reopened.inline_shapes[0].name == "Picture 1"
    assert reopened.inline_shapes[0].media == "word/media/image1.png"


def _corpus_file_with_media():
    """A real file that already carries an image, or None."""
    for candidate in sorted(CORPUS.glob("*.docx")):
        try:
            with zipfile.ZipFile(candidate) as zf:
                if any("/media/" in name for name in zf.namelist()):
                    return candidate
        except zipfile.BadZipFile:
            continue
    return None


def test_adding_a_picture_to_an_existing_document_keeps_it_intact():
    path = _corpus_file_with_media()
    if path is None:
        pytest.skip("no corpus file with media in this checkout")
    original = path.read_bytes()
    before = len(open_docx(original).inline_shapes)
    doc = Document(original)
    doc.add_picture(blobs.PNG_96, width=EMU_PER_INCH)
    saved = doc.to_bytes()
    assert len(open_docx(saved).inline_shapes) == before + 1
    with zipfile.ZipFile(io.BytesIO(original)) as a, \
            zipfile.ZipFile(io.BytesIO(saved)) as b:
        touched = set(b.namelist()) - set(a.namelist())
        touched |= {name for name in a.namelist()
                    if not name.endswith("/") and a.read(name) != b.read(name)}
    # the body, its relationships, the new media part — and nothing else
    assert all(name in ("word/document.xml", "word/_rels/document.xml.rels",
                        "[Content_Types].xml") or name.startswith("word/media/")
               for name in touched), touched
    assert any(name.startswith("word/media/") for name in touched), touched


def test_python_docx_can_read_a_picture_added_to_an_existing_document(tmp_path):
    path = _corpus_file_with_media()
    if path is None:
        pytest.skip("no corpus file with media in this checkout")
    # Compare python-docx to its own pre-edit baseline. docviews intentionally
    # sees pictures in text boxes that python-docx does not expose as
    # InlineShapes, so using docviews' count here makes the result depend on
    # which corpus file a platform's path ordering selects.
    theirs_before = len(docx.Document(str(path)).inline_shapes)
    doc = Document(path.read_bytes())
    doc.add_picture(blobs.PNG_96, width=EMU_PER_INCH)
    target = tmp_path / "with-picture.docx"
    doc.save(target)
    theirs = docx.Document(str(target))
    assert len(theirs.inline_shapes) == theirs_before + 1


def test_a_picture_can_be_put_inside_an_existing_run():
    doc = Document.new()
    paragraph = doc.add_paragraph("before ")
    run = paragraph.add_run("")
    doc.add_picture(blobs.GIF, run=run)
    assert len(doc.inline_shapes) == 1
    reopened = open_docx(doc.to_bytes())
    assert len(reopened.inline_shapes) == 1
