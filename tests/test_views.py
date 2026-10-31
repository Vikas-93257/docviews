"""The three views, exactly. Every expectation here is taken from the fixture's
own stated intention, not from what the implementation happens to do."""

from __future__ import annotations

import pytest

from docviews import DocxError, Limits, open_docx
from docviews.ooxml import PartTooLarge
from docviews.views import views_of

from conftest import STRICT_W, make_docx


def P(inner: str) -> str:
    return f"<w:p>{inner}</w:p>"


def R(text: str, preserve: bool = True) -> str:
    space = ' xml:space="preserve"' if preserve else ""
    return f"<w:r><w:t{space}>{text}</w:t></w:r>"


def INS(text: str) -> str:
    return f'<w:ins w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z">{R(text)}</w:ins>'


def DEL(text: str) -> str:
    return (f'<w:del w:id="2" w:author="A" w:date="2026-01-01T00:00:00Z">'
            f"<w:r><w:delText>{text}</w:delText></w:r></w:del>")


# --------------------------------------------------------------------- the core

def test_insertion_and_deletion_produce_three_different_texts():
    """The corpus' impossible text: reader mixes the two views, Word never does."""
    doc = open_docx(make_docx(P(R("We prefer ") + DEL("colour") + R(" ") + INS("color") + R(" spelling."))))
    assert doc.text("accepted") == "We prefer  color spelling."
    assert doc.text("original") == "We prefer colour  spelling."
    assert doc.text("markup") == "We prefer colour color spelling."


def test_deleted_text_is_absent_from_accepted_and_present_in_original():
    doc = open_docx(make_docx(P(R("This ") + DEL("old ") + R("sentence remains."))))
    assert doc.text("accepted") == "This sentence remains."
    assert doc.text("original") == "This old sentence remains."


def test_inserted_text_is_present_in_accepted_and_absent_from_original():
    """The reviewer's insertion leaves a doubled space in the draft — exactly the
    pattern the corpus documents contain."""
    doc = open_docx(make_docx(P(R("This sentence has an ") + INS(" inserted") + R(" word."))))
    assert doc.text("accepted") == "This sentence has an  inserted word."
    assert doc.text("original") == "This sentence has an  word."


def test_moves_are_revisions_too():
    body = P("<w:moveFrom w:id='7' w:author='A'>" + R("moved") + "</w:moveFrom>"
             + R(" plus ") + "<w:moveTo w:id='8' w:author='A'>" + R("here") + "</w:moveTo>")
    doc = open_docx(make_docx(body))
    assert doc.text("original") == "moved plus "
    assert doc.text("accepted") == " plus here"
    kinds = {r.kind for r in doc.revisions()}
    assert kinds == {"move_from", "move_to"}


def test_revisions_carry_author_date_and_position():
    doc = open_docx(make_docx(P(R("a")) + P(R("b ") + INS("c"))))
    revs = doc.revisions()
    assert len(revs) == 1
    assert (revs[0].kind, revs[0].author, revs[0].id) == ("insert", "A", "1")
    assert revs[0].date.startswith("2026-01-01")
    assert revs[0].paragraph_index == 1


# ------------------------------------------------------- everything else text lives in

def test_tabs_and_breaks_map_to_characters():
    body = P(R("a") + "<w:r><w:tab/></w:r>" + R("b") + "<w:r><w:br/></w:r>" + R("c"))
    doc = open_docx(make_docx(body))
    assert doc.text("accepted") == "a\tb\nc"


def test_hyperlink_text_is_document_text():
    body = P(R("see ") + "<w:hyperlink r:id='rId9'>" + R("the clause") + "</w:hyperlink>")
    doc = open_docx(make_docx(body))
    assert doc.text("accepted") == "see the clause"


def test_content_controls_are_read():
    body = P(R("before ") + "<w:sdt><w:sdtContent>" + R("inside") + "</w:sdtContent></w:sdt>")
    doc = open_docx(make_docx(body))
    assert doc.text("accepted") == "before inside"


def test_table_paragraphs_are_found_and_flagged():
    body = "<w:tbl><w:tr><w:tc>" + P(R("cell one")) + "</w:tc>" \
           "<w:tc>" + P(R("cell two")) + "</w:tc></w:tr></w:tbl>"
    doc = open_docx(make_docx(body))
    paras = list(doc.paragraphs("accepted"))
    assert [p.text for p in paras] == ["cell one", "cell two"]
    assert all(p.in_table for p in paras)


def test_text_boxes_are_separate_paragraphs_not_hidden_inside_the_anchor():
    """Text inside an anchored text box belongs to the box, not to its anchor."""
    body = P(R("anchor text") + "<w:r><w:drawing><w:txbxContent>"
             + P(R("boxed text")) + "</w:txbxContent></w:drawing></w:r>")
    doc = open_docx(make_docx(body))
    pairs = [(p.text, p.kind, p.in_textbox) for p in doc.paragraphs("accepted")]
    assert pairs == [("anchor text", "document", False),
                     ("boxed text", "textbox", True)]


def test_footnotes_endnotes_and_headers_are_read():
    extra = {
        "footnotes.xml":
            '<w:footnotes xmlns:w="{ns}"><w:footnote w:id="1">' + P(R("footnote text")) + "</w:footnote></w:footnotes>",
        "endnotes.xml":
            '<w:endnotes xmlns:w="{ns}"><w:endnote w:id="1">' + P(R("endnote text")) + "</w:endnote></w:endnotes>",
        "header1.xml":
            '<w:hdr xmlns:w="{ns}">' + P(R("header text")) + "</w:hdr>",
    }
    doc = open_docx(make_docx(P(R("body text")), extra_parts=extra))
    kinds = {p.kind: p.text for p in doc.paragraphs("accepted")}
    assert kinds["footnotes"] == "footnote text"
    assert kinds["endnotes"] == "endnote text"
    assert kinds["header"] == "header text"
    assert kinds["document"] == "body text"


# ------------------------------------------------------------- packages, guardrails

def test_strict_packages_are_read_where_python_docx_raises():
    body = P(R("strict text ") + INS("with a revision"))
    doc = open_docx(make_docx(body, strict=True))
    assert doc.strict is True
    assert doc.text("accepted") == "strict text with a revision"
    assert doc.text("original") == "strict text "


def test_expansion_guard_refuses_a_bomb_unless_limits_are_raised():
    bomb = P(R("x" * 200_000))
    data = make_docx(bomb)
    with pytest.raises(PartTooLarge):
        open_docx(data, limits=Limits(max_part_bytes=1000)).text("accepted")
    assert "x" * 100 in open_docx(data).text("accepted")


def test_unreadable_package_raises_a_named_error():
    with pytest.raises(DocxError):
        open_docx(b"this is not a zip file")


def test_missing_main_part_is_reported_not_swallowed():
    with pytest.raises(DocxError):
        open_docx(b"PK\x05\x06" + b"\x00" * 18)      # empty zip: no relationship


def test_views_of_works_on_any_element():
    from xml.etree import ElementTree as ET
    xml = f'<w:p xmlns:w="{STRICT_W}">' + R("keep ") + INS("new") + "</w:p>"
    v = views_of(ET.fromstring(xml))
    assert (v.original, v.accepted, v.markup) == ("keep ", "keep new", "keep new")
    assert v.diverge is True


# ------------------------------------------------- revisions that are not runs

def test_a_deleted_table_row_is_reported_as_a_row_scope_revision():
    """Word marks a deleted row with w:del inside w:trPr; the text stays put."""
    row = ('<w:tr><w:trPr><w:del w:id="9" w:author="A" '
           'w:date="2026-01-01T00:00:00Z"/></w:trPr>'
           f'<w:tc>{P(R("gone row"))}</w:tc></w:tr>')
    doc = open_docx(make_docx(f"<w:tbl>{row}</w:tbl>"))
    rev, = doc.revisions()
    assert (rev.kind, rev.scope, rev.paragraph_index) == ("delete", "row", -1)
    assert doc.text() == "gone row"          # the text is still there: that is the trap


def test_a_deleted_paragraph_mark_is_reported_as_a_paragraph_mark_revision():
    """w:del inside w:pPr/w:rPr deletes the paragraph mark: two paragraphs merge."""
    p = ('<w:p><w:pPr><w:rPr><w:del w:id="10" w:author="B" '
         'w:date="2026-01-02T00:00:00Z"/></w:rPr></w:pPr>'
         + R("merged here") + "</w:p>")
    doc = open_docx(make_docx(p + P(R("next paragraph"))))
    rev, = doc.revisions()
    assert (rev.kind, rev.scope, rev.paragraph_index) == ("delete", "paragraph-mark", -1)


def test_a_revision_inside_a_text_box_is_counted_once(corpus_like=None):
    """The text box is its own paragraph, so its revision is not the host's."""
    box = ('<w:txbxContent>' + P('<w:ins w:id="11" w:author="C" '
                                'w:date="2026-01-03T00:00:00Z">'
                                + R("boxed") + "</w:ins>") + "</w:txbxContent>")
    body = (P(R("host paragraph") + f'<w:r><w:drawing>{box}</w:drawing></w:r>'))
    doc = open_docx(make_docx(body))
    revs = doc.revisions()
    assert len(revs) == 1, revs
    assert revs[0].kind == "insert"
    paragraphs = {p.text: p for p in doc.iter_paragraphs("accepted")}
    assert paragraphs["boxed"].is_revised
    assert not paragraphs["host paragraph"].is_revised
