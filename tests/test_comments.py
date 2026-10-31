"""Comments: the range python-docx documents it cannot give you.

Every expectation is derived from the fixture's own XML, never from what the
implementation happens to return.
"""

from __future__ import annotations

import zipfile

import pytest

from docviews import open_docx

from conftest import TRANSITIONAL_W, make_docx

W = TRANSITIONAL_W
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
W15 = "http://schemas.microsoft.com/office/word/2010/11/wordml"


def P(inner: str) -> str:
    return f"<w:p>{inner}</w:p>"


def R(text: str) -> str:
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


def DEL(text: str) -> str:
    return (f'<w:del w:id="7" w:author="A" w:date="2026-01-01T00:00:00Z">'
            f"<w:r><w:delText>{text}</w:delText></w:r></w:del>")


def RANGE_START(cid: str) -> str:
    return f'<w:commentRangeStart w:id="{cid}"/>'


def RANGE_END(cid: str) -> str:
    return f'<w:commentRangeEnd w:id="{cid}"/>'


def REF(cid: str) -> str:
    return f'<w:r><w:commentReference w:id="{cid}"/></w:r>'


def comments_xml(*comments: tuple[str, str]) -> str:
    """comment bodies: (id, xml of the paragraphs)"""
    body = "".join(
        f'<w:comment w:id="{cid}" w:author="Reviewer {cid}" w:initials="R{cid}" '
        f'w:date="2026-02-0{cid}T00:00:00Z">{xml}</w:comment>'
        for cid, xml in comments)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:comments xmlns:w="{W}" xmlns:w14="{W14}" xmlns:w15="{W15}">'
            f'{body}</w:comments>')


def comment_body(text: str, para_id: str) -> str:
    return f'<w:p w14:paraId="{para_id}">{R(text)}</w:p>'


def extended_xml(*entries: tuple[str, str]) -> str:
    """entries: (paraId, extra attributes as an XML string)"""
    body = "".join(f'<w15:commentEx w15:paraId="{pid}" {attrs}/>'
                   for pid, attrs in entries)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w15:commentsEx xmlns:w15="{W15}">{body}</w15:commentsEx>')


# ------------------------------------------------------------------ the anchors

def test_range_inside_one_paragraph_lands_on_the_exact_characters():
    body = P(R("Alpha ") + RANGE_START("1") + R("beta gamma")
             + RANGE_END("1") + R(" delta") + REF("1"))
    doc = open_docx(make_docx(body, extra_parts={
        "comments.xml": comments_xml(("1", comment_body("Shorten this.", "A1B2C3D4")))}))
    body_text = list(doc.paragraphs())[0].text
    assert body_text == "Alpha beta gamma delta"
    # comment text is part of the package and stays reachable through text()
    assert doc.text() == "Alpha beta gamma delta\nShorten this."
    comment, = doc.comments()
    assert comment.text == "Shorten this."
    assert comment.author == "Reviewer 1"
    anchor, = comment.anchors
    assert (anchor.start, anchor.end) == (6, 16)
    assert anchor.text == "beta gamma"
    assert anchor.text == body_text[anchor.start:anchor.end]
    assert not anchor.is_point and not comment.is_orphan


def test_range_spanning_two_paragraphs_returns_one_segment_per_paragraph():
    body = (P(R("Intro ") + RANGE_START("2") + R("words"))
            + P(R("and more") + RANGE_END("2") + R(" after.") + REF("2")))
    doc = open_docx(make_docx(body, extra_parts={
        "comments.xml": comments_xml(("2", comment_body("Cut this.", "AAAA0002")))}))
    paragraphs = list(doc.paragraphs())
    comment, = doc.comments()
    assert len(comment.anchors) == 2
    first, second = comment.anchors
    assert (first.paragraph_index, first.start, first.end) == (0, 6, 11)
    assert first.text == paragraphs[0].text[first.start:first.end] == "words"
    assert (second.paragraph_index, second.start, second.end) == (1, 0, 8)
    assert second.text == paragraphs[1].text[second.start:second.end] == "and more"
    assert comment.anchor_text == "words and more"


def test_reference_without_a_range_is_a_point_anchor_not_a_crash():
    body = P(R("Text") + REF("3"))
    doc = open_docx(make_docx(body, extra_parts={
        "comments.xml": comments_xml(("3", comment_body("Note.", "AAAA0003")))}))
    comment, = doc.comments()
    anchor, = comment.anchors
    assert anchor.is_point and anchor.text == "" and anchor.start == anchor.end == 4
    assert len(anchor) == 0


def test_comment_with_no_marker_anywhere_is_an_orphan():
    body = P(R("Nothing here points at a comment."))
    doc = open_docx(make_docx(body, extra_parts={
        "comments.xml": comments_xml(("4", comment_body("Ghost.", "AAAA0004")))}))
    comment, = doc.comments()
    assert comment.is_orphan and comment.anchors == []
    assert comment.text == "Ghost."          # the text survives the anchor


def test_anchor_follows_the_view_it_was_asked_for():
    """Same document, two views: the comment is about different characters."""
    body = P(RANGE_START("5") + DEL("gone ") + R("kept") + RANGE_END("5") + REF("5"))
    parts = {"comments.xml": comments_xml(("5", comment_body("Both views.", "AAAA0005")))}
    accepted = open_docx(make_docx(body, extra_parts=parts))
    original = open_docx(make_docx(body, extra_parts=parts))
    a_anchor, = accepted.comments("accepted")[0].anchors
    o_anchor, = original.comments("original")[0].anchors
    a_body = list(accepted.paragraphs())[0].text
    o_body = list(original.paragraphs("original"))[0].text
    assert a_body == "kept" and o_body == "gone kept"
    assert a_anchor.text == "kept"
    assert o_anchor.text == "gone kept"
    assert a_anchor.text == a_body[a_anchor.start:a_anchor.end]


def test_anchor_in_a_header_is_attributed_to_the_header_part():
    body = P(R("Body text."))
    header = (f'<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
              + P(R("Head ") + RANGE_START("6") + R("office") + RANGE_END("6") + REF("6"))
              + "</w:hdr>")
    doc = open_docx(make_docx(body, extra_parts={
        "comments.xml": comments_xml(("6", comment_body("Header note.", "AAAA0006"))),
        "header1.xml": header}))
    comment, = doc.comments()
    anchor, = comment.anchors
    assert anchor.part.endswith("header1.xml")
    assert anchor.text == "office"
    header_paragraph = next(p for p in doc.paragraphs()
                            if p.part_name.endswith("header1.xml"))
    assert anchor.text == header_paragraph.text[anchor.start:anchor.end]


# ------------------------------------------------------- threads and resolution

def test_a_reply_thread_and_a_resolved_flag_come_from_comments_extended():
    # Word repeats the parent's range on the reply rather than pointing at the
    # parent, so that is what the fixture does (and what the corpus shows:
    # 380 replies, 380 of them carrying their own markers)
    body = P(R("Text") + RANGE_START("0") + R("marked") + RANGE_END("0")
             + RANGE_START("1") + R("marked") + RANGE_END("1")
             + REF("0") + REF("1"))
    doc = open_docx(make_docx(body, extra_parts={
        # the second body paragraph is the one commentsExtended keys on
        "comments.xml": comments_xml(
            ("0", comment_body("Parent.", "11110000") + comment_body("continued", "22220000")),
            ("1", comment_body("Reply.", "33330000"))),
        "commentsExtended.xml": extended_xml(
            ("22220000", 'w15:done="0"'),
            ("33330000", 'w15:paraIdParent="22220000" w15:done="1"'))}))
    parent, reply = doc.comments()
    assert parent.text == "Parent.\ncontinued"      # both paragraphs of the body
    assert parent.reply_ids == ["1"] and not parent.is_reply
    assert reply.parent_id == "0" and reply.is_reply
    assert reply.resolved and not parent.resolved
    assert [a.text for a in reply.anchors] == ["marked"]


def test_comments_are_empty_when_the_package_has_none():
    doc = open_docx(make_docx(P(R("No comments here."))))
    assert doc.comments() == []


# ------------------------------------------------------------- corpus regression

@pytest.mark.corpus
def test_every_anchor_in_the_corpus_lands_on_its_own_view_text(corpus_path):
    """999 comments in 89 real documents; every anchor must be the exact slice."""
    n_files = n_comments = n_anchors = n_bad = n_errors = 0
    for path in sorted(corpus_path.glob("*.docx")):
        try:
            with zipfile.ZipFile(path) as zf:
                if not any(n.endswith("comments.xml") for n in zf.namelist()):
                    continue
        except Exception:                                        # noqa: BLE001
            continue
        try:
            doc = open_docx(path)
            comments = doc.comments()
        except Exception:                                        # noqa: BLE001
            n_errors += 1                           # a raise here is a failure
            continue
        if not comments:
            continue                        # one fixture ships an empty part
        n_files += 1
        text_by_para = {(p.part_name, p.index): p.text for p in doc.paragraphs()}
        for comment in comments:
            n_comments += 1
            for anchor in comment.anchors:
                n_anchors += 1
                expected = text_by_para.get((anchor.part, anchor.paragraph_index))
                if expected is None:
                    n_bad += 1
                    continue
                if anchor.is_point:
                    continue
                if anchor.text != expected[anchor.start:anchor.end]:
                    n_bad += 1
    assert (n_files, n_comments, n_anchors) == (89, 999, 1189)
    assert (n_bad, n_errors) == (0, 0)
