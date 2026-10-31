"""Comments, with the thing python-docx documents it cannot give you.

python-docx 1.2.0 shipped comment *text*, and the release notes say plainly that
it "does not include the ability to directly access the comment 'reference' …
range of text selected". A comment without its range is a note with no idea what
it is about — the difference between "someone disagreed here" and "someone
disagreed about *this clause*".

The range is in the XML: ``<w:commentRangeStart w:id="1"/>`` … runs …
``<w:commentRangeEnd w:id="1"/>``. Ranges can span paragraphs, sit inside
revised text, or exist without an end marker at all, so the anchors are returned
as *segments* of (part, paragraph, character offsets) against a chosen view.
Threading and resolution live in ``commentsExtended.xml`` (``w15:done``,
``w15:paraIdParent``), keyed by the paragraph ids inside each comment.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, Iterator, List, Optional, Tuple
from xml.etree import ElementTree as ET

from .views import (PART_KINDS, _attr, _tag, part_paragraph_elements,
                    view_and_markers, views_of)


@dataclass
class Anchor:
    """Where a comment points, as character offsets into a paragraph's text."""

    part: str
    paragraph_index: int
    start: int
    end: int
    text: str
    is_point: bool = False          # a reference without a range

    def __len__(self) -> int:
        return max(0, self.end - self.start)


@dataclass
class Comment:
    id: str
    author: str = ""
    initials: str = ""
    date: str = ""
    text: str = ""
    part: str = ""
    paragraph_index: int = 0        # position within the comment, not the document
    anchors: List[Anchor] = field(default_factory=list)
    resolved: bool = False
    parent_id: str = ""
    reply_ids: List[str] = field(default_factory=list)
    para_ids: List[str] = field(default_factory=list, repr=False)

    @property
    def comment_id(self) -> str:
        """python-docx calls the id ``comment_id``; so does this."""
        return self.id

    @property
    def timestamp(self):
        """python-docx's name for the date, as a ``datetime`` when parseable."""
        from .authoring import _as_datetime
        return _as_datetime(self.date)

    @property
    def anchor_text(self) -> str:
        """Everything the comment is attached to, joined with spaces."""
        return " ".join(a.text for a in self.anchors if a.text)

    @property
    def is_reply(self) -> bool:
        return bool(self.parent_id)

    @property
    def is_orphan(self) -> bool:
        """A comment whose anchor was deleted from the document body."""
        return not self.anchors


def _comment_parts(package, extra=()) -> List:
    """The comments parts: the ones in the zip, plus any added since opening."""
    parts = [p for p in package.text_parts() if p.kind == "comments"]
    for part in extra:
        if all(existing.name != part.name for existing in parts):
            parts.append(part)
    return parts


def _extended_state(package) -> Dict[str, Tuple[bool, str]]:
    """paraId -> (resolved, parent paraId), from commentsExtended.xml."""
    state: Dict[str, Tuple[bool, str]] = {}
    for name in package.names():
        if not name.endswith("commentsExtended.xml"):
            continue
        try:
            root = ET.fromstring(package._zip.read(name))
        except Exception:
            continue
        for el in root:
            para_id = _attr(el, "paraId")
            if not para_id:
                continue
            state[para_id] = (_attr(el, "done") == "1", _attr(el, "paraIdParent"))
    return state


def _paragraph_ids(comment_el: ET.Element) -> List[str]:
    out = []
    for el in comment_el.iter():
        if _tag(el) == "p":
            pid = _attr(el, "paraId")          # w14:paraId, namespace-agnostic
            if pid:
                out.append(pid)
    return out


def collect_comments(document, view: str = "accepted",
                     extra=()) -> List[Comment]:
    """Every comment in the package, with anchors, threading and resolved state.

    ``extra`` is for parts created since the document was opened, so a comment
    added a moment ago reads back with the others.
    """
    package = document._package
    parts = _comment_parts(package, extra)
    if not parts:
        return []

    comments: List[Comment] = []
    by_id: Dict[str, Comment] = {}
    for part in parts:
        for el in part.root:
            if _tag(el) != "comment":
                continue
            cid = _attr(el, "id")
            body_paragraphs = [views_of(p).get(view) for p in el.iter()
                               if _tag(p) == "p"]
            comment = Comment(
                id=cid,
                author=_attr(el, "author"),
                initials=_attr(el, "initials"),
                date=_attr(el, "date"),
                text="\n".join(body_paragraphs) if body_paragraphs
                     else views_of(el).get(view),
                part=part.name,
            )
            comment.para_ids = _paragraph_ids(el)
            comments.append(comment)
            by_id[cid] = comment

    # threading and resolution
    extended = _extended_state(package)
    by_para_id: Dict[str, Comment] = {}
    for comment in comments:
        for pid in comment.para_ids:
            by_para_id.setdefault(pid, comment)
    for comment in comments:
        # commentsExtended keys the comment by the paraId of its *last*
        # paragraph, but older writers have used the first one: try the
        # paragraphs in reverse and take the first id the part knows about
        order = list(reversed(comment.para_ids))
        pid = next((p for p in order if p in extended), order[0] if order else "")
        resolved, parent_para = extended.get(pid, (False, ""))
        comment.resolved = resolved
        parent = by_para_id.get(parent_para)
        if parent is not None and parent is not comment:
            comment.parent_id = parent.id
            parent.reply_ids.append(comment.id)

    _attach_anchors(document, view, comments)
    return comments


def _attach_anchors(document, view: str, comments: List[Comment]) -> None:
    """Walk the text parts once and turn range markers into anchors.

    Segments are carried across paragraphs, because a reviewer can select a
    range that starts in clause 3 and ends in clause 5.
    """
    package = document._package
    by_id = {c.id: c for c in comments}
    open_segments: Dict[str, dict] = {}

    def emit(cid: str, segment: dict, end: int) -> None:
        comment = by_id.get(cid)
        if comment is None:
            return
        comment.anchors.append(Anchor(
            part=segment["part"], paragraph_index=segment["paragraph_index"],
            start=segment["start"], end=end,
            text=segment["text"][segment["start"]:end]))

    for part in package.text_parts():
        if part.kind == "comments":
            continue
        for paragraph, element in part_paragraph_elements(part, view):
            text, markers = view_and_markers(element, view)
            for cid, segment in list(open_segments.items()):
                if segment["part"] != part.name:
                    # a range does not survive into another part: close it at
                    # the end of the part it opened in rather than dragging the
                    # anchor into a header it was never about
                    emit(cid, segment, len(segment["text"]))
                    del open_segments[cid]
                    continue
                if segment["paragraph_index"] != paragraph.index:
                    emit(cid, segment, len(segment["text"]))
                    open_segments[cid] = {"part": part.name,
                                          "paragraph_index": paragraph.index,
                                          "start": 0, "text": text}
            for marker in markers:
                comment = by_id.get(marker.id)
                if comment is None:
                    continue                   # anchor for a comment that is gone
                if marker.kind == "range_start":
                    open_segments[marker.id] = {"part": part.name,
                                                "paragraph_index": paragraph.index,
                                                "start": marker.offset, "text": text}
                elif marker.kind == "range_end":
                    segment = open_segments.pop(marker.id, None)
                    if segment is not None:
                        emit(marker.id, segment, marker.offset)
                elif marker.kind == "reference" and not comment.anchors:
                    comment.anchors.append(Anchor(
                        part=part.name, paragraph_index=paragraph.index,
                        start=marker.offset, end=marker.offset, text="",
                        is_point=True))

    for cid, segment in open_segments.items():     # never closed: to the end
        emit(cid, segment, len(segment["text"]))
