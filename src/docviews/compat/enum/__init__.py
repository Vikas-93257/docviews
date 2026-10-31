"""``docviews.compat.enum`` — python-docx's enums, same objects, same protocol.

``WD_ALIGN_PARAGRAPH.CENTER`` is ``1``, ``int()``/``==``/sorting work on members,
``str()`` prints ``"CENTER (1)"``, ``list(WD_ALIGN_PARAGRAPH)`` iterates the
members, ``WD_ALIGN_PARAGRAPH(1)`` looks one up, ``.xml_value`` carries the
token Word's XML wants, ``.from_xml("center")``/``.to_xml(...)`` go both ways,
and ``isinstance(member, BaseXmlEnum)`` is true — because the members here *are*
instances of the same two base classes (see :mod:`.base`), built from the
generated tables in ``_generated.py``.

The classes are built, not stub objects, so the whole documented python-docx
enum protocol keeps working: ``__members__``, iteration, ``from_xml``, ``to_xml``.

``from docx.enum.text import WD_ALIGN_PARAGRAPH`` has a one-word answer too — the
same names are importable from ``docviews.compat.enum.text``, ``.style``,
``.section``, ``.table``, ``.shape`` and ``.dml``, in the modules python-docx
keeps them in.

Two deliberate differences from python-docx, both about the file rather than the
API, and both listed in ``COMPAT.md``:

* ``WD_COLOR_INDEX.AUTO``/``INHERITED`` carry the token ``"none"`` (what
  ``w:highlight`` accepts) where python-docx says ``"default"``;
* ``WD_BREAK_TYPE`` / ``WD_BREAK`` carry real tokens (``"page"``, ``"column"``,
  …) where python-docx's deprecated class has none, so
  ``run.add_break(WD_BREAK.PAGE)`` writes ``<w:br w:type="page"/>``; the members
  still print the way python-docx's do (``"WD_BREAK_TYPE.PAGE"``).

An ``ooxml`` property exposes the same token under the core API's name, so a
compat member can be handed straight to a ``docviews`` setter.
"""

from __future__ import annotations

import enum

from ._generated import ALIASES, BASES, ENUMS, HOME, TOKENS
from .base import BaseEnum, BaseXmlEnum, PlainXmlEnum

__all__ = sorted(set(ENUMS) | set(ALIASES)
                 | {"BaseEnum", "BaseXmlEnum", "PlainXmlEnum"})

# Members whose token we do not take from python-docx. Keyed by the canonical
# class name; every entry is a judgement about what the file wants, and each one
# is checked by tests/test_enums.py.
OVERRIDES = {
    # python-docx's legacy WD_BREAK_TYPE maps no XML value; these are the
    # tokens <w:br w:type="…"/> accepts.
    "WD_BREAK_TYPE": {2: "nextPage", 3: "continuous", 4: "evenPage",
                      5: "oddPage", 6: "textWrapping", 7: "page",
                      8: "column", 9: "lineClearLeft", 10: "lineClearRight",
                      11: "lineClearAll"},
    # "default" is python-docx's word for "auto"; w:highlight wants "none".
    "WD_COLOR_INDEX": {-1: "none", 0: "none"},
    "WD_UNDERLINE": {-1: "none"},
}

# the old name of OVERRIDES, kept because scripts import it
OOXML = OVERRIDES

# A class is built once, under python-docx's own name for it; names python-docx
# re-exports are pointed at the same class afterwards, so identity holds
# (``WD_ALIGN_PARAGRAPH is WD_PARAGRAPH_ALIGNMENT``).
CANONICAL = {name: ALIASES.get(name, name) for name in ENUMS}


def _token(cls_name: str, member: str, value) -> str | None:
    overrides = OVERRIDES.get(cls_name) or {}
    if isinstance(value, int) and value in overrides:
        return overrides[value]
    return (TOKENS.get(cls_name) or {}).get(member)


def _build(cls_name: str, values: dict) -> type:
    """A real enum class with python-docx's members, values and bases."""
    # python-docx's own base — unless we carry tokens for the class, in which
    # case the members need somewhere to keep them (its deprecated
    # WD_BREAK_TYPE maps no XML value of its own).
    kind = BASES.get(cls_name)
    if kind == "PlainEnum":
        base = PlainXmlEnum
    elif kind == "BaseXmlEnum" or OVERRIDES.get(cls_name) \
            or TOKENS.get(cls_name):
        base = BaseXmlEnum
    else:
        base = BaseEnum
    prepared = enum.EnumMeta.__prepare__(cls_name, (base,))
    prepared["__module__"] = __name__
    prepared["__qualname__"] = cls_name
    prepared["__doc__"] = f"{cls_name} — python-docx's members and values."
    for member, value in values.items():
        if not isinstance(value, int):
            raise TypeError(f"{cls_name}.{member}: python-docx value {value!r} "
                            f"is not an int")
        token = _token(cls_name, member, value)
        prepared[member] = ((value, token, "") if issubclass(base, BaseXmlEnum)
                            else (value, ""))
    return enum.EnumMeta(cls_name, (base,), prepared)


_CLASSES: dict = {}
for _name, _values in ENUMS.items():
    _class_name = CANONICAL[_name]
    if _class_name not in _CLASSES:
        _CLASSES[_class_name] = _build(_class_name, _values)
    globals()[_name] = _CLASSES[_class_name]
del _name, _values, _class_name
