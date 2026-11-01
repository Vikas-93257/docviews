"""``docviews.compat.enum`` — python-docx's enum names and protocol.

Members of the documented ``BaseEnum`` and ``BaseXmlEnum`` classes preserve
python-docx's integer values, ``str()``, ``repr()``, XML tokens (including its
``"UNMAPPED"`` sentinel), and ``from_xml``/``to_xml`` behavior. The deprecated
``WD_BREAK_TYPE`` is built as a plain :class:`enum.Enum`, like python-docx; its
writer behavior is implemented explicitly by ``Run.add_break`` rather than by
adding an ``xml_value`` property that the upstream enum does not have.

The published ``MSO_THEME_COLOR_INDEX`` page also lists ``MIXED`` although
python-docx 1.2.0 omits it at runtime. This package includes the documented
member with Office's value ``-2``; the runtime discrepancy is recorded in the
parity inventory.

``from docx.enum.text import WD_ALIGN_PARAGRAPH`` has a one-word answer too —
the same names are importable from ``docviews.compat.enum.text``, ``.style``,
``.section``, ``.table``, ``.shape`` and ``.dml``.

An ``ooxml`` property is a docviews alias for ``xml_value`` so a compat member
can be handed to the core writer. It does not replace or normalize the upstream
XML metadata.
"""

from __future__ import annotations

import enum

from ._generated import (ALIASES, BASES, DOCUMENTED_ONLY_MEMBER_DOCSTRINGS,
                        ENUMS, HOME, TOKENS)
from .base import BaseEnum, BaseXmlEnum, PlainXmlEnum

__all__ = sorted(set(ENUMS) | set(ALIASES)
                 | {"BaseEnum", "BaseXmlEnum", "PlainXmlEnum"})

# A class is built once, under python-docx's own name for it; names python-docx
# re-exports point to that same class afterwards (``WD_BREAK is WD_BREAK_TYPE``).
CANONICAL = {name: ALIASES.get(name, name) for name in ENUMS}


def _token(cls_name: str, member: str) -> str | None:
    """Return python-docx's exact ``xml_value``, including ``UNMAPPED``."""
    return (TOKENS.get(cls_name) or {}).get(member)


def _build(cls_name: str, values: dict) -> type:
    """A real enum class with python-docx's members, values and base shape."""
    kind = BASES.get(cls_name)
    if kind == "PlainEnum":
        base = enum.Enum
    elif kind == "BaseXmlEnum":
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
        docstring = (DOCUMENTED_ONLY_MEMBER_DOCSTRINGS.get(cls_name, {})
                     .get(member, ""))
        if base is BaseXmlEnum:
            prepared[member] = (value, _token(cls_name, member), docstring)
        elif base is enum.Enum:
            prepared[member] = value
        else:
            prepared[member] = (value, docstring)
    return enum.EnumMeta(cls_name, (base,), prepared)


_CLASSES: dict = {}
for _name, _values in ENUMS.items():
    _class_name = CANONICAL[_name]
    if _class_name not in _CLASSES:
        _CLASSES[_class_name] = _build(_class_name, _values)
    globals()[_name] = _CLASSES[_class_name]
del _name, _values, _class_name
