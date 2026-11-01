"""``docx.shared`` under the compat door — same objects, one implementation.

The units live in :mod:`docviews.units`; this module exists so that code doing
``from docx.shared import Pt`` finds it after ``import docviews.compat as docx``.

Two decorators python-docx also ships from ``docx.shared`` are here as well,
because ``from docx.shared import lazyproperty`` appears in code that mixes
python-docx internals with its own — ``lazyproperty`` (a property computed once,
then cached) and ``write_only_property`` (assignable, not readable). Both are
implemented here rather than copied; their behaviour is python-docx's.

``ElementProxy``, ``Parented``, ``StoryChild`` and ``TextAccumulator`` are not
provided: they exist to hand out lxml element proxies, the same boundary the
compat door draws at ``part``/``element``.
"""

from __future__ import annotations

import functools
from typing import Any, Callable

from ..units import (EMU_PER_CM, EMU_PER_INCH, EMU_PER_MM, EMU_PER_PT,
                     EMU_PER_TWIP, Length as _Length)


Length = _Length


class Inches(_Length):
    def __new__(cls, inches: float):
        return _Length.__new__(cls, int(inches * EMU_PER_INCH))


class Cm(_Length):
    def __new__(cls, cm: float):
        return _Length.__new__(cls, int(cm * EMU_PER_CM))


class Mm(_Length):
    def __new__(cls, mm: float):
        return _Length.__new__(cls, int(mm * EMU_PER_MM))


class Pt(_Length):
    def __new__(cls, points: float):
        return _Length.__new__(cls, int(points * EMU_PER_PT))


class Emu(_Length):
    def __new__(cls, emu: int):
        return _Length.__new__(cls, int(emu))


class Twips(_Length):
    def __new__(cls, twips: float):
        return _Length.__new__(cls, int(twips * EMU_PER_TWIP))


class RGBColor(tuple):
    """python-docx-compatible immutable RGB tuple."""

    def __new__(cls, r: int, g: int, b: int):
        message = "RGBColor() takes three integer values 0-255"
        for value in (r, g, b):
            if not isinstance(value, int):
                raise TypeError(message)
            if value < 0 or value > 255:
                raise ValueError(message)
        return super().__new__(cls, (r, g, b))

    def __repr__(self):
        return "RGBColor(0x%02x, 0x%02x, 0x%02x)" % self

    def __str__(self):
        return "%02X%02X%02X" % self

    @classmethod
    def from_string(cls, rgb_hex_str: str) -> RGBColor:
        r = int(rgb_hex_str[:2], 16)
        g = int(rgb_hex_str[2:4], 16)
        b = int(rgb_hex_str[4:], 16)
        return cls(r, g, b)


__all__ = ["Length", "Inches", "Cm", "Mm", "Pt", "Emu", "Twips", "RGBColor",
           "EMU_PER_INCH", "EMU_PER_CM", "EMU_PER_MM", "EMU_PER_PT",
           "EMU_PER_TWIP", "lazyproperty", "write_only_property"]

# The names python-docx defines here that are plumbing, not surface.
NOT_PROVIDED = ("ElementProxy", "Parented", "StoryChild", "TextAccumulator")


def write_only_property(fset: Callable[[Any, Any], None]):
    """``@write_only_property`` — assignment works, reading raises."""

    return property(fset=fset, doc=fset.__doc__)


class lazyproperty:
    """``@lazyproperty`` — evaluated on first access, then cached.

    A data descriptor, so the cached value cannot be overwritten by assignment;
    assigning raises ``AttributeError``, as it does in python-docx.
    """

    def __init__(self, fget: Callable[[Any], Any]) -> None:
        self._fget = fget
        self._name = fget.__name__
        functools.update_wrapper(self, fget)

    def __get__(self, obj: Any, owner: Any = None) -> Any:
        if obj is None:
            return self
        cached = obj.__dict__.get(self._name)
        if cached is None:
            cached = self._fget(obj)
            obj.__dict__[self._name] = cached
        return cached

    def __set__(self, obj: Any, value: Any) -> None:
        raise AttributeError(f"can't set attribute {self._name!r}")
