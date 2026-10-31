"""Units and colours: ``Pt``, ``Inches``, ``Cm``, ``Mm``, ``Emu``, ``RGBColor``.

A ``Length`` is an int of English Metric Units (914400 to the inch), which is
what python-docx uses too — so ``Inches(6) == 5486400`` and arithmetic against
python-docx code behaves identically. Every measurement this library takes or
returns is one of these, which is why ``doc.add_picture(path, width=Inches(4))``
needs no conversion.

``docviews.compat.shared`` re-exports these under python-docx's module name.
"""

from __future__ import annotations

EMU_PER_INCH = 914400
EMU_PER_CM = 360000
EMU_PER_MM = 36000
EMU_PER_PT = 12700
EMU_PER_TWIP = 635


class Length(int):
    """A distance in EMU. ``Length(914400)`` is one inch."""

    def __new__(cls, emu: int):
        return int.__new__(cls, emu)

    @property
    def inches(self) -> float:
        return self / EMU_PER_INCH

    @property
    def cm(self) -> float:
        return self / EMU_PER_CM

    @property
    def mm(self) -> float:
        return self / EMU_PER_MM

    @property
    def pt(self) -> float:
        return self / EMU_PER_PT

    @property
    def emu(self) -> int:
        return int(self)

    @property
    def twips(self) -> int:
        return int(round(self / EMU_PER_TWIP))

    def __repr__(self) -> str:
        return f"Length({int(self)} EMU = {self.pt:g} pt)"


def Inches(value: float) -> Length:
    return Length(int(round(value * EMU_PER_INCH)))


def Cm(value: float) -> Length:
    return Length(int(round(value * EMU_PER_CM)))


def Mm(value: float) -> Length:
    return Length(int(round(value * EMU_PER_MM)))


def Pt(value: float) -> Length:
    return Length(int(round(value * EMU_PER_PT)))


def Emu(value: int) -> Length:
    return Length(int(value))


def Twips(value: int) -> Length:
    return Length(int(value) * EMU_PER_TWIP)


class RGBColor(tuple):
    """An RGB colour, printable as the six hex digits OOXML wants."""

    def __new__(cls, red: int, green: int, blue: int):
        for value in (red, green, blue):
            if not 0 <= int(value) <= 255:
                raise ValueError("RGB values are 0..255")
        return super().__new__(cls, (int(red), int(green), int(blue)))

    @property
    def red(self) -> int:
        return self[0]

    @property
    def green(self) -> int:
        return self[1]

    @property
    def blue(self) -> int:
        return self[2]

    def __str__(self) -> str:
        return f"{self[0]:02X}{self[1]:02X}{self[2]:02X}"

    def __repr__(self) -> str:
        return f"RGBColor(0x{self[0]:02x}, 0x{self[1]:02x}, 0x{self[2]:02x})"


class Inches_(Length):
    pass
