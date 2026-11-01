"""python-docx's enum base classes, implemented locally.

python-docx numbers its enum members the way the MS API does and stores the
Word XML attribute value as ``xml_value``. Its two documented bases are part of
the public API (``docx.enum.base``), and code may use ``isinstance`` or call
``from_xml``/``to_xml``. ``ooxml`` is a docviews alias for ``xml_value``.

Legacy python-docx enums that derive directly from :class:`enum.Enum`, notably
``WD_BREAK_TYPE``, are built directly on that standard-library base. They have
no ``xml_value`` or integer behavior.
"""

from __future__ import annotations

import enum

__all__ = ["BaseEnum", "BaseXmlEnum", "PlainXmlEnum"]


class BaseEnum(int, enum.Enum):
    """Base for enums whose members do not map to an XML attribute value."""

    def __new__(cls, ms_api_value: int, docstr: str = ""):
        self = int.__new__(cls, ms_api_value)
        self._value_ = ms_api_value
        self.__doc__ = (docstr or "").strip()
        return self

    def __str__(self) -> str:
        return f"{self.name} ({self.value})"


class BaseXmlEnum(int, enum.Enum):
    """Base for enums whose members carry the XML token Word writes."""

    xml_value: str | None

    def __new__(cls, ms_api_value: int, xml_value: str | None = None,
                docstr: str = ""):
        self = int.__new__(cls, ms_api_value)
        self._value_ = ms_api_value
        self.xml_value = xml_value
        self.__doc__ = (docstr or "").strip()
        return self

    def __str__(self) -> str:
        return f"{self.name} ({self.value})"

    @property
    def ooxml(self) -> str | None:
        """The same token as ``xml_value``, under docviews' name for it."""
        return self.xml_value

    @classmethod
    def from_xml(cls, xml_value: str | None):
        """The member Word's ``xml_value`` token names."""
        member = next((m for m in cls if m.xml_value == xml_value), None)
        if member is None:
            raise ValueError(f"{cls.__name__} has no XML mapping for "
                             f"{xml_value!r}")
        return member

    @classmethod
    def to_xml(cls, value) -> str:
        """The token to write for a member (or a member's value)."""
        member = cls(value)
        if not member.xml_value:
            raise ValueError(f"{cls.__name__}.{member.name} has no XML "
                             f"representation")
        return member.xml_value


class PlainXmlEnum(BaseXmlEnum):
    """Legacy docviews helper: an integer XML enum with standard Enum ``str``."""

    def __str__(self) -> str:
        return f"{type(self).__name__}.{self.name}"
