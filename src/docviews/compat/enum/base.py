"""python-docx's enum base classes, same contract, our own implementation.

python-docx numbers its enum members the way the MS API does and hangs the
*Word* XML attribute value off the member as ``xml_value``; its two base classes
are part of the documented API (``docx.enum.base``), and code does reach for
them — ``isinstance(member, BaseXmlEnum)``, ``WD_ALIGN_PARAGRAPH.from_xml(...)``,
``SomeEnum.to_xml(value)``. All of that works on the members in
``docviews.compat.enum``, because those members are instances of these classes.

``ooxml`` is ours: the same token, under the name the core API uses for it, so a
compat member can be handed straight to a docviews setter.
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
    """A deprecated python-docx enum: standard-library ``str``, tokens kept.

    python-docx's ``WD_BREAK_TYPE`` derives from neither documented base, so its
    members print ``"WD_BREAK_TYPE.PAGE"``. Ours do too — and they also carry the
    token, which is what ``run.add_break(WD_BREAK.PAGE)`` writes.
    """

    def __str__(self) -> str:
        return f"{type(self).__name__}.{self.name}"
