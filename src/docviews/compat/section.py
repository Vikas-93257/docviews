"""python-docx-shaped section and header/footer proxies."""

from . import _HeaderFooter, _Section, _Sections

Section = _Section
Sections = _Sections
_Header = _HeaderFooter
_Footer = _HeaderFooter

__all__ = ["Section", "Sections", "_Header", "_Footer"]
