"""Exceptions used by the python-docx compatibility surface."""


class PythonDocxError(Exception):
    """Base class for errors exposed by python-docx's public exceptions module."""


class InvalidSpanError(PythonDocxError):
    """Raised when a requested table-cell merge region is not rectangular."""


class InvalidXmlError(PythonDocxError):
    """Raised when required WordprocessingML content is missing or invalid."""
