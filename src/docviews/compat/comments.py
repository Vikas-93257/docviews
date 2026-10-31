"""python-docx-shaped comment collection and comment proxies.

The implementations live in :mod:`docviews.compat` so they can wrap its
editable document, paragraph, table, and run objects. This module preserves the
usual ``from docx.comments import Comment, Comments`` import pattern as
``from docviews.compat.comments import Comment, Comments``.
"""

from . import Comment, Comments

__all__ = ["Comment", "Comments"]
