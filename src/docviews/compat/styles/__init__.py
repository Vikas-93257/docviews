"""python-docx-shaped style proxies."""

from .latent import LatentStyles
from .style import BaseStyle, CharacterStyle, ParagraphStyle
from .styles import Styles

__all__ = ["LatentStyles", "BaseStyle", "CharacterStyle", "ParagraphStyle", "Styles"]
