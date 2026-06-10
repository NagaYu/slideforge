"""SlideForge: Markdown -> professional .pptx generator."""

from .parser import parse_markdown
from .renderer import Renderer
from .themes import THEMES, get_theme, list_themes

__version__ = "0.2.0"
__all__ = ["parse_markdown", "Renderer", "THEMES", "get_theme",
           "list_themes", "__version__"]
