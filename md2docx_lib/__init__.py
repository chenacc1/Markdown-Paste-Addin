"""md2docx_lib — Markdown/HTML to Word document conversion library."""

__version__ = "2.0.0"

from .parser_markdown import parse_markdown
from .parser_html import parse_html
from .parser_math import extract_math, MathBlock
from .renderer_mermaid import render_mermaid
from .renderer_code import highlight_code
from .renderer_image import download_image
from .builder_docx import build_docx, DocumentBuilder
from .builder_toc import add_toc
from .builder_numbering import NumberingTracker
from .formatter import format_document, FORMAT, detect_paragraph_type, apply_format
from .template import load_template
