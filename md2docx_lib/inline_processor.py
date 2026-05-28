"""
Inline Markdown processor — strips leftover **, *, ` markers from text
and applies corresponding Word formatting within paragraphs.

Handles:
  - **bold text**      → bold run in Word
  - __bold text__      → bold run (alt syntax)
  - *italic text*      → italic run
  - _italic text_      → italic run (alt syntax)
  - `inline code`      → monospace run
  - ~~strikethrough~~  → strikethrough (if supported)
  - ***bold italic***  → bold+italic

Used when building Word paragraphs from text that may contain
leftover markdown syntax (e.g. from DeepSeek clipboard plain-text paste).
"""

import re
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ─── Token types ────────────────────────────────────────────────────────────────

TOKEN_PATTERNS = [
    # Bold+Italic (*** or ___)
    ("bold_italic", re.compile(r"\*\*\*(.+?)\*\*\*|___(.+?)___")),
    # Bold (** or __)
    ("bold",        re.compile(r"\*\*(.+?)\*\*|__(.+?)__")),
    # Italic (* or _) — but NOT ** or __ (must not be followed by another * or _)
    ("italic",      re.compile(r"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)|(?<!_)_(?!_)(.+?)(?<!_)_(?!_)")),
    # Inline code (`)
    ("code",        re.compile(r"`([^`]+)`")),
    # Strikethrough (~~)
    ("strikethrough", re.compile(r"~~(.+?)~~")),
]


def process_inline_paragraph(para, text: str):
    """Build a Word paragraph from text, applying inline markdown formatting.

    Args:
        para: python-docx Paragraph object
        text: Raw text potentially containing **, *, ` markers

    The paragraph will be populated with formatted runs.
    Each **bold** segment becomes a bold run, *italic* becomes italic, etc.
    Plain text between markers becomes normal runs.
    """
    tokens = _tokenize(text)
    para.clear()

    for token_type, token_text in tokens:
        run = para.add_run(token_text)
        if token_type == "bold":
            run.font.bold = True
        elif token_type == "italic":
            run.font.italic = True
        elif token_type == "bold_italic":
            run.font.bold = True
            run.font.italic = True
        elif token_type == "code":
            run.font.name = "Consolas"
            run.font.size = Pt(9.5)
            run.font.color.rgb = RGBColor(180, 60, 60)
        elif token_type == "strikethrough":
            _apply_strikethrough(run)


def clean_text(text: str) -> str:
    """Remove markdown formatting markers from plain text.
    Strips **, *, __, _, ``, ~~ while preserving the inner content.

    Useful for cleaning clipboard text before display or plain-text export.
    """
    result = text
    # Remove bold markers
    result = re.sub(r"\*\*(.+?)\*\*", r"\1", result)
    result = re.sub(r"__(.+?)__", r"\1", result)
    # Remove italic markers
    result = re.sub(r"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)", r"\1", result)
    result = re.sub(r"(?<!_)_(?!_)(.+?)(?<!_)_(?!_)", r"\1", result)
    # Remove code markers
    result = re.sub(r"`([^`]+)`", r"\1", result)
    # Remove strikethrough
    result = re.sub(r"~~(.+?)~~", r"\1", result)
    # Remove bold+italic
    result = re.sub(r"\*\*\*(.+?)\*\*\*", r"\1", result)
    result = re.sub(r"___(.+?)___", r"\1", result)

    return result


def is_clean_text(text: str) -> bool:
    """Check if text contains any markdown inline markers."""
    return not bool(re.search(r"\*\*|__|\*(?!\*)|\b_\b|`|~~", text))


# ─── Internal tokenizer ─────────────────────────────────────────────────────────

def _tokenize(text: str) -> list[tuple[str, str]]:
    """Break text into (type, content) tokens based on markdown markers.

    Returns list of (type, text) where type is one of:
      'plain', 'bold', 'italic', 'bold_italic', 'code', 'strikethrough'
    """
    tokens = []
    pos = 0
    n = len(text)

    while pos < n:
        earliest_match = None
        earliest_type = None
        earliest_start = n

        # Find the next marker
        for tok_type, pattern in TOKEN_PATTERNS:
            m = pattern.search(text, pos)
            if m:
                if m.start() < earliest_start:
                    earliest_start = m.start()
                    earliest_type = tok_type
                    earliest_match = m

        if earliest_match is None:
            # No more markers — rest is plain text
            if pos < n:
                tokens.append(("plain", text[pos:]))
            break

        # Add plain text before the match
        if earliest_start > pos:
            tokens.append(("plain", text[pos:earliest_start]))

        # Add the matched token
        content = earliest_match.group(1) or earliest_match.group(2) or ""
        tokens.append((earliest_type, content))

        pos = earliest_match.end()

    return tokens


def _apply_strikethrough(run):
    """Apply strikethrough formatting to a run via XML."""
    rPr = run._element.get_or_add_rPr()
    strike = OxmlElement("w:strike")
    strike.set(qn("w:val"), "true")
    rPr.append(strike)


# ─── Integration helper for builder ─────────────────────────────────────────────

def build_paragraph_from_text(doc_para, text: str):
    """Convenience: clear a paragraph and rebuild with inline formatting.

    Usage from builder_docx.py:
        from md2docx_lib.inline_processor import build_paragraph_from_text
        para = doc.add_paragraph()
        build_paragraph_from_text(para, raw_text)
    """
    # First, clean any double-processing issues
    text = text.strip()
    process_inline_paragraph(doc_para, text)
