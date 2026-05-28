"""Code syntax highlighting using Pygments — maps tokens to Word formatting."""

from docx.shared import Pt, RGBColor

_HAS_PYGMENTS = False
try:
    from pygments import highlight
    from pygments.lexers import get_lexer_by_name, TextLexer
    from pygments.token import Token
    from pygments.formatters import RawTokenFormatter
    _HAS_PYGMENTS = True
except ImportError:
    pass


# ─── Color scheme: Solarized-light-like ─────────────────────────────────────────

_TOKEN_STYLE = {
    Token.Keyword:       {"color": RGBColor(0x00, 0x77, 0x00), "bold": True},
    Token.Keyword.Type:  {"color": RGBColor(0x00, 0x77, 0x00), "bold": False},
    Token.Name.Builtin:  {"color": RGBColor(0x0A, 0x58, 0x9A)},
    Token.Name.Function: {"color": RGBColor(0x00, 0x5C, 0xBB)},
    Token.Name.Class:    {"color": RGBColor(0x00, 0x4E, 0x8C), "bold": True},
    Token.Name.Decorator: {"color": RGBColor(0xAF, 0x66, 0x00)},
    Token.String:        {"color": RGBColor(0xCB, 0x4B, 0x16)},
    Token.String.Doc:    {"color": RGBColor(0xCB, 0x4B, 0x16), "italic": True},
    Token.Number:        {"color": RGBColor(0x66, 0x33, 0xCC)},
    Token.Operator:      {"color": RGBColor(0x00, 0x00, 0x00)},
    Token.Comment:       {"color": RGBColor(0x99, 0x99, 0x99), "italic": True},
    Token.Generic.Heading: {"color": RGBColor(0x00, 0x00, 0x00), "bold": True},
    Token.Generic.Subheading: {"color": RGBColor(0x00, 0x00, 0x00), "bold": True},
    Token.Generic.Deleted: {"color": RGBColor(0xDC, 0x32, 0x2F)},
    Token.Generic.Inserted: {"color": RGBColor(0x00, 0x77, 0x00)},
    Token.Generic.Emph:  {"italic": True},
    Token.Generic.Strong: {"bold": True},
    Token.Error:         {"color": RGBColor(0xFF, 0x00, 0x00), "underline": True},
}


def highlight_code(code: str, language: str = "") -> list[list[dict]]:
    """Highlight code and return as lines of styled tokens.

    Each token is a dict: {'text': str, 'font': {...}}
    Returns list of lines, each line is a list of token dicts.

    If Pygments is not installed, returns plain text tokens.
    """
    tokens = _tokenize(code, language)
    return _tokens_to_lines(tokens)


def _tokenize(code: str, language: str):
    if not _HAS_PYGMENTS:
        return [("", code)]

    try:
        lexer = get_lexer_by_name(language or "text", stripall=True)
    except Exception:
        lexer = TextLexer(stripall=True)

    try:
        return list(lexer.get_tokens(code))
    except Exception:
        return [("", code)]


def _tokens_to_lines(pygments_tokens) -> list[list[dict]]:
    lines = []
    current_line = []

    for ttype, value in pygments_tokens:
        if "\n" in value:
            parts = value.split("\n")
            for j, part in enumerate(parts):
                if j > 0:
                    lines.append(current_line)
                    current_line = []
                if part:
                    style = _get_style(ttype)
                    current_line.append({"text": part, "style": style})
        else:
            style = _get_style(ttype)
            current_line.append({"text": value, "style": style})

    if current_line:
        lines.append(current_line)
    return lines


def _get_style(token_type) -> dict:
    """Find best-matching style for a token type."""
    # Walk up the token hierarchy
    t = token_type
    while t is not Token:
        if t in _TOKEN_STYLE:
            return _TOKEN_STYLE[t]
        t = t.parent
    return {}
