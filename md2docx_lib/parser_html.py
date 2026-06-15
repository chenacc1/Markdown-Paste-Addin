"""HTML content parser — tables, images, headings, lists, code blocks, math.

Produces the same chunk dict format as parser_markdown.
"""

import re
from urllib.parse import urljoin

_HAS_HTML = False
try:
    from bs4 import BeautifulSoup, Tag
    _HAS_HTML = True
except ImportError:
    pass

from .parser_math import extract_math


def parse_html(html: str, base_url: str = "") -> list[dict]:
    """Parse HTML into chunks of text, table, image, mermaid, code, etc."""
    if not _HAS_HTML:
        print("Error: beautifulsoup4 not installed. Run: pip install beautifulsoup4 html5lib")
        return [{"type": "text", "text": html}]

    # Pre-process: replace <br> tags with newlines so tree structures survive
    html = re.sub(r"<br\s*/?>", "\n", html, flags=re.IGNORECASE)

    soup = BeautifulSoup(html, "html5lib")
    chunks = []

    for tag in soup.find_all(["script", "style"]):
        tag.decompose()

    body = soup.find("body") or soup
    _walk_html(body, chunks, base_url)

    # Remove empty text chunks
    chunks = [c for c in chunks
              if c.get("type") != "text" or c.get("text", "").strip()]
    return chunks


def _walk_html(element: Tag, chunks: list[dict], base_url: str):
    for child in element.children:
        if isinstance(child, str):
            text = child.strip()
            if text:
                _add_text_with_math(text, chunks)
            continue

        if not isinstance(child, Tag):
            continue

        tag_name = child.name.lower() if child.name else ""

        # Table
        if tag_name == "table":
            table_data = _parse_html_table(child)
            if table_data:
                chunks.append({"type": "table", **table_data})
            continue

        # Image
        if tag_name == "img":
            src = child.get("src", "")
            alt = child.get("alt", "")
            if src:
                chunks.append({"type": "image", "src": urljoin(base_url, src) if base_url else src, "alt": alt})
            continue

        # Math elements (KaTeX/MathJax rendered HTML)
        if _is_math_element(child):
            latex = _extract_latex_from_math(child)
            display = "display" in child.get("class", []) or child.name in ("math", "mrow")
            chunks.append({"type": "math", "latex": latex, "display": display})
            continue

        # Headings
        if tag_name in ("h1", "h2", "h3", "h4", "h5", "h6"):
            level = int(tag_name[1])
            text = child.get_text(strip=True)
            if text:
                _add_text_with_math(text, chunks, default_type="heading", level=level)
            continue

        # Blockquote
        if tag_name == "blockquote":
            text = child.get_text(strip=True)
            if text:
                chunks.append({"type": "blockquote", "text": text})
            continue

        # Horizontal rule
        if tag_name == "hr":
            chunks.append({"type": "hr"})
            continue

        # Lists
        if tag_name == "ul":
            _parse_list(child, "ul", chunks)
            continue
        if tag_name == "ol":
            _parse_list(child, "ol", chunks)
            continue

        # Code blocks
        if tag_name == "pre":
            code_tag = child.find("code")
            code_text = code_tag.get_text() if code_tag else child.get_text()
            classes = code_tag.get("class", []) if code_tag else []
            lang = ""
            for cls in classes:
                if cls.startswith("language-"):
                    lang = cls[len("language-"):]

            if lang == "mermaid":
                chunks.append({"type": "mermaid", "code": code_text.strip()})
            else:
                chunks.append({"type": "code", "language": lang, "code": code_text})
            continue

        # Inline code
        if tag_name == "code" and not child.find_parent("pre"):
            text = child.get_text()
            if text:
                chunks.append({"type": "code", "language": "", "code": text,
                              "inline": True})
            continue

        # Paragraphs / divs / sections
        if tag_name in ("p", "div", "section", "article", "td", "th", "span",
                        "a", "strong", "em", "b", "i", "br"):
            if tag_name == "br":
                if chunks and chunks[-1]["type"] == "text":
                    chunks[-1]["text"] += "\n"
            elif tag_name == "p":
                text = child.get_text(strip=True)
                if text:
                    _add_text_with_math(text, chunks)
            else:
                _walk_html(child, chunks, base_url)
            continue

        _walk_html(child, chunks, base_url)


def _is_math_element(tag: Tag) -> bool:
    name = tag.name.lower() if tag.name else ""
    classes = tag.get("class", [])
    return ("katex" in str(classes).lower() or
            "mathjax" in str(classes).lower() or
            name in ("math", "mi", "mo", "mn", "mrow", "msup", "msub",
                     "mfrac", "msqrt", "mtable", "mtr", "mtd"))


def _extract_latex_from_math(tag: Tag) -> str:
    """Extract LaTeX source from a KaTeX/MathJax HTML element."""
    # KaTeX stores the source in <annotation> tag
    annotation = tag.find("annotation")
    if annotation:
        text = annotation.get_text()
        if text:
            return text.strip()
    # Fallback: use text content
    return tag.get_text().strip()


def _add_text_with_math(text: str, chunks: list[dict], **extra):
    """Add text to chunks, extracting any LaTeX math expressions first."""
    math_blocks = extract_math(text)
    if not math_blocks:
        # No math — add as plain text or heading
        if extra.get("default_type") == "heading":
            chunks.append({"type": "heading", "level": extra["level"], "text": text})
        else:
            chunks.append({"type": "text", "text": text})
        return

    # Replace math blocks with placeholders (from end to avoid offset shifts)
    for i in range(len(math_blocks) - 1, -1, -1):
        mb = math_blocks[i]
        placeholder = f"MATH_PLACEHOLDER_{i}"
        text = text[:mb.start] + placeholder + text[mb.end:]

    # Split by placeholders and emit alternating text/math chunks
    parts = text.split("MATH_PLACEHOLDER_")
    for idx, part in enumerate(parts):
        if not part:
            continue
        # Check if this part starts with a placeholder index
        # Format: "0rest_of_text" or just "rest_of_text"
        num_str = ""
        rest = part
        for ch in part:
            if ch.isdigit():
                num_str += ch
            else:
                break
        if num_str and len(num_str) < len(part):
            # Has a placeholder index followed by text
            math_idx = int(num_str)
            rest = part[len(num_str):]
            chunks.append({"type": "math", "latex": math_blocks[math_idx].text,
                          "display": math_blocks[math_idx].display})
            if rest:
                chunks.append({"type": "text", "text": rest})
        elif num_str and len(num_str) == len(part):
            # Entire part is a placeholder index
            math_idx = int(num_str)
            chunks.append({"type": "math", "latex": math_blocks[math_idx].text,
                          "display": math_blocks[math_idx].display})
        else:
            # Pure text
            if extra.get("default_type") == "heading":
                chunks.append({"type": "heading", "level": extra["level"], "text": part})
            else:
                chunks.append({"type": "text", "text": part})


def _parse_list(list_tag: Tag, list_type: str, chunks: list[dict], level: int = 0):
    """Recursively parse <ul>/<ol> into text chunks preserving nesting."""
    for li in list_tag.find_all("li", recursive=False):
        # Get text excluding nested lists
        text_parts = []
        nested = None
        for child in li.children:
            if isinstance(child, str):
                text_parts.append(child.strip())
            elif isinstance(child, Tag) and child.name in ("ul", "ol"):
                nested = child
                break
            elif isinstance(child, Tag):
                text_parts.append(child.get_text(strip=True))

        prefix = "  " * level + ("- " if list_type == "ul" else "1. ")
        text = " ".join(t for t in text_parts if t)
        if text:
            chunks.append({"type": "text", "text": prefix + text})

        if nested:
            _parse_list(nested, "ol" if nested.name == "ol" else "ul", chunks, level + 1)


def _parse_html_table(table_tag: Tag) -> dict | None:
    rows = []
    headers = []
    alignments = []

    thead = table_tag.find("thead")
    if thead:
        for tr in thead.find_all("tr"):
            headers = [td.get_text(strip=True) for td in tr.find_all(["th", "td"])]
            for td in tr.find_all(["td", "th"]):
                align = (td.get("align") or td.get("style", "")).lower()
                if "center" in align:
                    alignments.append("center")
                elif "right" in align:
                    alignments.append("right")
                else:
                    alignments.append("left")
            break

    tbody = table_tag.find("tbody") or table_tag
    for tr in tbody.find_all("tr"):
        if tr.find_parent("thead"):
            continue
        cells = [td.get_text(strip=True) for td in tr.find_all(["td", "th"])]
        if not cells:
            continue
        if not headers:
            headers = cells
            for td in tr.find_all(["td", "th"]):
                align = (td.get("align") or td.get("style", "")).lower()
                if "center" in align:
                    alignments.append("center")
                elif "right" in align:
                    alignments.append("right")
                else:
                    alignments.append("left")
            continue
        rows.append(cells)

    if not headers:
        return None
    if not alignments:
        alignments = ["left"] * len(headers)
    while len(alignments) < len(headers):
        alignments.append("left")

    return {"headers": headers, "alignments": alignments, "rows": rows}
