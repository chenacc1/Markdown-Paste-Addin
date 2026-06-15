"""Enhanced Markdown parser — tables, mermaid, math, code blocks, task lists,
blockquotes, nested lists, horizontal rules, images, headings, and text."""

import re
from .parser_math import extract_math, MathBlock


def parse_markdown(text: str) -> list[dict]:
    """Parse markdown text into typed chunks.

    Returns list of dicts with keys:
      - type: 'text', 'table', 'mermaid', 'image', 'math', 'code',
              'task_list', 'blockquote', 'hr', 'heading'
      - plus type-specific data
    """
    # Pre-extract math blocks and replace with placeholders to avoid
    # conflicts with other syntax (e.g. $ inside code blocks, tables)
    math_blocks = extract_math(text)
    math_placeholder_map = {}
    # Replace from end to start to preserve earlier positions
    for idx in range(len(math_blocks) - 1, -1, -1):
        mb = math_blocks[idx]
        placeholder = f"MATH_PLACEHOLDER_{idx}"
        math_placeholder_map[placeholder] = mb
        text = text[:mb.start] + placeholder + text[mb.end:]

    lines = text.split("\n")
    chunks = []
    i = 0

    while i < len(lines):
        line = lines[i]

        # Mermaid code block
        if _is_mermaid_start(line):
            code_lines = []
            i += 1
            while i < len(lines) and not lines[i].strip().startswith("```"):
                code_lines.append(lines[i])
                i += 1
            chunks.append({"type": "mermaid", "code": "\n".join(code_lines)})
            if i < len(lines):
                i += 1  # skip closing ```
            continue

        # Regular code block
        if _is_code_block_start(line):
            lang, code_lines = _read_code_block(lines, i)
            i += len(code_lines) + 2  # opening + content + closing ```
            chunks.append({"type": "code", "language": lang, "code": "\n".join(code_lines)})
            continue

        # Markdown table
        if _is_table_line(line):
            table_lines = [line]
            i += 1
            while i < len(lines) and _is_table_line(lines[i]):
                table_lines.append(lines[i])
                i += 1
            parsed = _parse_md_table(table_lines)
            if parsed:
                chunks.append({"type": "table", **parsed})
            continue

        # Horizontal rule
        if _is_hr(line):
            chunks.append({"type": "hr"})
            i += 1
            continue

        # Image syntax (standalone)
        img_match = re.match(r"^!\[([^\]]*)\]\(([^)]+)\)", line.strip())
        if img_match:
            chunks.append({"type": "image", "alt": img_match.group(1), "src": img_match.group(2)})
            i += 1
            continue

        # Blockquote
        if line.strip().startswith("> "):
            quote_lines = []
            while i < len(lines) and lines[i].strip().startswith("> "):
                quote_lines.append(lines[i].strip()[2:])
                i += 1
            chunks.append({"type": "blockquote", "text": "\n".join(quote_lines)})
            continue

        # Task list
        if _is_task_item(line):
            task_items = []
            while i < len(lines) and _is_task_item(lines[i]):
                m = re.match(r"^\s*(-|\*|\+)\s+\[([ xX])\]\s+(.*)", lines[i])
                if m:
                    task_items.append({
                        "checked": m.group(2).lower() == "x",
                        "text": m.group(3).strip(),
                        "indent": len(lines[i]) - len(lines[i].lstrip()),
                    })
                i += 1
            chunks.append({"type": "task_list", "items": task_items})
            continue

        # Heading detection (line starts with #)
        heading_match = re.match(r"^(#{1,6})\s+(.*)", line)
        if heading_match:
            level = len(heading_match.group(1))
            chunks.append({"type": "heading", "level": level,
                          "text": heading_match.group(2).strip()})
            i += 1
            continue

        # Accumulate text lines (including nested lists, which get inline markers)
        text_lines = [line]
        i += 1
        while i < len(lines):
            stripped = lines[i].strip()
            # Stop at any block-level start
            if (_is_mermaid_start(stripped) or _is_code_block_start(stripped) or
                _is_table_line(stripped) or _is_hr(stripped) or
                stripped.startswith("> ") or _is_task_item(stripped) or
                re.match(r"^(#{1,6})\s+", stripped) or
                re.match(r"^!\[([^\]]*)\]\(([^)]+)\)$", stripped)):
                break
            text_lines.append(lines[i])
            i += 1

        chunk_text = "\n".join(text_lines).strip()
        if chunk_text:
            _add_text_with_math(chunks, chunk_text, math_placeholder_map)

    return chunks


def _add_text_with_math(chunks: list[dict], text: str,
                        placeholder_map: dict[str, MathBlock]):
    """Split text containing math placeholders into alternating text/math chunks."""
    if not placeholder_map:
        chunks.append({"type": "text", "text": text})
        return

    # Build a regex that matches any placeholder
    pattern = "|".join(re.escape(k) for k in placeholder_map)
    parts = re.split(f"({pattern})", text)

    for part in parts:
        if not part.strip():
            continue
        if part in placeholder_map:
            mb = placeholder_map[part]
            chunks.append({
                "type": "math",
                "latex": mb.text,
                "display": mb.display,
            })
        else:
            chunks.append({"type": "text", "text": part})


# ─── Table parsing ──────────────────────────────────────────────────────────────

def _is_table_line(line: str) -> bool:
    s = line.strip()
    return s.startswith("|") and s.endswith("|")


def _parse_md_table(lines: list[str]) -> dict | None:
    rows = []
    header_row = None
    alignments = []

    for line in lines:
        stripped = line.strip()
        if not stripped.startswith("|"):
            continue
        cells = [c.strip() for c in stripped[1:-1].split("|")]

        if all(re.match(r"^[\-:]+$", c.strip()) for c in cells):
            alignments = []
            for c in cells:
                c = c.strip()
                if c.startswith(":") and c.endswith(":"):
                    alignments.append("center")
                elif c.endswith(":"):
                    alignments.append("right")
                else:
                    alignments.append("left")
            continue

        if not header_row:
            header_row = cells
        else:
            rows.append(cells)

    if not header_row:
        return None

    if not alignments:
        alignments = ["left"] * len(header_row)
    while len(alignments) < len(header_row):
        alignments.append("left")

    return {"headers": header_row, "alignments": alignments, "rows": rows}


# ─── Code block parsing ─────────────────────────────────────────────────────────

def _is_mermaid_start(line: str) -> bool:
    s = line.strip()
    return s.startswith("```mermaid") or s.startswith("``` mermaid")


def _is_code_block_start(line: str) -> bool:
    s = line.strip()
    return s.startswith("```") and not s.startswith("```mermaid") and not s.startswith("``` mermaid")


def _read_code_block(lines: list[str], start: int) -> tuple[str, list[str]]:
    opening = lines[start].strip()
    lang = opening[3:].strip()  # after ```

    code_lines = []
    i = start + 1
    while i < len(lines) and not lines[i].strip().startswith("```"):
        code_lines.append(lines[i])
        i += 1
    return lang, code_lines


# ─── Task list parsing ──────────────────────────────────────────────────────────

def _is_task_item(line: str) -> bool:
    return bool(re.match(r"^\s*(-|\*|\+)\s+\[([ xX])\]\s+", line))


# ─── Horizontal rule ────────────────────────────────────────────────────────────

def _is_hr(line: str) -> bool:
    s = line.strip()
    return bool(re.match(r"^(-{3,}|_{3,}|\*{3,})$", s))


# ─── Utility ────────────────────────────────────────────────────────────────────

def _detect_list_type(line: str) -> tuple[int, str, str]:
    """Return (indent_depth, marker_type, content) for a list line.

    marker_type: 'ul' for unordered (-, *, +), 'ol' for ordered (1., 2.)
    """
    m = re.match(r"^(\s*)([-*+]|\d+[.)])\s+(.*)", line)
    if not m:
        return 0, "", line
    indent = len(m.group(1))
    marker = m.group(2)
    content = m.group(3)
    if re.match(r"^\d+[.)]", marker):
        return indent, "ol", content
    return indent, "ul", content
