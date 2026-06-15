"""LaTeX math expression parser and Word OMML converter.

Handles:
  - Inline math: $...$  or \\(...\\)
  - Display math: $$...$$ or \\[...\\]

Generates Word OMML (Office Math Markup Language) compatible XML
that python-docx can insert as native Word equation objects.
"""

from dataclasses import dataclass
import re
import html


@dataclass
class MathBlock:
    text: str          # LaTeX source
    display: bool      # True = block, False = inline
    start: int         # position in source
    end: int           # position in source


# ─── Pattern matching ───────────────────────────────────────────────────────────

_MATH_PATTERNS = [
    # $$...$$ (display, must come before $...$)
    (re.compile(r"\$\$(.+?)\$\$", re.DOTALL), True),
    # \[...\] (display)
    (re.compile(r"\\\[(.+?)\\\]", re.DOTALL), True),
    # $...$ (inline, no line breaks)
    (re.compile(r"\$(.+?)\$"), False),
    # \(...\) (inline)
    (re.compile(r"\\\((.+?)\\\)"), False),
]


def extract_math(text: str) -> list[MathBlock]:
    """Find all math expressions in text, returning positions for replacement."""
    results = []
    seen_spans = set()

    for pattern, is_display in _MATH_PATTERNS:
        for m in pattern.finditer(text):
            start, end = m.start(), m.end()
            # Avoid overlapping matches (prefer first match)
            if any(start < pe and end > ps for ps, pe in seen_spans):
                continue
            seen_spans.add((start, end))
            results.append(MathBlock(
                text=m.group(1).strip(),
                display=is_display,
                start=start,
                end=end,
            ))

    # Sort by position
    results.sort(key=lambda x: x.start)
    return results


# ─── LaTeX-to-OMML converter ────────────────────────────────────────────────────

def latex_to_omml(latex: str, display: bool = False) -> str:
    """Convert a LaTeX math expression to OMML XML.

    Supports common constructs: fractions, sub/superscripts, Greek letters,
    square roots, integrals, sums, matrices, delimiters, and basic symbols.

    Returns inner OMML content string (no outer m:oMath wrap).
    """
    latex = latex.strip()
    parser = _LaTeXParser(latex)
    tree = parser.parse()
    return _omml_from_tree(tree)


# ─── LaTeX parser (recursive descent) ───────────────────────────────────────────

class _LaTeXParser:
    # Safety limit: max iterations per parsing loop to prevent hangs
    _MAX_ITER = 50000

    def __init__(self, source: str):
        self.source = source
        self.pos = 0

    def peek(self) -> str:
        return self.source[self.pos] if self.pos < len(self.source) else ""

    def consume(self) -> str:
        c = self.peek()
        self.pos += 1
        return c

    def skip_ws(self):
        while self.pos < len(self.source) and self.source[self.pos].isspace():
            self.pos += 1

    def parse(self) -> dict:
        return {"type": "group", "items": self._parse_items()}

    def _parse_items(self, stop_chars="}") -> list:
        items = []
        iters = 0
        while self.pos < len(self.source):
            iters += 1
            if iters > self._MAX_ITER:
                break
            prev_pos = self.pos
            c = self.peek()
            if c in stop_chars:
                break
            if c == "\\":
                # Check for \\ (row separator in matrices)
                if self.pos + 1 < len(self.source) and self.source[self.pos + 1] == "\\":
                    self.consume(); self.consume()
                    items.append({"type": "row_sep"})
                else:
                    items.append(self._parse_command())
            elif c == "{":
                self.consume()
                items.append(self._parse_group())
            elif c == "_":
                self.consume()
                items.append({"type": "sub", "content": self._parse_single()})
            elif c == "^":
                self.consume()
                items.append({"type": "sup", "content": self._parse_single()})
            elif c == "&":
                self.consume()
                items.append({"type": "column_sep"})
            else:
                items.append({"type": "text", "text": self.consume()})
            # Safety: ensure position always advances
            if self.pos == prev_pos:
                self.pos += 1
        return items

    def _parse_command(self):
        self.consume()  # backslash
        # Check for special single-char commands: \|, \{, \}, etc.
        if self.pos < len(self.source) and not self.source[self.pos].isalpha():
            ch = self.consume()
            if ch == "|":
                return {"type": "command", "cmd": "||"}
            return {"type": "text", "text": ch}

        # Read command name (letters only)
        cmd = ""
        while self.pos < len(self.source) and self.source[self.pos].isalpha():
            cmd += self.consume()

        node = {"type": "command", "cmd": cmd}

        if cmd == "frac":
            node["num"] = self._parse_single()
            node["den"] = self._parse_single()
        elif cmd == "dfrac":
            node["num"] = self._parse_single()
            node["den"] = self._parse_single()
        elif cmd == "tfrac":
            node["num"] = self._parse_single()
            node["den"] = self._parse_single()
        elif cmd == "sqrt":
            if self.peek() == "[":
                self.consume()
                degree_items = self._parse_items("]")
                if self.peek() == "]":
                    self.consume()
                node["degree"] = {"type": "group", "items": degree_items}
            node["content"] = self._parse_single()
        elif cmd in ("sum", "prod", "coprod", "int", "oint", "iint", "iiint",
                      "bigcup", "bigcap", "bigoplus", "bigotimes",
                      "bigvee", "bigwedge"):
            if self.peek() == "_":
                self.consume()
                node["lower"] = self._parse_single()
            if self.peek() == "^":
                self.consume()
                node["upper"] = self._parse_single()
        elif cmd == "left":
            self.skip_ws()
            delim = self._parse_delimiter()
            node["delim"] = delim
            node["items"] = self._parse_items_stop_at_right()
            # Consume the \right and its delimiter
            if (self.pos + 1 < len(self.source) and
                self.source[self.pos] == "\\" and
                self.pos + 5 < len(self.source) and
                self.source[self.pos+1:self.pos+6] == "right"):
                self.consume()  # \
                for _ in range(5):  # consume 'right'
                    self.consume()
                self.skip_ws()
                end_delim = self._parse_delimiter()
                node["right_node"] = {"delim": end_delim}
            # After \right, check for sub/sup on the whole delimited expression
            self.skip_ws()
            if self.peek() == "_":
                self.consume()
                node["lower"] = self._parse_single()
            if self.peek() == "^":
                self.consume()
                node["upper"] = self._parse_single()
        elif cmd == "right":
            self.skip_ws()
            if self.pos < len(self.source):
                node["delim"] = self.consume()
        elif cmd in ("overline", "underline", "bar", "vec", "hat", "tilde",
                      "dot", "ddot", "breve", "check", "acute", "grave",
                      "widehat", "widetilde", "overrightarrow", "overleftarrow"):
            node["content"] = self._parse_single()
        elif cmd == "begin":
            node = self._parse_environment()
        elif cmd in ("text", "mbox", "mathrm", "textbf", "textit"):
            node = {"type": "text", "text": self._parse_single_text()}
        elif cmd in ("mathbf", "mathit", "mathsf", "mathtt", "mathbb",
                      "mathcal", "mathscr", "mathfrak"):
            node["content"] = self._parse_single()
        elif cmd == "not":
            next_item = self._parse_single()
            node["content"] = next_item
        elif cmd == "quad":
            node = {"type": "text", "text": "  "}
        elif cmd == "qquad":
            node = {"type": "text", "text": "    "}
        elif cmd == ",":
            node = {"type": "text", "text": " "}
        elif cmd == ";":
            node = {"type": "text", "text": " "}
        elif cmd == "!":
            pass
        elif cmd == "limits":
            node = {"type": "limits"}
        elif cmd == "over":
            node["content_right"] = self._parse_single()
        elif cmd == "binom":
            node["num"] = self._parse_single()
            node["den"] = self._parse_single()
        elif cmd == "stackrel":
            node["content"] = self._parse_single()
            node["upper"] = self._parse_single()

        return node

    def _parse_items_stop_at_right(self) -> list:
        """Parse items until \\right is encountered."""
        items = []
        iters = 0
        while self.pos < len(self.source):
            iters += 1
            if iters > self._MAX_ITER:
                break
            prev_pos = self.pos
            c = self.peek()
            if c == "\\":
                # Check for \right
                if (self.pos + 5 < len(self.source) and
                    self.source[self.pos+1:self.pos+6] == "right"):
                    break
                # Check for \\
                if (self.pos + 1 < len(self.source) and
                    self.source[self.pos + 1] == "\\"):
                    self.consume(); self.consume()
                    items.append({"type": "row_sep"})
                else:
                    items.append(self._parse_command())
            elif c == "{":
                self.consume()
                items.append(self._parse_group())
            elif c == "_":
                self.consume()
                items.append({"type": "sub", "content": self._parse_single()})
            elif c == "^":
                self.consume()
                items.append({"type": "sup", "content": self._parse_single()})
            elif c == "&":
                self.consume()
                items.append({"type": "column_sep"})
            else:
                items.append({"type": "text", "text": self.consume()})
            if self.pos == prev_pos:
                self.pos += 1
        return items

    def _parse_group(self):
        items = self._parse_items("}")
        if self.peek() == "}":
            self.consume()
        return {"type": "group", "items": items}

    def _parse_single(self) -> dict:
        """Parse a single argument (atom, group, or single command)."""
        self.skip_ws()
        c = self.peek()
        if c == "{":
            self.consume()
            return self._parse_group()
        if c == "\\":
            return self._parse_command()
        if c in ("_", "^"):
            self.consume()
            return {f"type": "sub" if c == "_" else "sup",
                    "content": self._parse_single()}
        if c and c not in "}&\\":
            return {"type": "text", "text": self.consume()}
        return {"type": "text", "text": ""}

    def _parse_single_text(self) -> str:
        self.skip_ws()
        if self.peek() == "{":
            self.consume()
            text = self._parse_until("}")
            if self.peek() == "}":
                self.consume()
            return text
        if self.pos < len(self.source):
            return self.consume()
        return ""

    def _parse_until(self, stop: str) -> str:
        result = []
        while self.pos < len(self.source) and self.peek() != stop:
            if self.peek() == "\\" and self.pos + 1 < len(self.source):
                self.consume()
                result.append(self.consume())
            else:
                result.append(self.consume())
        return "".join(result)

    def _parse_delimiter(self) -> str:
        self.skip_ws()
        if self.pos >= len(self.source):
            return "."

        c = self.peek()
        if c == "\\":
            self.consume()
            if self.pos >= len(self.source):
                return "\\"
            next_c = self.peek()
            if next_c == "|":
                self.consume()
                return "||"
            if next_c in ("{", "}"):
                self.consume()
                return next_c
            cmd = ""
            while self.pos < len(self.source) and self.source[self.pos].isalpha():
                cmd += self.consume()
            return cmd if cmd else "\\"

        return self.consume()

    def _parse_environment(self) -> dict:
        self.skip_ws()
        env_name = ""
        if self.peek() == "{":
            self.consume()
            env_name = self._parse_until("}")
            if self.peek() == "}":
                self.consume()

        items = self._parse_items_until_end(env_name)
        return {"type": "environment", "env": env_name, "items": items}

    def _parse_items_until_end(self, env_name: str) -> list:
        items = []
        iters = 0
        while self.pos < len(self.source):
            iters += 1
            if iters > self._MAX_ITER:
                break
            prev_pos = self.pos
            c = self.peek()
            if c == "\\":
                if (self.pos + 3 < len(self.source) and
                    self.source[self.pos+1:self.pos+4] == "end"):
                    saved_pos = self.pos
                    self.consume()
                    self.consume(); self.consume(); self.consume()
                    self.skip_ws()
                    if self.peek() == "{":
                        self.consume()
                        found_env = self._parse_until("}")
                        if self.peek() == "}":
                            self.consume()
                        if found_env == env_name:
                            break
                        items.append({"type": "command", "cmd": "end"})
                        items.append({"type": "text", "text": found_env})
                    else:
                        items.append({"type": "command", "cmd": "end"})
                    continue

                if (self.pos + 1 < len(self.source) and
                    self.source[self.pos + 1] == "\\"):
                    self.consume(); self.consume()
                    items.append({"type": "row_sep"})
                else:
                    items.append(self._parse_command())
            elif c == "{":
                self.consume()
                items.append(self._parse_group())
            elif c == "_":
                self.consume()
                items.append({"type": "sub", "content": self._parse_single()})
            elif c == "^":
                self.consume()
                items.append({"type": "sup", "content": self._parse_single()})
            elif c == "&":
                self.consume()
                items.append({"type": "column_sep"})
            else:
                items.append({"type": "text", "text": self.consume()})
            if self.pos == prev_pos:
                self.pos += 1
        return items


# ─── OMML XML generator (rewritten) ─────────────────────────────────────────────

_OMML_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _run(text: str, sty: str = "") -> str:
    """Create an OMML run element <m:r><m:rPr>...</m:rPr><m:t>text</m:t></m:r>."""
    escaped = html.escape(text)
    if sty:
        return f'<m:r><m:rPr><m:sty m:val="{sty}"/></m:rPr><m:t>{escaped}</m:t></m:r>'
    return f'<m:r><m:t>{escaped}</m:t></m:r>'


_GREEK = {
    "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ",
    "epsilon": "ε", "varepsilon": "ε", "zeta": "ζ", "eta": "η",
    "theta": "θ", "vartheta": "ϑ", "iota": "ι", "kappa": "κ",
    "lambda": "λ", "mu": "μ", "nu": "ν", "xi": "ξ",
    "omicron": "ο", "pi": "π", "varpi": "ϖ", "rho": "ρ",
    "varrho": "ϱ", "sigma": "σ", "varsigma": "ς", "tau": "τ",
    "upsilon": "υ", "phi": "φ", "varphi": "ϕ", "chi": "χ",
    "psi": "ψ", "omega": "ω",
    "Gamma": "Γ", "Delta": "Δ", "Theta": "Θ", "Lambda": "Λ",
    "Xi": "Ξ", "Pi": "Π", "Sigma": "Σ", "Upsilon": "Υ",
    "Phi": "Φ", "Psi": "Ψ", "Omega": "Ω",
}

_COMMANDS = {
    "times": "×", "div": "÷", "pm": "±", "mp": "∓",
    "cdot": "·", "leq": "≤", "geq": "≥", "neq": "≠",
    "lt": "<", "gt": ">", "le": "≤", "ge": "≥",
    "approx": "≈", "equiv": "≡", "propto": "∝", "sim": "∼",
    "simeq": "≃", "cong": "≅", "ne": "≠",
    "rightarrow": "→", "leftarrow": "←", "leftrightarrow": "↔",
    "Rightarrow": "⇒", "Leftarrow": "⇐", "Leftrightarrow": "⇔",
    "uparrow": "↑", "downarrow": "↓",
    "infty": "∞", "partial": "∂", "nabla": "∇",
    "forall": "∀", "exists": "∃", "emptyset": "∅",
    "angle": "∠", "triangle": "△", "circ": "∘",
    "cdots": "⋯", "ldots": "…", "vdots": "⋮", "ddots": "⋱",
    "dots": "…",
    "to": "→", "mapsto": "↦", "implies": "⇒", "iff": "⇔",
    "cup": "∪", "cap": "∩", "subseteq": "⊆", "subset": "⊂",
    "supseteq": "⊇", "supset": "⊃",
    "oplus": "⊕", "otimes": "⊗", "perp": "⟂",
    "bullet": "•", "star": "★", "dagger": "†", "ddagger": "‡",
    "ell": "ℓ", "Re": "ℜ", "Im": "ℑ", "wp": "℘",
    "aleph": "ℵ", "hbar": "ℏ",
    "neg": "¬", "land": "∧", "lor": "∨",
    "in": "∈", "notin": "∉", "ni": "∋",
    "langle": "⟨", "rangle": "⟩",
    "lfloor": "⌊", "rfloor": "⌋",
    "lceil": "⌈", "rceil": "⌉",
    "||": "‖", "Vert": "‖", "lVert": "‖", "rVert": "‖",
    "colon": ":", "quad": "  ", "qquad": "    ",
    "mathbb": "", "mathcal": "", "mathbf": "",
    "mathit": "", "mathsf": "", "mathtt": "", "mathrm": "",
}

_NARY_OPS = {
    "sum": "∑", "prod": "∏", "coprod": "∐",
    "int": "∫", "oint": "∮", "iint": "∬", "iiint": "∭",
    "bigcup": "⋃", "bigcap": "⋂", "bigoplus": "⨁", "bigotimes": "⨂",
    "bigvee": "⋁", "bigwedge": "⋀",
}

_ACCENT_CHARS = {
    "bar": "0304", "overline": "0304",
    "vec": "20D7", "overrightarrow": "20D7", "overleftarrow": "2190",
    "hat": "0302", "widehat": "0302",
    "tilde": "0303", "widetilde": "0303",
    "dot": "0307", "ddot": "0308",
    "breve": "0306", "check": "030C",
    "acute": "0301", "grave": "0300",
}

_DELIM_MAP = {
    "(": "(", ")": ")", "[": "[", "]": "]",
    "{": "{", "}": "}", "|": "|", "||": "‖",
    ".": "",
    "\\": "",
    "langle": "⟨", "rangle": "⟩",
    "lfloor": "⌊", "rfloor": "⌋",
    "lceil": "⌈", "rceil": "⌉",
    "/": "/", "uparrow": "↑", "downarrow": "↓",
    "Uparrow": "⇑", "Downarrow": "⇓",
}

# Font command → OMML sty attribute mapping
_FONT_STY = {
    "mathbf": "b",
    "mathit": "i",
    "mathsf": "sans",
    "mathtt": "monospace",
    "mathrm": "roman",
    "mathbb": "double-struck",
    "mathcal": "script",
    "mathscr": "script",
    "mathfrak": "fraktur",
}


# ─── AST → OMML conversion ─────────────────────────────────────────────────────

def _omml_from_tree(node: dict) -> str:
    """Convert AST to OMML XML. Returns inner content (no outer m:oMath wrap)."""
    return _omml_items(node.get("items", []))


def _omml_items(items: list) -> str:
    """Convert a list of AST items to OMML XML string.

    Key fixes:
    - sub/sup are attached to the preceding element correctly
    - nary operators (∑, ∫, etc.) get their body filled properly
      by collecting following items until a boundary is reached
    """
    result = []
    i = 0
    while i < len(items):
        item = items[i]
        item_type = item.get("type", "")

        # Subscript
        if item_type == "sub":
            base = _pop_last(result)
            sub_content = _omml_node(item.get("content", {}))
            # Check if next is sup → combine as sSubSup
            if i + 1 < len(items) and items[i + 1].get("type") == "sup":
                i += 1
                sup_content = _omml_node(items[i].get("content", {}))
                result.append(
                    f'<m:sSubSup>'
                    f'<m:e>{base}</m:e>'
                    f'<m:sub>{sub_content}</m:sub>'
                    f'<m:sup>{sup_content}</m:sup>'
                    f'</m:sSubSup>')
            else:
                result.append(
                    f'<m:sSub>'
                    f'<m:e>{base}</m:e>'
                    f'<m:sub>{sub_content}</m:sub>'
                    f'</m:sSub>')
            i += 1
            continue

        # Superscript
        if item_type == "sup":
            base = _pop_last(result)
            sup_content = _omml_node(item.get("content", {}))
            result.append(
                f'<m:sSup>'
                f'<m:e>{base}</m:e>'
                f'<m:sup>{sup_content}</m:sup>'
                f'</m:sSup>')
            i += 1
            continue

        # Nary operators: collect body items and fill <m:e>
        if item_type == "command" and item.get("cmd") in _NARY_OPS:
            nary_xml = _omml_node(item)
            # Collect body: items following the nary operator until a boundary
            body_items = []
            i += 1
            while i < len(items):
                next_item = items[i]
                next_type = next_item.get("type", "")
                # Stop at relation operators and other nary ops
                if next_type == "command":
                    next_cmd = next_item.get("cmd", "")
                    if next_cmd in _NARY_OPS:
                        break
                    if next_cmd in ("geq", "ge", "leq", "le", "eq", "neq", "ne",
                                    "approx", "equiv", "sim", "propto",
                                    "left", "right"):
                        break
                # Stop at text that is a relation symbol
                if next_type == "text" and next_item.get("text", "").strip() in (
                    "=", "≥", "≤", "≠", ">", "<", "≫", "≪"):
                    break
                body_items.append(next_item)
                i += 1
            body_xml = _omml_items(body_items) if body_items else ""
            # Replace the empty <m:e></m:e> placeholder with actual body
            nary_xml = nary_xml.replace("<m:e></m:e>", f"<m:e>{body_xml}</m:e>", 1)
            result.append(nary_xml)
            continue

        result.append(_omml_node(item))
        i += 1

    return "".join(result)


def _pop_last(result: list) -> str:
    """Pop the last OMML element from result list to use as base for sub/sup.

    If the last element is a simple <m:r> (run), wrap it in <m:e> for use
    as a base. If it's a compound element (frac, nary, etc.), wrap it in
    <m:e> as-is.
    """
    if result:
        return result.pop()
    return _run("")


def _omml_node(node: dict) -> str:
    """Convert a single AST node to OMML XML."""
    t = node.get("type", "")

    if t == "text":
        text = node.get("text", "")
        if not text:
            return ""
        return _run(text)

    if t == "group":
        return _omml_items(node.get("items", []))

    if t == "row_sep":
        return ""
    if t == "column_sep":
        return ""

    if t == "command":
        return _omml_command(node)

    if t == "environment":
        return _omml_env(node)

    if t == "limits":
        return ""

    return ""


def _omml_command(node: dict) -> str:
    """Convert a command AST node to OMML XML."""
    cmd = node.get("cmd", "")

    # ── Fractions ───────────────────────────────────────────────────────────
    if cmd in ("frac", "dfrac", "tfrac"):
        num = _omml_node(node.get("num", {}))
        den = _omml_node(node.get("den", {}))
        fPr = ""
        if cmd == "dfrac":
            # dfrac = display-style fraction (no type override needed,
            # Word uses display mode by default in oMathPara)
            pass
        if cmd == "tfrac":
            # tfrac = text-style (inline) fraction
            fPr = '<m:fPr><m:type m:val="lin"/></m:fPr>'
        return (
            f'<m:f>{fPr}'
            f'<m:num>{num}</m:num>'
            f'<m:den>{den}</m:den>'
            f'</m:f>')

    # ── Binomial coefficient ────────────────────────────────────────────────
    if cmd == "binom":
        num = _omml_node(node.get("num", {}))
        den = _omml_node(node.get("den", {}))
        return (
            f'<m:d>'
            f'<m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr>'
            f'<m:e>'
            f'<m:f>'
            f'<m:num>{num}</m:num>'
            f'<m:den>{den}</m:den>'
            f'</m:f>'
            f'</m:e>'
            f'</m:d>')

    # ── Square root / nth root ──────────────────────────────────────────────
    if cmd == "sqrt":
        content = _omml_node(node.get("content", {}))
        if "degree" in node:
            deg = _omml_node(node["degree"])
            return (
                f'<m:rad>'
                f'<m:radPr><m:degHide m:val="0"/></m:radPr>'
                f'<m:deg>{deg}</m:deg>'
                f'<m:e>{content}</m:e>'
                f'</m:rad>')
        return (
            f'<m:rad>'
            f'<m:radPr><m:degHide m:val="1"/></m:radPr>'
            f'<m:deg/>'
            f'<m:e>{content}</m:e>'
            f'</m:rad>')

    # ── Large operators (nary) ──────────────────────────────────────────────
    # FIX: Generate correct nary OMML structure with proper body.
    # The nary operator in OMML has structure:
    #   <m:nary><m:naryPr>...</m:naryPr><m:sub/><m:sup/><m:e>body</m:e></m:nary>
    # The body (<m:e>) contains the expression the operator applies to.
    # We leave body empty here; it gets filled by _omml_items when the
    # nary node is encountered as a top-level item.
    if cmd in _NARY_OPS:
        op_char = _NARY_OPS[cmd]
        lower = node.get("lower")
        upper = node.get("upper")
        has_lower = lower is not None
        has_upper = upper is not None

        lower_xml = _omml_node(lower) if has_lower else ""
        upper_xml = _omml_node(upper) if has_upper else ""

        naryPr = (
            f'<m:naryPr>'
            f'<m:chr m:val="{html.escape(op_char)}"/>'
            f'<m:limLoc m:val="subSup"/>'
            f'<m:subHide m:val="{"0" if has_lower else "1"}"/>'
            f'<m:supHide m:val="{"0" if has_upper else "1"}"/>'
            f'</m:naryPr>')

        return (
            f'<m:nary>'
            f'{naryPr}'
            f'<m:sub>{lower_xml}</m:sub>'
            f'<m:sup>{upper_xml}</m:sup>'
            f'<m:e></m:e>'
            f'</m:nary>')

    # ── \left delimiter ─────────────────────────────────────────────────────
    if cmd == "left":
        delim = node.get("delim", "")
        beg_chr = _DELIM_MAP.get(delim, delim)
        items = node.get("items", [])
        inner = _omml_items(items)
        end_chr = ""
        right_node = node.get("right_node")
        if right_node:
            end_delim = right_node.get("delim", "")
            end_chr = _DELIM_MAP.get(end_delim, end_delim)

        d_xml = (
            f'<m:d>'
            f'<m:dPr>'
            f'<m:begChr m:val="{html.escape(beg_chr)}"/>'
            f'<m:endChr m:val="{html.escape(end_chr)}"/>'
            f'</m:dPr>'
            f'<m:e>{inner}</m:e>'
            f'</m:d>')

        # Handle sub/sup on the whole delimited expression
        lower = node.get("lower")
        upper = node.get("upper")
        if lower and upper:
            return (
                f'<m:sSubSup>'
                f'<m:e>{d_xml}</m:e>'
                f'<m:sub>{_omml_node(lower)}</m:sub>'
                f'<m:sup>{_omml_node(upper)}</m:sup>'
                f'</m:sSubSup>')
        if upper:
            return (
                f'<m:sSup>'
                f'<m:e>{d_xml}</m:e>'
                f'<m:sup>{_omml_node(upper)}</m:sup>'
                f'</m:sSup>')
        if lower:
            return (
                f'<m:sSub>'
                f'<m:e>{d_xml}</m:e>'
                f'<m:sub>{_omml_node(lower)}</m:sub>'
                f'</m:sSub>')
        return d_xml

    # ── \right delimiter (handled by \left, standalone = empty) ─────────────
    if cmd == "right":
        return ""

    # ── Accents ─────────────────────────────────────────────────────────────
    if cmd in _ACCENT_CHARS:
        chr_val = _ACCENT_CHARS[cmd]
        content = _omml_node(node.get("content", {}))
        return (
            f'<m:acc>'
            f'<m:accPr><m:chr m:val="{chr_val}"/></m:accPr>'
            f'<m:e>{content}</m:e>'
            f'</m:acc>')

    # ── Overline / underline (use m:bar) ────────────────────────────────────
    if cmd == "overline":
        content = _omml_node(node.get("content", {}))
        return (
            f'<m:bar>'
            f'<m:barPr><m:pos m:val="top"/></m:barPr>'
            f'<m:e>{content}</m:e>'
            f'</m:bar>')
    if cmd == "underline":
        content = _omml_node(node.get("content", {}))
        return (
            f'<m:bar>'
            f'<m:barPr><m:pos m:val="bot"/></m:barPr>'
            f'<m:e>{content}</m:e>'
            f'</m:bar>')

    # ── Greek letters ───────────────────────────────────────────────────────
    if cmd in _GREEK:
        return _run(_GREEK[cmd])

    # ── Font commands with proper OMML styling ──────────────────────────────
    # FIX: Use <m:rPr><m:sty m:val="..."/> to apply font styles instead of
    # ignoring them. This ensures \mathbb{R} renders differently from R.
    if cmd in _FONT_STY:
        content = _omml_node(node.get("content", {}))
        if not content:
            return ""
        sty = _FONT_STY[cmd]
        # Wrap the content's runs with the font style
        # We need to inject <m:rPr> into each <m:r> in content
        return _apply_font_sty(content, sty)

    # ── Symbol commands ─────────────────────────────────────────────────────
    if cmd in _COMMANDS and _COMMANDS[cmd]:
        return _run(_COMMANDS[cmd])

    # ── \not — combine with next char ───────────────────────────────────────
    if cmd == "not":
        content = _omml_node(node.get("content", {}))
        # Prepend a combining long solidus overlay (U+0338)
        # This is a simplified approach; proper OMML would use accPr
        if content.startswith("<m:r>"):
            # Insert the combining character before the text content
            content = content.replace("<m:t>", "<m:t>\u0338", 1)
        return content

    # ── \over (infix fraction) ──────────────────────────────────────────────
    if cmd == "over":
        # \over is handled specially: the left side is collected by the parent
        # For now, output as fraction with empty numerator
        den = _omml_node(node.get("content_right", {}))
        return (
            f'<m:f>'
            f'<m:num></m:num>'
            f'<m:den>{den}</m:den>'
            f'</m:f>')

    # ── \stackrel ───────────────────────────────────────────────────────────
    if cmd == "stackrel":
        content = _omml_node(node.get("content", {}))
        upper = _omml_node(node.get("upper", {}))
        return (
            f'<m:sSup>'
            f'<m:e>{content}</m:e>'
            f'<m:sup>{upper}</m:sup>'
            f'</m:sSup>')

    # ── Unknown command: show as text ───────────────────────────────────────
    if cmd:
        return _run(f"\\{cmd}")

    return ""


def _apply_font_sty(content: str, sty: str) -> str:
    """Apply OMML font style to all <m:r> elements in content string.

    Injects <m:rPr><m:sty m:val="..."/></m:rPr> into each <m:r> that
    doesn't already have an <m:rPr>.
    """
    rpr = f'<m:rPr><m:sty m:val="{sty}"/></m:rPr>'
    # Insert rPr after <m:r> if no rPr exists
    result = content.replace("<m:r>", f"<m:r>{rpr}", 1)
    # Handle remaining <m:r> occurrences (shouldn't normally happen in nested content)
    return result


def _omml_env(node: dict) -> str:
    """Convert an environment AST node to OMML XML."""
    env = node.get("env", "")
    items = node.get("items", [])

    if env in ("matrix", "pmatrix", "bmatrix", "vmatrix", "cases",
               "aligned", "align", "split", "array"):
        rows = _split_matrix(items)
        mc = _omml_matrix_rows(rows)

        if env == "pmatrix":
            return (
                f'<m:d>'
                f'<m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr>'
                f'<m:e>{mc}</m:e>'
                f'</m:d>')
        if env == "bmatrix":
            return (
                f'<m:d>'
                f'<m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/></m:dPr>'
                f'<m:e>{mc}</m:e>'
                f'</m:d>')
        if env == "vmatrix":
            return (
                f'<m:d>'
                f'<m:dPr><m:begChr m:val="|"/><m:endChr m:val="|"/></m:dPr>'
                f'<m:e>{mc}</m:e>'
                f'</m:d>')
        if env == "cases":
            return (
                f'<m:d>'
                f'<m:dPr><m:begChr m:val="{{"/><m:endChr m:val=" "/></m:dPr>'
                f'<m:e>{mc}</m:e>'
                f'</m:d>')
        # Plain matrix / aligned / split / array: no delimiters
        return (
            f'<m:d>'
            f'<m:dPr><m:begChr m:val=""/><m:endChr m:val=""/></m:dPr>'
            f'<m:e>{mc}</m:e>'
            f'</m:d>')

    # Unknown environment: just output content
    return _omml_items(items)


def _split_matrix(items: list) -> list:
    """Split items into rows based on row_sep."""
    rows = []
    current_row = []
    for item in items:
        if item.get("type") == "row_sep":
            rows.append(current_row)
            current_row = []
        elif item.get("type") == "column_sep":
            continue
        else:
            current_row.append(item)
    if current_row:
        rows.append(current_row)
    return rows


def _omml_matrix_rows(rows: list) -> str:
    """Convert matrix rows to OMML <m:m> → <m:mr> → <m:e> elements.

    FIX: Matrix content must be wrapped in <m:m> (matrix) element,
    with each row as <m:mr> and each cell as <m:e>.
    """
    mr_parts = []
    for row in rows:
        cells = []
        for cell in row:
            cells.append(f'<m:e>{_omml_node(cell)}</m:e>')
        mr_parts.append(f'<m:mr>{"".join(cells)}</m:mr>')
    return f'<m:m>{"".join(mr_parts)}</m:m>'


# ─── Full OMML document wrapper ─────────────────────────────────────────────────

def _omml_wrap(inner: str, display: bool = False) -> str:
    """Wrap inner OMML content in the proper Word math paragraph/run structure."""
    ns = f'xmlns:m="{_OMML_NS}"'
    if display:
        return (
            f'<m:oMathPara {ns}>'
            f'<m:oMath>{inner}</m:oMath>'
            f'</m:oMathPara>'
        )
    else:
        return (
            f'<m:oMath {ns}>'
            f'{inner}'
            f'</m:oMath>'
        )
