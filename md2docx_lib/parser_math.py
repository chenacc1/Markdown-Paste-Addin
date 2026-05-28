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


# ─── LaTeX-to-OMML converter (simplified) ────────────────────────────────────────

def latex_to_omml(latex: str, display: bool = False) -> str:
    """Convert a LaTeX math expression to OMML XML.

    Supports common constructs: fractions, sub/superscripts, Greek letters,
    square roots, integrals, sums, matrices, and basic symbols.

    Returns OMML XML string ready for Word m:oMathPara or m:oMath element.
    """
    latex = latex.strip()
    parser = _LaTeXParser(latex)
    tree = parser.parse()
    return _omml_from_tree(tree, display)


# ─── LaTeX parser (recursive descent) ───────────────────────────────────────────

class _LaTeXParser:
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

    def _parse_items(self) -> list:
        items = []
        while self.pos < len(self.source):
            c = self.peek()
            if c == "}":
                break
            if c == "\\":
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
            elif c == "\\" and self.pos + 1 < len(self.source) and self.source[self.pos + 1] == "\\":
                self.consume(); self.consume()
                items.append({"type": "row_sep"})
            else:
                items.append({"type": "text", "text": self.consume()})
        return items

    def _parse_command(self):
        self.consume()  # backslash
        # Read command name (letters only)
        cmd = ""
        while self.pos < len(self.source) and self.source[self.pos].isalpha():
            cmd += self.consume()

        node = {"type": "command", "cmd": cmd}

        if cmd in ("frac", "frac"):
            node["num"] = self._parse_single()
            node["den"] = self._parse_single()
        elif cmd == "sqrt":
            if self.peek() == "[":
                self.consume()
                node["degree"] = self._parse_until("]")
                self.consume()  # ]
            node["content"] = self._parse_single()
        elif cmd in ("sum", "prod", "int", "oint", "iint", "iiint"):
            if self.peek() == "_":
                self.consume()
                node["lower"] = self._parse_single()
            if self.peek() == "^":
                self.consume()
                node["upper"] = self._parse_single()
        elif cmd in ("left", "right"):
            delim = ""
            while self.pos < len(self.source) and self.peek().isspace():
                self.consume()
            if self.pos < len(self.source):
                delim = self.consume()
            node["delim"] = delim
            if cmd == "left":
                node["items"] = self._parse_items()
        elif cmd in ("overline", "underline", "bar", "vec", "hat", "tilde", "dot", "ddot"):
            node["content"] = self._parse_single()
        elif cmd in ("begin",):
            node = self._parse_environment()
        elif cmd in ("text", "mbox", "mathrm"):
            node = {"type": "text", "text": self._parse_single_text()}

        return node

    def _parse_group(self):
        items = self._parse_items()
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
            self.consume()
            return text
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

    def _parse_environment(self) -> dict:
        env_name = ""
        while self.pos < len(self.source) and self.source[self.pos].isalpha():
            env_name += self.consume()
        items = self._parse_items()
        return {"type": "environment", "env": env_name, "items": items}


# ─── OMML XML generator ─────────────────────────────────────────────────────────

_OMML_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _el(name: str, **attrs) -> str:
    attr_str = "".join(f' {k}="{html.escape(str(v))}"' for k, v in attrs.items())
    return f"<m:{name}{attr_str}>"


def _el_close(name: str) -> str:
    return f"</m:{name}>"


def _wrap(name: str, content: str, **attrs) -> str:
    return f"{_el(name, **attrs)}{content}{_el_close(name)}"


_GREEK = {
    "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ",
    "epsilon": "ε", "zeta": "ζ", "eta": "η", "theta": "θ",
    "iota": "ι", "kappa": "κ", "lambda": "λ", "mu": "μ",
    "nu": "ν", "xi": "ξ", "omicron": "ο", "pi": "π",
    "rho": "ρ", "sigma": "σ", "tau": "τ", "upsilon": "υ",
    "phi": "φ", "chi": "χ", "psi": "ψ", "omega": "ω",
    "Gamma": "Γ", "Delta": "Δ", "Theta": "Θ", "Lambda": "Λ",
    "Xi": "Ξ", "Pi": "Π", "Sigma": "Σ", "Phi": "Φ",
    "Psi": "Ψ", "Omega": "Ω",
    "varepsilon": "ε", "varphi": "φ", "vartheta": "ϑ",
}

_COMMANDS = {
    "times": "×", "div": "÷", "pm": "±", "mp": "∓",
    "cdot": "·", "leq": "≤", "geq": "≥", "neq": "≠",
    "approx": "≈", "equiv": "≡", "propto": "∝", "sim": "∼",
    "rightarrow": "→", "leftarrow": "←", "leftrightarrow": "↔",
    "Rightarrow": "⇒", "Leftarrow": "⇐",
    "infty": "∞", "partial": "∂", "nabla": "∇",
    "forall": "∀", "exists": "∃", "emptyset": "∅",
    "angle": "∠", "triangle": "△", "circ": "∘",
    "cdots": "⋯", "ldots": "…", "vdots": "⋮", "ddots": "⋱",
    "to": "→", "mapsto": "↦", "implies": "⇒", "iff": "⇔",
    "cup": "∪", "cap": "∩", "subseteq": "⊆", "subset": "⊂",
    "oplus": "⊕", "otimes": "⊗", "perp": "⟂",
    "bullet": "•", "star": "★", "dagger": "†",
}


def _omml_from_tree(node: dict, display: bool = False) -> str:
    """Convert AST to OMML XML. Returns inner content (no outer wrap)."""
    items = _omml_items(node.get("items", []))
    return items


def _omml_items(items: list) -> str:
    result = []
    i = 0
    while i < len(items):
        item = items[i]

        # sub/sup pairs: if next item is sup, combine
        if item.get("type") == "sub":
            sub_content = _omml_single(item.get("content", {}))
            sup_content = ""
            if i + 1 < len(items) and items[i + 1].get("type") == "sup":
                sup_content = _omml_single(items[i + 1].get("content", {}))
                i += 1
            result.append(_wrap("sSubSup",
                _wrap("e", _omml_prev(result)) + _wrap("sub", sub_content) + _wrap("sup", sup_content)))
            i += 1
            continue

        if item.get("type") == "sup":
            base = _omml_prev(result)
            sup_content = _omml_single(item.get("content", {}))
            result.append(_wrap("sSup",
                _wrap("e", base) + _wrap("sup", sup_content)))
            i += 1
            continue

        result.append(_omml_item(item))
        i += 1

    # Collapse into runs
    if not result:
        return ""
    return "".join(result)


def _omml_prev(items: list) -> str:
    """Get the last item's content as a base element."""
    if items:
        last = items.pop()
        return last
    return _wrap("r", _wrap("t", ""))


def _omml_item(item: dict) -> str:
    t = item.get("type", "")
    if t == "text":
        return _wrap("r", _wrap("t", html.escape(item.get("text", ""))))
    if t == "group":
        inner = _omml_items(item.get("items", []))
        return _wrap("d", inner)
    if t == "row_sep":
        return ""  # handled at matrix level
    if t == "column_sep":
        return ""  # handled at matrix level

    if t == "command":
        return _omml_command(item)

    return ""


def _omml_single(node: dict) -> str:
    if node.get("type") == "text":
        return _wrap("r", _wrap("t", html.escape(node.get("text", ""))))
    if node.get("type") == "group":
        return _wrap("d", _omml_items(node.get("items", [])))
    if node.get("type") == "command":
        return _omml_command(node)
    return ""


def _omml_command(node: dict) -> str:
    cmd = node.get("cmd", "")

    # Fractions
    if cmd == "frac":
        num = _omml_single(node.get("num", {}))
        den = _omml_single(node.get("den", {}))
        return _wrap("f", _wrap("num", num) + _wrap("den", den))

    # Square root
    if cmd == "sqrt":
        content = _omml_single(node.get("content", {}))
        if "degree" in node:
            deg = _omml_single(node["degree"])
            return _wrap("rad", _wrap("deg", deg) + _wrap("e", content))
        return _wrap("rad", _wrap("e", content))

    # Large operators
    if cmd in ("sum", "prod", "int", "oint", "iint", "iiint"):
        op_map = {"sum": "∑", "prod": "∏", "int": "∫",
                   "oint": "∮", "iint": "∬", "iiint": "∭"}
        func_chr = _wrap("r", _wrap("t", op_map.get(cmd, cmd)))

        lower = _wrap("limLow", _wrap("r", _wrap("t", _omml_single(node.get("lower", {}))))) if "lower" in node else ""
        upper = _wrap("limUpp", _wrap("r", _wrap("t", _omml_single(node.get("upper", {}))))) if "upper" in node else ""

        return _wrap("func", func_chr + lower + upper)

    # Accents
    if cmd == "bar":
        return _wrap("acc", _wrap("accPr", _el("chr", val="0304") + _el_close("chr") + _el_close("accPr"))
                     + _wrap("e", _omml_single(node.get("content", {}))))
    if cmd == "vec":
        return _wrap("acc", _wrap("accPr", _el("chr", val="20D7") + _el_close("chr") + _el_close("accPr"))
                     + _wrap("e", _omml_single(node.get("content", {}))))
    if cmd == "hat":
        return _wrap("acc", _wrap("accPr", _el("chr", val="0302") + _el_close("chr") + _el_close("accPr"))
                     + _wrap("e", _omml_single(node.get("content", {}))))
    if cmd == "tilde":
        return _wrap("acc", _wrap("accPr", _el("chr", val="0303") + _el_close("chr") + _el_close("accPr"))
                     + _wrap("e", _omml_single(node.get("content", {}))))
    if cmd == "dot":
        return _wrap("acc", _wrap("accPr", _el("chr", val="0307") + _el_close("chr") + _el_close("accPr"))
                     + _wrap("e", _omml_single(node.get("content", {}))))
    if cmd == "ddot":
        return _wrap("acc", _wrap("accPr", _el("chr", val="0308") + _el_close("chr") + _el_close("accPr"))
                     + _wrap("e", _omml_single(node.get("content", {}))))

    # Overline / underline
    if cmd == "overline":
        return _wrap("bar", _wrap("barPr", _el("pos", val="top") + _el_close("pos") + _el_close("barPr"))
                     + _wrap("e", _omml_single(node.get("content", {}))))
    if cmd == "underline":
        return _wrap("bar", _wrap("barPr", _el("pos", val="bot") + _el_close("pos") + _el_close("barPr"))
                     + _wrap("e", _omml_single(node.get("content", {}))))

    # Greek
    if cmd in _GREEK:
        return _wrap("r", _wrap("t", _GREEK[cmd]))

    # Symbol commands
    if cmd in _COMMANDS:
        return _wrap("r", _wrap("t", _COMMANDS[cmd]))

    # Matrix / cases environments
    if node.get("type") == "environment":
        return _omml_env(node)

    # Fallback: show command name
    return _wrap("r", _wrap("t", cmd))


def _omml_env(node: dict) -> str:
    env = node.get("env", "")
    items = node.get("items", [])

    if env in ("matrix", "pmatrix", "bmatrix", "vmatrix"):
        rows = _split_matrix(items)
        mc = _wrap("m", _omml_matrix_rows(rows))
        if env == "pmatrix":
            return _wrap("d",
                _wrap("dPr", _el("begChr", val="(") + _el_close("begChr") + _el("endChr", val=")") + _el_close("endChr") + _el_close("dPr")) + mc)
        if env == "bmatrix":
            return _wrap("d",
                _wrap("dPr", _el("begChr", val="[") + _el_close("begChr") + _el("endChr", val="]") + _el_close("endChr") + _el_close("dPr")) + mc)
        if env == "vmatrix":
            return _wrap("d",
                _wrap("dPr", _el("begChr", val="|") + _el_close("begChr") + _el("endChr", val="|") + _el_close("endChr") + _el_close("dPr")) + mc)
        return mc

    return _wrap("r", _wrap("t", f"[{env}]"))


def _split_matrix(items: list) -> list:
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
    result = []
    for row in rows:
        cells = []
        for cell in row:
            cells.append(_omml_single(cell))
        result.append(_wrap("mr", "".join(cells)))
    return "".join(result)


# ─── Full OMML document wrapper ─────────────────────────────────────────────────

def _omml_wrap(inner: str, display: bool = False) -> str:
    """Wrap inner OMML content in the proper Word math paragraph/run structure."""
    if display:
        return (
            f'<m:oMathPara {_el("", xmlns_m=_OMML_NS)[:-1]}>'
            f'<m:oMath>{inner}</m:oMath>'
            f'</m:oMathPara>'
        )
    else:
        return (
            f'<m:oMath {_el("", xmlns_m=_OMML_NS)[:-1]}>'
            f'{inner}'
            f'</m:oMath>'
        )
