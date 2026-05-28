#!/usr/bin/env python3
"""
Word Document Format Unifier (v2.0)
Auto-detect heading/caption/body types and apply consistent formatting.

Usage:
  python format_docx.py document.docx                   # format in-place
  python format_docx.py document.docx output.docx       # output to new file
  python format_docx.py --active                        # format active Word doc
  python format_docx.py --show                          # show current presets
  python format_docx.py --lang en document.docx         # English formatting
  python format_docx.py --list-presets                  # list preset names
"""

import argparse
import os
import sys

from docx import Document

from md2docx_lib.formatter import (
    format_document, FORMAT, detect_paragraph_type, apply_format, set_default_font
)


def main():
    parser = argparse.ArgumentParser(
        description="Word document format unifier (v2.0)"
    )
    parser.add_argument("input", nargs="?", help="Input .docx file path")
    parser.add_argument("output", nargs="?", help="Output path (default: overwrite)")
    parser.add_argument("--active", "-a", action="store_true",
                        help="Format active Word document")
    parser.add_argument("--show", "-s", action="store_true",
                        help="Show current format presets")
    parser.add_argument("--lang", default="cn",
                        help="Language mode: cn (Chinese) or en (English)")
    parser.add_argument("--list-presets", action="store_true",
                        help="List available preset names")

    args = parser.parse_args()

    if args.show:
        _show_presets()
        return

    if args.list_presets:
        print("\n".join(FORMAT.keys()))
        return

    if args.active:
        _format_active(args.lang)
        return

    if not args.input:
        parser.print_help()
        sys.exit(1)

    _format_file(args.input, args.output, args.lang)


def _format_file(input_path: str, output_path: str, lang: str):
    if not os.path.exists(input_path):
        print(f"Error: file not found — {input_path}")
        sys.exit(1)

    out = output_path or input_path
    print(f"Formatting: {input_path} (lang={lang})\n")
    doc = Document(input_path)
    stats = format_document(doc, lang)
    doc.save(out)

    print("Done:")
    for name, count in stats.items():
        if count > 0:
            print(f"  {name}: {count}")
    print(f"\nSaved: {out}")


def _format_active(lang: str):
    try:
        import win32com.client
        word = win32com.client.GetObject(Class="Word.Application")
    except Exception:
        print("Error: cannot connect to Word. Make sure Word is open with a document.")
        return

    doc = word.ActiveDocument
    if doc is None:
        print("Error: no document open in Word.")
        return

    for i in range(1, doc.Paragraphs.Count + 1):
        para = doc.Paragraphs(i)
        text = para.Range.Text.strip()
        if not text:
            continue

        ptype = _detect_com(para, lang)
        if ptype == "空行":
            continue

        fmt = FORMAT.get(ptype, FORMAT["正文"])
        _apply_com(para, fmt)
        if i % 50 == 0:
            print(f"  Processing... {i}/{doc.Paragraphs.Count}")

    print(f"  Processed {doc.Paragraphs.Count} paragraphs. Done.")


def _detect_com(para, lang: str) -> str:
    text = para.Range.Text.strip()
    if not text:
        return "空行"
    return detect_paragraph_type(_ComAdapter(para), lang)


def _apply_com(para, fmt: dict):
    pf = para.Format
    if "space_before" in fmt:
        pf.SpaceBefore = fmt["space_before"]
    if "space_after" in fmt:
        pf.SpaceAfter = fmt["space_after"]
    if "line_spacing" in fmt:
        pf.LineSpacingRule = 1
    if "alignment" in fmt:
        para.Alignment = fmt["alignment"]
    if "first_line_indent" in fmt:
        pf.FirstLineIndent = fmt["first_line_indent"]

    rng = para.Range
    if "font_name" in fmt:
        rng.Font.Name = fmt["font_name"]
    if "font_size" in fmt:
        rng.Font.Size = fmt["font_size"]
    if "bold" in fmt:
        rng.Font.Bold = fmt["bold"]
    if "color" in fmt:
        rng.Font.Color = fmt["color"]


class _ComAdapter:
    """Adapter to make COM paragraph look like python-docx paragraph."""
    def __init__(self, com_para):
        self.text = com_para.Range.Text
        self.style = type("Style", (), {"name": com_para.Style.NameLocal})()
        self._runs = [_ComRunAdapter(com_para.Range.Font)]

    @property
    def runs(self):
        return self._runs


class _ComRunAdapter:
    def __init__(self, font):
        self.bold = font.Bold
        self.font = type("Font", (), {"size": font.Size})()


def _show_presets():
    print("\nCurrent Format Presets:\n")
    for name, fmt in FORMAT.items():
        print(f"  [{name}]")
        print(f"    Font: {fmt.get('font_name','-')}  Size: {_pt(fmt.get('font_size'))}")
        print(f"    Bold: {fmt.get('bold',False)}  Align: {_al(fmt.get('alignment'))}")
        print(f"    Before: {_pt(fmt.get('space_before'))}  After: {_pt(fmt.get('space_after'))}")
        print(f"    Line spacing: {fmt.get('line_spacing','-')}")
        if fmt.get("first_line_indent"):
            print(f"    First line indent: yes")
        print()


def _pt(val) -> str:
    if val is None:
        return "-"
    from docx.shared import Pt
    if isinstance(val, Pt):
        return f"{int(val / 12700)}pt"
    return str(val)


def _al(val) -> str:
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    if val == WD_ALIGN_PARAGRAPH.CENTER:
        return "center"
    elif val == WD_ALIGN_PARAGRAPH.RIGHT:
        return "right"
    return "left"


if __name__ == "__main__":
    main()
