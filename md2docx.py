#!/usr/bin/env python3
"""
md2docx — Clipboard/File to Word (.docx) Converter (v2.0)

Copy content from DeepSeek, web pages, or Markdown files and convert to Word with:
  - Markdown/HTML tables    → Word-native tables
  - Mermaid diagrams        → embedded PNG images
  - Embedded images (URL)   → downloaded and embedded
  - LaTeX math ($...$)      → Word native OMML equations
  - Code blocks             → syntax-highlighted (Pygments)
  - Task lists (- [ ])      → checkbox lists
  - Blockquotes (>)         → styled quotes
  - Horizontal rules        → Word dividers
  - Nested lists            → multi-level list styles
  - Auto-captions           → Figure/Table numbering
  - Table of Contents       → TOC field code

Usage:
  python md2docx.py output.docx                      # from clipboard
  python md2docx.py -c output.docx                   # from clipboard (explicit)
  python md2docx.py input.md output.docx             # from file
  python md2docx.py input.md output.docx --toc        # with table of contents
  python md2docx.py input.md output.docx --template T.docx  # with template
"""

import argparse
import os
import sys
from pathlib import Path

from md2docx_lib import parse_markdown, parse_html, build_docx, DocumentBuilder
from md2docx_lib.clipboard import read_clipboard
from md2docx_lib.formatter import format_document


def main():
    parser = argparse.ArgumentParser(
        description="Convert clipboard or Markdown file to Word .docx (v3.2)"
    )
    parser.add_argument("arg1", nargs="?",
                        help="Output .docx path (clipboard mode) OR Input .md file")
    parser.add_argument("arg2", nargs="?",
                        help="Output .docx path (when arg1 is input file)")
    parser.add_argument("-c", "--clipboard", action="store_true",
                        help="Read from clipboard")
    parser.add_argument("--toc", action="store_true",
                        help="Add table of contents")
    parser.add_argument("--title", default="",
                        help="Document title")
    parser.add_argument("--template", default="",
                        help="Word template .docx to use as base")
    parser.add_argument("--no-captions", action="store_true",
                        help="Disable auto figure/table numbering")
    parser.add_argument("--format", action="store_true",
                        help="Apply format_docx style after conversion")
    parser.add_argument("--image-width", type=float, default=5.5,
                        help="Max image width in inches (default: 5.5)")

    args = parser.parse_args()

    # Determine mode
    use_clipboard = args.clipboard or not args.arg1 or (
        args.arg1 and not args.arg1.endswith(".md"))
    if args.arg1 and args.arg1.endswith(".md") and not args.clipboard:
        use_clipboard = False

    if use_clipboard:
        output_path = args.arg1 or "clipboard-output.docx"
        _process_clipboard(output_path, args)
    else:
        input_path = Path(args.arg1)
        if not input_path.exists():
            print(f"Error: file not found: {args.arg1}")
            sys.exit(1)
        output_path = args.arg2 or str(input_path.with_suffix(".docx"))
        _process_file(str(input_path), output_path, args)


def _process_clipboard(output_path: str, args):
    print("Reading clipboard...")
    data = read_clipboard()
    if not data:
        print("Error: clipboard is empty or has unsupported format")
        sys.exit(1)

    print(f"  Format: {data['format']} ({len(data['content'])} chars)")

    if data["format"] == "html":
        print("Parsing HTML...")
        chunks = parse_html(data["content"])
    else:
        print("Parsing Markdown...")
        chunks = parse_markdown(data["content"])

    _build(chunks, output_path, args)


def _process_file(input_path: str, output_path: str, args):
    print(f"Reading: {input_path}")
    text = Path(input_path).read_text(encoding="utf-8")

    if text.strip().startswith("<"):
        print("Parsing HTML...")
        chunks = parse_html(text)
    else:
        print("Parsing Markdown...")
        chunks = parse_markdown(text)

    _build(chunks, output_path, args)


def _build(chunks: list, output_path: str, args):
    stats = {}
    for c in chunks:
        stats[c.get("type", "text")] = stats.get(c.get("type", "text"), 0) + 1
    parts = [f"{v} {k}s" for k, v in stats.items()]
    print(f"  Found: {', '.join(parts) if parts else 'nothing'}")

    print("Building Word document...")
    builder = DocumentBuilder(
        template_path=args.template,
        auto_captions=not args.no_captions,
        add_toc=args.toc,
        image_max_width=args.image_width,
        title=args.title,
    )
    builder.build(chunks, output_path)

    if args.format:
        print("Applying formatting...")
        from docx import Document
        doc = Document(output_path)
        format_document(doc)
        doc.save(output_path)

    print(f"\nDone: {output_path}")
    print(f"  Open: start \"\" \"{output_path}\"")


if __name__ == "__main__":
    main()
