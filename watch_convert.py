#!/usr/bin/env python3
"""Watch Markdown file and auto-convert on changes.

Usage:
  python watch_convert.py input.md output.docx
  python watch_convert.py input.md output.docx --interval 3
"""

import argparse
import os
import time
import sys
from pathlib import Path

from md2docx_lib import parse_markdown, build_docx


def watch_and_convert(input_path: str, output_path: str, interval: float = 2.0):
    if not os.path.isfile(input_path):
        print(f"Error: file not found — {input_path}")
        sys.exit(1)

    last_mtime = 0
    print(f"Watching: {input_path}")
    print(f"Output:   {output_path}")
    print(f"Press Ctrl+C to stop.\n")

    while True:
        try:
            mtime = os.path.getmtime(input_path)
            if mtime != last_mtime:
                last_mtime = mtime
                print(f"[{time.strftime('%H:%M:%S')}] File changed, converting...")
                text = Path(input_path).read_text(encoding="utf-8")
                chunks = parse_markdown(text)
                build_docx(chunks, output_path)
                print(f"  Done: {output_path}")

            time.sleep(interval)
        except KeyboardInterrupt:
            print("\nStopped.")
            break


def main():
    parser = argparse.ArgumentParser(
        description="Watch a Markdown file and auto-convert to Word on changes."
    )
    parser.add_argument("input", help="Input .md file to watch")
    parser.add_argument("output", nargs="?", help="Output .docx path")
    parser.add_argument("--interval", "-i", type=float, default=2.0,
                        help="Poll interval in seconds (default: 2.0)")

    args = parser.parse_args()

    if not args.output:
        args.output = Path(args.input).with_suffix(".docx")

    watch_and_convert(args.input, str(args.output), args.interval)


if __name__ == "__main__":
    main()
