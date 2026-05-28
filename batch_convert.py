#!/usr/bin/env python3
"""Batch Markdown-to-Word converter.

Usage:
  python batch_convert.py *.md output_dir/
  python batch_convert.py input_dir/ output_dir/
  python batch_convert.py --watch input_dir/ output_dir/
"""

import argparse
import os
import sys
import time
from pathlib import Path

from md2docx_lib import parse_markdown, build_docx


def convert_file(input_path: str, output_path: str) -> bool:
    try:
        text = Path(input_path).read_text(encoding="utf-8")
        chunks = parse_markdown(text)
        build_docx(chunks, output_path)
        print(f"  OK: {input_path} → {output_path}")
        return True
    except Exception as e:
        print(f"  FAIL: {input_path}: {e}")
        return False


def batch_convert(input_pattern: str, output_dir: str):
    """Convert all matching .md files to .docx."""
    import glob
    files = sorted(glob.glob(input_pattern))
    if not files:
        print(f"No files found matching: {input_pattern}")
        return

    os.makedirs(output_dir, exist_ok=True)
    success = 0
    failed = 0

    print(f"Batch converting {len(files)} file(s)...\n")
    for f in files:
        name = Path(f).stem
        out = os.path.join(output_dir, name + ".docx")
        if convert_file(f, out):
            success += 1
        else:
            failed += 1

    print(f"\nDone: {success} succeeded, {failed} failed.")


def watch_dir(input_dir: str, output_dir: str, poll_interval: float = 2.0):
    """Watch a directory for .md file changes and auto-convert."""
    os.makedirs(output_dir, exist_ok=True)
    known = {}

    print(f"Watching: {input_dir}")
    print(f"Output:   {output_dir}")
    print("Press Ctrl+C to stop.\n")

    while True:
        try:
            current_files = {}
            for f in Path(input_dir).glob("*.md"):
                mtime = f.stat().st_mtime
                current_files[str(f)] = mtime

            # New or changed
            for fpath, mtime in current_files.items():
                if fpath not in known or mtime != known[fpath]:
                    name = Path(fpath).stem
                    out = os.path.join(output_dir, name + ".docx")
                    print(f"[{time.strftime('%H:%M:%S')}] Converting: {Path(fpath).name}")
                    convert_file(fpath, out)

            known = current_files
            time.sleep(poll_interval)
        except KeyboardInterrupt:
            print("\nStopped.")
            break


def main():
    parser = argparse.ArgumentParser(description="Batch Markdown to Word converter")
    parser.add_argument("input", help="Input .md file pattern, directory, or file")
    parser.add_argument("output", nargs="?", default=".",
                        help="Output directory (default: current)")
    parser.add_argument("--watch", "-w", action="store_true",
                        help="Watch mode: auto-convert on file changes")
    parser.add_argument("--interval", "-i", type=float, default=2.0,
                        help="Watch poll interval in seconds (default: 2.0)")

    args = parser.parse_args()

    if args.watch:
        watch_dir(args.input, args.output, args.interval)
    else:
        # Support single file, glob, or directory
        if os.path.isdir(args.input):
            pattern = os.path.join(args.input, "*.md")
        elif "*" in args.input or "?" in args.input:
            pattern = args.input
        elif os.path.isfile(args.input):
            name = Path(args.input).stem
            out = os.path.join(args.output, name + ".docx")
            convert_file(args.input, out)
            return
        else:
            print(f"Error: not found — {args.input}")
            sys.exit(1)

        batch_convert(pattern, args.output)


if __name__ == "__main__":
    main()
