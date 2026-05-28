#!/usr/bin/env python3
"""
MarkdownPasteAddin Bridge Server
Local HTTP server that receives Markdown/HTML from Chrome extension and returns .docx.

Usage:
  python bridge_server.py                      # default: http://localhost:9876
  python bridge_server.py --port 9877          # custom port
  python bridge_server.py --no-clipboard       # disable password mode (use if sharing network)
"""

import argparse
import io
import json
import os
import sys
import tempfile
import threading
from http.server import HTTPServer, BaseHTTPRequestHandler
from pathlib import Path
from urllib.parse import parse_qs, urlparse

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent))

from md2docx_lib import parse_markdown, parse_html
from md2docx_lib.builder_docx import DocumentBuilder
from md2docx_lib.formatter import format_document

VERSION = "2.0.0"
DEFAULT_PORT = 9876


class BridgeHandler(BaseHTTPRequestHandler):
    """HTTP request handler for the bridge server."""

    def log_message(self, format, *args):
        print(f"  [{self.log_date_time_string()}] {args[0]}")

    # ── CORS headers (extension runs from chrome-extension:// origin) ───────────

    def _set_cors(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def do_OPTIONS(self):
        self.send_response(204)
        self._set_cors()
        self.end_headers()

    # ── GET /api/health ─────────────────────────────────────────────────────────

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/api/health":
            self.send_response(200)
            self._set_cors()
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps({
                "status": "ok",
                "version": VERSION,
                "features": ["tables", "mermaid", "math", "code_highlight",
                            "task_lists", "blockquotes", "toc", "auto_captions"]
            }).encode())
        else:
            self.send_response(404)
            self.end_headers()

    # ── POST /api/convert ───────────────────────────────────────────────────────

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path != "/api/convert":
            self.send_response(404)
            self.end_headers()
            return

        try:
            # Read request body
            content_length = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(content_length)
            data = json.loads(body)

            content = data.get("content", "")
            fmt = data.get("format", "markdown")
            options = data.get("options", {})

            if not content.strip():
                self._error(400, "Empty content")
                return

            # Strip leftover markdown markers from clipboard plain text
            from md2docx_lib.inline_processor import clean_text
            content = clean_text(content)

            print(f"  Converting: {len(content)} chars, format={fmt}")
            print(f"  Options: {options}")

            # Parse content
            if fmt == "html":
                chunks = parse_html(content)
            else:
                chunks = parse_markdown(content)

            # Stats
            stats = {}
            for c in chunks:
                t = c.get("type", "text")
                stats[t] = stats.get(t, 0) + 1
            parts = [f"{v} {k}" for k, v in sorted(stats.items())]
            print(f"  Found: {', '.join(parts)}")

            # Build document
            builder = DocumentBuilder(
                template_path=options.get("template", ""),
                auto_captions=options.get("auto_captions", True),
                add_toc=options.get("toc", False),
                image_max_width=options.get("image_width", 5.5),
                title=options.get("title", ""),
                show_progress=False,
            )

            # Write to temp file
            with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
                tmp_path = f.name

            try:
                builder.build(chunks, tmp_path)

                # Apply formatting if requested
                if options.get("format_after", False):
                    from docx import Document
                    doc = Document(tmp_path)
                    format_document(doc)
                    doc.save(tmp_path)

                # Apply report standard if requested
                if options.get("preset") and options.get("preset") != "none":
                    from docx import Document
                    from md2docx_lib.report_standards import apply_report_standard, PRESETS
                    preset_key = options["preset"]
                    if preset_key in PRESETS:
                        doc = Document(tmp_path)
                        apply_report_standard(doc, preset=preset_key,
                            title=options.get("title", ""),
                            subtitle=options.get("subtitle", ""),
                            author=options.get("author", ""),
                            org=options.get("org", ""))
                        doc.save(tmp_path)

                # Read back the file
                with open(tmp_path, "rb") as f:
                    docx_bytes = f.read()

                # Send response
                filename = self._sanitize_filename(
                    options.get("title", "export")) + ".docx"

                self.send_response(200)
                self._set_cors()
                self.send_header("Content-Type",
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                self.send_header("Content-Disposition",
                    f'attachment; filename="{filename}"')
                self.send_header("Content-Length", str(len(docx_bytes)))
                self.end_headers()
                self.wfile.write(docx_bytes)

                print(f"  OK: {filename} ({len(docx_bytes)} bytes)")

            finally:
                try:
                    os.unlink(tmp_path)
                except OSError:
                    pass

        except json.JSONDecodeError:
            self._error(400, "Invalid JSON")
        except Exception as e:
            print(f"  ERROR: {e}")
            import traceback
            traceback.print_exc()
            self._error(500, str(e))

    def _error(self, code: int, message: str):
        self.send_response(code)
        self._set_cors()
        self.send_header("Content-Type", "application/json")
        self.end_headers()
        self.wfile.write(json.dumps({"error": message}).encode())

    def _sanitize_filename(self, name: str) -> str:
        import re
        name = name.strip() or "export"
        # Remove all non-ASCII characters (HTTP headers are ASCII-only)
        name = name.encode('ascii', errors='ignore').decode('ascii')
        name = re.sub(r'[<>:"/\\|?*\s]', '_', name)
        name = name.strip('_') or "export"
        return name[:80]


def main():
    parser = argparse.ArgumentParser(
        description="MarkdownPasteAddin Bridge Server v" + VERSION
    )
    parser.add_argument("--port", "-p", type=int, default=DEFAULT_PORT,
                        help=f"Port to listen on (default: {DEFAULT_PORT})")
    parser.add_argument("--host", default="127.0.0.1",
                        help="Host to bind to (default: 127.0.0.1)")

    args = parser.parse_args()

    server = HTTPServer((args.host, args.port), BridgeHandler)

    print(f"""
╔══════════════════════════════════════════════════════════╗
║     MarkdownPasteAddin Bridge Server v{VERSION:<12}      ║
╠══════════════════════════════════════════════════════════╣
║  URL:       http://{args.host}:{args.port:<5}                      ║
║  Health:    http://{args.host}:{args.port}/api/health              ║
║  Convert:   POST http://{args.host}:{args.port}/api/convert        ║
╠══════════════════════════════════════════════════════════╣
║  Press Ctrl+C to stop                                   ║
╚══════════════════════════════════════════════════════════╝
""")

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down...")
        server.shutdown()


if __name__ == "__main__":
    main()
