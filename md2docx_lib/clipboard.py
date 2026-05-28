"""Windows clipboard reading with HTML and plain text support."""

import re
import sys

_HAS_CLIPBOARD = False
try:
    import win32clipboard
    _HAS_CLIPBOARD = True
except ImportError:
    pass


def read_clipboard() -> dict | None:
    """Read clipboard, return {'format': 'html'|'text', 'content': str} or None."""
    if not _HAS_CLIPBOARD:
        print("Error: pywin32 not installed. Run: pip install pywin32")
        return None

    try:
        win32clipboard.OpenClipboard()
        try:
            cf_html = win32clipboard.RegisterClipboardFormat("HTML Format")
            if win32clipboard.IsClipboardFormatAvailable(cf_html):
                raw = win32clipboard.GetClipboardData(cf_html)
                return {"format": "html", "content": _decode_clipboard_html(raw)}

            if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
                text = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
                return {"format": "text", "content": text}

            if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_TEXT):
                text = win32clipboard.GetClipboardData(win32clipboard.CF_TEXT)
                if isinstance(text, bytes):
                    text = text.decode("utf-8", errors="replace")
                return {"format": "text", "content": text}

            return None
        finally:
            win32clipboard.CloseClipboard()
    except Exception as e:
        print(f"Error reading clipboard: {e}")
        return None


def _decode_clipboard_html(raw: bytes | str) -> str:
    """Decode CF_HTML format from Windows clipboard."""
    if isinstance(raw, bytes):
        raw = raw.decode("utf-8", errors="replace")
    m = re.search(r"<html[^>]*>", raw, re.IGNORECASE)
    if not m:
        m = re.search(r"<[^>]+>", raw)
    if m:
        return raw[m.start():]
    return raw
