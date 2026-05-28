"""Image downloader with data: URI and HTTP/HTTPS support."""

import base64
import re
import requests


def download_image(url: str) -> bytes | None:
    """Download image from URL or decode data: URI. Returns raw bytes or None."""
    if url.startswith("data:"):
        m = re.match(r"data:[^;]*;base64,(.+)", url)
        if m:
            try:
                return base64.b64decode(m.group(1))
            except Exception:
                pass
        return None

    try:
        resp = requests.get(url, timeout=15,
                           headers={"User-Agent": "Mozilla/5.0"})
        if resp.status_code == 200:
            return resp.content
        print(f"  [warn] Image download failed: {url} (status {resp.status_code})")
        return None
    except Exception as e:
        print(f"  [warn] Image download error: {url} ({e})")
        return None
