"""Mermaid diagram renderer — local mmdc with online fallback."""

import base64
import shutil
import subprocess
import tempfile
from pathlib import Path

import requests


def _find_npx() -> str:
    node_path = shutil.which("node") or shutil.which("node.exe")
    if node_path:
        node_dir = Path(node_path).parent
        for name in ("npx.cmd", "npx.exe", "npx"):
            candidate = node_dir / name
            if candidate.exists():
                return str(candidate)
    for name in ("npx.cmd", "npx.exe", "npx"):
        found = shutil.which(name)
        if found:
            return found
    return "npx"


def render_mermaid_mmdc(code: str, scale: int = 2) -> bytes | None:
    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            inp = Path(tmpdir) / "diagram.mmd"
            out = Path(tmpdir) / "diagram.png"
            inp.write_text(code, encoding="utf-8")

            npx = _find_npx()
            result = subprocess.run(
                [npx, "--yes", "@mermaid-js/mermaid-cli",
                 "-i", str(inp), "-o", str(out),
                 "-s", str(scale), "-b", "transparent"],
                capture_output=True, text=True, timeout=30,
                cwd=tmpdir,
            )
            if result.returncode == 0 and out.exists():
                return out.read_bytes()
            print(f"  [warn] mmdc failed (exit {result.returncode}): {result.stderr[:200]}")
            return None
    except FileNotFoundError:
        print("  [warn] Node.js not found, trying online API...")
        return None
    except subprocess.TimeoutExpired:
        print("  [warn] mmdc timed out")
        return None
    except Exception as e:
        print(f"  [warn] mmdc error: {e}")
        return None


def render_mermaid_online(code: str) -> bytes | None:
    try:
        encoded = base64.urlsafe_b64encode(code.encode()).decode()
        url = f"https://mermaid.ink/img/{encoded}?type=png"
        resp = requests.get(url, timeout=15)
        if resp.status_code == 200:
            return resp.content
        print(f"  [warn] Online API returned {resp.status_code}")
        return None
    except Exception as e:
        print(f"  [warn] Online API error: {e}")
        return None


def render_mermaid(code: str, scale: int = 2) -> bytes | None:
    print(f"  Rendering diagram ({len(code.split(chr(10)))} lines)...")
    result = render_mermaid_mmdc(code, scale)
    if result:
        print("  OK (local mmdc)")
        return result
    print("  Falling back to online API...")
    result = render_mermaid_online(code)
    if result:
        print("  OK (online API)")
        return result
    print("  FAILED to render diagram")
    return None
