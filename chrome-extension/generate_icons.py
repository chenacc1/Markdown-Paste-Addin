#!/usr/bin/env python3
"""Generate simple PNG icons for the Chrome extension using Pillow (if available)
or create SVG-based placeholder text."""

import os
import sys

ICON_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icons")

SIZES = [16, 48, 128]

# ── Try Pillow ──
try:
    from PIL import Image, ImageDraw, ImageFont

    def generate_icon(size):
        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        # Blue circle background
        margin = size // 8
        draw.ellipse(
            [margin, margin, size - margin, size - margin],
            fill=(37, 99, 235, 255)  # #2563eb
        )

        # White "W" letter
        try:
            font_size = size * 3 // 5
            font = ImageFont.truetype("C:\\Windows\\Fonts\\segoeui.ttf", font_size)
        except Exception:
            font = ImageFont.load_default()

        text = "W"
        bbox = draw.textbbox((0, 0), text, font=font)
        tw = bbox[2] - bbox[0]
        th = bbox[3] - bbox[1]
        x = (size - tw) // 2
        y = (size - th) // 2
        draw.text((x, y), text, fill=(255, 255, 255, 255), font=font)

        return img

    def main():
        os.makedirs(ICON_DIR, exist_ok=True)
        for size in SIZES:
            img = generate_icon(size)
            path = os.path.join(ICON_DIR, f"icon{size}.png")
            img.save(path)
            print(f"  Created: {path} ({os.path.getsize(path)} bytes)")
        print("Done!")

except ImportError:
    print("Pillow not installed. Creating SVG-based placeholder icons...")
    print("Tip: pip install Pillow for proper PNG icon generation")

    import codecs

    def main():
        os.makedirs(ICON_DIR, exist_ok=True)

        # Create a simple 1x1 PNG placeholder as base64
        # Minimal valid PNG (1x1 blue pixel)
        # This is a proper PNG file header
        import base64
        import struct
        import zlib

        for size in SIZES:
            # Create a minimal valid PNG with a blue square
            def create_png(w, h, r, g, b):
                def chunk(chunk_type, data):
                    c = chunk_type + data
                    crc = struct.pack(">I", zlib.crc32(c) & 0xffffffff)
                    return struct.pack(">I", len(data)) + c + crc

                # IHDR
                ihdr = struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)
                # IDAT
                raw = b""
                for y in range(h):
                    raw += b"\x00" + bytes([r, g, b]) * w
                idat = zlib.compress(raw)

                return (b"\x89PNG\r\n\x1a\n" +
                        chunk(b"IHDR", ihdr) +
                        chunk(b"IDAT", idat) +
                        chunk(b"IEND", b""))

            png_bytes = create_png(size, size, 37, 99, 235)
            path = os.path.join(ICON_DIR, f"icon{size}.png")
            with open(path, "wb") as f:
                f.write(png_bytes)
            print(f"  Created placeholder: {path} ({len(png_bytes)} bytes)")
        print("Done! (placeholder icons — use Pillow for proper text)")
        print("  pip install Pillow")
        print(f"  python {__file__}")


if __name__ == "__main__":
    main()
