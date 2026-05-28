#!/usr/bin/env python3
"""
Generate Chrome Web Store & Edge Add-ons promotional images v2.
Professional designs with proper typography and visual hierarchy.
Requires: pip install Pillow
"""

import os, sys
try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError:
    print("Pillow not installed. Run: pip install Pillow")
    sys.exit(1)

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "store-assets")
for d in ["screenshots", "chrome", "edge"]:
    os.makedirs(os.path.join(OUTPUT_DIR, d), exist_ok=True)

# ─── Brand ──────────────────────────────────────────────────────────────────────
BLUE  = (37, 99, 235)
BLUE_D = (29, 78, 216)
WHITE = (255, 255, 255)
DARK  = (15, 23, 42)
GRAY  = (100, 116, 139)
LIGHT = (241, 245, 249)
GREEN = (34, 197, 94)
RED   = (239, 68, 68)
CARD  = (255, 255, 255)
BORDER = (226, 232, 240)

# ─── Fonts ──────────────────────────────────────────────────────────────────────
FONT_DIR = "C:\\Windows\\Fonts"
_loaded = {}

def _font(name, size):
    key = (name, size)
    if key in _loaded: return _loaded[key]
    paths = [
        os.path.join(FONT_DIR, name),
        os.path.join(FONT_DIR, name + ".ttf"),
        os.path.join(FONT_DIR, name + ".ttc"),
    ]
    for p in paths:
        if os.path.exists(p):
            try: _loaded[key] = ImageFont.truetype(p, size); return _loaded[key]
            except: pass
    _loaded[key] = ImageFont.load_default(); return _loaded[key]

def f_bold(size):    return _font("segoeuib", size)
def f_regular(size): return _font("segoeui", size)
def f_cjk(size):     return _font("msyh", size) or _font("simhei", size) or f_regular(size)

# ─── Drawing helpers ────────────────────────────────────────────────────────────
class D:
    def __init__(self, draw, w, h): self.d, self.w, self.h = draw, w, h
    def text(self, xy, t, size=14, bold=False, cjk=False, color=DARK, anchor=None, font=None):
        fn = font or (f_bold(size) if bold else (f_cjk(size) if cjk else f_regular(size)))
        kw = {"fill": color, "font": fn}
        if anchor: kw["anchor"] = anchor
        self.d.text(xy, t, **kw)
    def rect(self, xy, fill, radius=0, outline=None):
        x1,y1,x2,y2 = xy
        if radius and hasattr(self.d, 'rounded_rectangle'):
            self.d.rounded_rectangle(xy, radius, fill=fill, outline=outline or fill)
        else:
            self.d.rectangle(xy, fill=fill, outline=outline)
    def circle(self, xy, r, fill): self.d.ellipse([xy[0]-r,xy[1]-r,xy[0]+r,xy[1]+r], fill=fill)
    def line(self, *args):
        # Accept both: line(x1,y1,x2,y2,color,width) and line((x1,y1),(x2,y2),color,width)
        if len(args) >= 5 and isinstance(args[0], (int,float)):
            x1,y1,x2,y2,color = args[:5]
            width = args[5] if len(args) >= 6 else 1
            self.d.line([(x1,y1),(x2,y2)], fill=color, width=width)
        elif len(args) >= 3 and isinstance(args[0], tuple):
            self.d.line(args[:2], fill=args[2], width=args[3] if len(args)>=4 else 1)
        else:
            self.d.line(*args)

# ─── Screenshot 1: Main Interface ──────────────────────────────────────────────
def make_screenshot1():
    W, H = 1280, 800
    img = Image.new("RGB", (W,H), (248,250,252))
    d = D(ImageDraw.Draw(img), W, H)

    # --- Browser chrome ---
    d.rect((0,0,W,36), (30,30,35))
    d.circle((18,18), 6, (237,67,55))
    d.circle((36,18), 6, (246,190,32))
    d.circle((54,18), 6, (38,206,76))
    d.rect((72,8,W-240,28), (50,50,55), radius=8)
    d.text((80,10), "deepseek.com/chat", size=11, color=(200,200,200))
    d.rect((W-120,8,W-96,28), BLUE, radius=4)
    d.text((W-108,12), "W", size=11, bold=True, color=WHITE)
    d.text((W-84,12), "MarkdownPasteAddin", size=10, color=(200,200,210))

    # --- Content area ---
    y = 56
    # Title
    d.text((40,y), "# 系统架构设计报告", size=22, bold=True, cjk=True)
    y += 40
    d.text((40,y), "以下为系统整体架构及各模块说明：", size=13, cjk=True, color=GRAY)
    y += 32

    # ASCII Tree → transformed to diagram
    d.rect((40,y,40+560,y+180), CARD, radius=10, outline=BORDER)
    # Label badge
    d.rect((52, y+10, 200, y+30), (239,246,255), radius=6)
    d.text((62,y+13), "Text Diagram → Mermaid", size=10, color=BLUE)
    # Mock diagram nodes
    nodes = [("系统架构", 280, y+50, BLUE),
             ("前端层", 180, y+90, (59,130,246)),
             ("Web应用", 100, y+130, (99,160,246)),
             ("移动端", 260, y+130, (99,160,246)),
             ("后端层", 380, y+90, (59,130,246)),
             ("API网关", 340, y+130, (99,160,246)),
             ("数据库", 420, y+130, (99,160,246))]
    for label, nx, ny, clr in nodes:
        d.rect((nx-40, ny-14, nx+40, ny+14), clr, radius=14)
        d.text((nx, ny-2), label, size=10, bold=True, color=WHITE, anchor="mm", cjk=True)
    # Connector lines
    d.line(260, y+64, 180, y+78, (180,190,210), 2)
    d.line(260, y+64, 380, y+78, (180,190,210), 2)
    d.line(180, y+104, 100, y+118, (180,190,210), 2)
    d.line(180, y+104, 260, y+118, (180,190,210), 2)
    d.line(380, y+104, 340, y+118, (180,190,210), 2)
    d.line(380, y+104, 420, y+118, (180,190,210), 2)

    y += 195

    # Table
    d.text((40,y), "## 技术栈对比", size=16, bold=True, cjk=True)
    y += 28
    headers = [("技术",80), ("版本",200), ("用途",320), ("状态",440)]
    for h, hx in headers:
        d.rect((hx-10,y,hx+100,y+28), (217,226,243), outline=(200,210,220))
        d.text((hx,y+5), h, size=11, bold=True, cjk=True)
    rows = [("React", "18.2", "前端框架", "✓"), ("Python", "3.13", "文档生成", "✓"), ("Mermaid", "10.x", "图表渲染", "✓")]
    for ri, row in enumerate(rows):
        yy = y + 28 + ri*26
        bg = CARD if ri % 2 == 0 else LIGHT
        for ci, (h, hx) in enumerate(headers):
            d.rect((hx-10, yy, hx+100, yy+26), bg, outline=BORDER)
            d.text((hx, yy+3), row[ci], size=11, cjk=True, color=DARK if row[ci]!="✓" else GREEN)

    y += 140

    # Bottom features bar
    d.rect((40, y, W-40, y+44), CARD, radius=10, outline=BORDER)
    features = ["Markdown表格 → Word表格", "Mermaid流程图 → PNG", "LaTeX公式 → OMML",
                "代码块 → 语法高亮", "框架图 → 渲染图", "3种报告格式标准"]
    fx = 60
    for feat in features:
        d.rect((fx, y+10, fx+len(feat)*8+16, y+34), LIGHT, radius=6)
        d.text((fx+8, y+15), feat, size=10, cjk=True, color=GRAY)
        fx += len(feat)*8 + 24

    # Floating button
    bx, by = W-60, H-100
    d.circle((bx,by), 26, BLUE)
    d.text((bx,by), "W", size=18, bold=True, color=WHITE, anchor="mm")
    d.circle((bx+16,by-20), 5, GREEN)

    img.save(os.path.join(OUTPUT_DIR, "screenshots", "screenshot-1-main.png"))
    print("  screenshot-1-main.png (main interface)")


# ─── Screenshot 2: Popup ───────────────────────────────────────────────────────
def make_screenshot2():
    W, H = 1280, 800
    img = Image.new("RGB", (W,H), (248,250,252))
    d = D(ImageDraw.Draw(img), W, H)

    # Browser frame
    d.rect((0,0,W,36), (30,30,35))
    d.circle((18,18), 6, (237,67,55))
    d.circle((36,18), 6, (246,190,32))
    d.circle((54,18), 6, (38,206,76))
    d.rect((72,8,W-240,28), (50,50,55), radius=8)
    d.text((80,10), "deepseek.com/chat", size=11, color=(200,200,200))
    d.rect((W-120,8,W-96,28), BLUE, radius=4)
    d.text((W-108,12), "W", size=11, bold=True, color=WHITE)

    # Page content (dimmed)
    y = 56
    d.text((40,y), "# 系统架构设计报告", size=18, color=(180,185,195), cjk=True)
    y += 28
    for _ in range(8):
        d.rect((40,y,W-40,y+10), (220,225,235), radius=4)
        y += 16

    # Popup card overlay
    px, py = W - 370, 56
    pw, ph = 340, 560

    # Shadow
    d.rect((px+4, py+4, px+pw+4, py+ph+4), (0,0,0,80) if False else (220,225,235), radius=16)

    # Popup body
    d.rect((px, py, px+pw, py+ph), WHITE, radius=16, outline=BORDER)

    # Header
    d.rect((px+14, py+10, px+44, py+40), BLUE, radius=10)
    d.text((px+29, py+22), "W", size=18, bold=True, color=WHITE, anchor="mm")
    d.text((px+54, py+14), "MarkdownPasteAddin", size=14, bold=True, cjk=True)
    d.text((px+54, py+32), "v3.0 · Chrome & Edge", size=10, color=GRAY)

    # Status
    d.rect((px+14, py+52, px+pw-14, py+74), LIGHT, radius=8)
    d.circle((px+24, py+63), 4, GREEN)
    d.text((px+34, py+59), "Connected · v2.0.0", size=11, color=GRAY)

    # Export button
    d.rect((px+14, py+86, px+pw-14, py+120), BLUE, radius=10)
    d.text((px+pw//2, py+98), "Export Page to Word", size=14, bold=True, color=WHITE, anchor="mm", font=f_regular(14))

    # Secondary buttons
    d.rect((px+14, py+128, px+pw-14, py+158), WHITE, radius=10, outline=BORDER)
    d.text((px+pw//2, py+139), "Export Selection to Word", size=12, color=GRAY, anchor="mm", font=f_regular(12))
    d.rect((px+14, py+164, px+pw-14, py+194), WHITE, radius=10, outline=BORDER)
    d.text((px+pw//2, py+175), "Open Side Panel", size=12, color=GRAY, anchor="mm", font=f_regular(12))

    # Format preset
    d.text((px+16, py+210), "Report Format", size=10, color=GRAY)
    d.rect((px+14, py+224, px+pw-14, py+248), WHITE, radius=6, outline=BORDER)
    d.text((px+22, py+231), "Business Report", size=12, color=DARK)
    d.text((px+pw-30, py+231), "▾", size=10, color=GRAY)

    # Option chips
    d.text((px+16, py+262), "Options", size=10, color=GRAY)
    chip_y = py + 276
    for label, checked in [("TOC", False), ("Numbering", True), ("Format", True), ("Cover", False)]:
        cw = len(label) * 8 + 36
        bg = (239,246,255) if checked else WHITE
        bc = BLUE if checked else BORDER
        d.rect((px+14, chip_y, px+14+cw, chip_y+24), bg, radius=16, outline=bc)
        clr = BLUE if checked else GRAY
        d.text((px+22, chip_y+5), label, size=10, color=clr)
        chip_y += 28

    # Footer
    d.line(px+16, py+ph-28, px+pw-16, py+ph-28, BORDER, 1)
    d.text((px+20, py+ph-20), "Settings", size=10, color=BLUE)
    d.text((px+pw-40, py+ph-20), "v3.0", size=10, color=GRAY)

    img.save(os.path.join(OUTPUT_DIR, "screenshots", "screenshot-2-popup.png"))
    print("  screenshot-2-popup.png (extension popup)")


# ─── Screenshot 3: Word Output ─────────────────────────────────────────────────
def make_screenshot3():
    W, H = 1280, 800
    img = Image.new("RGB", (W,H), (238,238,242))
    d = D(ImageDraw.Draw(img), W, H)

    # Word window frame
    d.rect((0,0,W,40), (30,30,35))
    d.text((16,10), "export.docx — Word", size=11, color=(200,200,200))
    d.rect((W-40,8,W-28,32), (237,67,55), radius=4)
    d.rect((W-66,8,W-48,32), (80,80,85), radius=4)

    # Ribbon
    d.rect((0,40,W,84), (243,243,245))
    tabs = ["File", "Home", "Insert", "Design", "Layout", "References", "Mailings", "Review", "View"]
    for i, tab in enumerate(tabs):
        d.text((16+i*66, 46), tab, size=11, color=DARK if i==1 else GRAY)

    # Content paper
    px, py = 80, 104
    pw, ph = W-160, H-140
    d.rect((px, py, px+pw, py+ph), WHITE, outline=(210,210,215))
    # Page shadow
    d.rect((px+4, py+4, px+pw+4, py+ph+4), (200,200,210), outline=(190,190,200))

    # Cover page preview
    d.rect((px, py, px+pw, py+ph), WHITE)
    cy = py + 40
    # Title
    d.text((px+pw//2, cy), "系统架构设计报告", size=24, bold=True, cjk=True, anchor="mm")
    cy += 36
    d.text((px+pw//2, cy), "MarkdownPasteAddin v3.0 生成", size=13, cjk=True, color=GRAY, anchor="mm")
    cy += 50
    # Decorative line
    d.rect((px+80, cy, px+pw-80, cy+1), (210,215,220))
    cy += 20

    # TOC
    d.text((px+60, cy), "目录", size=16, bold=True, cjk=True)
    cy += 28
    toc_items = ["1. 系统概述", "  1.1 技术栈对比", "  1.2 模块详解", "2. 架构设计", "  2.1 整体框架图", "3. 部署说明"]
    for item in toc_items:
        indent = 20 if item.startswith("  ") else 0
        d.text((px+60+indent, cy), item.strip(), size=12, cjk=True, color=BLUE if not indent else GRAY)
        cy += 20

    cy += 16

    # Section 1
    d.text((px+60, cy), "1. 系统概述", size=16, bold=True, cjk=True)
    cy += 26
    d.rect((px+60, cy, px+pw-60, cy+12), LIGHT, radius=4)
    d.rect((px+62, cy+2, px+180, cy+10), LIGHT, radius=3)
    cy += 20
    d.rect((px+60, cy, px+pw-60, cy+12), LIGHT, radius=4)
    cy += 20

    # Table preview
    d.text((px+60, cy), "1.1 技术栈对比", size=14, bold=True, cjk=True)
    cy += 24
    headers = ["技术", "版本", "用途", "状态"]
    col_w = (pw-120)//4
    for i, h in enumerate(headers):
        d.rect((px+60+i*col_w, cy, px+60+(i+1)*col_w, cy+24), (217,226,243), outline=(190,200,215))
        d.text((px+68+i*col_w, cy+4), h, size=10, bold=True, cjk=True)
    rows = [("React", "18.2", "前端框架", "✓"), ("Python", "3.13", "文档生成", "✓"), ("Mermaid", "10.x", "图表渲染", "✓")]
    for ri, row in enumerate(rows):
        cy2 = cy+24+ri*22
        for i, cell in enumerate(row):
            bg = WHITE if ri%2==0 else LIGHT
            d.rect((px+60+i*col_w, cy2, px+60+(i+1)*col_w, cy2+22), bg, outline=BORDER)
            clr = GREEN if cell == "✓" else DARK
            d.text((px+68+i*col_w, cy2+3), cell, size=10, cjk=True, color=clr)
    cy += 24 + 3*22 + 16

    # Diagram preview
    d.text((px+60, cy), "1.2 整体框架图", size=14, bold=True, cjk=True)
    cy += 24
    d.rect((px+60, cy, px+560, cy+120), LIGHT, radius=8, outline=BORDER)
    # Mini diagram nodes
    nodes = [("系统架构", px+280, cy+20), ("前端层", px+180, cy+55), ("后端层", px+380, cy+55),
             ("Web", px+130, cy+90), ("App", px+230, cy+90), ("API", px+350, cy+90), ("DB", px+410, cy+90)]
    for label, nx, ny in nodes:
        clr = BLUE if label == "系统架构" else (59,130,246)
        d.rect((nx-30, ny-10, nx+30, ny+10), clr, radius=10)
        d.text((nx, ny-1), label, size=9, bold=True, color=WHITE, anchor="mm", cjk=True)
    # Lines
    for src_x, src_y, dst_x, dst_y in [(280,30,180,47),(280,30,380,47),(180,65,130,82),(180,65,230,82),(380,65,350,82),(380,65,410,82)]:
        d.line(src_x, src_y, dst_x, dst_y, (180,190,210), 2)

    d.text((px+580, cy+44), "文本框架图", size=11, cjk=True, color=GRAY)
    d.text((px+580, cy+66), "↓", size=14, color=BLUE)
    d.text((px+580, cy+84), "Mermaid 渲染图", size=11, cjk=True, color=GRAY)

    cy += 136

    # Footer with page number
    d.text((px+pw//2, py+ph-40), "— 1 —", size=10, color=GRAY, anchor="mm")
    d.rect((px+60, py+ph-34, px+pw-60, py+ph-33), BORDER)

    img.save(os.path.join(OUTPUT_DIR, "screenshots", "screenshot-3-result.png"))
    print("  screenshot-3-result.png (Word output)")


# ─── Promotional images ────────────────────────────────────────────────────────
def make_promo(size, edge=False):
    W, H = size
    img = Image.new("RGB", (W,H), DARK)
    d = D(ImageDraw.Draw(img), W, H)

    # Gradient bg
    for i in range(H):
        r = int(15 + (i/H)*10)
        g = int(23 + (i/H)*15)
        b = int(42 + (i/H)*12)
        d.line((0,i), (W,i), (r,g,b), 1)

    # Large icon
    icon_r = min(W,H) // 5
    cx = W // 3 if W > 500 else W // 2
    cy = H // 2
    d.circle((cx, cy), icon_r, BLUE)
    d.circle((cx, cy), icon_r - 4, (30,85,210))
    d.text((cx, cy), "W", size=icon_r*3//4, bold=True, color=WHITE, anchor="mm", font=f_bold(icon_r*3//4))

    # Text
    tx = cx + icon_r + 40 if W > 500 else W // 2
    ty = cy - 50 if W > 500 else cy + icon_r + 20

    d.text((tx, ty), "MarkdownPasteAddin", size=36 if W>500 else 24, bold=True, color=WHITE, cjk=True)
    ty += 44 if W>500 else 30
    d.text((tx, ty), "将网页内容一键导出为专业 Word 文档", size=18 if W>500 else 13, color=(180,190,200), cjk=True)
    ty += 30 if W>500 else 22
    features = "表格 · 流程图 · 数学公式 · 代码高亮 · 框架图 · 报告标准"
    d.text((tx, ty), features, size=13 if W>500 else 10, color=(130,140,155), cjk=True)

    # Bottom accent bar
    d.rect((0, H-4, W, H), BLUE)

    # Edge-specific: "Microsoft Edge" badge if needed
    if edge and W <= 300:
        pass  # small size, keep clean

    return img


# ─── Main ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("Generating store assets v2...\n")

    print("[Screenshots]")
    make_screenshot1()
    make_screenshot2()
    make_screenshot3()

    print("\n[Chrome Web Store]")
    for size in [(440,280), (920,680), (1400,560)]:
        img = make_promo(size)
        path = os.path.join(OUTPUT_DIR, "chrome", f"promo-{size[0]}x{size[1]}.png")
        img.save(path)
        print(f"  promo-{size[0]}x{size[1]}.png ({os.path.getsize(path):,} bytes)")

    print("\n[Edge Add-ons]")
    for size in [(300,300), (440,280), (1400,560)]:
        img = make_promo(size, edge=True)
        path = os.path.join(OUTPUT_DIR, "edge", f"promo-{size[0]}x{size[1]}.png")
        img.save(path)
        print(f"  promo-{size[0]}x{size[1]}.png ({os.path.getsize(path):,} bytes)")

    print(f"\nDone! All assets in: {OUTPUT_DIR}")
