#!/usr/bin/env python3
"""
MarkdownPasteAddin GUI v3.0 — Modern desktop interface

Features:
  - Clean two-panel layout (edit | preview)
  - Format preset selector (政府公文 / 学术论文 / 商务报告 / 自定义)
  - One-click clipboard paste + convert
  - Live preview of parsed content
  - Drag-and-drop .md file support
  - Status bar with progress
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from md2docx_lib import parse_markdown
from md2docx_lib.builder_docx import DocumentBuilder
from md2docx_lib.clipboard import read_clipboard
from md2docx_lib.parser_html import parse_html
from md2docx_lib.inline_processor import clean_text
from md2docx_lib.report_standards import PRESETS, apply_report_standard


# ─── Theme colors ───────────────────────────────────────────────────────────────

COLORS = {
    "bg":           "#f0f2f5",
    "surface":      "#ffffff",
    "primary":      "#2563eb",
    "primary_hover":"#1d4ed8",
    "text":         "#1e293b",
    "text_secondary":"#64748b",
    "border":       "#e2e8f0",
    "success":      "#22c55e",
    "warning":      "#f59e0b",
    "error":        "#ef4444",
    "code_bg":      "#1e293b",
    "code_text":    "#e2e8f0",
}

FONTS = {
    "heading": ("Segoe UI", 14, "bold"),
    "body":    ("Segoe UI", 11),
    "mono":    ("Consolas", 11),
    "small":   ("Segoe UI", 9),
    "title":   ("Segoe UI", 18, "bold"),
}


class ModernButton(ttk.Button):
    """Styled button with hover effect."""
    pass


class MarkdownPasteGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("MarkdownPasteAddin v3.0 — Smart Markdown to Word")
        self.root.geometry("1050x720")
        self.root.minsize(800, 500)
        self.root.configure(bg=COLORS["bg"])

        # ── Vars ──
        self.output_path = tk.StringVar(value=os.path.expanduser("~\\Desktop\\export.docx"))
        self.format_preset = tk.StringVar(value="business")
        self.auto_captions = tk.BooleanVar(value=True)
        self.add_toc = tk.BooleanVar(value=True)
        self.add_cover = tk.BooleanVar(value=True)
        self.image_width = tk.DoubleVar(value=5.5)
        self.status_text = tk.StringVar(value="Ready — paste content or open a Markdown file")

        # ── Style ttk ──
        self._setup_ttk_style()

        # ── Build UI ──
        self._build_ui()

        # ── Auto-paste ──
        self.root.after(300, self._auto_paste)

    # ═══════════════════════════════════════════════════════════════════════════
    # Style
    # ═══════════════════════════════════════════════════════════════════════════

    def _setup_ttk_style(self):
        style = ttk.Style()
        style.theme_use("clam")

        # Configure clam theme colors
        style.configure(".",
            background=COLORS["bg"],
            foreground=COLORS["text"],
            font=FONTS["body"],
        )

        style.configure("TFrame", background=COLORS["bg"])
        style.configure("TLabel", background=COLORS["bg"], foreground=COLORS["text"])
        style.configure("TLabelframe", background=COLORS["bg"])
        style.configure("TLabelframe.Label", background=COLORS["bg"],
                       foreground=COLORS["text_secondary"], font=FONTS["small"])

        # Primary button
        style.configure("Primary.TButton",
            background=COLORS["primary"],
            foreground="white",
            borderwidth=0,
            padding=(20, 10),
            font=FONTS["body"],
        )
        style.map("Primary.TButton",
            background=[("active", COLORS["primary_hover"])],
        )

        # Tool button
        style.configure("Tool.TButton",
            padding=(12, 6),
            font=FONTS["body"],
        )

        # Entry
        style.configure("TEntry", fieldbackground=COLORS["surface"],
                       borderwidth=1, relief="solid")

        # Checkbutton
        style.configure("TCheckbutton", background=COLORS["bg"])

        # Combobox
        style.configure("TCombobox", fieldbackground=COLORS["surface"])

    # ═══════════════════════════════════════════════════════════════════════════
    # UI Layout
    # ═══════════════════════════════════════════════════════════════════════════

    def _build_ui(self):
        # ── Top bar ──
        topbar = ttk.Frame(self.root, padding=(16, 10))
        topbar.pack(fill=tk.X)

        # Logo + title
        logo_frame = tk.Frame(topbar, bg=COLORS["bg"])
        logo_frame.pack(side=tk.LEFT)

        logo = tk.Label(logo_frame, text="W", font=("Segoe UI", 22, "bold"),
                       fg="white", bg=COLORS["primary"],
                       padx=10, pady=2)
        logo.pack(side=tk.LEFT, padx=(0, 10))

        title_label = tk.Label(logo_frame, text="MarkdownPasteAddin",
                              font=FONTS["title"], fg=COLORS["text"], bg=COLORS["bg"])
        title_label.pack(side=tk.LEFT)
        ver_label = tk.Label(logo_frame, text="v3.0", font=FONTS["small"],
                           fg=COLORS["text_secondary"], bg=COLORS["bg"])
        ver_label.pack(side=tk.LEFT, padx=(6, 0))

        # Action buttons
        actions = ttk.Frame(topbar)
        actions.pack(side=tk.RIGHT, pady=4)

        ttk.Button(actions, text="Paste", command=self.paste_clipboard,
                  style="Tool.TButton").pack(side=tk.LEFT, padx=3)
        ttk.Button(actions, text="Open File", command=self.open_file,
                  style="Tool.TButton").pack(side=tk.LEFT, padx=3)
        ttk.Button(actions, text="Clear", command=self.clear,
                  style="Tool.TButton").pack(side=tk.LEFT, padx=3)

        ttk.Separator(actions, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=8)

        self.quick_btn = ttk.Button(actions, text="Paste & Convert",
                                     command=self.quick_paste_convert, style="Primary.TButton")
        self.quick_btn.pack(side=tk.LEFT, padx=3)

        self.convert_btn = ttk.Button(actions, text="Convert to Word",
                                      command=self.convert, style="Tool.TButton")
        self.convert_btn.pack(side=tk.LEFT, padx=3)

        # ── Main content (two-panel) ──
        main = ttk.Frame(self.root)
        main.pack(fill=tk.BOTH, expand=True, padx=16, pady=(6, 0))

        # Left: Editor
        left_frame = ttk.LabelFrame(main, text="Markdown Content", padding=(8, 4))
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Editor toolbar
        editor_tb = ttk.Frame(left_frame)
        editor_tb.pack(fill=tk.X, pady=(0, 4))

        self.char_count = tk.StringVar(value="0 chars")
        ttk.Label(editor_tb, textvariable=self.char_count,
                 font=FONTS["small"], foreground=COLORS["text_secondary"]
                 ).pack(side=tk.LEFT)
        ttk.Button(editor_tb, text="Preview", command=self._refresh_preview,
                  style="Tool.TButton").pack(side=tk.RIGHT)

        # Editor text area
        self.editor_frame = tk.Frame(left_frame, bg=COLORS["surface"],
                                     highlightbackground=COLORS["border"],
                                     highlightthickness=1)
        self.editor_frame.pack(fill=tk.BOTH, expand=True)

        self.editor = tk.Text(self.editor_frame, wrap=tk.WORD,
                              font=FONTS["mono"],
                              bg=COLORS["surface"], fg=COLORS["text"],
                              insertbackground=COLORS["primary"],
                              selectbackground=COLORS["primary"],
                              selectforeground="white",
                              relief=tk.FLAT,
                              padx=12, pady=12,
                              undo=True, maxundo=100)
        self.editor.pack(fill=tk.BOTH, expand=True)
        self.editor.bind("<<Modified>>", self._on_editor_modified)

        # Editor scrollbar
        editor_scroll = ttk.Scrollbar(self.editor, orient=tk.VERTICAL,
                                      command=self.editor.yview)
        self.editor.configure(yscrollcommand=editor_scroll.set)
        editor_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Right: Panel
        right_frame = ttk.Frame(main, width=320)
        right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(12, 0))
        right_frame.pack_propagate(False)

        # Preview
        preview_frame = ttk.LabelFrame(right_frame, text="Parsed Preview", padding=(8, 4))
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        self.preview_text = tk.Text(preview_frame, wrap=tk.WORD,
                                    font=FONTS["mono"],
                                    bg="#f8fafc", fg=COLORS["text"],
                                    relief=tk.FLAT,
                                    padx=10, pady=10,
                                    state=tk.DISABLED,
                                    height=10)
        self.preview_text.pack(fill=tk.BOTH, expand=True)

        # Settings
        settings_frame = ttk.LabelFrame(right_frame, text="Export Settings", padding=(8, 4))
        settings_frame.pack(fill=tk.X, pady=(0, 4))

        # Format preset
        ttk.Label(settings_frame, text="Report Format:",
                 font=FONTS["small"]).pack(anchor=tk.W, pady=(4, 2))
        preset_names = [f"{k} - {PRESETS[k].name.split('(')[0].strip()}" for k in PRESETS]
        preset_keys = list(PRESETS.keys())
        self.preset_combo = ttk.Combobox(settings_frame, values=preset_names,
                                         state="readonly", textvariable=self.format_preset)
        self.preset_combo.pack(fill=tk.X)
        self.preset_combo.current(preset_keys.index("business"))
        self.preset_combo.bind("<<ComboboxSelected>>", lambda e: self._update_preset())

        # Options
        opts = ttk.Frame(settings_frame)
        opts.pack(fill=tk.X, pady=(8, 2))

        ttk.Checkbutton(opts, text="TOC", variable=self.add_toc).grid(row=0, column=0, sticky=tk.W)
        ttk.Checkbutton(opts, text="Cover Page", variable=self.add_cover).grid(row=0, column=1, sticky=tk.W)
        ttk.Checkbutton(opts, text="Numbering", variable=self.auto_captions).grid(row=1, column=0, sticky=tk.W)

        # Output path
        ttk.Label(settings_frame, text="Output:",
                 font=FONTS["small"]).pack(anchor=tk.W, pady=(8, 2))
        out_row = ttk.Frame(settings_frame)
        out_row.pack(fill=tk.X)
        ttk.Entry(out_row, textvariable=self.output_path,
                 font=FONTS["small"]).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(out_row, text="...", width=3,
                  command=self._browse_output).pack(side=tk.LEFT, padx=(4, 0))

        # Image width
        size_row = ttk.Frame(settings_frame)
        size_row.pack(fill=tk.X, pady=(8, 2))
        ttk.Label(size_row, text="Image width:",
                 font=FONTS["small"]).pack(side=tk.LEFT)
        ttk.Spinbox(size_row, from_=1.0, to=8.0, increment=0.5,
                    textvariable=self.image_width, width=5,
                    font=FONTS["small"]).pack(side=tk.RIGHT)

        # ── Status bar ──
        status_frame = tk.Frame(self.root, bg=COLORS["surface"], height=32)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM, padx=16, pady=(0, 10))

        self.status_label = tk.Label(status_frame, textvariable=self.status_text,
                                     font=FONTS["small"],
                                     fg=COLORS["text_secondary"],
                                     bg=COLORS["surface"],
                                     anchor=tk.W, padx=10)
        self.status_label.pack(fill=tk.X, side=tk.LEFT)

        # Status indicator dot
        self.status_dot = tk.Canvas(status_frame, width=16, height=16,
                                    bg=COLORS["surface"], highlightthickness=0)
        self.status_dot.pack(side=tk.RIGHT, padx=(0, 8))
        self._dot = self.status_dot.create_oval(2, 4, 14, 16, fill="#94a3b8", outline="")

    # ═══════════════════════════════════════════════════════════════════════════
    # Actions
    # ═══════════════════════════════════════════════════════════════════════════

    def _auto_paste(self):
        try:
            data = read_clipboard()
            if data and data.get("content") and len(data["content"]) > 30:
                self.set_content(data["content"])
                self._set_status(f"Auto-pasted from clipboard ({len(data['content'])} chars, {data['format']})", "success")
                self._refresh_preview()
        except Exception:
            self._set_status("Ready — paste content or open a Markdown file", "idle")

    def paste_clipboard(self):
        try:
            data = read_clipboard()
            if not data:
                messagebox.showinfo("Empty", "Clipboard has no text content.")
                return
            content = data["content"]
            self.set_content(content)
            self._set_status(f"Pasted ({len(content)} chars, format: {data['format']})", "success")
            self._refresh_preview()
        except Exception as e:
            self._set_status(f"Clipboard error: {e}", "error")

    def open_file(self):
        path = filedialog.askopenfilename(
            title="Open Markdown File",
            filetypes=[("Markdown", "*.md"), ("Text", "*.txt"), ("All", "*.*")]
        )
        if path:
            try:
                with open(path, "r", encoding="utf-8") as f:
                    content = f.read()
                self.set_content(content)
                base = os.path.splitext(path)[0]
                self.output_path.set(base + ".docx")
                self._set_status(f"Opened: {os.path.basename(path)}", "success")
                self._refresh_preview()
            except Exception as e:
                self._set_status(f"File error: {e}", "error")

    def clear(self):
        self.editor.delete("1.0", tk.END)
        self._update_preview_display("")
        self._set_status("Cleared", "idle")

    def set_content(self, text: str):
        self.editor.delete("1.0", tk.END)
        self.editor.insert("1.0", text)
        self.editor.edit_modified = False
        self._update_char_count()

    # ── One-click Paste & Convert (replicates 粘贴转Word.bat) ──

    def quick_paste_convert(self):
        """Paste clipboard content and convert to Word in one click."""
        try:
            data = read_clipboard()
            if not data or not data.get("content"):
                messagebox.showinfo("Clipboard Empty",
                    "No content in clipboard.\n\n"
                    "Copy content from DeepSeek/Web first (Ctrl+C),\n"
                    "then click 'Paste & Convert' again.")
                return

            content = data["content"]
            self.set_content(content)
            self._set_status(f"Pasted {len(content):,} chars. Converting...", "running")
            self._refresh_preview()
        except Exception as e:
            self._set_status(f"Clipboard read failed: {e}", "error")
            messagebox.showerror("Error", f"Cannot read clipboard:\n{e}")
            return

        # Convert immediately
        self.convert()

    # ── Conversion ──

    def convert(self):
        content = self.editor.get("1.0", tk.END).strip()
        if not content:
            messagebox.showwarning("No content", "Enter or paste Markdown first.")
            return

        output = self.output_path.get()
        if not output:
            messagebox.showwarning("No output", "Set an output path.")
            return

        self._set_status("Converting...", "running")
        self.convert_btn.configure(state="disabled")
        self.root.config(cursor="watch")
        self.root.update()

        def _run():
            try:
                if content.strip().startswith("<") and "</" in content:
                    chunks = parse_html(content)
                else:
                    cleaned = clean_text(content)
                    chunks = parse_markdown(cleaned)

                preset_key = self.format_preset.get().split(" - ")[0]

                builder = DocumentBuilder(
                    auto_captions=self.auto_captions.get(),
                    image_max_width=self.image_width.get(),
                    add_toc=self.add_toc.get(),
                    show_progress=False,
                )
                builder.build(chunks, output)

                # Apply report standard
                if self.add_cover.get() or self.add_toc.get():
                    from docx import Document
                    doc = Document(output)
                    apply_report_standard(doc, preset=preset_key,
                        title=os.path.splitext(os.path.basename(output))[0])
                    doc.save(output)

                self.root.after(0, lambda: self._on_done(output))
            except Exception as e:
                self.root.after(0, lambda: self._on_error(str(e)))

        threading.Thread(target=_run, daemon=True).start()

    def _on_done(self, path):
        self.root.config(cursor="")
        self.convert_btn.configure(state="normal")
        self._set_status(f"Saved: {os.path.basename(path)}", "success")

        if messagebox.askyesno("Done", f"Document saved:\n{path}\n\nOpen in Word?"):
            os.startfile(path)

    def _on_error(self, msg):
        self.root.config(cursor="")
        self.convert_btn.configure(state="normal")
        self._set_status(f"Error: {msg}", "error")
        messagebox.showerror("Conversion Failed", msg)

    # ── Preview ──

    def _refresh_preview(self):
        content = self.editor.get("1.0", tk.END).strip()
        if not content:
            self._update_preview_display("")
            return

        try:
            if content.strip().startswith("<") and "</" in content:
                chunks = parse_html(content)
            else:
                chunks = parse_markdown(content)
        except Exception:
            chunks = []

        stats = {}
        for c in chunks:
            t = c.get("type", "text")
            stats[t] = stats.get(t, 0) + 1

        lines = ["Parsed Content:\n"]
        type_names = {
            "heading": "Heading", "text": "Text", "table": "Table",
            "mermaid": "Mermaid", "code": "Code", "math": "Math Formula",
            "task_list": "Task List", "blockquote": "Quote", "hr": "Divider",
            "image": "Image",
        }
        for t, n in sorted(stats.items()):
            label = type_names.get(t, t)
            bar = "█" * min(n, 20)
            lines.append(f"  {label:<16} {n:>3}  {bar}")

        lines.append(f"\n  {'Total chunks':<16} {len(chunks):>3}")

        # First few items preview
        if chunks:
            lines.append(f"\nItems:")
            for i, c in enumerate(chunks[:8]):
                t = c.get("type", "?")
                preview = ""
                if t == "text":
                    preview = c.get("text", "")[:50]
                elif t == "heading":
                    preview = c.get("text", "")
                elif t == "table":
                    preview = f"{len(c.get('headers',[]))}×{len(c.get('rows',[]))}"
                elif t == "code":
                    preview = f"{c.get('language','')} ({len(c.get('code',''))} chars)"
                elif t == "mermaid":
                    preview = f"({len(c.get('code',''))} chars)"
                lines.append(f"    [{t}] {preview}")
            if len(chunks) > 8:
                lines.append(f"    ... and {len(chunks) - 8} more")

        self._update_preview_display("\n".join(lines))

    def _update_preview_display(self, text: str):
        self.preview_text.configure(state=tk.NORMAL)
        self.preview_text.delete("1.0", tk.END)
        if text:
            self.preview_text.insert("1.0", text)
        self.preview_text.configure(state=tk.DISABLED)

    # ── Helpers ──

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            title="Save Word Document",
            defaultextension=".docx",
            filetypes=[("Word Document", "*.docx")]
        )
        if path:
            self.output_path.set(path)

    def _on_editor_modified(self, event=None):
        if self.editor.edit_modified:
            self._update_char_count()
            self.editor.edit_modified = False

    def _update_char_count(self):
        count = len(self.editor.get("1.0", tk.END).strip())
        self.char_count.set(f"{count:,} chars")

    def _update_preset(self):
        key = self.format_preset.get().split(" - ")[0]
        self.format_preset.set(key)

    def _set_status(self, msg: str, level: str = "idle"):
        self.status_text.set(msg)
        colors = {"idle": "#94a3b8", "success": "#22c55e", "error": "#ef4444",
                  "running": "#2563eb", "warning": "#f59e0b"}
        self.status_dot.itemconfig(self._dot, fill=colors.get(level, "#94a3b8"))


def main():
    app = MarkdownPasteGUI()
    app.root.mainloop()


if __name__ == "__main__":
    main()
