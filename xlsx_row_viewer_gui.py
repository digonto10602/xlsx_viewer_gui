#!/usr/bin/env python3
"""
Xlsx Row Viewer (Tkinter GUI)

Requested:
- Font: Palatino (best-effort) + 1.5x larger than Tk defaults
- Window: 1920x1080
- Values: larger, mouse-selectable, right-click -> Copy, then paste anywhere
- Sheet selector
- File picker if no file passed
- URL values: show an "Open URL" button

Notes on Palatino:
- Windows commonly has "Palatino Linotype"
- macOS commonly has "Palatino"
- Linux may have "URW Palladio L" or "TeX Gyre Pagella" (similar)
"""

import argparse
import platform
import sys
import webbrowser
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, font as tkfont


def enable_hidpi_awareness():
    if platform.system().lower() == "windows":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(2)  # per-monitor aware
        except Exception:
            try:
                import ctypes
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass


def pick_file() -> str | None:
    return filedialog.askopenfilename(
        title="Select an Excel file",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    ) or None


def list_sheets(path: str) -> list[str]:
    xls = pd.ExcelFile(path, engine="openpyxl")
    return list(xls.sheet_names)


def load_sheet(path: str, sheet_name: str) -> pd.DataFrame:
    return pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")


def choose_palatino_family() -> str:
    candidates = [
        "Palatino Linotype",   # Windows
        "Palatino",            # macOS
        "URW Palladio L",      # Linux common
        "TeX Gyre Pagella",    # Palatino-like
        "Book Antiqua",        # Windows alternative
        "Times New Roman",
        "DejaVu Serif",
        "Liberation Serif",
        "Serif",
    ]
    available = set(tkfont.families())
    for c in candidates:
        if c in available:
            return c
    return "Serif"


class ScrollFrame(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.canvas = tk.Canvas(self, highlightthickness=0)
        self.vbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vbar.set)

        self.vbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.inner = ttk.Frame(self.canvas)
        self.window_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")

        self.inner.bind("<Configure>", self._on_inner_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        # Mouse wheel scrolling
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)  # Windows
        self.canvas.bind_all("<Button-4>", self._on_mousewheel)    # Linux up
        self.canvas.bind_all("<Button-5>", self._on_mousewheel)    # Linux down

    def _on_inner_configure(self, _event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.window_id, width=event.width)

    def _on_mousewheel(self, event):
        if hasattr(event, "delta") and event.delta:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        else:
            if getattr(event, "num", None) == 4:
                self.canvas.yview_scroll(-3, "units")
            elif getattr(event, "num", None) == 5:
                self.canvas.yview_scroll(3, "units")


class App(tk.Tk):
    def __init__(self, path: str):
        enable_hidpi_awareness()
        super().__init__()

        self.path = path
        self.sheets = list_sheets(path)
        if not self.sheets:
            raise RuntimeError("No sheets found in this file.")

        self.df = pd.DataFrame()
        self.row = 0
        self.focused_url: str | None = None

        self.title("Xlsx Row Viewer")
        self.geometry("1920x1080")
        self.minsize(1100, 700)

        # Theme
        self.style = ttk.Style(self)
        for t in ("vista", "xpnative", "clam"):
            if t in self.style.theme_names():
                self.style.theme_use(t)
                break

        # Fonts
        self.family = choose_palatino_family()
        self._apply_fonts(scale=1.5)

        # Right-click context menu for values
        self.value_menu = tk.Menu(self, tearoff=0)
        self.value_menu.add_command(label="Copy", command=self._copy_from_active_value)
        self._active_value_widget: tk.Text | None = None

        # UI
        self._build_topbar()
        self._build_header()
        self._build_scroll_area()
        self._build_controls()
        self._bind_keys()

        self.load_current_sheet()

    def _apply_fonts(self, scale: float):
        base = tkfont.nametofont("TkDefaultFont")
        base_size = base.cget("size")
        if base_size == 0:
            base_size = 12

        def scaled(sz):
            sgn = -1 if sz < 0 else 1
            return int(abs(sz) * scale) * sgn

        # Update named fonts to Palatino-ish family and scaled size
        for name in (
            "TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont",
            "TkCaptionFont", "TkSmallCaptionFont", "TkIconFont", "TkTooltipFont"
        ):
            try:
                f = tkfont.nametofont(name)
                f.configure(family=self.family, size=scaled(f.cget("size")))
            except Exception:
                pass

        # Custom fonts
        self.font_bold = tkfont.Font(family=self.family, size=scaled(base_size), weight="bold")
        self.font_header = tkfont.Font(family=self.family, size=scaled(base_size + 5), weight="bold")

        # Value font (explicitly larger for the printed values)
        self.font_value = tkfont.Font(family=self.family, size=scaled(base_size + 2), weight="normal")

    def _build_topbar(self):
        top = ttk.Frame(self, padding=(16, 14))
        top.pack(fill="x")

        ttk.Label(top, text="File:", font=self.font_bold).pack(side="left")
        self.file_lbl = ttk.Label(top, text=self.path)
        self.file_lbl.pack(side="left", padx=(10, 14), fill="x", expand=True)

        ttk.Button(top, text="Change file…", command=self.change_file).pack(side="left", padx=(0, 14))

        ttk.Label(top, text="Sheet:", font=self.font_bold).pack(side="left")
        self.sheet_var = tk.StringVar(value=self.sheets[0])
        self.sheet_combo = ttk.Combobox(
            top, textvariable=self.sheet_var, values=self.sheets, state="readonly", width=34
        )
        self.sheet_combo.pack(side="left", padx=(10, 0))
        self.sheet_combo.bind("<<ComboboxSelected>>", lambda e: self.on_sheet_change())

    def _build_header(self):
        self.header = ttk.Label(self, text="", font=self.font_header)
        self.header.pack(fill="x", padx=16, pady=(0, 10), anchor="w")

    def _build_scroll_area(self):
        self.scroll = ScrollFrame(self)
        self.scroll.pack(fill="both", expand=True, padx=16, pady=12)

    def _build_controls(self):
        controls = ttk.Frame(self, padding=(16, 0))
        controls.pack(fill="x")

        ttk.Button(controls, text="◀ Prev", command=self.prev_row).pack(side="left")
        ttk.Button(controls, text="Next ▶", command=self.next_row).pack(side="left", padx=(12, 0))
        ttk.Button(controls, text="Open all URLs in row", command=self.open_all_urls).pack(side="right")

        hint = ttk.Label(
            self,
            text="Keys: Left/Right changes row. Right-click any value to Copy. Ctrl+C also works after selecting text.",
        )
        hint.pack(fill="x", padx=16, pady=(10, 16), anchor="w")

    def _bind_keys(self):
        self.bind("<Left>", lambda e: self.prev_row())
        self.bind("<Right>", lambda e: self.next_row())
        self.bind("<Return>", lambda e: self.open_focused_url())

    def set_focus_url(self, url: str):
        self.focused_url = url

    def open_focused_url(self):
        if self.focused_url:
            webbrowser.open(self.focused_url, new=2)

    def change_file(self):
        root = tk.Tk()
        root.withdraw()
        new_path = pick_file()
        root.destroy()
        if not new_path:
            return

        self.path = new_path
        self.file_lbl.config(text=self.path)

        try:
            self.sheets = list_sheets(self.path)
        except Exception as e:
            messagebox.showerror("Error", f"Could not read sheets:\n{e}")
            return

        if not self.sheets:
            messagebox.showinfo("No sheets", "No sheets found in that file.")
            return

        self.sheet_combo["values"] = self.sheets
        self.sheet_var.set(self.sheets[0])
        self.load_current_sheet()

    def on_sheet_change(self):
        self.load_current_sheet()

    def load_current_sheet(self):
        sheet = self.sheet_var.get()
        try:
            self.df = load_sheet(self.path, sheet).reset_index(drop=True)
        except Exception as e:
            messagebox.showerror("Error", f"Could not load sheet '{sheet}':\n{e}")
            self.df = pd.DataFrame()

        self.row = 0
        self.focused_url = None
        self.render_row()

    def prev_row(self):
        if len(self.df) == 0:
            return
        if self.row > 0:
            self.row -= 1
            self.render_row()

    def next_row(self):
        if len(self.df) == 0:
            return
        if self.row < len(self.df) - 1:
            self.row += 1
            self.render_row()

    def open_all_urls(self):
        if len(self.df) == 0:
            return
        urls = []
        r = self.df.iloc[self.row]
        for col in self.df.columns:
            v = r[col]
            if pd.isna(v):
                continue
            s = str(v).strip()
            if s.startswith("www."):
                s = "https://" + s
            if s.lower().startswith("http://") or s.lower().startswith("https://"):
                urls.append(s)
        if not urls:
            messagebox.showinfo("No URLs", "No URLs found in this row.")
            return
        for u in urls:
            webbrowser.open(u, new=2)

    # ---------- Copy / context menu ----------
    def _copy_from_active_value(self):
        w = self._active_value_widget
        if not w:
            return
        try:
            # Copy selected text if any; otherwise copy full content
            sel = w.selection_get()
            txt = sel
        except Exception:
            txt = w.get("1.0", "end-1c")
        self.clipboard_clear()
        self.clipboard_append(txt)

    def _show_value_menu(self, widget: tk.Text, event):
        self._active_value_widget = widget
        try:
            widget.focus_set()
        except Exception:
            pass
        self.value_menu.tk_popup(event.x_root, event.y_root)

    def _add_selectable_value(self, parent, text: str):
        """
        Read-only, selectable Text widget with:
        - bigger value font
        - right-click Copy menu
        - Ctrl+C support
        """
        txt = tk.Text(
            parent,
            wrap="word",
            borderwidth=1,
            relief="solid",
            highlightthickness=0,
            padx=10,
            pady=8,
            font=self.font_value,
        )

        content = text if (text is not None and text != "") else "(empty)"
        txt.insert("1.0", content)

        # Height heuristics
        lines = max(1, content.count("\n") + 1)
        if len(content) > 160:
            lines = max(lines, 3)
        txt.configure(height=lines)

        # Keep selectable but read-only: toggle enable during selection/copy
        txt.configure(state="disabled")

        def enable(_e=None):
            txt.configure(state="normal")

        def disable(_e=None):
            txt.configure(state="disabled")

        txt.bind("<Button-1>", enable)
        txt.bind("<B1-Motion>", enable)
        txt.bind("<ButtonRelease-1>", disable)

        def on_copy(_e=None):
            # copy selection if exists else full
            self._active_value_widget = txt
            self._copy_from_active_value()
            return "break"

        txt.bind("<Control-c>", on_copy)
        txt.bind("<Control-C>", on_copy)
        txt.bind("<Command-c>", on_copy)
        txt.bind("<Command-C>", on_copy)

        # Right-click menu
        txt.bind("<Button-3>", lambda e, w=txt: self._show_value_menu(w, e))  # Windows/Linux
        txt.bind("<Button-2>", lambda e, w=txt: self._show_value_menu(w, e))  # macOS

        txt.pack(anchor="w", fill="x", expand=True)
        return txt

    def render_row(self):
        for w in self.scroll.inner.winfo_children():
            w.destroy()

        sheet = self.sheet_var.get()
        if len(self.df) == 0:
            self.header.config(text=f"Sheet: {sheet} | Empty or failed to load.")
            return

        self.header.config(text=f"Sheet: {sheet} | Row {self.row + 1}/{len(self.df)}")
        self.focused_url = None

        r = self.df.iloc[self.row]

        for col in self.df.columns:
            ttk.Label(self.scroll.inner, text=f"{col}:", font=self.font_bold).pack(anchor="w", pady=(14, 6))

            v = r[col]
            if pd.isna(v):
                self._add_selectable_value(self.scroll.inner, "(empty)")
                continue

            s = str(v).strip()
            url = None
            if s.startswith("www."):
                url = "https://" + s
            elif s.lower().startswith("http://") or s.lower().startswith("https://"):
                url = s

            if url:
                block = ttk.Frame(self.scroll.inner)
                block.pack(fill="x", anchor="w")

                self._add_selectable_value(block, url)

                btn = ttk.Button(block, text="Open URL", command=lambda u=url: webbrowser.open(u, new=2))
                btn.pack(anchor="w", pady=(10, 0))

                btn.bind("<Enter>", lambda e, u=url: self.set_focus_url(u))
                btn.bind("<FocusIn>", lambda e, u=url: self.set_focus_url(u))
            else:
                self._add_selectable_value(self.scroll.inner, s)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx", nargs="?", default=None, help="Path to .xlsx (optional; opens picker if omitted)")
    args = ap.parse_args()

    path = args.xlsx
    if not path:
        root = tk.Tk()
        root.withdraw()
        path = pick_file()
        root.destroy()
        if not path:
            return

    try:
        app = App(path)
        app.mainloop()
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        raise


if __name__ == "__main__":
    main()