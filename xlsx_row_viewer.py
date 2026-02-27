#!/usr/bin/env python3
"""
Xlsx Row Viewer (Windows-friendly)

Features:
- Interactive row browsing (LEFT/RIGHT)
- TAB cycles URL fields, ENTER opens selected URL in default browser
- If no XLSX path is provided, opens a file picker
"""

import argparse
import re
import sys
import webbrowser
from typing import List, Tuple, Optional

import pandas as pd

# curses is stdlib on macOS/Linux; on Windows you need `windows-curses` at build time and runtime (but runtime is bundled into the exe).
try:
    import curses
except Exception as e:
    print(
        "ERROR: curses is not available.\n"
        "On Windows (when running as .py), install:  pip install windows-curses\n"
        f"Details: {e}",
        file=sys.stderr,
    )
    raise


URL_RE = re.compile(r"(?i)\bhttps?://[^\s<>\"]+|www\.[^\s<>\"]+")


def is_nan(x) -> bool:
    try:
        return pd.isna(x)
    except Exception:
        return False


def to_str(x) -> str:
    if is_nan(x):
        return ""
    try:
        if isinstance(x, float) and x.is_integer():
            return str(int(x))
        return str(x)
    except Exception:
        return repr(x)


def normalize_url(s: str) -> Optional[str]:
    s = (s or "").strip()
    if not s:
        return None
    m = URL_RE.search(s)
    if not m:
        return None
    url = m.group(0)
    if url.lower().startswith("www."):
        url = "https://" + url
    return url


def pick_file_dialog() -> Optional[str]:
    # Use Tk file picker only if user didn't pass a file
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception:
        return None

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    path = filedialog.askopenfilename(
        title="Select an Excel file",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
    )
    root.destroy()
    return path or None


def load_xlsx(path: str, sheet: str) -> pd.DataFrame:
    sheet_name = int(sheet) if str(sheet).isdigit() else sheet
    return pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")


def row_lines(df: pd.DataFrame, row_idx: int) -> Tuple[List[Tuple[str, str, Optional[str]]], List[int]]:
    row = df.iloc[row_idx]
    lines = []
    url_line_indices = []
    for col in df.columns:
        val_s = to_str(row[col])
        url = normalize_url(val_s)
        shown = val_s if val_s != "" else "(empty)"
        lines.append((str(col), shown, url))
        if url:
            url_line_indices.append(len(lines) - 1)
    return lines, url_line_indices


def draw(stdscr, title: str, df: pd.DataFrame, row_idx: int, url_focus_i: int) -> int:
    stdscr.erase()
    h, w = stdscr.getmaxyx()

    header = f"{title} | Row {row_idx+1}/{len(df)}  (LEFT/RIGHT rows, TAB URL, ENTER open, q quit)"
    stdscr.addnstr(0, 0, header, w - 1, curses.A_BOLD)

    if len(df) == 0:
        stdscr.addnstr(2, 0, "DataFrame is empty.", w - 1)
        stdscr.refresh()
        return 0

    lines, url_line_indices = row_lines(df, row_idx)

    focused_line = -1
    if url_line_indices:
        url_focus_i = max(0, min(url_focus_i, len(url_line_indices) - 1))
        focused_line = url_line_indices[url_focus_i]

    y = 2
    for i, (col, shown, url) in enumerate(lines):
        if y >= h - 2:
            stdscr.addnstr(y, 0, "... (more columns not shown) ...", w - 1)
            y += 1
            break

        prefix = f"{col}: "
        suffix = "  [URL]" if url else ""
        line = prefix + shown + suffix

        attr = curses.A_REVERSE if (i == focused_line and url is not None) else curses.A_NORMAL
        stdscr.addnstr(y, 0, line, w - 1, attr)
        y += 1

    footer = (
        "No URL in this row."
        if not url_line_indices
        else f"Selected URL field: {lines[url_line_indices[url_focus_i]][0]}  |  ENTER opens in your browser"
    )
    stdscr.addnstr(h - 1, 0, footer, w - 1, curses.A_DIM)

    stdscr.refresh()
    return url_focus_i


def interactive(stdscr, df: pd.DataFrame, title: str):
    curses.curs_set(0)
    stdscr.keypad(True)

    if len(df) == 0:
        draw(stdscr, title, df, 0, 0)
        stdscr.getch()
        return

    row_idx = 0
    url_focus_i = 0

    while True:
        url_focus_i = draw(stdscr, title, df, row_idx, url_focus_i)
        ch = stdscr.getch()

        if ch in (ord("q"), ord("Q")):
            break

        if ch == curses.KEY_RIGHT:
            if row_idx < len(df) - 1:
                row_idx += 1
                url_focus_i = 0

        elif ch == curses.KEY_LEFT:
            if row_idx > 0:
                row_idx -= 1
                url_focus_i = 0

        elif ch in (curses.KEY_TAB, 9):  # TAB
            _, url_line_indices = row_lines(df, row_idx)
            if url_line_indices:
                url_focus_i = (url_focus_i + 1) % len(url_line_indices)

        elif ch in (curses.KEY_ENTER, 10, 13):  # ENTER
            lines, url_line_indices = row_lines(df, row_idx)
            if url_line_indices:
                focused_line = url_line_indices[url_focus_i]
                _, _, url = lines[focused_line]
                if url:
                    webbrowser.open(url, new=2)


def main():
    ap = argparse.ArgumentParser(description="Xlsx Row Viewer (interactive).")
    ap.add_argument("xlsx", nargs="?", default=None, help="Path to .xlsx file (optional; opens picker if omitted)")
    ap.add_argument("--sheet", default="0", help="Sheet name or 0-based sheet index (default: 0)")
    ap.add_argument("--dropna-rows", action="store_true", help="Drop rows that are entirely empty")
    args = ap.parse_args()

    path = args.xlsx
    if not path:
        path = pick_file_dialog()
        if not path:
            print("No file selected. Exiting.")
            return

    df = load_xlsx(path, args.sheet)
    if args.dropna_rows:
        df = df.dropna(how="all").reset_index(drop=True)

    title = f"{path} | sheet={args.sheet}"
    curses.wrapper(lambda stdscr: interactive(stdscr, df, title))


if __name__ == "__main__":
    main()