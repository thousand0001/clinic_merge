# -*- coding: utf-8 -*-
"""
讀取健保署上傳文字檔（CP950 編碼）

用法：
    python 讀取健保署上傳檔_cp950.py
    python 讀取健保署上傳檔_cp950.py /path/to/FM.txt

預設會開啟簡單文字視窗。若只想在終端機預覽：
    python 讀取健保署上傳檔_cp950.py /path/to/FM.txt --preview
"""

from __future__ import annotations

import argparse
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext
from typing import Optional


ENCODING = "cp950"
DEFAULT_DIR = Path.home() / "Desktop" / "健保署115年上傳名單"


def read_upload_text(path: Path) -> str:
    with path.open("r", encoding=ENCODING) as f:
        return f.read()


def pick_file() -> Optional[Path]:
    root = tk.Tk()
    root.withdraw()
    try:
        initial_dir = str(DEFAULT_DIR) if DEFAULT_DIR.exists() else None
        filename = filedialog.askopenfilename(
            title="選擇健保署上傳文字檔",
            initialdir=initial_dir,
            filetypes=[
                ("文字檔", "*.txt"),
                ("所有檔案", "*.*"),
            ],
        )
    finally:
        root.destroy()
    return Path(filename) if filename else None


def show_text_window(path: Path, text: str) -> None:
    root = tk.Tk()
    root.title(f"健保署上傳檔閱讀器 - {path.name}")
    root.geometry("1100x700")

    lines = text.splitlines()
    info = tk.Label(
        root,
        text=f"檔案：{path}\n編碼：{ENCODING}　字元數：{len(text)}　行數：{len(lines)}",
        anchor="w",
        justify="left",
    )
    info.pack(fill="x", padx=8, pady=6)

    box = scrolledtext.ScrolledText(root, wrap="none", font=("Menlo", 12))
    box.pack(fill="both", expand=True, padx=8, pady=(0, 8))
    box.insert("1.0", text)
    box.configure(state="disabled")

    root.mainloop()


def print_preview(path: Path, text: str, lines: int) -> None:
    all_lines = text.splitlines()
    print(f"讀取成功：{path}")
    print(f"編碼：{ENCODING}")
    print(f"字元數：{len(text)}")
    print(f"行數：{len(all_lines)}")
    print(f"前 {min(lines, len(all_lines))} 行：")
    for line in all_lines[:lines]:
        print(line)


def main(argv: Optional[list[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="讀取 CP950 編碼的健保署上傳文字檔")
    parser.add_argument("file", nargs="?", help="健保署上傳文字檔路徑；未指定時開啟選檔視窗")
    parser.add_argument("--preview", action="store_true", help="只在終端機顯示前幾行，不開文字視窗")
    parser.add_argument("--lines", type=int, default=10, help="--preview 顯示行數，預設 10")
    args = parser.parse_args(argv)

    path = Path(args.file).expanduser() if args.file else pick_file()
    if path is None:
        print("未選擇檔案，程式已取消。")
        return 0
    if not path.exists() or not path.is_file():
        print(f"找不到檔案：{path}", file=sys.stderr)
        return 2

    try:
        text = read_upload_text(path)
    except UnicodeDecodeError as exc:
        msg = f"讀取失敗：檔案不是有效的 {ENCODING} 編碼。\n{exc}"
        print(msg, file=sys.stderr)
        messagebox.showerror("讀取失敗", msg)
        return 3
    except OSError as exc:
        msg = f"讀取失敗：{exc}"
        print(msg, file=sys.stderr)
        messagebox.showerror("讀取失敗", msg)
        return 4

    if args.preview:
        print_preview(path, text, max(args.lines, 0))
    else:
        show_text_window(path, text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
