#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path
import sys
import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from backup.csv_loader import detect_encoding


def read_csv_with_fallback(csv_path: Path) -> pd.DataFrame:
    try:
        encoding = detect_encoding(csv_path)
        return pd.read_csv(csv_path, encoding=encoding)
    except Exception as exc:
        raise RuntimeError(f"無法讀取 {csv_path}: {exc}") from exc


def confirm_overwrite(root: tk.Tk, xlsx_path: Path) -> bool:
    return messagebox.askyesno(
        "檔案已存在",
        f"已存在同名檔案：\n{xlsx_path}\n\n是否覆蓋？",
        parent=root,
    )


def convert_folder(folder: Path, root: tk.Tk) -> tuple[list[Path], list[Path]]:
    if not folder.exists() or not folder.is_dir():
        raise NotADirectoryError(f"資料夾不存在: {folder}")

    pattern = "**/*.csv"
    csv_files = sorted(folder.glob(pattern))
    converted_files: list[Path] = []
    skipped_files: list[Path] = []

    for csv_path in csv_files:
        xlsx_path = csv_path.with_suffix(".xlsx")
        if xlsx_path.exists() and not confirm_overwrite(root, xlsx_path):
            skipped_files.append(xlsx_path)
            continue

        dataframe = read_csv_with_fallback(csv_path)
        dataframe.to_excel(xlsx_path, index=False, engine="openpyxl")
        converted_files.append(xlsx_path)

    return converted_files, skipped_files


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="將資料夾內的 CSV 批次轉成 XLSX。")
    parser.add_argument("folder", nargs="?", help="CSV 所在資料夾路徑")
    return parser


def create_root() -> tk.Tk:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    return root


def choose_folder_with_dialog(root: tk.Tk) -> Path | None:
    selected = filedialog.askdirectory(title="選擇要轉換 CSV 的資料夾")
    return Path(selected).resolve() if selected else None


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    root = create_root()
    try:
        folder: Path | None = None
        if args.folder:
            folder = Path(args.folder).expanduser().resolve()
        else:
            folder = choose_folder_with_dialog(root)
            if folder is None:
                print("未選擇資料夾，已取消。")
                return 1

        converted_files, skipped_files = convert_folder(folder, root)

        if not converted_files and not skipped_files:
            messagebox.showinfo("完成", "這個資料夾和子資料夾內沒有找到 CSV 檔案。", parent=root)
            print("這個資料夾和子資料夾內沒有找到 CSV 檔案。")
            return 0

        summary = [f"已轉換 {len(converted_files)} 個檔案。"]
        if skipped_files:
            summary.append(f"略過 {len(skipped_files)} 個已存在的 XLSX。")
        messagebox.showinfo("完成", "\n".join(summary), parent=root)

        print(f"已轉換 {len(converted_files)} 個檔案:")
        for path in converted_files:
            print(path)
        if skipped_files:
            print(f"略過 {len(skipped_files)} 個檔案:")
            for path in skipped_files:
                print(path)
        return 0
    finally:
        root.destroy()


if __name__ == "__main__":
    raise SystemExit(main())
