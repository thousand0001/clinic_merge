#!/usr/bin/env python3
from __future__ import annotations

import tkinter as tk
import shutil
from pathlib import Path
from tkinter import filedialog, messagebox

SOURCE_ENCODING = "cp950"
SOURCE_ERRORS = "replace"


def iter_csv_files(folder: Path) -> list[Path]:
    return sorted(
        p for p in folder.rglob("*")
        if p.is_file() and p.suffix.lower() == ".csv"
    )


def backup_file(path: Path) -> Path:
    backup_path = path.with_suffix(path.suffix + ".bak")
    counter = 1
    while backup_path.exists():
        backup_path = path.with_suffix(path.suffix + f".bak{counter}")
        counter += 1
    shutil.copy2(path, backup_path)
    return backup_path


def convert_csv_file(
    csv_path: Path,
    target_encoding: str,
    backup: bool = True,
) -> tuple[str, Path | None]:
    backup_path = backup_file(csv_path) if backup else None

    text = csv_path.read_bytes().decode(SOURCE_ENCODING, errors=SOURCE_ERRORS)

    csv_path.write_text(text, encoding=target_encoding, newline="")

    return SOURCE_ENCODING, backup_path


def create_root() -> tk.Tk:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    return root


def choose_folder_with_dialog(root: tk.Tk) -> Path | None:
    selected = filedialog.askdirectory(title="選擇要轉換 CSV 編碼的資料夾")
    return Path(selected).resolve() if selected else None


def confirm_convert(root: tk.Tk, folder: Path, count: int, target_encoding: str) -> bool:
    return messagebox.askyesno(
        "確認轉換",
        f"資料夾：\n{folder}\n\n找到 {count} 個 CSV 檔案。\n"
        f"來源編碼固定使用：{SOURCE_ENCODING}\n"
        f"將轉成編碼：{target_encoding}\n"
        "並先建立原始檔備份（.bak）。\n\n是否開始？",
        parent=root,
    )


def main() -> int:
    root = create_root()
    try:
        folder = choose_folder_with_dialog(root)
        if folder is None:
            print("未選擇資料夾，已取消。")
            return 1
        if not folder.exists() or not folder.is_dir():
            raise NotADirectoryError(f"資料夾不存在: {folder}")

        target_encoding = "utf-8-sig"
        csv_files = iter_csv_files(folder)
        if not csv_files:
            messagebox.showinfo("完成", "這個資料夾和子資料夾內沒有找到 CSV 檔案。", parent=root)
            print("這個資料夾和子資料夾內沒有找到 CSV 檔案。")
            return 0

        if not confirm_convert(root, folder, len(csv_files), target_encoding):
            print("已取消。")
            return 1

        converted_files: list[Path] = []
        for csv_path in csv_files:
            detected, backup_path = convert_csv_file(
                csv_path,
                target_encoding=target_encoding,
                backup=True,
            )
            converted_files.append(csv_path)
            print(f"[OK] {csv_path} | {detected} -> {target_encoding} | backup={backup_path}")

        messagebox.showinfo(
            "完成",
            f"已轉換 {len(converted_files)} 個 CSV 檔案為 {target_encoding}，並建立 .bak 備份。",
            parent=root,
        )
        return 0
    finally:
        root.destroy()


if __name__ == "__main__":
    raise SystemExit(main())
