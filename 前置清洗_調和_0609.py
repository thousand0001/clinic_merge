# -*- coding: utf-8 -*-
"""
調和前置清洗 + 共用核心包裝（0609）

目前調和資料已可直接交給共用核心，
這支前置器先作為模組化入口，後續若有格式差異再往這裡收。
"""

from __future__ import annotations

import csv
import importlib.util
import os
import re
import shutil
import sys
import tempfile
from pathlib import Path
from typing import Optional


SCRIPT_DIR = Path(__file__).resolve().parent


def _find_generic_script(script_dir: Path) -> Path:
    def date_key(path: Path) -> tuple[int, str]:
        match = re.search(r"(\d{4})(?=\.py$)", path.name)
        return (int(match.group(1)) if match else -1, path.name)

    candidates = sorted(
        script_dir.glob("選會員_共用核心_*.py"),
        key=date_key,
        reverse=True,
    )
    if not candidates:
        raise RuntimeError("找不到共用核心 選會員_共用核心_*.py")
    return candidates[0]

GENERIC_SCRIPT = _find_generic_script(SCRIPT_DIR)
ILLEGAL_XLSX_CHARS_RE = re.compile(r"[\x00-\x08\x0B-\x0C\x0E-\x1F]")


def _load_generic_module():
    spec = importlib.util.spec_from_file_location("member_merge_core", GENERIC_SCRIPT)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"無法載入共用核心：{GENERIC_SCRIPT}")
    mod = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = mod
    spec.loader.exec_module(mod)
    return mod


def _sanitize_csv_files(source_dir: Path) -> int:
    sanitized = 0
    csv_paths = sorted({
        path
        for path in source_dir.rglob("*")
        if path.is_file() and path.suffix.lower() == ".csv"
    })
    for path in csv_paths:
        rows = None
        for encoding in ("utf-8-sig", "utf-16", "cp950"):
            try:
                with path.open("r", encoding=encoding, errors="strict", newline="") as file:
                    rows = list(csv.reader(file))
                break
            except (UnicodeDecodeError, UnicodeError):
                continue
        if rows is None:
            with path.open("r", encoding="cp950", errors="replace", newline="") as file:
                rows = list(csv.reader(file))
        cleaned = [
            [ILLEGAL_XLSX_CHARS_RE.sub("", value) for value in row]
            for row in rows
        ]
        with path.open("w", encoding="utf-8-sig", newline="") as file:
            csv.writer(file).writerows(cleaned)
        sanitized += 1
    return sanitized


def process_excel(source_path: str, template_path: Optional[str] = None) -> str:
    generic = _load_generic_module()
    source_dir = Path(source_path).resolve()
    if not source_dir.is_dir():
        raise ValueError("請選擇資料夾。")

    template = template_path or generic._find_template(str(SCRIPT_DIR))
    if not os.path.exists(template):
        raise ValueError(f"找不到模板檔：{template}")

    temp_root = Path(tempfile.mkdtemp(prefix="tiaohe_clean_"))
    temp_source = temp_root / source_dir.name
    try:
        shutil.copytree(source_dir, temp_source)
        sanitized = _sanitize_csv_files(temp_source)
        print(f"已清理 CSV {sanitized} 個", flush=True)
        temp_output = Path(generic.process_excel(str(temp_source), template))
        final_output = source_dir.parent / temp_output.name
        if final_output.exists():
            final_output.unlink()
        shutil.move(str(temp_output), str(final_output))
        return str(final_output)
    finally:
        shutil.rmtree(temp_root, ignore_errors=True)


def main() -> None:
    generic = _load_generic_module()
    import tkinter as tk
    from tkinter import filedialog, messagebox

    root = tk.Tk()
    root.withdraw()

    src = filedialog.askdirectory(title="選擇調和來源資料夾")
    if not src:
        return

    template = generic._find_template(str(SCRIPT_DIR))

    try:
        out = process_excel(src, template)
        messagebox.showinfo("完成", f"已輸出：\n{out}")
        generic.open_file_cross_platform(out)
    except Exception as e:
        messagebox.showerror("錯誤", str(e))


if __name__ == "__main__":
    main()
