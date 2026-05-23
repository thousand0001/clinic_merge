# -*- coding: utf-8 -*-
"""
藍主仕前置清洗 + 通用主程式包裝

用途：
- 將「主次代碼/*.CSV」轉成通用版可辨識的月份 CSV
- 再呼叫 run_merge_通用（自動偵測最新版）完成輸出
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
from typing import List, Optional


SCRIPT_DIR = Path(__file__).resolve().parent


def _find_generic_script(script_dir: Path) -> Path:
    candidates = sorted(script_dir.glob("run_merge_通用_*.py"), reverse=True)
    if not candidates:
        raise RuntimeError("找不到通用程式 run_merge_通用_*.py")
    return candidates[0]

GENERIC_SCRIPT = _find_generic_script(SCRIPT_DIR)
CP950 = "cp950"


def _load_generic_module():
    spec = importlib.util.spec_from_file_location("run_merge_generic", GENERIC_SCRIPT)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"無法載入通用程式：{GENERIC_SCRIPT}")
    mod = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = mod
    spec.loader.exec_module(mod)
    return mod


def _month_code_from_name(name: str) -> Optional[str]:
    match = re.search(r"(?<!\d)(1(?:14|15)\d{2})(?!\d)", name)
    return match.group(1) if match else None


def _parse_csv_rows(csv_path: Path) -> List[List[str]]:
    with csv_path.open("r", encoding=CP950, errors="replace", newline="") as f:
        return list(csv.reader(f))


def _header_index_map(header: List[str]) -> dict[str, int]:
    return {str(name).strip(): idx for idx, name in enumerate(header)}


def _pick(row: List[str], mapping: dict[str, int], *candidates: str) -> str:
    for candidate in candidates:
        idx = mapping.get(candidate)
        if idx is not None and idx < len(row):
            return str(row[idx]).strip()
    return ""


def _convert_main_code_csv(csv_path: Path, output_dir: Path) -> Optional[Path]:
    rows = _parse_csv_rows(csv_path)
    if len(rows) < 2:
        return None

    header = rows[0]
    mapping = _header_index_map(header)
    out_rows: List[List[str]] = [["ID", "姓名", "日期", "次數", "申請金額"]]

    for row in rows[1:]:
        pid = _pick(row, mapping, "身分證號", "身份證號", "ID")
        name = _pick(row, mapping, "病患姓名", "姓名", "會員姓名")
        dt = _pick(row, mapping, "就醫日", "日期")
        amount = _pick(row, mapping, "申請金額", "申請額")
        if not pid or not dt:
            continue
        out_rows.append([pid, name, dt, "1", amount])

    if len(out_rows) <= 1:
        return None

    month_code = _month_code_from_name(csv_path.stem) or csv_path.stem
    out_path = output_dir / f"{month_code}.csv"
    with out_path.open("w", encoding="utf-8-sig", newline="") as f:
        csv.writer(f).writerows(out_rows)
    return out_path


def _preclean_main_codes(source_dir: Path) -> List[Path]:
    src_dir = source_dir / "主次代碼"
    if not src_dir.is_dir():
        return []

    generated: List[Path] = []
    for csv_path in sorted(src_dir.glob("*.CSV")):
        converted = _convert_main_code_csv(csv_path, source_dir)
        if converted is not None:
            generated.append(converted)
    return generated


def process_excel(source_path: str, template_path: Optional[str] = None) -> str:
    generic = _load_generic_module()
    source_dir = Path(source_path).resolve()
    if not source_dir.is_dir():
        raise ValueError("請選擇資料夾。")

    template = template_path or generic._find_template(str(SCRIPT_DIR))
    if not os.path.exists(template):
        raise ValueError(f"找不到模板檔：{template}")

    temp_root = Path(tempfile.mkdtemp(prefix="lan_clean_"))
    temp_source = temp_root / source_dir.name

    try:
        shutil.copytree(source_dir, temp_source)
        generated = _preclean_main_codes(temp_source)
        print(f"已產生清洗後月份檔 {len(generated)} 個", flush=True)

        # 避免原始 CSV 與清洗後 CSV 同時被掃到造成重複/干擾
        for folder_name in ("主次代碼", "次數費用"):
            raw_dir = temp_source / folder_name
            if raw_dir.exists():
                shutil.rmtree(raw_dir, ignore_errors=True)

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

    src = filedialog.askdirectory(title="選擇藍主仕來源資料夾")
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
