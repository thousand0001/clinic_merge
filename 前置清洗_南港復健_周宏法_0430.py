# -*- coding: utf-8 -*-
"""
南港 / 周宏法 前置清洗 + 通用主程式包裝

用途：
- 先把 TXT 月份檔轉成通用版可辨識的 CSV
- 再呼叫 run_merge_通用（自動偵測最新版）完成輸出

適用資料：
- 南港復健科醫聖：R11440次數/*.txt
- 周宏法醫聖：次數費用/*.txt
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
from typing import Dict, List, Optional


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


def _normalize_header_text(text: str) -> str:
    return re.sub(r"\s+", "", text.strip())


def _parse_txt_lines(txt_path: Path) -> List[List[str]]:
    with txt_path.open("r", encoding=CP950, errors="replace", newline="") as f:
        raw_lines = [line.rstrip("\r\n") for line in f if line.strip()]

    rows: List[List[str]] = []
    for line in raw_lines:
        if "費用年月:" in line:
            continue
        parts = [part.strip() for part in line.split(",")]
        while parts and parts[-1] == "":
            parts.pop()
        if parts:
            rows.append(parts)
    return rows


def _header_index_map(header: List[str]) -> Dict[str, int]:
    return {_normalize_header_text(name): idx for idx, name in enumerate(header)}


def _pick(row: List[str], mapping: Dict[str, int], *candidates: str) -> str:
    for candidate in candidates:
        idx = mapping.get(candidate)
        if idx is not None and idx < len(row):
            return row[idx].strip()
    return ""


def _convert_txt_to_monthly_csv(txt_path: Path) -> Optional[Path]:
    rows = _parse_txt_lines(txt_path)
    if len(rows) < 2:
        return None

    header = rows[0]
    mapping = _header_index_map(header)
    data_rows = rows[1:]

    out_rows: List[List[str]] = [["ID", "姓名", "生日", "日期", "次數", "申請金額"]]

    for row in data_rows:
        pid = _pick(row, mapping, "身分證", "身分證號", "身份證號", "ID")
        dt = _pick(row, mapping, "看診日", "日期", "就醫日", "最後就診日")
        name = _pick(row, mapping, "姓名", "病患姓名", "會員姓名")
        bday = _pick(row, mapping, "生日", "出生日期")

        if not pid or not dt:
            continue

        amount = _pick(row, mapping, "申請額", "申請金額", "申請額小計")
        out_rows.append([pid, name, bday, dt, "1", amount])

    if len(out_rows) <= 1:
        return None

    month_code = _month_code_from_name(txt_path.stem) or _month_code_from_name(txt_path.name)
    out_name = f"{month_code or txt_path.stem}.csv"
    out_path = txt_path.with_name(out_name)
    with out_path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerows(out_rows)
    return out_path


def _preclean_txt_monthlies(source_dir: Path) -> List[Path]:
    generated: List[Path] = []
    for txt_path in source_dir.rglob("*.txt"):
        converted = _convert_txt_to_monthly_csv(txt_path)
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

    temp_root = Path(tempfile.mkdtemp(prefix="nk_whf_"))
    temp_source = temp_root / source_dir.name

    try:
        shutil.copytree(source_dir, temp_source)
        generated = _preclean_txt_monthlies(temp_source)
        print(f"已產生清洗後月份檔 {len(generated)} 個", flush=True)

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

    src = filedialog.askdirectory(title="選擇南港/周宏法來源資料夾")
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
