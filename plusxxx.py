#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""plusxxx.py

從視窗選資料夾，讀取資料夾內的Excel/ODS (xls/xlsx/ods)
對包含欄位名稱包含 "ID" 或 "身份" 或 "身分" 的欄位進行更新：
如果值是英文 + 9個數字 (如 A123450000)，則數字 + 20513 -> A123470513

輸出到原資料夾下的新資料夾 [data20513]，每個檔名後綴 _13.xlsx
"""

import os
import re
import sys
from datetime import date, datetime
from pathlib import Path

import pandas as pd

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox
except ImportError:
    tk = None

PATTERN = re.compile(r"^([A-Za-z])(\d{9})$")
ID_KEYWORDS = ["ID", "身份", "身分"]
HEADER_HINTS = ID_KEYWORDS + ["生日", "出生"]
DATE_TEXT_PATTERN = re.compile(
    r"^\s*(\d{4}[-/.]\d{1,2}[-/.]\d{1,2}"
    r"|\d{4}年\d{1,2}月\d{1,2}日)"
    r"(\s+\d{1,2}:\d{2}(:\d{2})?)?\s*$"
)


def normalize_value(v):
    if isinstance(v, pd.Series):
        return v.apply(normalize_value)
    if pd.isna(v):
        return v
    s = str(v).strip()
    m = PATTERN.match(s)
    if not m:
        return v
    prefix = m.group(1)
    numeric = int(m.group(2))
    numeric += 20513
    if numeric < 0:
        return v
    new_numeric = str(numeric).zfill(9)
    return f"{prefix}{new_numeric}"


def normalize_date_value(v):
    if pd.isna(v):
        return v
    if isinstance(v, pd.Timestamp):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, str) and DATE_TEXT_PATTERN.match(v):
        dt = pd.to_datetime(v, errors="coerce")
        if not pd.isna(dt):
            return dt.strftime("%Y-%m-%d")
    return v


def clean_id_text(v):
    if pd.isna(v):
        return ""
    return re.sub(r"[^A-Za-z0-9]", "", str(v).strip()).upper()


def is_id_like(v):
    return bool(PATTERN.fullmatch(clean_id_text(v)))


def detect_header_and_data_start(raw_df: pd.DataFrame):
    for header_idx, row in raw_df.iterrows():
        header = row.fillna("").astype(str).tolist()
        if not any(any(hint in col_name for hint in HEADER_HINTS) for col_name in header):
            continue

        matched_cols = [
            idx for idx, col_name in enumerate(header)
            if any(k in col_name for k in ID_KEYWORDS)
        ]
        if not matched_cols:
            continue

        for data_idx in range(header_idx + 1, len(raw_df)):
            for col_idx in matched_cols:
                if is_id_like(raw_df.iat[data_idx, col_idx]):
                    return header_idx, data_idx, matched_cols

    fallback_data_start = 1 if len(raw_df) > 1 else 0
    fallback_header = raw_df.iloc[0].fillna("").astype(str).tolist() if len(raw_df) else []
    fallback_cols = [
        idx for idx, col_name in enumerate(fallback_header)
        if any(k in col_name for k in ID_KEYWORDS)
    ]
    return 0, fallback_data_start, fallback_cols


def process_file(input_path: Path, output_path: Path):
    extensions = {".xls", ".xlsx", ".ods"}
    if input_path.suffix.lower() not in extensions:
        return False

    try:
        raw_sheets = pd.read_excel(input_path, sheet_name=None, header=None, engine=None)
    except Exception as e:
        print(f"無法讀取 {input_path.name}：{e}")
        return False

    changed = False
    writer = pd.ExcelWriter(output_path, engine="openpyxl")

    for sheet_name, raw in raw_sheets.items():
        raw = raw.copy().map(normalize_date_value)
        header_row, data_start_row, matched_col_indexes = detect_header_and_data_start(raw)
        if header_row >= len(raw):
            header_row = 0
        if data_start_row <= header_row:
            data_start_row = header_row + 1

        if matched_col_indexes:
            for col_idx in matched_col_indexes:
                for row_idx in range(data_start_row, len(raw)):
                    old_value = raw.iat[row_idx, col_idx]
                    new_value = normalize_value(old_value)
                    if new_value != old_value:
                        raw.iat[row_idx, col_idx] = new_value
                        changed = True

        raw.to_excel(writer, sheet_name=sheet_name, header=False, index=False)

    writer.close()

    if changed:
        print(f"已處理並輸出：{output_path}")
        return True
    else:
        print(f"已輸出（無欄位匹配）：{output_path}")
        return True


def open_output_folder(folder_path: Path):
    try:
        if sys.platform.startswith("darwin"):
            os.system(f'open "{folder_path}"')
        elif os.name == "nt":
            os.startfile(str(folder_path))
        else:
            os.system(f'xdg-open "{folder_path}"')
    except Exception as e:
        print(f"無法自動開啟資料夾：{e}")


def choose_folder_and_run():
    if tk is None:
        print("tkinter 無法使用，請確保已安裝 tkinter。")
        return

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    folder = filedialog.askdirectory(title="選擇要處理的資料夾")
    if not folder:
        print("未選擇資料夾，結束。")
        return

    folder_path = Path(folder)
    output_dir = folder_path / "data20513"
    output_dir.mkdir(exist_ok=True)

    files = list(folder_path.iterdir())
    if not files:
        messagebox.showinfo("提示", "資料夾沒有檔案。")
        return

    success_count = 0
    fail_count = 0

    for item in files:
        if item.is_file() and item.suffix.lower() in [".xls", ".xlsx", ".ods"]:
            out_name = f"{item.stem}_13.xlsx"
            out_path = output_dir / out_name
            ok = process_file(item, out_path)
            if ok:
                success_count += 1
            else:
                fail_count += 1

    messagebox.showinfo("完成", f"完成：成功 {success_count} 個，失敗 {fail_count} 個。\n輸出資料夾：{output_dir}")
    open_output_folder(output_dir)


if __name__ == "__main__":
    choose_folder_and_run()
