# -*- coding: utf-8 -*-
"""
醫聖前置清洗 + 通用主程式包裝（0513）

用途：
- 若資料夾內含 TXT 月份檔，先轉成通用版可辨識的 CSV
- 其餘 Excel 型醫聖資料則直接交給通用版
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
from typing import Dict, List, Optional


SCRIPT_DIR = Path(__file__).resolve().parent


def _find_generic_script(script_dir: Path) -> Path:
    candidates = sorted(script_dir.glob("run_merge_通用_*.py"), reverse=True)
    if not candidates:
        raise RuntimeError("找不到通用程式 run_merge_通用_*.py")
    return candidates[0]

GENERIC_SCRIPT = _find_generic_script(SCRIPT_DIR)
CP950 = "cp950"
TW_ID_RE = re.compile(r"^[A-Z][12]\d{8}$")
AGE_TEXT_RE = re.compile(r"^\d{2,3}歲\d{1,2}月\d{1,2}天$")
MAX_REASONABLE_CLAIM_AMOUNT = 50000


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
    if match:
        return match.group(1)
    match = re.search(r"(?<!\d)(1(?:14|15))[./-](\d{2})[./-]\d{2}", name)
    if match:
        return f"{match.group(1)}{match.group(2)}"
    return None


def _normalize_existing_month_workbooks(source_dir: Path) -> List[Path]:
    generated: List[Path] = []
    month_dir = source_dir / "醫聖月份費用xlsx"
    if not month_dir.is_dir():
        return generated

    for path in sorted(month_dir.glob("*.xlsx")):
        code = _month_code_from_name(path.stem) or _month_code_from_name(path.name)
        if not code:
            continue
        target = month_dir / f"{code}.xlsx"
        if target.exists() or target == path:
            continue
        shutil.copy2(path, target)
        generated.append(target)
    return generated


def _existing_standard_month_codes(source_dir: Path) -> set[str]:
    codes: set[str] = set()
    for path in source_dir.rglob("*"):
        if not path.is_file():
            continue
        if path.suffix.lower() not in {".xlsx", ".xls", ".csv", ".ods"}:
            continue
        code = _month_code_from_name(path.stem) or _month_code_from_name(path.name)
        if code:
            codes.add(code)
    return codes


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


def _find_txt_header_row(rows: List[List[str]]) -> int:
    for idx, row in enumerate(rows[:5]):
        normalized = {_normalize_header_text(v) for v in row}
        if any(token in normalized for token in ("身分證", "身分證號", "身份證號", "姓名", "看診日", "病歷號")):
            return idx
    return 0


def _header_index_map(header: List[str]) -> Dict[str, int]:
    return {_normalize_header_text(name): idx for idx, name in enumerate(header)}


def _pick(row: List[str], mapping: Dict[str, int], *candidates: str) -> str:
    for candidate in candidates:
        idx = mapping.get(candidate)
        if idx is not None and idx < len(row):
            return row[idx].strip()
    return ""


def _looks_like_tw_id(text: str) -> bool:
    return bool(TW_ID_RE.fullmatch(text.strip().upper()))


def _looks_like_date(text: str) -> bool:
    s = text.strip()
    return bool(re.fullmatch(r"\d{2,3}[./-]\d{1,2}[./-]\d{1,2}", s) or re.fullmatch(r"\d{4}[./-]\d{1,2}[./-]\d{1,2}", s))


def _looks_like_name(text: str) -> bool:
    s = text.strip()
    if not s:
        return False
    if _looks_like_tw_id(s) or _looks_like_date(s):
        return False
    if re.fullmatch(r"\d+", s):
        return False
    if "歲" in s and "天" in s:
        return False
    return True


def _clean_claim_amount(text: str) -> str:
    s = str(text).strip().replace(",", "")
    if not re.fullmatch(r"\d+(?:\.0+)?", s):
        return ""
    amount = int(float(s))
    if amount < 0 or amount > MAX_REASONABLE_CLAIM_AMOUNT:
        return ""
    return str(amount)


def _extract_shifted_claim_amount(row: List[str]) -> str:
    """
    醫聖 TXT 的主訴欄可能含逗號，但原始檔沒有 CSV 引號保護。
    line.split(",") 後，主訴後面的欄位會整段右移；此時以前 10 欄為主訴前欄位，
    從資料列尾端回推 12 個固定欄位，申請額固定落在倒數第 7 欄。
    """
    if len(row) >= 22:
        amount = _clean_claim_amount(row[-7])
        if amount:
            return amount

    for idx, value in enumerate(row):
        if AGE_TEXT_RE.fullmatch(str(value).strip()) and idx + 4 < len(row):
            amount = _clean_claim_amount(row[idx + 4])
            if amount:
                return amount
    return ""


def _extract_row_fields(row: List[str], mapping: Dict[str, int]) -> tuple[str, str, str, str, str]:
    pid = _pick(row, mapping, "身分證", "身分證號", "身份證號", "ID")
    name = _pick(row, mapping, "姓名", "病患姓名", "會員姓名")
    dt = _pick(row, mapping, "看診日", "日期", "就醫日", "最後就診日")
    bday = _pick(row, mapping, "生日", "出生日期")
    amount = _pick(row, mapping, "申請額", "申請金額", "申請額小計")
    amount = _clean_claim_amount(amount) or _extract_shifted_claim_amount(row)

    if _looks_like_tw_id(pid) and _looks_like_name(name) and _looks_like_date(dt):
        return pid, name, bday, dt, amount

    id_idx = next((i for i, v in enumerate(row) if _looks_like_tw_id(str(v))), None)
    date_indexes = [i for i, v in enumerate(row) if _looks_like_date(str(v))]

    if id_idx is not None:
        pid = str(row[id_idx]).strip()
    if (not dt or not _looks_like_date(dt)) and date_indexes:
        dt = str(row[date_indexes[0]]).strip()
    if (not bday or not _looks_like_date(bday) or bday == dt) and len(date_indexes) >= 2:
        bday = str(row[date_indexes[1]]).strip()

    if not _looks_like_name(name):
        dt_idx = next((i for i, v in enumerate(row) if str(v).strip() == dt), None)
        if dt_idx is not None and dt_idx > 0 and _looks_like_name(str(row[dt_idx - 1])):
            name = str(row[dt_idx - 1]).strip()
        elif id_idx is not None:
            for cand in (id_idx - 1, id_idx - 2, id_idx - 3, id_idx - 4):
                if 0 <= cand < len(row) and _looks_like_name(str(row[cand])):
                    name = str(row[cand]).strip()
                    break

    if not amount:
        amount = _extract_shifted_claim_amount(row)

    return pid, name, bday, dt, amount


def _convert_txt_to_monthly_csv(txt_path: Path) -> Optional[Path]:
    rows = _parse_txt_lines(txt_path)
    if len(rows) < 2:
        return None

    header_row = _find_txt_header_row(rows)
    header = rows[header_row]
    mapping = _header_index_map(header)
    data_rows = rows[header_row + 1:]

    out_rows: List[List[str]] = [["ID", "姓名", "生日", "日期", "次數", "申請金額"]]

    for row in data_rows:
        pid, name, bday, dt, amount = _extract_row_fields(row, mapping)

        if not pid or not dt:
            continue

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
    generated.extend(_normalize_existing_month_workbooks(source_dir))
    existing_codes = _existing_standard_month_codes(source_dir)
    for txt_path in source_dir.rglob("*.txt"):
        month_code = _month_code_from_name(txt_path.stem) or _month_code_from_name(txt_path.name)
        if month_code and month_code in existing_codes:
            continue
        converted = _convert_txt_to_monthly_csv(txt_path)
        if converted is not None:
            generated.append(converted)
            code = _month_code_from_name(converted.stem) or _month_code_from_name(converted.name)
            if code:
                existing_codes.add(code)
    return generated


def process_excel(source_path: str, template_path: Optional[str] = None) -> str:
    generic = _load_generic_module()
    source_dir = Path(source_path).resolve()
    if not source_dir.is_dir():
        raise ValueError("請選擇資料夾。")

    template = template_path or generic._find_template(str(SCRIPT_DIR))
    if not os.path.exists(template):
        raise ValueError(f"找不到模板檔：{template}")

    temp_root = Path(tempfile.mkdtemp(prefix="ys_clean_"))
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

    src = filedialog.askdirectory(title="選擇醫聖來源資料夾")
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
