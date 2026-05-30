# -*- coding: utf-8 -*-
"""
品安安展望雲端前置清洗 + 通用主程式包裝

用途：
- 先用展望讀檔邏輯收斂原始資料。
- 轉成通用版可辨識的標準來源資料夾。
- 再呼叫 run_merge_通用，避免輸出後硬修格式。

重要規則：
- 百分位名單仍以通用版規則為準：需有疾病樣態才納入糖心腎 KPI 候選。
- 不修改、不刪除原始資料。
"""

from __future__ import annotations

import importlib.util
import os
import re
import shutil
import sys
import tempfile
from pathlib import Path
from typing import Any, Dict, Iterable, Optional, Sequence

from openpyxl import Workbook


SCRIPT_DIR = Path(__file__).resolve().parent
PINAN_SCRIPT = SCRIPT_DIR / "run_merge_品安安_展望雲端_0525.py"


def _find_generic_script(script_dir: Path) -> Path:
    candidates = sorted(script_dir.glob("run_merge_通用_*.py"), reverse=True)
    if not candidates:
        raise RuntimeError("找不到通用程式 run_merge_通用_*.py")
    return candidates[0]


GENERIC_SCRIPT = _find_generic_script(SCRIPT_DIR)


def _load_module(path: Path, module_name: str):
    spec = importlib.util.spec_from_file_location(module_name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"無法載入程式：{path}")
    mod = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = mod
    spec.loader.exec_module(mod)
    return mod


GENERIC = _load_module(GENERIC_SCRIPT, "run_merge_generic")


def _select_reader(source_dir: Path):
    return _load_module(PINAN_SCRIPT, "run_merge_pinanan")


def _normalize_disease(value: Any) -> Any:
    text = GENERIC.normalize_text(value).upper()
    if not text:
        return None
    if "DKD" in text or "糖尿病腎" in text:
        return 3
    if "CKD" in text or "腎臟病" in text:
        return 2
    if "DM" in text or "糖尿病" in text:
        return 1
    if "ASCVD" in text:
        return 4
    parsed = GENERIC.parse_disease_code(value)
    return parsed.value if parsed else None


def _date_value(value: Any) -> Any:
    parsed = GENERIC.parse_date(value)
    return parsed if parsed else value


def _number_or_blank(value: Any) -> Any:
    num = GENERIC.parse_float(value)
    if num is None or num == 0:
        return None
    return int(num) if float(num).is_integer() else num


def _is_x115_source(text: Any) -> bool:
    compact = re.sub(r"[\s\-_()（）\[\]{}]+", "", str(text or "").strip()).lower()
    return bool(compact and ("115x" in compact or "不選" in compact or "不要" in compact))


def _append_rows(ws, headers: Sequence[str], rows: Iterable[Sequence[Any]]) -> None:
    ws.append(list(headers))
    for row in rows:
        ws.append(list(row))


def _save_workbook(path: Path, sheets: Dict[str, tuple[Sequence[str], Iterable[Sequence[Any]]]]) -> None:
    wb = Workbook()
    first = True
    for title, (headers, rows) in sheets.items():
        ws = wb.active if first else wb.create_sheet()
        ws.title = title
        first = False
        _append_rows(ws, headers, rows)
    wb.save(path)


def _collect_x115_rows(source_dir: Path, reader) -> list[list[Any]]:
    rows_out: list[list[Any]] = []
    seen: set[str] = set()
    for path in sorted(source_dir.iterdir()):
        if not path.is_file() or path.name.startswith("~$") or path.suffix.lower() not in (".xlsx", ".ods"):
            continue
        if not _is_x115_source(path.stem):
            continue
        for sheet_name, rows in reader.iter_workbook_rows(path):
            if not _is_x115_source(sheet_name) and not _is_x115_source(path.stem):
                continue
            header_index = reader.find_header(
                rows,
                [["姓名", "會員姓名", "病患姓名"], ["ID", "身分證", "身分證號", "身分證號碼", "身份證號", "身份證號碼"]],
            )
            if header_index is None:
                continue
            header = rows[header_index]
            name_col = reader.find_col(header, ["姓名", "會員姓名", "病患姓名"])
            id_col = reader.find_col(header, ["ID", "身分證", "身分證號", "身分證號碼", "身份證", "身份證號", "身份證號碼"])
            bday_col = reader.find_col(header, ["生日", "BIRTHDAY", "出生日期", "出生年月日"])
            if name_col is None or id_col is None:
                continue
            for row in rows[header_index + 1:]:
                if len(row) <= max(name_col, id_col):
                    continue
                pid = GENERIC.normalize_id(row[id_col])
                name = GENERIC.normalize_text(row[name_col])
                if not pid or pid in seen:
                    continue
                bday = row[bday_col] if bday_col is not None and len(row) > bday_col else None
                rows_out.append([pid, name, _date_value(bday)])
                seen.add(pid)
    return rows_out


def _screening_rows(records, key: str):
    for rec in records:
        value = rec.screenings.get(key)
        parsed = GENERIC.parse_date(value)
        if rec.pid and parsed:
            yield [rec.pid, parsed]


def _month_code(path: Path) -> Optional[str]:
    match = re.fullmatch(r"1(?:14|15)\d{2}", path.stem)
    return match.group(0) if match else None


def _month_date(code: str) -> str:
    return f"{code[:3]}/{code[3:5]}/01"


def _record_lookup_maps(records, reader):
    chart_to_pid: Dict[str, set[str]] = {}
    name_to_pid: Dict[str, set[str]] = {}
    for rec in records:
        if not rec.pid:
            continue
        for chart in getattr(rec, "charts", set()):
            normalized_chart = reader.normalize_chart(chart)
            if normalized_chart:
                chart_to_pid.setdefault(normalized_chart, set()).add(rec.pid)
        name = reader.normalize_name(rec.name)
        if name:
            name_to_pid.setdefault(name, set()).add(rec.pid)
    return chart_to_pid, name_to_pid


def _resolve_pid(chart: Any, name: Any, chart_to_pid: Dict[str, set[str]], name_to_pid: Dict[str, set[str]], reader) -> str:
    chart_key = reader.normalize_chart(chart)
    if chart_key:
        ids = chart_to_pid.get(chart_key, set())
        if len(ids) == 1:
            return next(iter(ids))
    name_key = reader.normalize_name(name)
    if name_key:
        ids = name_to_pid.get(name_key, set())
        if len(ids) == 1:
            return next(iter(ids))
    return ""


def _collect_monthly_claim_rows(source_dir: Path, records, reader) -> Dict[str, list[list[Any]]]:
    chart_to_pid, name_to_pid = _record_lookup_maps(records, reader)
    monthly: Dict[str, Dict[str, Dict[str, float]]] = {}

    count_dir = source_dir / "次數"
    if count_dir.is_dir():
        for path in sorted(count_dir.glob("*.xls")):
            code = _month_code(path)
            if not code:
                continue
            rows = reader.read_html_table(path)
            header_index = reader.find_header(
                rows,
                [["病歷號"], ["身分證", "身分證號", "身份證號"], ["看診次數", "次數"]],
            )
            if header_index is None:
                continue
            header = rows[header_index]
            chart_col = reader.find_col(header, ["病歷號", "病歷號碼"])
            name_col = reader.find_col(header, ["姓名", "病患姓名"])
            id_col = reader.find_col(header, ["身分證", "身分證號", "身份證號", "ID"])
            count_col = reader.find_col(header, ["看診次數", "次數", "就診次數"])
            if None in (chart_col, name_col, id_col, count_col):
                continue
            for row in rows[header_index + 1:]:
                if len(row) <= max(chart_col, name_col, id_col, count_col):
                    continue
                pid = GENERIC.normalize_id(row[id_col])
                if not pid:
                    pid = _resolve_pid(row[chart_col], row[name_col], chart_to_pid, name_to_pid, reader)
                if not pid:
                    continue
                count = reader.parse_number(row[count_col])
                if not count:
                    continue
                data = monthly.setdefault(code, {}).setdefault(pid, {"count": 0.0, "amount": 0.0})
                data["count"] += count

    fee_dir = source_dir / "費用"
    if fee_dir.is_dir():
        for path in sorted(fee_dir.glob("*.xls")):
            code = _month_code(path)
            if not code:
                continue
            rows = reader.read_html_table(path)
            header_index = reader.find_header(rows, [["病歷號"], ["姓名"], ["掛帳費"]])
            if header_index is None:
                continue
            header = rows[header_index]
            chart_col = reader.find_col(header, ["病歷號", "病歷號碼"])
            name_col = reader.find_col(header, ["姓名", "病患姓名"])
            amount_col = reader.find_col(header, ["掛帳費"])
            if None in (chart_col, name_col, amount_col):
                continue
            for row in rows[header_index + 1:]:
                if len(row) <= max(chart_col, name_col, amount_col):
                    continue
                pid = _resolve_pid(row[chart_col], row[name_col], chart_to_pid, name_to_pid, reader)
                if not pid:
                    continue
                amount = reader.parse_number(row[amount_col])
                if not amount:
                    continue
                data = monthly.setdefault(code, {}).setdefault(pid, {"count": 0.0, "amount": 0.0})
                data["amount"] += amount

    return {
        code: [
            [pid, _month_date(code), _number_or_blank(values["count"]), _number_or_blank(values["amount"])]
            for pid, values in sorted(member_map.items())
            if values["count"] or values["amount"]
        ]
        for code, member_map in sorted(monthly.items())
    }


def _fallback_claim_rows(records, year: int, q1_only: bool = False):
    for rec in records:
        if not rec.pid:
            continue
        if year == 114:
            count = rec.count_114_q1 if q1_only else max(rec.count_114 - rec.count_114_q1, 0)
            amount = rec.amount_114_q1 if q1_only else max(rec.amount_114 - rec.amount_114_q1, 0)
        else:
            count = rec.count_115_q1 if q1_only else max(rec.count_115 - rec.count_115_q1, 0)
            amount = rec.amount_115_q1 if q1_only else max(rec.amount_115 - rec.amount_115_q1, 0)
        if count or amount:
            month = 1 if q1_only else 5
            yield [rec.pid, f"{year}/{month:02d}/01", _number_or_blank(count), _number_or_blank(amount)]


def _write_clean_source(source_dir: Path, clean_dir: Path) -> None:
    reader = _select_reader(source_dir)
    records = sorted(
        reader.collect_data(source_dir).values(),
        key=lambda rec: (0 if rec.pid else 1, reader.normalize_name(rec.name), rec.pid),
    )

    member_rows = []
    health_rows = []
    ascvd_rows = []
    self_select_rows = []
    x115_rows = []
    for rec in records:
        disease = _normalize_disease(rec.dmk_raw)
        ascvd = rec.ascvd if rec.ascvd not in (None, "") else None
        member_rows.append([
            rec.pid,
            rec.name,
            _date_value(rec.bday),
            rec.sex,
            disease,
            rec.last_visit,
            "A" if rec.pid else None,
        ])
        health_rows.append([
            rec.pid,
            _number_or_blank(rec.health.get("hba_val")),
            _date_value(rec.health.get("hba_dt")),
            _number_or_blank(rec.health.get("ldl_val")),
            _date_value(rec.health.get("ldl_dt")),
            _number_or_blank(rec.health.get("uacr_val")),
            _date_value(rec.health.get("uacr_dt")),
        ])
        if ascvd is not None:
            ascvd_rows.append([rec.pid, ascvd, rec.last_visit])
        if getattr(rec, "is_self_select", False):
            self_select_rows.append([rec.pid, rec.name, _date_value(rec.bday)])
        if getattr(rec, "is_115x", False):
            x115_rows.append([rec.pid, rec.name, _date_value(rec.bday)])
    for row in _collect_x115_rows(source_dir, reader):
        if row[0] not in {existing[0] for existing in x115_rows}:
            x115_rows.append(row)

    _save_workbook(
        clean_dir / "展望清洗_主檔.xlsx",
        {
            "會員名單": (
                ["ID", "姓名", "生日", "性別", "疾病樣態", "最後就診日", "會員別"],
                member_rows,
            ),
            "HealthCase": (
                [
                    "家醫收案會員ID",
                    "最近一次HbA1c檢查結果(%)",
                    "最近一次HbA1c檢查日期",
                    "最近一次LDL檢查結果(mg/dL)",
                    "最近一次LDL檢查日期",
                    "最近一次UACR檢查結果(mg/gm)",
                    "最近一次UACR檢查日期",
                ],
                health_rows,
            ),
            "ascvd": (["ID", "ASCVD", "最後就診日"], ascvd_rows),
            "自選會員": (["ID", "姓名", "生日"], self_select_rows),
            "115X": (["ID", "姓名", "生日"], x115_rows),
        },
    )

    screening_map = {
        "成人健檢": "成人健檢",
        "子宮抹片": "子宮抹片",
        "老人流感": "老人流感",
        "糞便潛血": "糞便潛血",
        "肝炎篩檢": "BC肝炎",
    }
    _save_workbook(
        clean_dir / "展望清洗_篩檢.xlsx",
        {
            title: (["ID", "最後篩檢日期"], list(_screening_rows(records, key)))
            for title, key in screening_map.items()
        },
    )

    monthly_claim_rows = _collect_monthly_claim_rows(source_dir, records, reader)
    if not monthly_claim_rows:
        monthly_claim_rows = {
            "11401": list(_fallback_claim_rows(records, 114, q1_only=True)),
            "11404": list(_fallback_claim_rows(records, 114, q1_only=False)),
            "11501": list(_fallback_claim_rows(records, 115, q1_only=True)),
            "11504": list(_fallback_claim_rows(records, 115, q1_only=False)),
        }
    _save_workbook(
        clean_dir / "展望清洗_申報統計.xlsx",
        {
            code: (["ID", "日期", "次數", "申請金額"], rows)
            for code, rows in monthly_claim_rows.items()
        },
    )


def process_excel(source_path: str, template_path: Optional[str] = None) -> str:
    source_dir = Path(source_path).expanduser().resolve()
    if not source_dir.is_dir():
        raise ValueError("請選擇品安安展望雲端資料夾。")

    template = template_path or GENERIC._find_template(str(SCRIPT_DIR))
    if not os.path.exists(template):
        raise ValueError(f"找不到模板檔：{template}")

    temp_root = Path(tempfile.mkdtemp(prefix="zhanwang_clean_"))
    clean_dir = temp_root / source_dir.name
    clean_dir.mkdir(parents=True, exist_ok=True)
    try:
        _write_clean_source(source_dir, clean_dir)
        temp_output = Path(GENERIC.process_excel(str(clean_dir), template))
        final_output = source_dir.parent / temp_output.name
        if final_output.exists():
            final_output.unlink()
        shutil.move(str(temp_output), str(final_output))
        return str(final_output)
    finally:
        shutil.rmtree(temp_root, ignore_errors=True)


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = list(argv if argv is not None else sys.argv[1:])
    if args:
        out = process_excel(args[0], args[1] if len(args) > 1 else None)
        print(f"已輸出：{out}")
        return 0

    import tkinter as tk
    from tkinter import filedialog, messagebox

    root = tk.Tk()
    root.withdraw()
    src = filedialog.askdirectory(title="選擇品安安展望雲端資料夾")
    if not src:
        return 1
    try:
        out = process_excel(src)
        messagebox.showinfo("完成", f"已輸出：\n{out}")
        GENERIC.open_file_cross_platform(out)
    except Exception as exc:
        messagebox.showerror("錯誤", str(exc))
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
