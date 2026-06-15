# -*- coding: utf-8 -*-
"""
方鼎前置清洗 + 最新共用核心包裝（0616）

保留方鼎 XLS 相容、月份檔名與姓名補 ID 規則，
自選／不選會員交由最新共用核心依通用 E1/E2 規則處理。
"""

from __future__ import annotations

import argparse
import importlib.util
import os
import re
import sys
import datetime
import shutil
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple


SCRIPT_DIR = Path(__file__).resolve().parent


def _find_common_core(script_dir: Path) -> Path:
    candidates = sorted(script_dir.glob("選會員_共用核心_*.py"), reverse=True)
    if not candidates:
        raise RuntimeError("找不到共用核心：選會員_共用核心_*.py")
    return candidates[0]

COMMON_CORE = _find_common_core(SCRIPT_DIR)


def _load_common_core():
    spec = importlib.util.spec_from_file_location("fangding_common_core", COMMON_CORE)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"無法載入共用核心：{COMMON_CORE}")
    mod = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = mod
    spec.loader.exec_module(mod)
    return mod


@dataclass
class FangdingMonthlyClaimSheetScan:
    sheet_name: str
    header_row: int
    id_col: Optional[int]
    name_col: Optional[int]
    date_col: Optional[int]
    count_col: Optional[int]
    amount_col: Optional[int]
    year_bucket: Optional[int]
    month: Optional[int]


GENERIC = _load_common_core()


class FangdingProfile(GENERIC.ProcessingProfile):

    def __init__(self) -> None:
        super().__init__()
        self.claim_fee_months_115: List[int] = []

    def load_xls_as_workbook(self, xls_path: str):
        # 方鼎匯出的 .xls 檔有時實際上是 xlsx 格式，xlrd 無法讀取，需改用 openpyxl。
        try:
            return GENERIC._load_xls_as_workbook(xls_path)
        except Exception as exc:
            if "xlsx file; not supported" not in str(exc).lower():
                raise
            fd, temp_xlsx_path = tempfile.mkstemp(prefix="fangding_compat_", suffix=".xlsx")
            os.close(fd)
            try:
                shutil.copyfile(xls_path, temp_xlsx_path)
                return GENERIC.openpyxl.load_workbook(temp_xlsx_path, data_only=True)
            finally:
                try:
                    os.unlink(temp_xlsx_path)
                except OSError:
                    pass

    def parse_date(self, value: Any) -> Optional[datetime.date]:
        if value is None:
            return None
        s = str(value).strip()
        m = re.match(r"^(\d{4})\.(\d{1,2})\.(\d{1,2})$", s)
        if m:
            y, mo, d = map(int, m.groups())
            try:
                return datetime.date(y, mo, d)
            except ValueError:
                return None
        return super().parse_date(value)

    def extract_monthly_sheet_code(self, text: Any) -> Optional[str]:
        raw = GENERIC.normalize_text(text)
        if not raw:
            return None
        match = re.search(r"(?<!\d)(1(?:14|15)\d{2})(?!\d)", raw)
        return match.group(1) if match else None

    def monthly_sheet_kind_token(self, *texts: Any) -> str:
        normalized = "".join(GENERIC._normalize_sheet_lookup(t) for t in texts if t is not None)
        if any(token in normalized for token in ("費用統計", "費用明細", "費用", "申請額")):
            return "fee"
        if any(token in normalized for token in ("就診次數", "看診次數", "申報筆數", "次數統計", "次數")):
            return "count"
        return "monthly"

    def sheet_year_bucket(self, title: str) -> Optional[int]:
        code = self.extract_monthly_sheet_code(title)
        if not code:
            return None
        if code.startswith("114"):
            return 114
        if code.startswith("115"):
            return 115
        return None

    def sheet_month(self, title: str) -> Optional[int]:
        code = self.extract_monthly_sheet_code(title)
        return int(code[3:5]) if code else None

    def canonical_source_sheet_name(self, sheet_name: str, file_path: str, single_sheet: bool, src_ws: Any = None) -> str:
        file_stem = os.path.splitext(os.path.basename(file_path))[0]
        monthly_code = self.extract_monthly_sheet_code(file_stem) or self.extract_monthly_sheet_code(sheet_name)
        if monthly_code:
            kind = self.monthly_sheet_kind_token(file_stem, sheet_name)
            if single_sheet:
                return f"{monthly_code}_{kind}"
            sheet_token = GENERIC._normalize_sheet_lookup(sheet_name) or "sheet"
            return f"{monthly_code}_{kind}_{sheet_token}"[:31]
        return GENERIC._canonical_source_sheet_name(sheet_name, file_path, single_sheet, src_ws)

    def find_monthly_claim_header_row(self, sheet: Any, search_rows: int = 30) -> Optional[int]:
        id_aliases = ["ID", "身分證號", "身分證號碼", "身份證號", "身份證號碼", "身分證字號", "身份證字號"]
        date_aliases = ["看診日", "日期", "最後看診日期", "最後就診日", "最後看診日"]
        amount_aliases = ["醫療費用總計", "申請金額", "申請額", "申報總金額", "總金額", "總額"]
        count_aliases = ["申報筆數", "看診次數", "就診次數", "門診次數", "次數", "件數"]
        has_sheet_month = self.sheet_year_bucket(sheet.title) in (114, 115) and self.sheet_month(sheet.title) is not None

        for r in range(1, min(search_rows, sheet.max_row) + 1):
            hmap = GENERIC.build_header_map(sheet, r)
            id_col_cand = GENERIC.find_column_exact(hmap, id_aliases)
            id_col = GENERIC.find_id_col_by_content(sheet, r, id_col_cand)
            name_col = GENERIC.find_column_exact(hmap, ["姓名", "會員姓名", "個案姓名", "病患姓名"])
            if id_col is None and name_col is None:
                continue

            date_col = GENERIC.find_column_exact(hmap, date_aliases)
            count_col = GENERIC.find_column_exact(hmap, count_aliases)
            amount_col = None
            for alias in amount_aliases:
                amount_col = GENERIC.find_col_by_keywords_any_row(sheet, r, [alias])
                if amount_col:
                    break
            if (count_col or amount_col) and (date_col or has_sheet_month):
                return r
        return None

    def scan_monthly_claim_sheet(self, sheet_name: str, sheet: Any) -> Optional[FangdingMonthlyClaimSheetScan]:
        id_aliases = ["ID", "身分證號", "身分證號碼", "身份證號", "身份證號碼", "身分證字號", "身份證字號", "身分證", "身份證"]
        date_aliases = ["看診日", "日期", "最後看診日期", "最後就診日", "最後看診日"]
        amount_aliases = ["醫療費用總計", "申請金額", "申請額", "申報總金額", "總金額", "總額"]
        count_aliases = ["申報筆數", "看診次數", "就診次數", "門診次數", "次數", "件數"]
        header_row = self.find_monthly_claim_header_row(sheet, search_rows=30)
        if header_row is None:
            return None
        hmap = GENERIC.build_header_map(sheet, header_row)
        id_col = GENERIC.find_id_col_by_content(sheet, header_row, GENERIC.find_column_exact(hmap, id_aliases))
        name_col = GENERIC.find_column_exact(hmap, ["姓名", "會員姓名", "個案姓名", "病患姓名"])
        date_col = GENERIC.find_column_exact(hmap, date_aliases)
        count_col = GENERIC.find_column_exact(hmap, count_aliases)
        amount_col = None
        for alias in amount_aliases:
            amount_col = GENERIC.find_col_by_keywords_any_row(sheet, header_row, [alias])
            if amount_col:
                break
        has_sheet_month = self.sheet_year_bucket(sheet_name) in (114, 115) and self.sheet_month(sheet_name) is not None
        if (id_col is None and name_col is None) or (count_col is None and amount_col is None) or (date_col is None and not has_sheet_month):
            return None
        return FangdingMonthlyClaimSheetScan(
            sheet_name=sheet_name,
            header_row=header_row,
            id_col=id_col,
            name_col=name_col,
            date_col=date_col,
            count_col=count_col,
            amount_col=amount_col,
            year_bucket=self.sheet_year_bucket(sheet_name),
            month=self.sheet_month(sheet_name),
        )

    def collect_monthly_claim_summaries(
        self,
        wb_src: Any,
        monthly_scans: Optional[Dict[str, FangdingMonthlyClaimSheetScan]] = None,
    ) -> Tuple[Dict[str, Dict[str, float]], List[int]]:
        all_members = GENERIC.collect_all_members(wb_src)
        member_name_to_ids: Dict[str, List[str]] = {}
        for pid, info in all_members.items():
            name = GENERIC.normalize_text(info.get("name"))
            if name:
                member_name_to_ids.setdefault(name, []).append(pid)

        out: Dict[str, Dict[str, float]] = {}
        seen_115_count_months: set = set()
        seen_115_fee_months: set = set()
        scans = monthly_scans or {
            sheet_name: scan
            for sheet_name in wb_src.sheetnames
            if (scan := self.scan_monthly_claim_sheet(sheet_name, wb_src[sheet_name])) is not None
        }

        for sheet_name, scan in scans.items():
            sh = wb_src[sheet_name]
            matched_any_115_count = False
            matched_any_115_fee = False
            for r in range(scan.header_row + 1, sh.max_row + 1):
                pid = ""
                if scan.id_col:
                    raw_pid = sh.cell(r, scan.id_col).value
                    pid = GENERIC.normalize_text(raw_pid).upper()
                    if not GENERIC.is_valid_tw_id(pid):
                        pid = ""
                if not pid and scan.name_col:
                    name = GENERIC.normalize_text(sh.cell(r, scan.name_col).value)
                    matched_ids = member_name_to_ids.get(name, [])
                    if len(matched_ids) == 1:
                        pid = matched_ids[0]
                if not pid:
                    continue

                dt = self.parse_date(sh.cell(r, scan.date_col).value) if scan.date_col else None
                cnt = GENERIC.parse_float(sh.cell(r, scan.count_col).value) if scan.count_col else None
                amt = GENERIC.parse_float(sh.cell(r, scan.amount_col).value) if scan.amount_col else None
                if dt is None and cnt is None and amt is None:
                    continue

                year_bucket = GENERIC._roc_year_from_date(dt)
                if year_bucket not in (114, 115):
                    year_bucket = scan.year_bucket
                if year_bucket not in (114, 115):
                    continue

                month = dt.month if dt is not None else scan.month
                if month is None:
                    continue

                is_q1 = month <= 4
                bucket = out.setdefault(pid, GENERIC._empty_claim_bucket())
                if dt is not None:
                    bucket["last_visit_ord"] = max(bucket.get("last_visit_ord", 0.0), float(dt.toordinal()))

                prefix = str(year_bucket)
                if cnt is not None:
                    month_key = GENERIC._claim_month_key(year_bucket, "cnt", month)
                    bucket[month_key] = bucket.get(month_key, 0.0) + cnt
                    if year_bucket == 115:
                        bucket["115_cnt"] += cnt
                        if is_q1:
                            bucket["115_cnt_q1"] += cnt
                            matched_any_115_count = True
                    elif is_q1:
                        bucket[f"{prefix}_cnt"] += cnt
                    if year_bucket == 114:
                        bucket["114_cnt_full"] += cnt
                if amt is not None:
                    month_key = GENERIC._claim_month_key(year_bucket, "amt", month)
                    bucket[month_key] = bucket.get(month_key, 0.0) + amt
                    if is_q1:
                        bucket[f"{prefix}_amt"] += amt
                        if year_bucket == 115:
                            matched_any_115_fee = True
                    bucket[f"{prefix}_amt_total"] += amt
            if scan.month is not None:
                if matched_any_115_count:
                    seen_115_count_months.add(scan.month)
                if matched_any_115_fee:
                    seen_115_fee_months.add(scan.month)

        self.claim_fee_months_115 = sorted(seen_115_fee_months)
        fee_month_count_115 = float(len(seen_115_fee_months))
        for bucket in out.values():
            bucket["115_months"] = fee_month_count_115
        return out, sorted(seen_115_count_months)


def process_excel(source_path: str, template_path: Optional[str] = None) -> str:
    profile = FangdingProfile()
    template = template_path or GENERIC._find_template(str(SCRIPT_DIR))
    if not os.path.exists(template):
        raise ValueError(f"找不到模板檔：{template}")

    original_populate_doctor_sheet = GENERIC.populate_doctor_sheet
    original_months_summary = GENERIC._format_115_months_summary
    original_supplement_claim_counts = GENERIC._supplement_claim_counts_from_hisb

    def supplement_claim_counts_with_fee_months(*args: Any, **kwargs: Any):
        claim_sums, count_months = original_supplement_claim_counts(*args, **kwargs)
        fee_month_count = float(len(profile.claim_fee_months_115))
        for bucket in claim_sums.values():
            bucket["115_months"] = fee_month_count
        return claim_sums, count_months

    def populate_doctor_sheet_with_fee_months(*args: Any, **kwargs: Any) -> None:
        original_populate_doctor_sheet(*args, **kwargs)
        wb_tpl = args[0]
        fee_months = profile.claim_fee_months_115
        if GENERIC.DOCTOR_SHEET_NAME not in wb_tpl.sheetnames:
            return
        ws_doc = wb_tpl[GENERIC.DOCTOR_SHEET_NAME]
        if fee_months:
            month_text = GENERIC._format_month_span(fee_months)
            ws_doc["O2"] = (
                f"115年{month_text}\n月平均費用\n"
                f"（實際費用月份總額 ÷ {len(fee_months)}）"
            )
        else:
            ws_doc["O2"] = "115年月平均費用\n（未偵測到費用月份）"

    def format_months_summary(count_months: List[int]) -> str:
        count_text = (
            f"{GENERIC._format_month_span(count_months)}，共{len(count_months)}個月"
            if count_months
            else "未偵測到月份資料"
        )
        fee_months = profile.claim_fee_months_115
        fee_text = (
            f"{GENERIC._format_month_span(fee_months)}，共{len(fee_months)}個月"
            if fee_months
            else "未偵測到費用月份"
        )
        return (
            f"115年次數月份：{count_text}；醫生看 L/M 欄依次數月份統計。"
            f"115年費用月份：{fee_text}；醫生看 O 欄依費用月份計算月平均。"
        )

    GENERIC.populate_doctor_sheet = populate_doctor_sheet_with_fee_months
    GENERIC._format_115_months_summary = format_months_summary
    GENERIC._supplement_claim_counts_from_hisb = supplement_claim_counts_with_fee_months
    try:
        return GENERIC.process_excel(source_path, template, profile=profile)
    finally:
        GENERIC.populate_doctor_sheet = original_populate_doctor_sheet
        GENERIC._format_115_months_summary = original_months_summary
        GENERIC._supplement_claim_counts_from_hisb = original_supplement_claim_counts


def main(argv: Optional[List[str]] = None) -> int:
    import tkinter as tk
    from tkinter import filedialog, messagebox

    parser = argparse.ArgumentParser(description="方鼎前置清洗與選會員輸出")
    parser.add_argument("source_dir", nargs="?", help="來源資料夾；未指定時開啟選擇視窗")
    parser.add_argument("--no-open", action="store_true", help="完成後不自動開啟 Excel")
    parser.add_argument("--no-dialog", action="store_true", help="完成後不顯示完成視窗")
    args = parser.parse_args(argv)

    root = None
    src = args.source_dir
    if not src:
        root = tk.Tk()
        root.withdraw()
        src = filedialog.askdirectory(title="選擇方鼎來源資料夾")
    if not src:
        if root is not None:
            root.destroy()
        print("未選擇來源資料夾，程式已取消。")
        return 0
    src = str(Path(src).expanduser().resolve())
    if not os.path.isdir(src):
        raise FileNotFoundError(f"找不到來源資料夾：{src}")

    template = GENERIC._find_template(str(SCRIPT_DIR))

    try:
        out = process_excel(src, template)
        print(f"完成：{out}")
        if not args.no_dialog:
            if root is None:
                root = tk.Tk()
                root.withdraw()
            messagebox.showinfo("完成", f"已輸出：\n{out}", parent=root)
        if not args.no_open:
            GENERIC.open_file_cross_platform(out)
    except Exception as e:
        print(f"錯誤：{e}")
        if not args.no_dialog:
            if root is None:
                root = tk.Tk()
                root.withdraw()
            messagebox.showerror("錯誤", str(e), parent=root)
        return 2
    finally:
        if root is not None:
            root.destroy()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
