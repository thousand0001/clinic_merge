# -*- coding: utf-8 -*-
"""
書田前置清洗 + 通用主程式包裝

書田資料大多可直接交給通用版。
此入口保留書田 R11440 月報的月份判定差異：
月分頁中的日期欄常是會員最後一次看診日，不一定屬於該月，
因此月份歸屬優先以分頁名稱判定。
"""

from __future__ import annotations

import importlib.util
import os
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple


SCRIPT_DIR = Path(__file__).resolve().parent


def _find_generic_script(script_dir: Path) -> Path:
    candidates = sorted(script_dir.glob("run_merge_通用_*.py"), reverse=True)
    if not candidates:
        raise RuntimeError("找不到通用程式 run_merge_通用_*.py")
    return candidates[0]


GENERIC_SCRIPT = _find_generic_script(SCRIPT_DIR)


def _load_generic_module():
    spec = importlib.util.spec_from_file_location("run_merge_generic", GENERIC_SCRIPT)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"無法載入通用程式：{GENERIC_SCRIPT}")
    mod = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = mod
    spec.loader.exec_module(mod)
    return mod


GENERIC = _load_generic_module()


class ShutianProfile(GENERIC.ProcessingProfile):
    def collect_monthly_claim_summaries(
        self,
        wb_src: Any,
        monthly_scans: Optional[Dict[str, Any]] = None,
    ) -> Tuple[Dict[str, Dict[str, float]], List[int]]:
        out: Dict[str, Dict[str, float]] = {}
        seen_115_months: set = set()
        scans = monthly_scans or {
            sheet_name: scan
            for sheet_name in wb_src.sheetnames
            if (scan := self.scan_monthly_claim_sheet(sheet_name, wb_src[sheet_name])) is not None
        }

        for sheet_name, scan in scans.items():
            sh = wb_src[sheet_name]
            sheet_year_bucket = self.sheet_year_bucket(sheet_name)
            sheet_month = self.sheet_month(sheet_name)
            for r in range(scan.header_row + 1, sh.max_row + 1):
                pid_raw = sh.cell(r, scan.id_col).value if scan.id_col else None
                pid = GENERIC.normalize_text(pid_raw).upper()
                if not pid or not GENERIC.is_valid_tw_id(pid):
                    continue

                dt = self.parse_date(sh.cell(r, scan.date_col).value) if scan.date_col else None
                cnt = GENERIC.parse_float(sh.cell(r, scan.count_col).value) if scan.count_col else None
                amt = GENERIC.parse_float(sh.cell(r, scan.amount_col).value) if scan.amount_col else None
                if dt is None and cnt is None and amt is None:
                    continue

                year_bucket = (
                    sheet_year_bucket
                    if sheet_year_bucket in (114, 115)
                    else GENERIC._roc_year_from_date(dt)
                )
                if year_bucket not in (114, 115):
                    continue

                month = sheet_month if sheet_month is not None else (dt.month if dt is not None else None)
                if month is None:
                    continue

                is_q1 = month <= 4
                if year_bucket == 115 and is_q1:
                    seen_115_months.add(month)

                bucket = out.setdefault(pid, GENERIC._empty_claim_bucket())
                if dt is not None:
                    bucket["last_visit_ord"] = max(bucket.get("last_visit_ord", 0.0), float(dt.toordinal()))

                prefix = str(year_bucket)
                if cnt is not None:
                    if year_bucket == 115:
                        bucket["115_cnt"] += cnt
                        if is_q1:
                            bucket["115_cnt_q1"] += cnt
                    elif is_q1:
                        bucket[f"{prefix}_cnt"] += cnt
                    if year_bucket == 114:
                        bucket["114_cnt_full"] += cnt
                if amt is not None:
                    if is_q1:
                        bucket[f"{prefix}_amt"] += amt
                    bucket[f"{prefix}_amt_total"] += amt

        month_count_115 = float(len(seen_115_months))
        for bucket in out.values():
            bucket["115_months"] = month_count_115
        return out, sorted(seen_115_months)


def process_excel(source_path: str, template_path: Optional[str] = None) -> str:
    template = template_path or GENERIC._find_template(str(SCRIPT_DIR))
    if not os.path.exists(template):
        raise ValueError(f"找不到模板檔：{template}")
    return GENERIC.process_excel(source_path, template, profile=ShutianProfile())


def main() -> None:
    import tkinter as tk
    from tkinter import filedialog, messagebox

    root = tk.Tk()
    root.withdraw()

    src = filedialog.askdirectory(title="選擇書田來源資料夾")
    if not src:
        return

    template = GENERIC._find_template(str(SCRIPT_DIR))

    try:
        out = process_excel(src, template)
        messagebox.showinfo("完成", f"已輸出：\n{out}")
        GENERIC.open_file_cross_platform(out)
    except Exception as e:
        messagebox.showerror("錯誤", str(e))


if __name__ == "__main__":
    main()
