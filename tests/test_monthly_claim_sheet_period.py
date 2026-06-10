import datetime
import importlib.util
import sys
from pathlib import Path

from openpyxl import Workbook


MODULE_PATH = Path(__file__).resolve().parents[1] / "選會員_共用核心_0610.py"
SPEC = importlib.util.spec_from_file_location("member_core_0610", MODULE_PATH)
assert SPEC is not None and SPEC.loader is not None
CORE = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = CORE
SPEC.loader.exec_module(CORE)


def test_month_sheet_name_has_priority_over_last_visit_date():
    wb = Workbook()
    ws = wb.active
    ws.title = "11401"
    ws.append(["姓名", "身分證號", "日期", "次數", "總額"])
    ws.append(["測試者", "A123456789", datetime.date(2026, 3, 5), 2, 300])

    claim_sums, months_115 = CORE.collect_monthly_claim_summaries(wb)

    assert months_115 == []
    assert claim_sums["A123456789"]["114_cnt_full"] == 2
    assert claim_sums["A123456789"]["114_amt_total"] == 300
    assert claim_sums["A123456789"]["115_cnt"] == 0
    assert claim_sums["A123456789"]["115_amt_total"] == 0
