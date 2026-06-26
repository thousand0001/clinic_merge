from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Set


SUPPORTED_EXTENSIONS = {
    ".xlsx",
    ".xlsm",
    ".xls",
    ".csv",
    ".ods",
    ".txt",
    ".pdf",
}


ID_ALIASES = (
    "ID",
    "身分證",
    "身分證號",
    "身分證號碼",
    "身分證字號",
    "身份證",
    "身份證號",
    "身份證號碼",
    "身份證字號",
    "家醫收案會員ID",
    "會員ID",
)
NAME_ALIASES = ("姓名", "會員姓名", "病患姓名", "名字")
BIRTHDAY_ALIASES = ("生日", "出生日期", "出生年月日", "BIRTHDAY")
DATE_ALIASES = ("日期", "看診日期", "看診日", "就診日期", "就診日", "就醫日", "最後篩檢日期", "最後就診日")
AMOUNT_ALIASES = ("申請金額", "申請額", "總額", "費用", "申報總金額", "掛帳費", "藥費")
COUNT_ALIASES = ("次數", "看診次數", "就診次數")
CHART_ALIASES = ("病歷號", "病歷號碼", "掛號證", "掛號証")


FIELD_ALIASES: Dict[str, Sequence[str]] = {
    "id": ID_ALIASES,
    "name": NAME_ALIASES,
    "birthday": BIRTHDAY_ALIASES,
    "date": DATE_ALIASES,
    "amount": AMOUNT_ALIASES,
    "count": COUNT_ALIASES,
    "chart_no": CHART_ALIASES,
    "indicator": ("指標名稱", "指標"),
    "disease": ("疾病樣態", "疾病樣態分類", "DM/CKD/DKD/ASCVD", "ASCVD"),
    "p4p_plan": ("P4P收案計畫",),
    "p4p_status": ("P4P收案狀態", "收案狀態"),
    "hba1c": ("最近一次HbA1c檢查結果(%)", "HbA1c"),
    "hba1c_date": ("最近一次HbA1c檢查日期", "HbA1c檢查日期"),
    "ldl": ("最近一次LDL檢查結果(mg/dL)", "最近一次LDL檢查結果", "LDL"),
    "ldl_date": ("最近一次LDL檢查日期", "LDL檢查日期"),
    "uacr": ("最近一次UACR檢查結果(mg/gm)", "最近一次UACR檢查結果", "UACR"),
    "uacr_date": ("最近一次UACR檢查日期", "UACR檢查日期"),
}


@dataclass
class SheetRows:
    file_path: Path
    sheet_name: str
    rows: List[List[Any]]


@dataclass
class SourceFinding:
    file_path: Path
    extension: str
    sheet_name: str = ""
    data_type: str = "unknown"
    status: str = "ok"
    reason: str = ""
    header_row: Optional[int] = None
    row_count: int = 0
    data_row_count: int = 0
    valid_id_rows: int = 0
    unique_id_count: int = 0
    date_rows: int = 0
    amount_rows: int = 0
    count_rows: int = 0
    columns: Dict[str, str] = field(default_factory=dict)
    matched_output_ids: Optional[int] = None
    missing_output_ids: Optional[int] = None
    examples: List[str] = field(default_factory=list)
    source_ids: Set[str] = field(default_factory=set, repr=False)

    def relative_file(self, root: Path) -> str:
        try:
            return str(self.file_path.relative_to(root))
        except ValueError:
            return str(self.file_path)
