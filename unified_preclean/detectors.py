from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, Optional, Sequence

from .schema import FIELD_ALIASES, SourceFinding, SheetRows
from .utils import compact, find_col, find_header_row, is_valid_id, parse_date, parse_number


def _filename_token(path: Path) -> str:
    return re.sub(r"[\s\-_()（）\[\]{}]+", "", path.stem).lower()


def _classify_by_name(path: Path, sheet_name: str) -> str:
    token = _filename_token(path)
    sheet = compact(sheet_name).lower()
    combined = token + sheet
    if any(key in sheet for key in ("百分位", "計算", "統計")):
        return "non_data_report"
    if "115x" in token or "不選" in token or "不要" in token:
        return "exclude_select"
    if ("會員" in token or "選" in token) and any(key in token for key in ("自選", "預選", "要選", "a115")):
        return "self_select"
    if any(key in combined for key in ("healthcase", "健康管理", "個案健康管理")):
        return "healthcase"
    if any(key in combined for key in ("p4p收案", "p4pcase", "收案管理")):
        return "p4p_enroll"
    if any(key in combined for key in ("p4p追蹤", "p4ptrack", "追蹤管理")):
        return "p4p_track"
    if any(key in combined for key in ("需照護名單", "需要照護名單", "家醫計畫", "ascvd", "三高")):
        return "member_roster"
    if re.fullmatch(r"1(?:14|15)\d{2}", path.stem) or re.fullmatch(r"1(?:14|15)\d{2}", sheet_name):
        return "monthly_claim"
    if any(key in combined for key in ("r11440", "門診診療次數", "次數月報")):
        return "monthly_claim"
    if any(key in token for key in ("成健", "成人預防保健", "成人健檢")):
        return "screening_adult"
    if any(key in token for key in ("bc肝", "b肝c肝", "肝炎")):
        return "screening_hep"
    if any(key in token for key in ("子抹", "抹片")):
        return "screening_pap"
    if any(key in token for key in ("糞篩", "糞便潛血", "潛血")):
        return "screening_fit"
    if any(key in token for key in ("老流", "老人流感", "流感")):
        return "screening_flu"
    if any(key in combined for key in ("主次代碼", "診斷代碼", "主次診斷")):
        return "diagnosis"
    if any(key in token for key in ("費用", "收入", "掛帳")):
        return "claim_fee"
    if "次數" in token:
        return "claim_count"
    return "unknown"


def _find_best_header(rows: Sequence[Sequence[object]], guessed_type: str) -> Optional[int]:
    if guessed_type.startswith("screening_"):
        return (
            find_header_row(rows, [FIELD_ALIASES["id"], FIELD_ALIASES["date"]], 30)
            or find_header_row(rows, [FIELD_ALIASES["id"], ("最後篩檢日期",)], 30)
            or find_header_row(rows, [FIELD_ALIASES["indicator"], FIELD_ALIASES["id"]], 30)
        )
    if guessed_type in {"monthly_claim", "claim_fee"}:
        return (
            find_header_row(rows, [FIELD_ALIASES["id"], FIELD_ALIASES["date"], FIELD_ALIASES["amount"]], 30)
            or find_header_row(rows, [FIELD_ALIASES["id"], FIELD_ALIASES["date"]], 30)
            or find_header_row(rows, [FIELD_ALIASES["chart_no"], FIELD_ALIASES["date"], FIELD_ALIASES["amount"]], 30)
        )
    if guessed_type == "claim_count":
        return find_header_row(rows, [FIELD_ALIASES["id"], FIELD_ALIASES["count"]], 30) or find_header_row(
            rows, [FIELD_ALIASES["name"], FIELD_ALIASES["count"]], 30
        )
    if guessed_type in {"self_select", "exclude_select", "member_roster"}:
        return find_header_row(rows, [FIELD_ALIASES["id"]], 30) or find_header_row(
            rows, [FIELD_ALIASES["name"], FIELD_ALIASES["birthday"]], 30
        )
    if guessed_type == "healthcase":
        return find_header_row(rows, [FIELD_ALIASES["id"], FIELD_ALIASES["hba1c"]], 30) or find_header_row(
            rows, [FIELD_ALIASES["id"], FIELD_ALIASES["ldl"]], 30
        )
    if guessed_type.startswith("p4p_"):
        return find_header_row(rows, [FIELD_ALIASES["id"], FIELD_ALIASES["p4p_plan"]], 30)
    if guessed_type == "diagnosis":
        return find_header_row(rows, [FIELD_ALIASES["id"]], 30)
    return find_header_row(rows, [FIELD_ALIASES["id"]], 30)


def detect_sheet(sheet: SheetRows) -> SourceFinding:
    ext = sheet.file_path.suffix.lower()
    finding = SourceFinding(
        file_path=sheet.file_path,
        extension=ext,
        sheet_name=sheet.sheet_name,
        row_count=len(sheet.rows),
    )
    if ext == ".pdf":
        finding.data_type = "pdf_text"
        finding.data_row_count = len(sheet.rows)
        finding.reason = "PDF first-stage text inventory only"
        return finding
    if not sheet.rows:
        finding.status = "skipped"
        finding.reason = "empty sheet"
        return finding

    guessed_type = _classify_by_name(sheet.file_path, sheet.sheet_name)
    if guessed_type == "non_data_report":
        finding.data_type = guessed_type
        finding.status = "skipped"
        finding.reason = "statistics/report sheet, not a source table"
        finding.data_row_count = max(len(sheet.rows) - 1, 0)
        return finding
    header_index = _find_best_header(sheet.rows, guessed_type)
    if header_index is None:
        finding.data_type = guessed_type
        finding.status = "unparsed"
        finding.reason = "header not detected"
        finding.data_row_count = max(len(sheet.rows) - 1, 0)
        return finding

    header = sheet.rows[header_index]
    columns: Dict[str, int] = {}
    for field, aliases in FIELD_ALIASES.items():
        col = find_col(header, aliases)
        if col is not None:
            columns[field] = col
            finding.columns[field] = str(header[col])

    if guessed_type == "unknown":
        guessed_type = _classify_by_columns(columns)
    finding.data_type = guessed_type
    finding.header_row = header_index + 1

    unique_ids = set()
    examples = []
    for row in sheet.rows[header_index + 1 :]:
        if not any(value not in (None, "") for value in row):
            continue
        finding.data_row_count += 1
        pid = _cell(row, columns.get("id"))
        if is_valid_id(pid):
            finding.valid_id_rows += 1
            unique_ids.add(str(pid).strip().upper())
        if parse_date(_cell(row, columns.get("date"))) is not None:
            finding.date_rows += 1
        if parse_number(_cell(row, columns.get("amount"))) is not None:
            finding.amount_rows += 1
        if parse_number(_cell(row, columns.get("count"))) is not None:
            finding.count_rows += 1
        if len(examples) < 3:
            name = _cell(row, columns.get("name"))
            if pid or name:
                examples.append("/".join(str(part) for part in (pid, name) if part not in (None, "")))
    finding.unique_id_count = len(unique_ids)
    finding.source_ids = unique_ids
    finding.examples = examples
    if finding.data_row_count and not finding.valid_id_rows and "id" in columns:
        finding.status = "warning"
        finding.reason = "id column found but no valid IDs"
    return finding


def _classify_by_columns(columns: Dict[str, int]) -> str:
    if {"hba1c", "ldl"} & set(columns):
        return "healthcase"
    if "p4p_plan" in columns:
        return "p4p"
    if "indicator" in columns and "date" in columns:
        return "screening"
    if "amount" in columns and "date" in columns:
        return "monthly_claim"
    if "count" in columns:
        return "claim_count"
    if "disease" in columns:
        return "member_roster"
    if "id" in columns:
        return "id_list"
    return "unknown"


def _cell(row: Sequence[object], index: Optional[int]) -> object:
    if index is None or index >= len(row):
        return None
    return row[index]
