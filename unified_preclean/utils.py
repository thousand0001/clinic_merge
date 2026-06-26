from __future__ import annotations

import datetime as dt
import math
import re
from typing import Any, Iterable, Optional, Sequence


ID_RE = re.compile(r"(?:[A-Z][1289]\d{8}|[A-Z][A-D]\d{8})")


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return text


def compact(value: Any) -> str:
    return re.sub(r"[\s　\-_()（）\[\]{}:：]+", "", normalize_text(value)).upper()


def normalize_id(value: Any) -> str:
    text = normalize_text(value).upper().strip("'")
    return re.sub(r"\s+", "", text)


def is_valid_id(value: Any) -> bool:
    return bool(ID_RE.fullmatch(normalize_id(value)))


def parse_number(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)
    text = normalize_text(value).replace(",", "")
    if text in {"", "-", "—", "–", "None", "nan"}:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def parse_date(value: Any) -> Optional[dt.date]:
    if value is None:
        return None
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    text = normalize_text(value)
    if text in {"", "-", "—", "–", "None", "nan"}:
        return None
    text = text.split()[0]
    match = re.fullmatch(r"(\d{2,4})[./\-](\d{1,2})[./\-](\d{1,2})", text)
    if match:
        year = int(match.group(1))
        if year < 1911:
            year += 1911
        try:
            return dt.date(year, int(match.group(2)), int(match.group(3)))
        except ValueError:
            return None
    digits = re.sub(r"\D", "", text)
    try:
        if len(digits) == 7:
            return dt.date(int(digits[:3]) + 1911, int(digits[3:5]), int(digits[5:7]))
        if len(digits) == 8:
            year = int(digits[:4])
            if year < 1911:
                year += 1911
            return dt.date(year, int(digits[4:6]), int(digits[6:8]))
        if len(digits) == 6:
            return dt.date(int(digits[:2]) + 1911, int(digits[2:4]), int(digits[4:6]))
    except ValueError:
        return None
    return None


def find_col(header: Sequence[Any], aliases: Iterable[str]) -> Optional[int]:
    compact_header = [compact(value) for value in header]
    targets = {compact(alias) for alias in aliases}
    for index, value in enumerate(compact_header):
        if value in targets:
            return index
    return None


def find_header_row(rows: Sequence[Sequence[Any]], required_groups: Sequence[Iterable[str]], scan_rows: int = 30) -> Optional[int]:
    for row_index, row in enumerate(rows[:scan_rows]):
        if all(find_col(row, group) is not None for group in required_groups):
            return row_index
    return None

