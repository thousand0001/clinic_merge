from __future__ import annotations

import datetime as dt
import hashlib
import json
import math
import re
from decimal import Decimal, InvalidOperation
from typing import Any, Iterable, Optional


MISSING_TEXT = {"", "-", "—", "–", "none", "nan", "null"}


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    text = str(value).strip()
    return "" if text.lower() in MISSING_TEXT else text


def normalize_id(value: Any) -> str:
    text = normalize_text(value).upper().replace(" ", "")
    if text.endswith(".0"):
        text = text[:-2]
    return text


def normalize_name(value: Any) -> str:
    return re.sub(r"\s+", "", normalize_text(value))


def normalize_phone(value: Any) -> str:
    digits = re.sub(r"\D", "", normalize_text(value))
    if digits and not digits.startswith("0"):
        digits = "0" + digits
    return digits


def parse_decimal(value: Any) -> Decimal:
    text = normalize_text(value).replace(",", "")
    if not text:
        return Decimal("0")
    try:
        return Decimal(text)
    except InvalidOperation:
        return Decimal("0")


def parse_date(value: Any) -> Optional[dt.date]:
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value

    text = normalize_text(value)
    if not text:
        return None
    text = text.split(" ")[0]
    digits = re.sub(r"\D", "", text)

    year = month = day = None
    parts = [part for part in re.split(r"[./-]", text) if part]
    if len(parts) == 3:
        try:
            year, month, day = (int(part) for part in parts)
        except ValueError:
            return None
    elif len(digits) == 7:
        year, month, day = int(digits[:3]), int(digits[3:5]), int(digits[5:7])
    elif len(digits) == 8:
        year, month, day = int(digits[:4]), int(digits[4:6]), int(digits[6:8])
    else:
        return None

    if year < 1911:
        year += 1911
    try:
        return dt.date(year, month, day)
    except ValueError:
        return None


def stable_row_hash(values: Iterable[Any]) -> str:
    payload = [normalize_text(value) for value in values]
    raw = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


# NHI 分區業務組別代碼：1=臺北 2=北區 3=中區 4=南區 5=高屏 6=東區
_BRANCH_PREFIXES: list[tuple[int, tuple[str, ...]]] = [
    (1, ("臺北市", "台北市")),
    (2, ("新北市", "基隆市", "宜蘭縣", "宜蘭市",
         "桃園市", "桃園縣", "新竹市", "新竹縣", "苗栗縣")),
    (3, ("臺中市", "台中市", "台中縣", "彰化縣", "南投縣")),
    (4, ("嘉義市", "嘉義縣", "雲林縣", "臺南市", "台南市", "台南縣")),
    (5, ("高雄市", "高雄縣", "屏東縣")),
    (6, ("花蓮縣", "臺東縣", "台東縣", "澎湖縣", "金門縣", "連江縣")),
]


def branch_from_address(address: str) -> Optional[int]:
    """由診所地址推算 NHI 分區業務組別代碼（1–6），無法判斷回傳 None。"""
    if not address:
        return None
    for code, prefixes in _BRANCH_PREFIXES:
        if any(address.startswith(p) for p in prefixes):
            return code
    return None

