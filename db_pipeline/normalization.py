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
        year, month, day = (int(part) for part in parts)
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

