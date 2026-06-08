# -*- coding: utf-8 -*-
"""
DatasetBundle → member dict

`build_from_bundle()` 等同舊流程 `build_members()`，
但來源改為 DatasetBundle（可來自 FolderDataProvider 或 PostgresDataProvider）。

產出的 member dict key 與舊流程 `write_output()` 完全相容：
- AW:BJ（ws_doc col 49-62）：中文 key，直接對應 `designated_fields` 列表
- L/M/N/O（ws_doc col 12-15）：114_count_q1 / 115_count / 114_avg_amount / 115_avg_amount
"""
from __future__ import annotations

import datetime as dt
from collections import defaultdict
from typing import Any, Dict, Optional

from db_pipeline.datasets.models import DatasetBundle

# ── 疾病樣態映射（與舊流程 disease_code_text 相同） ──────────────────────────
_DISEASE_MAP = {
    "1": "DM", "2": "HTN", "3": "CKD", "4": "DKD",
    "5": "HLP", "6": "ASCVD",
}


def _disease_code_text(value: Any) -> str:
    text = str(value).strip() if value else ""
    return _DISEASE_MAP.get(text, "None")


def _disease_class_text(disease_code: str, ascvd_value: Any) -> str:
    ascvd = str(ascvd_value).strip().lower() if ascvd_value else ""
    if disease_code in {"DM", "DKD"}:
        return "糖尿病"
    if disease_code == "CKD":
        return "慢性腎臟病"
    if disease_code in {"HTN", "HLP"}:
        return "高血壓/高血脂"
    if ascvd and ascvd not in {"0", ""}:
        return "ASCVD"
    return ""


def _norm(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        return text[:-2]
    return text


def _merge_if_empty(base: Dict[str, Any], updates: Dict[str, Any]) -> None:
    for key, value in updates.items():
        if value not in (None, "") and not base.get(key):
            base[key] = value


# ── 公開函式 ──────────────────────────────────────────────────────────────────
def build_from_bundle(bundle: DatasetBundle) -> Dict[str, Dict[str, Any]]:
    """
    DatasetBundle → Dict[person_id, member_dict]

    member_dict key 對應 `資料庫輸出0601.py` 的 write_output() 所用欄位。
    """
    members: Dict[str, Dict[str, Any]] = {}

    # ── A. 照護名單（成員 + AW:BJ 欄位） ────────────────────────────────────
    for rec in bundle.members:
        pid = rec.person_id
        m = members.setdefault(pid, {"id": pid})
        _merge_if_empty(m, {
            "name":          rec.name,
            "birth":         rec.birth_date.isoformat() if rec.birth_date else "",
            "sex":           rec.sex,
            "phone":         rec.phone,
            "mobile":        rec.mobile,
            "address":       rec.address,
            # AW:BJ — 中文 key 供 write_output() 直接取用
            "個案類別":                    _norm(rec.case_category),
            "論質名單":                    _norm(rec.quality_roster),
            "65歲以上多重慢性病註記":      _norm(rec.multi_chronic_65),
            "高診次註記":                  _norm(rec.high_visit),
            "慢性病註記":                  _norm(rec.chronic_mark),
            "非慢性病註記":                _norm(rec.non_chronic_mark),
            "與前一年家醫收案診所相同":    _norm(rec.same_clinic_previous_year),
            "疾病樣態":                    _norm(rec.disease_pattern),
            "ASCVD":                       _norm(rec.ascvd),
            "三高":                        _norm(rec.three_highs),
            "高血壓":                      _norm(rec.hypertension),
            "高血脂":                      _norm(rec.hyperlipidemia),
            "高血糖":                      _norm(rec.hyperglycemia),
        })
        # disease_code / disease_class 由 disease_pattern 推算（只在尚未設定時寫入）
        if not m.get("disease_code"):
            disease_code = _disease_code_text(rec.disease_pattern)
            m["disease_code"] = disease_code
            ascvd_val = rec.ascvd if disease_code == "None" else "1"
            m["ASCVD"] = _norm(rec.ascvd) if disease_code == "None" else "1"
            m["disease_class"] = _disease_class_text(disease_code, ascvd_val)

    # ── B. 每月就診（L/M/N/O） ────────────────────────────────────────────────
    claims_agg: Dict[str, dict] = defaultdict(lambda: {
        "114_count_q1":     0.0,
        "114_count_full":   0.0,
        "115_count":        0.0,
        "114_amount_total": 0.0,
        "115_amount_q1":    0.0,
        "115_amount_total": 0.0,
        "last_visit":       "",
    })
    claim_months_115: set = set()

    for rec in bundle.monthly_claims:
        pid = rec.person_id
        agg = claims_agg[pid]
        count  = float(rec.visit_count)
        amount = float(rec.amount)
        if rec.roc_year == 114:
            agg["114_count_full"]   += count
            agg["114_amount_total"] += amount
            if rec.month <= 4:
                agg["114_count_q1"] += count
        elif rec.roc_year == 115:
            claim_months_115.add(rec.month)
            agg["115_count"]        += count
            agg["115_amount_total"] += amount
            if rec.month <= 4:
                agg["115_amount_q1"] += amount
        # last_visit
        if rec.last_visit_date:
            iso = rec.last_visit_date.isoformat()
            if iso > agg["last_visit"]:
                agg["last_visit"] = iso

    month_count_115 = max(len(claim_months_115), 1)
    for pid, sums in claims_agg.items():
        m = members.setdefault(pid, {"id": pid})
        m["114_count_q1"]   = sums["114_count_q1"] or None      # L
        m["114_count"]      = sums["114_count_full"] or None
        m["114_amount"]     = sums["114_amount_total"] or None
        m["115_count"]      = sums["115_count"] or None          # M
        m["115_amount"]     = sums["115_amount_total"] or None
        m["114_avg_amount"] = (                                   # N
            round(sums["114_amount_total"] / 12, 2)
            if sums["114_amount_total"] else None
        )
        m["115_avg_amount"] = (                                   # O
            round(sums["115_amount_q1"] / month_count_115, 2)
            if sums["115_amount_q1"] else None
        )
        if sums["last_visit"] and not m.get("last_visit"):
            m["last_visit"] = sums["last_visit"]

    # ── C. 名單旗標（designated / self_select / exclude_select） ─────────────
    for rec in bundle.member_selections:
        pid = rec.person_id
        m = members.setdefault(pid, {"id": pid})
        if rec.selection_type == "designated_114":
            m["designated"]    = "✔"
            m["is_114_member"] = "✔"
        elif rec.selection_type == "self_select":
            m["self_select"] = "✔"
        elif rec.selection_type == "exclude_select":
            m["exclude_select"] = "✔"

    # ── D. 檢驗結果（HbA1c / LDL / UACR） ────────────────────────────────────
    _LAB_KEYS: Dict[str, tuple] = {
        "HbA1c": ("hba1c",  "hba1c_date"),
        "LDL":   ("ldl",    "ldl_date"),
        "UACR":  ("uacr",   "uacr_date"),
    }
    for rec in bundle.lab_results:
        pid = rec.person_id
        m = members.setdefault(pid, {"id": pid})
        keys = _LAB_KEYS.get(rec.test_code)
        if keys:
            val_key, date_key = keys
            _merge_if_empty(m, {
                val_key:  rec.result_value or None,
                date_key: rec.tested_at,
            })

    # ── E. 篩檢（成人健檢 / 子宮抹片 / 老人流感 / 糞便潛血 / 肝炎篩檢） ──────
    _SCREENING_KEYS: Dict[str, str] = {
        "成人健檢": "adult",
        "子宮抹片": "pap",
        "老人流感": "flu",
        "糞便潛血": "fit",
        "肝炎篩檢": "bc",
    }
    for rec in bundle.screenings:
        pid = rec.person_id
        m = members.setdefault(pid, {"id": pid})
        key = _SCREENING_KEYS.get(rec.screening_type)
        if key:
            _merge_if_empty(m, {key: rec.screened_at})

    # ── F. P4P 收案 ────────────────────────────────────────────────────────────
    for rec in bundle.p4p_cases:
        pid = rec.person_id
        m = members.setdefault(pid, {"id": pid})
        _merge_if_empty(m, {
            "p4p_plan":         rec.plan,
            "p4p_status":       rec.status,
            "p4p_enroll_date":  rec.enrolled_at,
        })

    # ── G. P4P 追蹤 ───────────────────────────────────────────────────────────
    for rec in bundle.p4p_tracks:
        pid = rec.person_id
        m = members.setdefault(pid, {"id": pid})
        _merge_if_empty(m, {
            "p4p_last_track": rec.last_tracked_at,
            "p4p_next_track": rec.next_track_at,
            "p4p_overdue":    rec.overdue or None,
        })

    # ── 補齊 sex（無來源時從身分證推算） ────────────────────────────────────────
    for pid, m in members.items():
        if not m.get("sex") and len(pid) >= 2:
            digit = pid[1:2]
            m["sex"] = "M" if digit == "1" else ("F" if digit == "2" else "")
        if m.get("disease_code") and not m.get("disease_class"):
            m["disease_class"] = _disease_class_text(
                m["disease_code"], m.get("ASCVD", "")
            )

    return members
