# -*- coding: utf-8 -*-
"""
PostgreSQL staging 寫入層

DatasetBundle → staging.* 單一 transaction

流程：
1. 查 meta.clinics 取得 clinic_id
2. 建立 meta.import_batches 批次記錄
3. 刪除同 batch_id + clinic_id 的舊 staging 資料（冪等）
4. COPY FROM stdin 批次寫入各 staging 表
5. 更新 batch status → 'validated'
6. 失敗時 ROLLBACK，batch status → 'failed'

使用 psql subprocess（與現有 資料庫輸出0601.py 相同模式）。
不依賴 psycopg2，不寫死路徑。
"""
from __future__ import annotations

import csv
import io
import json
import os
import shutil
import subprocess
from decimal import Decimal, InvalidOperation
from typing import Any, List, Optional, Sequence

from db_pipeline.datasets.models import (
    DatasetBundle,
    LabResultRecord,
    MemberRecord,
    MemberSelectionRecord,
    MonthlyClaimRecord,
    P4PCaseRecord,
    P4PTrackRecord,
    ScreeningRecord,
)
from db_pipeline.storage.contracts import StageResult
from db_pipeline.validation.models import ValidationReport


# ── psql 路徑偵測（跨平台） ───────────────────────────────────────────────────
def _find_psql() -> str:
    env_val = os.getenv("CLINIC_PSQL")
    if env_val:
        return env_val
    found = shutil.which("psql")
    if found:
        return found
    win_default = r"C:\Program Files\PostgreSQL\16\bin\psql.exe"
    if os.path.isfile(win_default):
        return win_default
    return "psql"


def _db_env() -> dict:
    env = os.environ.copy()
    env.setdefault("PGPASSWORD", os.getenv("CLINIC_DB_PASSWORD", ""))
    return env


def _psql_args() -> List[str]:
    return [
        _find_psql(),
        "-h", os.getenv("CLINIC_DB_HOST", "localhost"),
        "-p", os.getenv("CLINIC_DB_PORT", "5432"),
        "-U", os.getenv("CLINIC_DB_USER", "thousand0001"),
        "-d", os.getenv("CLINIC_DB_NAME", "clinic_merge"),
        "-v", "ON_ERROR_STOP=1",
    ]


def _run_sql(sql: str) -> str:
    """執行 SQL（含 COPY FROM stdin），失敗時拋 RuntimeError。"""
    result = subprocess.run(
        _psql_args(),
        input=sql,
        text=True,
        capture_output=True,
        env=_db_env(),
    )
    if result.returncode != 0:
        raise RuntimeError(
            f"psql 執行失敗：\n{result.stderr.strip() or result.stdout.strip()}"
        )
    return result.stdout


def _run_query(sql: str) -> str:
    """執行單行查詢並回傳 stdout（-t -A 模式）。"""
    result = subprocess.run(
        _psql_args() + ["-t", "-A", "-c", sql],
        text=True,
        capture_output=True,
        env=_db_env(),
    )
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip())
    return result.stdout.strip()


# ── CSV 產生工具 ──────────────────────────────────────────────────────────────
def _clean_str(v: Any) -> str:
    """轉字串並移除可能讓 COPY 出錯的換行與控制字元。"""
    s = str(v)
    # 將換行符號換成空格（Excel 欄位可能含 \n）
    s = s.replace("\r\n", " ").replace("\r", " ").replace("\n", " ")
    return s


def _csv_block(rows: List[List[Any]]) -> str:
    """rows → CSV 字串（無表頭）。
    None → 未引用空欄位（PostgreSQL CSV 模式視為 NULL），
    bool → t/f（引用），其餘全部引用並逸出內部雙引號。
    """
    lines = []
    for row in rows:
        fields = []
        for v in row:
            if v is None:
                fields.append("")
            elif isinstance(v, bool):
                fields.append('"t"' if v else '"f"')
            else:
                s = _clean_str(v).replace('"', '""')
                fields.append(f'"{s}"')
        lines.append(",".join(fields))
    return "\n".join(lines) + ("\n" if lines else "")


def _dt(value: Any) -> Optional[str]:
    if value is None:
        return None
    return value.isoformat() if hasattr(value, "isoformat") else str(value)


def _dec(value: Any) -> Optional[str]:
    if value is None:
        return None
    return str(value) if isinstance(value, Decimal) else str(value)


def _esc(text: str) -> str:
    return str(text).replace("'", "''")


# ── 資料列轉換 ────────────────────────────────────────────────────────────────
def _member_row(batch_id: str, clinic_id: int, rec: MemberRecord) -> List[Any]:
    raw = json.dumps({
        "quality_roster":            rec.quality_roster,
        "multi_chronic_65":          rec.multi_chronic_65,
        "high_visit":                rec.high_visit,
        "chronic_mark":              rec.chronic_mark,
        "non_chronic_mark":          rec.non_chronic_mark,
        "same_clinic_previous_year": rec.same_clinic_previous_year,
        "three_highs":               rec.three_highs,
        "hypertension":              rec.hypertension,
        "hyperlipidemia":            rec.hyperlipidemia,
        "hyperglycemia":             rec.hyperglycemia,
        "source_file":               rec.trace.source_file,
        "source_system":             rec.trace.source_system,
    }, ensure_ascii=False)
    return [
        batch_id, clinic_id, rec.trace.source_row,
        rec.person_id, rec.person_id, "",          # patient_id, normalized, chart_no
        rec.name, _dt(rec.birth_date), rec.sex,
        rec.phone, rec.mobile, rec.address,
        rec.case_category,                          # member_type
        rec.disease_pattern, rec.disease_pattern,   # disease_code, disease_text
        rec.ascvd, None,                            # ascvd, last_visit_date
        raw, rec.trace.raw_row_hash,
    ]


def _claim_row(batch_id: str, clinic_id: int, rec: MonthlyClaimRecord) -> List[Any]:
    raw = json.dumps({
        "source_file":   rec.trace.source_file,
        "source_system": rec.trace.source_system,
    }, ensure_ascii=False)
    natural_key = f"{clinic_id}_{rec.person_id}_{rec.roc_year}{rec.month:02d}"
    return [
        batch_id, clinic_id, rec.trace.source_row,
        rec.person_id, rec.person_id, "", "", None,  # id, normalized, chart, name, birth_date
        _dt(rec.last_visit_date),                  # service_date
        rec.roc_year, rec.month,
        _dec(rec.visit_count), _dec(rec.amount),
        "summary",                                  # data_level
        raw, natural_key, rec.trace.raw_row_hash,
    ]


def _flag_row(batch_id: str, clinic_id: int, rec: MemberSelectionRecord) -> List[Any]:
    raw = json.dumps({
        "source_file":   rec.trace.source_file,
        "source_system": rec.trace.source_system,
    }, ensure_ascii=False)
    return [
        batch_id, clinic_id,
        rec.person_id, rec.selection_type, True,
        raw, rec.trace.raw_row_hash,
    ]


def _p4p_case_row(batch_id: str, clinic_id: int, rec: P4PCaseRecord) -> List[Any]:
    raw = json.dumps({"source_file": rec.trace.source_file, "source_system": rec.trace.source_system}, ensure_ascii=False)
    return [batch_id, clinic_id, rec.person_id, rec.plan, rec.status,
            _dt(rec.enrolled_at), None, None, "", raw, rec.trace.raw_row_hash]


def _p4p_track_row(batch_id: str, clinic_id: int, rec: P4PTrackRecord) -> List[Any]:
    raw = json.dumps({"source_file": rec.trace.source_file, "source_system": rec.trace.source_system}, ensure_ascii=False)
    return [batch_id, clinic_id, rec.person_id, rec.plan, "",
            None, _dt(rec.last_tracked_at), _dt(rec.next_track_at), rec.overdue,
            raw, rec.trace.raw_row_hash]


def _screening_row(batch_id: str, clinic_id: int, rec: ScreeningRecord) -> List[Any]:
    raw = json.dumps({"source_file": rec.trace.source_file, "source_system": rec.trace.source_system}, ensure_ascii=False)
    return [batch_id, clinic_id, rec.person_id, rec.screening_type, _dt(rec.screened_at), raw, rec.trace.raw_row_hash]


def _lab_row(batch_id: str, clinic_id: int, rec: LabResultRecord) -> List[Any]:
    raw = json.dumps({"source_file": rec.trace.source_file, "source_system": rec.trace.source_system}, ensure_ascii=False)
    try:
        result_val = str(Decimal(str(rec.result_value))) if rec.result_value else None
    except (InvalidOperation, Exception):
        result_val = None
    return [batch_id, clinic_id, rec.person_id, rec.test_code, result_val, _dt(rec.tested_at), raw, rec.trace.raw_row_hash]


# ── COPY 區塊 ─────────────────────────────────────────────────────────────────
def _copy_block(table: str, columns: Sequence[str], rows: List[List[Any]]) -> str:
    if not rows:
        return ""
    col_list = ", ".join(columns)
    csv_data = _csv_block(rows)
    return (
        f"COPY {table} ({col_list}) FROM stdin WITH (FORMAT csv, HEADER false);\n"
        f"{csv_data}"
        f"\\.\n\n"
    )


# ── 主寫入器 ──────────────────────────────────────────────────────────────────
class PostgresStagingWriter:
    """DatasetBundle → staging.* 單一 transaction。"""

    def get_clinic_id(self, clinic_code: str) -> int:
        result = _run_query(
            f"SELECT clinic_id FROM meta.clinics WHERE clinic_code = '{_esc(clinic_code)}' LIMIT 1;"
        )
        if not result:
            raise ValueError(f"找不到診所代碼：{clinic_code}")
        return int(result)

    def stage(
        self,
        clinic_id: int,
        batch_id: str,
        bundle: DatasetBundle,
        validation_report: ValidationReport,
        source_system: str = "",
        source_root: str = "",
        requested_by: str = "",
    ) -> StageResult:
        sql = self._build_sql(clinic_id, batch_id, bundle, source_system, source_root, requested_by)
        try:
            _run_sql(sql)
        except RuntimeError as exc:
            try:
                _run_query(
                    f"UPDATE meta.import_batches SET status='failed', "
                    f"message='{_esc(str(exc)[:500])}' WHERE batch_id='{_esc(batch_id)}';"
                )
            except Exception:
                pass
            raise
        return StageResult(
            batch_id=batch_id,
            staged_counts=bundle.counts(),
            validation_report=validation_report,
        )

    def _build_sql(
        self,
        clinic_id: int,
        batch_id: str,
        bundle: DatasetBundle,
        source_system: str,
        source_root: str,
        requested_by: str,
    ) -> str:
        bid = _esc(batch_id)

        member_rows    = [_member_row(batch_id, clinic_id, r) for r in bundle.members]
        claim_rows     = [_claim_row(batch_id, clinic_id, r) for r in bundle.monthly_claims]
        flag_rows      = [_flag_row(batch_id, clinic_id, r) for r in bundle.member_selections]
        p4p_rows       = (
            [_p4p_case_row(batch_id, clinic_id, r) for r in bundle.p4p_cases] +
            [_p4p_track_row(batch_id, clinic_id, r) for r in bundle.p4p_tracks]
        )
        screening_rows = [_screening_row(batch_id, clinic_id, r) for r in bundle.screenings]
        lab_rows       = [_lab_row(batch_id, clinic_id, r) for r in bundle.lab_results]

        return "".join([
            "BEGIN;\n\n",

            f"INSERT INTO meta.import_batches "
            f"(batch_id, clinic_id, source_system, import_mode, scope, status, source_root, requested_by) "
            f"VALUES ('{bid}', {clinic_id}, '{_esc(source_system)}', 'replace_scope', '{{}}', 'staged', "
            f"'{_esc(source_root)}', '{_esc(requested_by)}') "
            f"ON CONFLICT (batch_id) DO UPDATE SET status='staged', started_at=now(), message=NULL;\n\n",

            f"DELETE FROM staging.members     WHERE batch_id='{bid}' AND clinic_id={clinic_id};\n",
            f"DELETE FROM staging.claims      WHERE batch_id='{bid}' AND clinic_id={clinic_id};\n",
            f"DELETE FROM staging.member_flags WHERE batch_id='{bid}' AND clinic_id={clinic_id};\n",
            f"DELETE FROM staging.p4p_records  WHERE batch_id='{bid}' AND clinic_id={clinic_id};\n",
            f"DELETE FROM staging.screenings   WHERE batch_id='{bid}' AND clinic_id={clinic_id};\n",
            f"DELETE FROM staging.lab_results  WHERE batch_id='{bid}' AND clinic_id={clinic_id};\n\n",

            _copy_block("staging.members", [
                "batch_id", "clinic_id", "row_no",
                "patient_id", "patient_id_normalized", "chart_no",
                "name", "birth_date", "sex", "phone", "mobile", "address",
                "member_type", "disease_code", "disease_text", "ascvd",
                "last_visit_date", "raw_data", "row_hash",
            ], member_rows),

            _copy_block("staging.claims", [
                "batch_id", "clinic_id", "row_no",
                "patient_id", "patient_id_normalized", "chart_no",
                "name", "birth_date", "service_date",
                "roc_year", "month", "visit_count", "claim_amount", "data_level",
                "raw_data", "natural_key", "row_hash",
            ], claim_rows),

            _copy_block("staging.member_flags", [
                "batch_id", "clinic_id",
                "patient_id_normalized", "flag_type", "flag_value",
                "raw_data", "row_hash",
            ], flag_rows),

            _copy_block("staging.p4p_records", [
                "batch_id", "clinic_id", "patient_id_normalized",
                "plan_name", "status",
                "enroll_date", "last_track_date", "next_track_date", "overdue_status",
                "raw_data", "row_hash",
            ], p4p_rows),

            _copy_block("staging.screenings", [
                "batch_id", "clinic_id", "patient_id_normalized",
                "screening_type", "screening_date",
                "raw_data", "row_hash",
            ], screening_rows),

            _copy_block("staging.lab_results", [
                "batch_id", "clinic_id", "patient_id_normalized",
                "test_type", "result_value", "result_date",
                "raw_data", "row_hash",
            ], lab_rows),

            f"UPDATE meta.import_batches SET status='validated', validated_at=now() "
            f"WHERE batch_id='{bid}';\n\n",

            "COMMIT;\n",
        ])


# ── 便利函式 ──────────────────────────────────────────────────────────────────
def stage_bundle(
    bundle: DatasetBundle,
    validation_report: ValidationReport,
    clinic_code: str,
    batch_id: str,
    source_system: str = "",
    source_root: str = "",
    requested_by: str = "",
) -> StageResult:
    """
    一行呼叫完成 clinic_id 查詢 + staging 寫入。

    範例：
        from db_pipeline.storage import stage_bundle
        result = stage_bundle(bundle, report, "3501110080", "my-batch-uuid")
    """
    writer = PostgresStagingWriter()
    clinic_id = writer.get_clinic_id(clinic_code)
    return writer.stage(
        clinic_id=clinic_id,
        batch_id=batch_id,
        bundle=bundle,
        validation_report=validation_report,
        source_system=source_system,
        source_root=source_root,
        requested_by=requested_by,
    )
