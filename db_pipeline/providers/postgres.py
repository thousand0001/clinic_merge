from __future__ import annotations

import datetime as dt
import json
import uuid
from dataclasses import dataclass
from decimal import Decimal
from typing import Any, Dict, Iterable, List, Optional

from db_pipeline.datasets.models import (
    DatasetBundle,
    LabResultRecord,
    MemberRecord,
    MemberSelectionRecord,
    MonthlyClaimRecord,
    P4PCaseRecord,
    P4PTrackRecord,
    ScreeningRecord,
    SourceTrace,
)
from db_pipeline.storage import _run_query


def _esc(value: str) -> str:
    return value.replace("'", "''")


def _date(value: Any) -> Optional[dt.date]:
    if not value:
        return None
    return dt.date.fromisoformat(str(value)[:10])


def _decimal(value: Any) -> Decimal:
    return Decimal(str(value or 0))


def _rows(sql: str) -> List[Dict[str, Any]]:
    output = _run_query(sql)
    if not output:
        return []
    return [json.loads(line) for line in output.splitlines() if line.strip()]


def _trace(
    clinic_code: str,
    batch_id: str,
    source_system: str,
    row: Dict[str, Any],
) -> SourceTrace:
    raw = row.get("raw_data") or {}
    if isinstance(raw, str):
        raw = json.loads(raw)
    return SourceTrace(
        clinic_code=clinic_code,
        batch_id=batch_id,
        source_system=str(raw.get("source_system") or source_system),
        source_file=str(raw.get("source_file") or ""),
        source_sheet=str(raw.get("source_sheet") or ""),
        source_row=int(row.get("row_no") or 0),
        raw_row_hash=str(row.get("row_hash") or ""),
    )


def _json_query(table: str, columns: Iterable[str], where: str, order: str) -> str:
    selected = ", ".join(columns)
    return (
        "SELECT row_to_json(provider_row)::text FROM ("
        f"SELECT {selected} FROM {table} WHERE {where} ORDER BY {order}"
        ") AS provider_row;"
    )


@dataclass(frozen=True)
class PostgresDataProvider:
    clinic_code: str
    batch_id: str

    def load_bundle(self) -> DatasetBundle:
        batch_id = str(uuid.UUID(self.batch_id))
        clinic_code = _esc(self.clinic_code)
        metadata = _run_query(
            "SELECT b.clinic_id, b.source_system "
            "FROM meta.import_batches b "
            "JOIN meta.clinics c ON c.clinic_id=b.clinic_id "
            f"WHERE b.batch_id='{batch_id}'::uuid "
            f"AND c.clinic_code='{clinic_code}' "
            "AND b.status IN ('validated','published','superseded') LIMIT 1;"
        )
        if not metadata:
            raise ValueError(
                f"找不到已驗證批次：clinic_code={self.clinic_code}, "
                f"batch_id={batch_id}"
            )
        clinic_id_text, source_system = metadata.split("|", 1)
        clinic_id = int(clinic_id_text)
        where = f"batch_id='{batch_id}'::uuid AND clinic_id={clinic_id}"
        bundle = DatasetBundle()

        member_rows = _rows(_json_query(
            "staging.members",
            (
                "row_no", "patient_id_normalized", "name", "birth_date",
                "sex", "phone", "mobile", "address", "member_type",
                "disease_code", "ascvd", "raw_data", "row_hash",
            ),
            where,
            "row_no, patient_id_normalized",
        ))
        for row in member_rows:
            raw = row.get("raw_data") or {}
            bundle.members.append(MemberRecord(
                trace=_trace(self.clinic_code, batch_id, source_system, row),
                person_id=str(row["patient_id_normalized"]),
                name=str(row.get("name") or ""),
                birth_date=_date(row.get("birth_date")),
                sex=str(row.get("sex") or ""),
                phone=str(row.get("phone") or ""),
                mobile=str(row.get("mobile") or ""),
                address=str(row.get("address") or ""),
                case_category=str(row.get("member_type") or ""),
                disease_pattern=str(row.get("disease_code") or ""),
                ascvd=str(row.get("ascvd") or ""),
                quality_roster=str(raw.get("quality_roster") or ""),
                multi_chronic_65=str(raw.get("multi_chronic_65") or ""),
                high_visit=str(raw.get("high_visit") or ""),
                chronic_mark=str(raw.get("chronic_mark") or ""),
                non_chronic_mark=str(raw.get("non_chronic_mark") or ""),
                same_clinic_previous_year=str(
                    raw.get("same_clinic_previous_year") or ""),
                three_highs=str(raw.get("three_highs") or ""),
                hypertension=str(raw.get("hypertension") or ""),
                hyperlipidemia=str(raw.get("hyperlipidemia") or ""),
                hyperglycemia=str(raw.get("hyperglycemia") or ""),
            ))

        claim_rows = _rows(_json_query(
            "staging.claims",
            (
                "row_no", "patient_id_normalized", "service_date",
                "roc_year", "month", "visit_count", "claim_amount",
                "raw_data", "row_hash",
            ),
            where,
            "roc_year, month, patient_id_normalized, row_no",
        ))
        for row in claim_rows:
            bundle.monthly_claims.append(MonthlyClaimRecord(
                trace=_trace(self.clinic_code, batch_id, source_system, row),
                person_id=str(row["patient_id_normalized"]),
                roc_year=int(row["roc_year"]),
                month=int(row["month"]),
                visit_count=_decimal(row.get("visit_count")),
                amount=_decimal(row.get("claim_amount")),
                last_visit_date=_date(row.get("service_date")),
            ))

        flag_rows = _rows(_json_query(
            "staging.member_flags",
            (
                "patient_id_normalized", "flag_type", "raw_data", "row_hash",
            ),
            where,
            "patient_id_normalized, flag_type",
        ))
        for row in flag_rows:
            bundle.member_selections.append(MemberSelectionRecord(
                trace=_trace(self.clinic_code, batch_id, source_system, row),
                person_id=str(row["patient_id_normalized"]),
                selection_type=str(row["flag_type"]),
            ))

        p4p_rows = _rows(_json_query(
            "staging.p4p_records",
            (
                "patient_id_normalized", "plan_name", "status",
                "enroll_date", "last_track_date", "next_track_date",
                "overdue_status", "raw_data", "row_hash",
            ),
            where,
            "patient_id_normalized, plan_name, status, last_track_date",
        ))
        for row in p4p_rows:
            trace = _trace(self.clinic_code, batch_id, source_system, row)
            if row.get("status") or row.get("enroll_date"):
                bundle.p4p_cases.append(P4PCaseRecord(
                    trace=trace,
                    person_id=str(row["patient_id_normalized"]),
                    plan=str(row.get("plan_name") or ""),
                    status=str(row.get("status") or ""),
                    enrolled_at=_date(row.get("enroll_date")),
                ))
            else:
                bundle.p4p_tracks.append(P4PTrackRecord(
                    trace=trace,
                    person_id=str(row["patient_id_normalized"]),
                    plan=str(row.get("plan_name") or ""),
                    last_tracked_at=_date(row.get("last_track_date")),
                    next_track_at=_date(row.get("next_track_date")),
                    overdue=str(row.get("overdue_status") or ""),
                ))

        screening_rows = _rows(_json_query(
            "staging.screenings",
            (
                "patient_id_normalized", "screening_type", "screening_date",
                "raw_data", "row_hash",
            ),
            where,
            "patient_id_normalized, screening_type, screening_date",
        ))
        for row in screening_rows:
            bundle.screenings.append(ScreeningRecord(
                trace=_trace(self.clinic_code, batch_id, source_system, row),
                person_id=str(row["patient_id_normalized"]),
                screening_type=str(row["screening_type"]),
                screened_at=_date(row.get("screening_date")),
            ))

        lab_rows = _rows(_json_query(
            "staging.lab_results",
            (
                "patient_id_normalized", "test_type", "result_value",
                "result_date", "raw_data", "row_hash",
            ),
            where,
            "patient_id_normalized, test_type, result_date",
        ))
        for row in lab_rows:
            bundle.lab_results.append(LabResultRecord(
                trace=_trace(self.clinic_code, batch_id, source_system, row),
                person_id=str(row["patient_id_normalized"]),
                test_code=str(row["test_type"]),
                result_value=str(row.get("result_value") or ""),
                tested_at=_date(row.get("result_date")),
            ))
        return bundle
