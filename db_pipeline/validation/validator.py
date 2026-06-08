from __future__ import annotations

import re
from typing import Iterable

from db_pipeline.datasets.models import DatasetBundle, SourceTrace
from db_pipeline.validation.models import ValidationIssue, ValidationReport


TW_ID_RE = re.compile(r"^[A-Z][12][0-9]{8}$")


def _issue_for_trace(
    severity: str,
    dataset: str,
    code: str,
    message: str,
    trace: SourceTrace,
) -> ValidationIssue:
    return ValidationIssue(
        severity=severity,
        dataset=dataset,
        code=code,
        message=message,
        source_file=trace.source_file,
        source_sheet=trace.source_sheet,
        source_row=trace.source_row,
    )


def _validate_person_ids(
    report: ValidationReport,
    dataset: str,
    records: Iterable[object],
) -> None:
    for record in records:
        person_id = getattr(record, "person_id", "")
        trace = getattr(record, "trace")
        if not TW_ID_RE.fullmatch(person_id):
            severity = "warning" if person_id.startswith("chart:") else "error"
            message = (
                f"病歷號暫代識別碼（身分證號待補）：{person_id!r}"
                if person_id.startswith("chart:")
                else f"身分證號格式不符（應為英文字母開頭＋9位數字）：{person_id!r}"
            )
            report.issues.append(
                _issue_for_trace(
                    severity,
                    dataset,
                    "invalid_person_id",
                    message,
                    trace,
                )
            )


def validate_bundle(bundle: DatasetBundle) -> ValidationReport:
    report = ValidationReport(dataset_counts=bundle.counts())
    for dataset, records in (
        ("members", bundle.members),
        ("monthly_claims", bundle.monthly_claims),
        ("p4p_cases", bundle.p4p_cases),
        ("p4p_tracks", bundle.p4p_tracks),
        ("lab_results", bundle.lab_results),
        ("screenings", bundle.screenings),
        ("member_selections", bundle.member_selections),
    ):
        _validate_person_ids(report, dataset, records)

    for record in bundle.monthly_claims:
        if record.roc_year not in (114, 115):
            report.issues.append(
                _issue_for_trace(
                    "error",
                    "monthly_claims",
                    "invalid_roc_year",
                    f"民國年度超出支援範圍（僅接受 114、115）：{record.roc_year}",
                    record.trace,
                )
            )
        if not 1 <= record.month <= 12:
            report.issues.append(
                _issue_for_trace(
                    "error",
                    "monthly_claims",
                    "invalid_month",
                    f"月份超出範圍（需為 1～12）：{record.month}",
                    record.trace,
                )
            )
    return report

