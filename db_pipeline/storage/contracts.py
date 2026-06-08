from __future__ import annotations

from dataclasses import dataclass
from typing import Protocol

from db_pipeline.datasets.models import DatasetBundle
from db_pipeline.validation.models import ValidationReport


@dataclass(frozen=True)
class StageResult:
    batch_id: str
    staged_counts: dict
    validation_report: ValidationReport


class StagingWriter(Protocol):
    """標準資料寫入 staging 的邊界。

    實作必須在單一 transaction 中完成；若驗證失敗，不得更新 current。
    """

    def stage(
        self,
        clinic_id: int,
        batch_id: str,
        bundle: DatasetBundle,
        validation_report: ValidationReport,
    ) -> StageResult:
        ...

