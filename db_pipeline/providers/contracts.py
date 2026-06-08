from __future__ import annotations

from typing import Protocol

from db_pipeline.datasets.models import DatasetBundle


class DataProvider(Protocol):
    def load_bundle(self) -> DatasetBundle:
        """讀取單一診所批次並回傳標準資料集。"""
        ...
