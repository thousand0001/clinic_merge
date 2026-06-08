from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from db_pipeline.config.models import ClinicConfig
from db_pipeline.datasets.models import DatasetBundle
from db_pipeline.parsers import get_parser


@dataclass(frozen=True)
class FolderDataProvider:
    source_dir: Path
    config: ClinicConfig
    batch_id: str

    def load_bundle(self) -> DatasetBundle:
        source_dir = self.source_dir.expanduser().resolve()
        if not source_dir.is_dir():
            raise FileNotFoundError(f"來源資料夾不存在：{source_dir}")
        parser = get_parser(self.config.source_system)
        result = parser.parse(source_dir, self.config, self.batch_id)
        errors = [issue for issue in result.issues if issue.severity == "error"]
        if errors:
            messages = "；".join(issue.message for issue in errors)
            raise ValueError(f"資料夾解析失敗：{messages}")
        return result.bundle
