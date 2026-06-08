from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Protocol

from db_pipeline.config.models import ClinicConfig
from db_pipeline.datasets.models import DatasetBundle
from db_pipeline.validation.models import ValidationIssue


@dataclass
class ParseCoverage:
    discovered_files: int = 0
    parsed_files: int = 0
    skipped_files: Dict[str, str] = field(default_factory=dict)
    parsed_rows: Dict[str, int] = field(default_factory=dict)
    unmatched_rows: Dict[str, int] = field(default_factory=dict)
    unlinked_rows: Dict[str, int] = field(default_factory=dict)


@dataclass
class ParseResult:
    bundle: DatasetBundle
    coverage: ParseCoverage
    issues: List[ValidationIssue] = field(default_factory=list)


class SourceParser(Protocol):
    source_system: str

    def parse(
        self,
        source_dir: Path,
        config: ClinicConfig,
        batch_id: str,
    ) -> ParseResult:
        ...
