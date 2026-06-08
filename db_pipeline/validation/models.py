from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List


@dataclass(frozen=True)
class ValidationIssue:
    severity: str
    dataset: str
    code: str
    message: str
    source_file: str = ""
    source_sheet: str = ""
    source_row: int = 0


@dataclass
class ValidationReport:
    issues: List[ValidationIssue] = field(default_factory=list)
    dataset_counts: Dict[str, int] = field(default_factory=dict)

    @property
    def is_valid(self) -> bool:
        return not any(issue.severity == "error" for issue in self.issues)

