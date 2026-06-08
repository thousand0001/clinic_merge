from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Optional, Sequence, Tuple

from db_pipeline.config.models import ClinicConfig


SUPPORTED_FILE_SUFFIXES = {".xlsx", ".xlsm", ".xls", ".ods", ".csv", ".txt", ".pdf"}


@dataclass(frozen=True)
class DetectionCandidate:
    source_system: str
    score: int
    reasons: Tuple[str, ...]


@dataclass(frozen=True)
class DetectionResult:
    source_system: Optional[str]
    confidence: str
    candidates: Tuple[DetectionCandidate, ...]


def _source_files(source_dir: Path) -> List[Path]:
    return sorted(
        path
        for path in source_dir.rglob("*")
        if path.is_file()
        and not path.name.startswith(("~$", "."))
        and path.suffix.lower() in SUPPORTED_FILE_SUFFIXES
    )


def _score_config(
    config: ClinicConfig,
    source_dir: Path,
    files: Sequence[Path],
) -> DetectionCandidate:
    score = 0
    reasons: List[str] = []
    names = [path.name.lower() for path in files]
    source_text = str(source_dir).lower()

    if config.clinic_code.lower() in source_text:
        score += 100
        reasons.append(f"路徑包含醫事機構代碼 {config.clinic_code}")
    if config.clinic_name.lower() in source_text:
        score += 40
        reasons.append(f"路徑包含診所名稱 {config.clinic_name}")

    for token in config.detection.file_name_contains:
        token_lower = token.lower()
        matches = sum(token_lower in name for name in names)
        if matches:
            score += min(matches, 3) * 10
            reasons.append(f"檔名包含 {token} ({matches})")

    for pattern in config.detection.file_name_regex:
        regex = re.compile(pattern, re.IGNORECASE)
        matches = sum(bool(regex.search(name)) for name in names)
        if matches:
            score += min(matches, 3) * 12
            reasons.append(f"檔名符合 {pattern} ({matches})")

    return DetectionCandidate(
        source_system=config.source_system,
        score=score,
        reasons=tuple(reasons),
    )


def detect_source_system(
    source_dir: Path,
    configs: Iterable[ClinicConfig],
    explicit_source_system: Optional[str] = None,
) -> DetectionResult:
    if explicit_source_system:
        return DetectionResult(
            source_system=explicit_source_system,
            confidence="explicit",
            candidates=(),
        )

    files = _source_files(source_dir)
    candidates = sorted(
        (_score_config(config, source_dir, files) for config in configs),
        key=lambda item: (-item.score, item.source_system),
    )
    positive = [candidate for candidate in candidates if candidate.score > 0]
    if not positive:
        return DetectionResult(None, "unknown", tuple(candidates))
    if len(positive) > 1 and positive[0].score == positive[1].score:
        return DetectionResult(None, "ambiguous", tuple(candidates))
    confidence = "high" if positive[0].score >= 20 else "low"
    return DetectionResult(positive[0].source_system, confidence, tuple(candidates))
