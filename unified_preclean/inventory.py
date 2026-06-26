from __future__ import annotations

from pathlib import Path
from typing import List, Optional, Set

from .detectors import detect_sheet
from .readers import discover_files, read_any
from .schema import SourceFinding


def inventory_folder(source_dir: Path, output_ids: Optional[Set[str]] = None) -> List[SourceFinding]:
    findings: List[SourceFinding] = []
    for path in discover_files(source_dir):
        try:
            sheets = read_any(path)
        except Exception as exc:
            findings.append(
                SourceFinding(
                    file_path=path,
                    extension=path.suffix.lower(),
                    status="error",
                    reason=f"read failed: {exc}",
                )
            )
            continue
        for sheet in sheets:
            finding = detect_sheet(sheet)
            if output_ids is not None and finding.unique_id_count:
                matched = finding.source_ids & output_ids
                finding.matched_output_ids = len(matched)
                finding.missing_output_ids = len(finding.source_ids - output_ids)
            findings.append(finding)
    return findings
