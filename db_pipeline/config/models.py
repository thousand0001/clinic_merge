from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional


SUPPORTED_SOURCE_SYSTEMS = {
    "sm",
    "new_sm",
    "prospect",
    "hongcheng",
    "medical_saint",
    "tiaohe",
    "custom",
}


@dataclass(frozen=True)
class DetectionRule:
    file_name_contains: List[str] = field(default_factory=list)
    file_name_regex: List[str] = field(default_factory=list)
    sheet_name_contains: List[str] = field(default_factory=list)
    header_contains: List[str] = field(default_factory=list)


@dataclass(frozen=True)
class ClinicConfig:
    clinic_code: str
    clinic_name: str
    source_system: str
    detection: DetectionRule = field(default_factory=DetectionRule)
    encodings: List[str] = field(
        default_factory=lambda: ["utf-8-sig", "utf-16", "cp950", "big5"]
    )
    sheet_aliases: Dict[str, List[str]] = field(default_factory=dict)
    column_aliases: Dict[str, List[str]] = field(default_factory=dict)
    options: Dict[str, Any] = field(default_factory=dict)

    def validate(self) -> None:
        if not self.clinic_code:
            raise ValueError("clinic_code 不可為空")
        if not self.clinic_name:
            raise ValueError("clinic_name 不可為空")
        if self.source_system not in SUPPORTED_SOURCE_SYSTEMS:
            raise ValueError(
                "不支援的 source_system："
                f"{self.source_system}；可用值為 {sorted(SUPPORTED_SOURCE_SYSTEMS)}"
            )


def load_clinic_config(path: Path) -> ClinicConfig:
    data = json.loads(path.read_text(encoding="utf-8"))
    detection_data = data.pop("detection", {})
    config = ClinicConfig(
        detection=DetectionRule(**detection_data),
        **data,
    )
    config.validate()
    return config

