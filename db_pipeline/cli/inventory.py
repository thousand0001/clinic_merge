from __future__ import annotations

import argparse
import json
from collections import Counter
from pathlib import Path
from typing import Optional, Sequence

from db_pipeline.config.models import load_clinic_config
from db_pipeline.detection.detector import SUPPORTED_FILE_SUFFIXES, detect_source_system


def build_inventory(source_dir: Path) -> dict:
    files = [
        path
        for path in source_dir.rglob("*")
        if path.is_file() and path.suffix.lower() in SUPPORTED_FILE_SUFFIXES
    ]
    suffix_counts = Counter(path.suffix.lower() for path in files)
    return {
        "source_dir": str(source_dir),
        "file_count": len(files),
        "suffix_counts": dict(sorted(suffix_counts.items())),
        "files": [str(path.relative_to(source_dir)) for path in sorted(files)],
    }


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="只讀盤點診所來源資料夾，不寫入資料庫")
    parser.add_argument("source_dir", type=Path)
    parser.add_argument("--config", type=Path, action="append", default=[])
    parser.add_argument("--source-system")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args(argv)

    source_dir = args.source_dir.expanduser().resolve()
    if not source_dir.is_dir():
        raise ValueError(f"找不到來源資料夾：{source_dir}")

    configs = [load_clinic_config(path) for path in args.config]
    detection = detect_source_system(
        source_dir,
        configs,
        explicit_source_system=args.source_system,
    )
    result = build_inventory(source_dir)
    result["detection"] = {
        "source_system": detection.source_system,
        "confidence": detection.confidence,
        "candidates": [
            {
                "source_system": candidate.source_system,
                "score": candidate.score,
                "reasons": list(candidate.reasons),
            }
            for candidate in detection.candidates
        ],
    }

    text = json.dumps(result, ensure_ascii=False, indent=2)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(text + "\n", encoding="utf-8")
    else:
        print(text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

