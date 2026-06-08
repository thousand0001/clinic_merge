from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Optional, Sequence

from db_pipeline.config.models import load_clinic_config
from db_pipeline.parsers.sm import SmParser
from db_pipeline.validation.validator import validate_bundle


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        description="旁路解析診所資料並輸出覆蓋摘要，不寫入資料庫"
    )
    parser.add_argument("source_dir", type=Path)
    parser.add_argument("--config", type=Path, required=True)
    parser.add_argument("--batch-id", default="dry-run")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args(argv)

    source_dir = args.source_dir.expanduser().resolve()
    if not source_dir.is_dir():
        parser.error(f"來源資料夾不存在：{source_dir}")
    config = load_clinic_config(args.config)
    parsers = {"sm": SmParser()}
    source_parser = parsers.get(config.source_system)
    if source_parser is None:
        raise ValueError(f"尚未實作解析器：{config.source_system}")

    result = source_parser.parse(source_dir, config, args.batch_id)
    validation = validate_bundle(result.bundle)
    issues = result.issues + validation.issues
    is_valid = not any(issue.severity == "error" for issue in issues)
    summary = {
        "source_dir": str(source_dir),
        "source_system": config.source_system,
        "dataset_counts": result.bundle.counts(),
        "coverage": {
            "discovered_files": result.coverage.discovered_files,
            "parsed_files": result.coverage.parsed_files,
            "skipped_files": result.coverage.skipped_files,
            "parsed_rows": result.coverage.parsed_rows,
            "unmatched_rows": result.coverage.unmatched_rows,
        },
        "validation": {
            "is_valid": is_valid,
            "issue_count": len(issues),
            "issues": [
                {
                    "severity": issue.severity,
                    "dataset": issue.dataset,
                    "code": issue.code,
                    "message": issue.message,
                    "source_file": issue.source_file,
                    "source_sheet": issue.source_sheet,
                    "source_row": issue.source_row,
                }
                for issue in issues[:100]
            ],
        },
    }
    text = json.dumps(summary, ensure_ascii=False, indent=2)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(text + "\n", encoding="utf-8")
    else:
        print(text)
    return 0 if is_valid else 2


if __name__ == "__main__":
    raise SystemExit(main())
