from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Optional

from .coverage import load_output_ids, summarize, write_csv_report, write_xlsx_report
from .inventory import inventory_folder


def main(argv: Optional[list[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="統一前置清洗旁路盤點工具（只讀，不產正式 Excel）")
    parser.add_argument("source_dir", help="來源資料夾")
    parser.add_argument("--output", "-o", help="報告輸出路徑，預設 /private/tmp/unified_preclean_report.xlsx")
    parser.add_argument("--csv", action="store_true", help="輸出 CSV 而不是 xlsx")
    parser.add_argument("--existing-output", help="既有正式 Excel；目前只讀取會員總表 ID，供後續覆蓋比對擴充")
    parser.add_argument("--json-summary", action="store_true", help="在終端機輸出 JSON 摘要")
    args = parser.parse_args(argv)

    source_dir = Path(args.source_dir).expanduser().resolve()
    if not source_dir.is_dir():
        raise FileNotFoundError(f"找不到來源資料夾：{source_dir}")

    output_path = Path(args.output).expanduser().resolve() if args.output else Path(
        "/private/tmp/unified_preclean_report.csv" if args.csv else "/private/tmp/unified_preclean_report.xlsx"
    )
    output_ids = load_output_ids(Path(args.existing_output).expanduser().resolve()) if args.existing_output else None
    findings = inventory_folder(source_dir, output_ids=output_ids)

    if args.csv:
        written = write_csv_report(source_dir, findings, output_path)
    else:
        written = write_xlsx_report(source_dir, findings, output_path)

    summary = summarize(findings)
    print(f"來源資料夾：{source_dir}")
    print(f"盤點項目：{summary['files_or_sheets']}")
    print(f"報告：{written}")
    if args.json_summary:
        print(json.dumps(summary, ensure_ascii=False, indent=2, sort_keys=True))
    else:
        print("資料類型：" + "；".join(f"{key}={value}" for key, value in summary["by_type"].items()))
        print("狀態：" + "；".join(f"{key}={value}" for key, value in summary["status"].items()))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

