# -*- coding: utf-8 -*-
"""
db_pipeline 一鍵產出 CLI

使用方式：
    # 指定診所代碼，自動取最新 validated 批次
    python -m db_pipeline.cli.output --clinic-code 3501186011 --template 選會員模板0526.xlsx

    # 指定批次 ID（用於跨診所比對或回溯）
    python -m db_pipeline.cli.output --batch-id 2912a96f-... --template 選會員模板0526.xlsx

    # 指定輸出路徑
    python -m db_pipeline.cli.output --clinic-code 3501186011 \\
        --template 選會員模板0526.xlsx \\
        --dest /tmp/鈞安診所_0609.xlsx
"""
from __future__ import annotations

import argparse
import datetime
import importlib.util
import os
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Optional, Sequence

# 將專案根加入 sys.path
_PROJECT_DIR = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_PROJECT_DIR))

from db_pipeline.output.member_builder import build_from_bundle
from db_pipeline.providers.postgres import PostgresDataProvider


def _psql_args() -> list:
    psql = os.getenv("CLINIC_PSQL") or shutil.which("psql") or "psql"
    return [
        psql,
        "-h", os.getenv("CLINIC_DB_HOST", "localhost"),
        "-p", os.getenv("CLINIC_DB_PORT", "5432"),
        "-U", os.getenv("CLINIC_DB_USER", "thousand0001"),
        "-d", os.getenv("CLINIC_DB_NAME", "clinic_merge"),
        "-t", "-A",
    ]


def _query(sql: str) -> str:
    res = subprocess.run(_psql_args() + ["-c", sql],
                         capture_output=True, text=True)
    if res.returncode != 0:
        raise RuntimeError(res.stderr.strip())
    return res.stdout.strip()


def _latest_validated_batch(clinic_code: str) -> tuple[str, str, str]:
    """回傳 (batch_id, clinic_name, source_system)，找不到時拋 ValueError。"""
    row = _query(
        f"SELECT ib.batch_id, c.clinic_name, ib.source_system "
        f"FROM meta.import_batches ib "
        f"JOIN meta.clinics c ON ib.clinic_id = c.clinic_id "
        f"WHERE c.clinic_code = '{clinic_code}' "
        f"  AND ib.status IN ('validated', 'published') "
        f"ORDER BY ib.validated_at DESC NULLS LAST, ib.started_at DESC "
        f"LIMIT 1;"
    )
    if not row:
        raise ValueError(f"找不到診所 {clinic_code!r} 的已驗證批次。")
    parts = row.split("|")
    return parts[0], parts[1], parts[2]


def _batch_info(batch_id: str) -> tuple[str, str, str, str]:
    """回傳 (batch_id, clinic_code, clinic_name, source_system)。"""
    row = _query(
        f"SELECT ib.batch_id, c.clinic_code, c.clinic_name, ib.source_system "
        f"FROM meta.import_batches ib "
        f"JOIN meta.clinics c ON ib.clinic_id = c.clinic_id "
        f"WHERE ib.batch_id = '{batch_id}'::uuid;"
    )
    if not row:
        raise ValueError(f"找不到批次 {batch_id!r}。")
    parts = row.split("|")
    return parts[0], parts[1], parts[2], parts[3]


def _load_write_output():
    """動態載入舊流程 write_output()（檔名含中文）。"""
    legacy_path = _PROJECT_DIR / "資料庫輸出0601.py"
    if not legacy_path.exists():
        raise FileNotFoundError(f"找不到舊流程模組：{legacy_path}")
    spec = importlib.util.spec_from_file_location("legacy_output", legacy_path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod.write_output


def generate(
    batch_id: str,
    clinic_code: str,
    template_path: Path,
    dest_path: Path,
    *,
    quiet: bool = False,
) -> int:
    """
    指定批次 → Excel。回傳寫入行數。
    """
    _, _, clinic_name, _ = _batch_info(batch_id)

    provider = PostgresDataProvider(clinic_code=clinic_code, batch_id=batch_id)
    bundle = provider.load_bundle()

    if not quiet:
        counts = bundle.counts()
        print(f"  members={counts['members']}  claims={counts['monthly_claims']}  "
              f"p4p={counts['p4p_cases']}  lab={counts['lab_results']}  "
              f"screenings={counts['screenings']}")

    members = build_from_bundle(bundle)
    write_output = _load_write_output()
    count = write_output(template_path, dest_path, clinic_name, members)
    return count


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        description="從 staging 批次產出選會員 Excel（一鍵產出）"
    )
    grp = parser.add_mutually_exclusive_group(required=True)
    grp.add_argument("--clinic-code", metavar="CODE",
                     help="診所代碼（自動取最新 validated 批次）")
    grp.add_argument("--batch-id", metavar="UUID",
                     help="指定批次 UUID")
    parser.add_argument("--template", type=Path, required=True,
                        help="Excel 範本路徑")
    parser.add_argument("--dest", type=Path,
                        help="輸出 Excel 路徑（預設：<診所名稱>_<日期>.xlsx）")
    parser.add_argument("--quiet", action="store_true",
                        help="不顯示詳細計數")
    args = parser.parse_args(argv)

    template = args.template.expanduser().resolve()
    if not template.exists():
        print(f"[錯誤] 找不到範本：{template}", file=sys.stderr)
        return 1

    # 取得 batch_id 與基本資訊
    try:
        if args.clinic_code:
            batch_id, clinic_name, source_system = _latest_validated_batch(args.clinic_code)
            clinic_code = args.clinic_code
            print(f"診所：{clinic_name}（{clinic_code}）  來源系統：{source_system}")
        else:
            batch_id, clinic_code, clinic_name, source_system = _batch_info(args.batch_id)
            batch_id = args.batch_id
            print(f"診所：{clinic_name}（{clinic_code}）  來源系統：{source_system}")
    except (ValueError, RuntimeError) as exc:
        print(f"[錯誤] {exc}", file=sys.stderr)
        return 1

    print(f"批次：{batch_id}")

    # 決定輸出路徑
    if args.dest:
        dest = args.dest.expanduser().resolve()
    else:
        today = datetime.date.today().strftime("%Y%m%d")
        dest = Path(f"{clinic_name}_{today}.xlsx").resolve()

    dest.parent.mkdir(parents=True, exist_ok=True)

    try:
        count = generate(batch_id, clinic_code, template, dest, quiet=args.quiet)
    except Exception as exc:
        print(f"[錯誤] {exc}", file=sys.stderr)
        return 1

    print(f"✓ 已產出 {dest}（{count} 筆）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
