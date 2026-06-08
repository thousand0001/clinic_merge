# -*- coding: utf-8 -*-
"""
db_pipeline 批次清單 CLI

使用方式：
    # 列出所有批次（最近 20 筆）
    python -m db_pipeline.cli.batches

    # 指定診所
    python -m db_pipeline.cli.batches --clinic-code 3501186011

    # 顯示更多
    python -m db_pipeline.cli.batches --limit 50

    # 只顯示 validated
    python -m db_pipeline.cli.batches --status validated
"""
from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
from typing import Optional, Sequence


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


def list_batches(
    clinic_code: Optional[str] = None,
    status: Optional[str] = None,
    limit: int = 20,
) -> list[dict]:
    """回傳批次摘要列表。"""
    where_parts = ["1=1"]
    if clinic_code:
        where_parts.append(f"c.clinic_code = '{clinic_code}'")
    if status:
        where_parts.append(f"ib.status = '{status}'")
    where = " AND ".join(where_parts)

    sql = (
        f"SELECT "
        f"  c.clinic_code, c.clinic_name, "
        f"  ib.batch_id, ib.source_system, ib.status, "
        f"  to_char(ib.started_at AT TIME ZONE 'Asia/Taipei', 'YYYY-MM-DD HH24:MI') AS started_at, "
        f"  to_char(ib.validated_at AT TIME ZONE 'Asia/Taipei', 'YYYY-MM-DD HH24:MI') AS validated_at "
        f"FROM meta.import_batches ib "
        f"JOIN meta.clinics c ON ib.clinic_id = c.clinic_id "
        f"WHERE {where} "
        f"ORDER BY ib.started_at DESC "
        f"LIMIT {limit};"
    )
    raw = _query(sql)
    if not raw:
        return []

    rows = []
    for line in raw.splitlines():
        parts = line.split("|")
        if len(parts) < 7:
            continue
        rows.append({
            "clinic_code":   parts[0],
            "clinic_name":   parts[1],
            "batch_id":      parts[2],
            "source_system": parts[3],
            "status":        parts[4],
            "started_at":    parts[5],
            "validated_at":  parts[6],
        })
    return rows


def _status_label(status: str) -> str:
    return {
        "validated": "✓ validated",
        "published": "★ published",
        "failed":    "✗ failed",
        "staged":    "○ staged",
    }.get(status, status)


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        description="列出 staging 批次（診所、狀態、時間）"
    )
    parser.add_argument("--clinic-code", metavar="CODE",
                        help="只顯示指定診所")
    parser.add_argument("--status", metavar="STATUS",
                        choices=["validated", "published", "failed", "staged"],
                        help="只顯示特定狀態")
    parser.add_argument("--limit", type=int, default=20,
                        help="顯示筆數上限（預設 20）")
    args = parser.parse_args(argv)

    try:
        rows = list_batches(
            clinic_code=args.clinic_code,
            status=args.status,
            limit=args.limit,
        )
    except RuntimeError as exc:
        print(f"[錯誤] {exc}", file=sys.stderr)
        return 1

    if not rows:
        print("（沒有符合條件的批次）")
        return 0

    # 欄寬
    name_w  = max(len(r["clinic_name"])   for r in rows)
    sys_w   = max(len(r["source_system"]) for r in rows)
    st_w    = max(len(_status_label(r["status"])) for r in rows)

    header = (f"{'診所名稱':<{name_w}}  {'代碼':<12}  {'系統':<{sys_w}}  "
              f"{'狀態':<{st_w}}  {'批次 ID':<36}  {'開始時間':<16}  驗證時間")
    print(header)
    print("-" * len(header))
    for r in rows:
        print(
            f"{r['clinic_name']:<{name_w}}  {r['clinic_code']:<12}  "
            f"{r['source_system']:<{sys_w}}  {_status_label(r['status']):<{st_w}}  "
            f"{r['batch_id']:<36}  {r['started_at']:<16}  {r['validated_at']}"
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
