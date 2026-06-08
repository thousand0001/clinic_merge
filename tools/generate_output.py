# -*- coding: utf-8 -*-
"""
從 staging DB 產出 Excel（新流程輸出）

使用方式：
    python tools/generate_output.py <batch_id> \\
        --template "115年6月指定名單格式/選會員模板0526.xlsx" \\
        --dest /tmp/new_output.xlsx

批次 ID 必須已存在且 status = 'validated'。
診所名稱自動從 meta.clinics 查詢。
"""
from __future__ import annotations

import argparse
import importlib.util
import sys
from pathlib import Path
from typing import Optional, Sequence

# ── 將專案根目錄加入 sys.path 以便 import ──────────────────────────────────────
PROJECT_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_DIR))

from db_pipeline.datasets.models import DatasetBundle
from db_pipeline.output.member_builder import build_from_bundle
from db_pipeline.providers.postgres import PostgresDataProvider


def _load_legacy():
    """動態載入 資料庫輸出0601.py（檔名含中文，無法直接 import）。"""
    legacy_path = PROJECT_DIR / "資料庫輸出0601.py"
    if not legacy_path.exists():
        raise FileNotFoundError(f"找不到舊流程模組：{legacy_path}")
    spec = importlib.util.spec_from_file_location("legacy_output", legacy_path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def generate(
    batch_id: str,
    template_path: Path,
    dest_path: Path,
) -> int:
    """
    1. 從 DB 讀取 batch_id 的 staging 資料
    2. 建立 member dict（build_from_bundle）
    3. 呼叫舊流程 write_output() 產出 Excel
    回傳寫入行數。
    """
    # 查詢診所代碼與名稱
    import subprocess, os, shutil

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
        res = subprocess.run(
            _psql_args() + ["-c", sql],
            capture_output=True, text=True,
        )
        if res.returncode != 0:
            raise RuntimeError(res.stderr.strip())
        return res.stdout.strip()

    row = _query(
        f"SELECT c.clinic_code, c.clinic_name "
        f"FROM meta.import_batches ib "
        f"JOIN meta.clinics c ON ib.clinic_id = c.clinic_id "
        f"WHERE ib.batch_id = '{batch_id}'::uuid LIMIT 1;"
    )
    if not row:
        raise ValueError(f"找不到批次：{batch_id}")
    clinic_code, clinic_name = row.split("|")

    # 讀取 staging → DatasetBundle
    provider = PostgresDataProvider(clinic_code=clinic_code, batch_id=batch_id)
    bundle: DatasetBundle = provider.load_bundle()

    # 建立 member dict
    members = build_from_bundle(bundle)
    print(f"members: {len(members)}")

    # 載入舊流程的 write_output
    legacy = _load_legacy()
    count = legacy.write_output(template_path, dest_path, clinic_name, members)
    return count


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="從 staging 批次產出選會員 Excel")
    parser.add_argument("batch_id", help="meta.import_batches.batch_id (validated)")
    parser.add_argument("--template", required=True, help="Excel 範本路徑")
    parser.add_argument("--dest", required=True, help="輸出 Excel 路徑")
    args = parser.parse_args(argv)

    template = Path(args.template)
    dest = Path(args.dest)

    if not template.exists():
        print(f"[錯誤] 找不到範本：{template}", file=sys.stderr)
        return 1

    dest.parent.mkdir(parents=True, exist_ok=True)

    count = generate(args.batch_id, template, dest)
    print(f"已產出 {dest}（{count} 筆）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
