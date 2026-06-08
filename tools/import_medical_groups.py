# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import csv
import os
import subprocess
import tempfile
from pathlib import Path
from typing import Any, Dict

import openpyxl


def normalize_code(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return text


def header_map(row) -> Dict[str, int]:
    return {str(value).strip(): idx for idx, value in enumerate(row) if value is not None and str(value).strip()}


def _find_psql() -> str:
    """跨平台尋找 psql 執行檔路徑。優先使用環境變數 CLINIC_PSQL，其次在 PATH 搜尋。"""
    import shutil
    env_val = os.getenv("CLINIC_PSQL")
    if env_val:
        return env_val
    found = shutil.which("psql")
    if found:
        return found
    win_default = r"C:\Program Files\PostgreSQL\16\bin\psql.exe"
    if os.path.isfile(win_default):
        return win_default
    return "psql"


def run_psql(sql_path: Path) -> None:
    cmd = [
        _find_psql(),
        "-h",
        os.getenv("CLINIC_DB_HOST", "localhost"),
        "-p",
        os.getenv("CLINIC_DB_PORT", "5432"),
        "-U",
        os.getenv("CLINIC_DB_USER", "thousand0001"),
        "-d",
        os.getenv("CLINIC_DB_NAME", "clinic_merge"),
        "-v",
        "ON_ERROR_STOP=1",
        "-f",
        str(sql_path),
    ]
    subprocess.run(cmd, check=True)


def import_medical_groups(xlsx_path: Path) -> int:
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    try:
        ws = wb.active
        rows = ws.iter_rows(values_only=True)
        next(rows, None)
        headers = header_map(next(rows))
        required = ["分區業務組別代碼", "醫事機構代碼", "醫事機構名稱", "機構地址", "原始合約起始日期", "官方名稱"]
        missing = [name for name in required if name not in headers]
        if missing:
            raise ValueError("醫療群檔缺少欄位：" + "、".join(missing))

        with tempfile.TemporaryDirectory(prefix="clinic_medical_groups_") as tmp_dir:
            tmp_path = Path(tmp_dir)
            csv_path = tmp_path / "medical_groups.csv"
            sql_path = tmp_path / "import_medical_groups.sql"
            count = 0
            with csv_path.open("w", newline="", encoding="utf-8") as fh:
                writer = csv.writer(fh)
                writer.writerow(["group_name", "clinic_code", "clinic_name", "address", "official_name", "contract_start"])
                for row in rows:
                    group_name = str(row[headers["分區業務組別代碼"]] or "").strip()
                    clinic_code = normalize_code(row[headers["醫事機構代碼"]])
                    clinic_name = str(row[headers["醫事機構名稱"]] or "").strip()
                    address = str(row[headers["機構地址"]] or "").strip()
                    official_name = str(row[headers["官方名稱"]] or "").strip()
                    contract_start = normalize_code(row[headers["原始合約起始日期"]])
                    if not clinic_code:
                        continue
                    writer.writerow([group_name, clinic_code, clinic_name, address, official_name, contract_start])
                    count += 1

            sql_path.write_text(
                f"""
SET clinic.actor = 'import_medical_groups';
CREATE TEMP TABLE tmp_medical_groups (
  group_name TEXT,
  clinic_code TEXT,
  clinic_name TEXT,
  address TEXT,
  official_name TEXT,
  contract_start TEXT
);
\\copy tmp_medical_groups FROM '{csv_path}' WITH (FORMAT csv, HEADER true)

INSERT INTO meta.medical_groups (group_name)
SELECT DISTINCT group_name
FROM tmp_medical_groups
WHERE group_name <> ''
ON CONFLICT (group_name) DO UPDATE SET updated_at = now();

INSERT INTO meta.clinics (
  clinic_code, clinic_name, source_system, medical_group_id,
  official_name, institution_address, contract_start_date
)
SELECT
  t.clinic_code,
  COALESCE(NULLIF(t.clinic_name, ''), NULLIF(t.official_name, ''), t.clinic_code),
  '衛福部',
  g.medical_group_id,
  t.official_name,
  t.address,
  t.contract_start
FROM tmp_medical_groups t
LEFT JOIN meta.medical_groups g ON g.group_name = t.group_name
ON CONFLICT (clinic_code) DO UPDATE
SET medical_group_id = EXCLUDED.medical_group_id,
    clinic_name = COALESCE(NULLIF(meta.clinics.clinic_name, ''), EXCLUDED.clinic_name),
    official_name = EXCLUDED.official_name,
    institution_address = EXCLUDED.institution_address,
    contract_start_date = EXCLUDED.contract_start_date,
    updated_at = now();
""",
                encoding="utf-8",
            )
            run_psql(sql_path)
            return count
    finally:
        wb.close()


def main() -> None:
    parser = argparse.ArgumentParser(description="匯入衛福部醫療群資料到 clinic_merge PostgreSQL")
    parser.add_argument("xlsx", nargs="?", default="醫療群_衛福部資料.xlsx")
    args = parser.parse_args()
    count = import_medical_groups(Path(args.xlsx))
    print(f"已匯入/更新醫療群診所資料 {count} 筆")


if __name__ == "__main__":
    main()
