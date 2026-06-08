# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import importlib.util
from pathlib import Path
from typing import Optional


PROJECT_DIR = Path(__file__).resolve().parent
DEFAULT_SOURCE_DIR = Path.home()  # 資料夾由使用者透過選擇視窗指定，此為初始目錄提示
COMMON_SCRIPT = PROJECT_DIR / "資料庫輸出0601.py"


def load_common_module():
    spec = importlib.util.spec_from_file_location("clinic_common_output_0601", COMMON_SCRIPT)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"無法載入通用程式：{COMMON_SCRIPT}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def choose_source_dir(initial_dir: Path) -> Optional[Path]:
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.update()
        selected = filedialog.askdirectory(
            title="選擇陳森豊診所資料夾",
            initialdir=str(initial_dir if initial_dir.exists() else PROJECT_DIR),
            mustexist=True,
        )
        root.destroy()
        return Path(selected) if selected else None
    except Exception:
        text = input(f"請輸入資料夾路徑（直接 Enter 使用預設：{initial_dir}）：").strip()
        return Path(text) if text else initial_dir


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="選擇診所資料夾，上傳 PostgreSQL 並產生會員 Excel")
    parser.add_argument(
        "--source-dir",
        type=Path,
        help="診所原始資料資料夾；未指定時會跳出資料夾選擇視窗",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        help="Excel 輸出資料夾；未指定時預設為來源資料夾的上一層",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    source_dir = args.source_dir or choose_source_dir(DEFAULT_SOURCE_DIR)
    if source_dir is None:
        print("已取消，沒有選擇資料夾。")
        return 1
    source_dir = source_dir.expanduser().resolve()
    if not source_dir.exists() or not source_dir.is_dir():
        raise FileNotFoundError(f"找不到資料夾：{source_dir}")
    output_dir = (args.output_dir.expanduser().resolve() if args.output_dir else source_dir.parent)

    print(f"資料來源資料夾：{source_dir}")
    print(f"輸出資料夾：{output_dir}")
    print("資料庫：clinic_merge（可用 CLINIC_DB_NAME 環境變數覆蓋）")

    common = load_common_module()
    output_path = common.process(
        source_dir,
        clinic_code="3501110080",
        clinic_name="陳森豊",
        output_dir=output_dir,
        actor="thousand0001",
    )
    print(output_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
