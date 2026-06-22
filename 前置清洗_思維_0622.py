# -*- coding: utf-8 -*-
"""
思維診所專用前置清洗（0622）

思維的醫聖月份 TXT 在電話後多一個空白欄，申請額實際位於
"年齡"欄後一欄。本程式只在思維流程修正此欄位，再交由醫聖
前置清洗與共用核心產生正式 Excel。
"""

from __future__ import annotations

import importlib.util
import re
import sys
from pathlib import Path
from types import ModuleType
from typing import Optional


SCRIPT_DIR = Path(__file__).resolve().parent
TW_ID_PATTERN = re.compile(r"(?:[A-Z][1289]\d{8}|[A-Z][A-D]\d{8})")


def _find_medical_saint_script() -> Path:
    def date_key(path: Path) -> tuple[int, str]:
        match = re.search(r"(\d{4})(?=\.py$)", path.name)
        return (int(match.group(1)) if match else -1, path.name)

    candidates = sorted(
        SCRIPT_DIR.glob("前置清洗_醫聖_*.py"),
        key=date_key,
        reverse=True,
    )
    if not candidates:
        raise RuntimeError("找不到醫聖前置清洗程式。")
    return candidates[0]


def _load_medical_saint_module() -> ModuleType:
    script_path = _find_medical_saint_script()
    spec = importlib.util.spec_from_file_location("medical_saint_preclean_for_siwei", script_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"無法載入醫聖前置清洗程式：{script_path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def _apply_siwei_rules(module: ModuleType) -> None:
    module.TW_ID_RE = TW_ID_PATTERN

    def extract_siwei_claim_amount(row: list[str]) -> str:
        for idx, value in enumerate(row):
            if not module.AGE_TEXT_RE.fullmatch(str(value).strip()):
                continue
            if idx + 1 >= len(row):
                return ""
            return module._clean_claim_amount(row[idx + 1])
        return ""

    module._extract_shifted_claim_amount = extract_siwei_claim_amount


def process_excel(source_path: str, template_path: Optional[str] = None) -> str:
    module = _load_medical_saint_module()
    _apply_siwei_rules(module)
    return module.process_excel(source_path, template_path)


def main() -> None:
    module = _load_medical_saint_module()
    _apply_siwei_rules(module)

    import tkinter as tk
    from tkinter import filedialog, messagebox

    root = tk.Tk()
    root.withdraw()

    source_dir = filedialog.askdirectory(title="選擇思維診所來源資料夾")
    if not source_dir:
        messagebox.showinfo("提示", "未選擇資料夾，程式已結束。")
        return

    try:
        output_path = module.process_excel(source_dir)
        messagebox.showinfo("完成", f"已輸出：\n{output_path}")
        module._load_generic_module().open_file_cross_platform(output_path)
    except Exception as exc:
        messagebox.showerror("錯誤", str(exc))


if __name__ == "__main__":
    main()
