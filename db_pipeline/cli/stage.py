# -*- coding: utf-8 -*-
"""
db_pipeline 上傳入口（parse → validate → stage）

支援：
- 命令列模式：python -m db_pipeline.cli.stage --source-dir ... --config ...
- GUI 模式：直接執行，跳出資料夾選擇視窗
- --dry-run：只解析不寫入 DB

流程：
1. 選擇診所來源資料夾
2. 讀取診所設定（JSON）
3. 解析原始資料 → DatasetBundle
4. 驗證 DatasetBundle
5. 寫入 staging.* (單一 transaction)
6. 顯示結果摘要
"""
from __future__ import annotations

import argparse
import datetime
import json
import socket
import uuid
from pathlib import Path
from typing import Optional, Sequence

from db_pipeline.config.models import ClinicConfig, load_clinic_config
from db_pipeline.parsers import PARSER_REGISTRY, get_parser
from db_pipeline.raw import collect_raw_sources
from db_pipeline.storage import PostgresStagingWriter
from db_pipeline.validation.models import ValidationIssue
from db_pipeline.validation.validator import validate_bundle

TZ_TW = datetime.timezone(datetime.timedelta(hours=8))
SCRIPT_DIR = Path(__file__).resolve().parent.parent  # db_pipeline/


# ── GUI 工具 ──────────────────────────────────────────────────────────────────
def _pick_directory(title: str, initial: Optional[Path] = None) -> Optional[Path]:
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        root.update()
        path = filedialog.askdirectory(
            title=title,
            initialdir=str(initial or Path.home()),
            mustexist=True,
        )
        root.destroy()
        return Path(path) if path else None
    except Exception:
        text = input(f"{title}（路徑）：").strip()
        return Path(text) if text else None


def _pick_file(title: str, initial: Optional[Path] = None) -> Optional[Path]:
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        root.update()
        path = filedialog.askopenfilename(
            title=title,
            initialdir=str(initial or SCRIPT_DIR),
            filetypes=[("JSON 設定檔", "*.json"), ("所有檔案", "*.*")],
        )
        root.destroy()
        return Path(path) if path else None
    except Exception:
        text = input(f"{title}（路徑）：").strip()
        return Path(text) if text else None


def _show_result(summary: dict, dry_run: bool) -> None:
    try:
        import tkinter as tk
        from tkinter import messagebox

        mode = "【試跑模式，未寫入 DB】" if dry_run else "【已寫入 staging】"
        counts = summary.get("dataset_counts", {})
        issues = summary.get("validation", {}).get("issues", [])
        skipped = summary.get("coverage", {}).get("skipped_files", {})
        errors   = [i for i in issues if i["severity"] == "error"]
        warnings = [i for i in issues if i["severity"] == "warning"]
        total_errors   = summary.get("validation", {}).get("error_count", len(errors))
        total_warnings = summary.get("validation", {}).get("warning_count", len(warnings))

        lines = [
            mode,
            f"診所：{summary.get('clinic_name', '')}（{summary.get('clinic_code', '')}）",
            f"系統：{summary.get('source_system', '')}",
            f"批次：{summary.get('batch_id', '')}",
            "",
            "── 解析結果 ──",
            f"  發現檔案：{summary['coverage']['discovered_files']}",
            f"  成功解析：{summary['coverage']['parsed_files']}",
            (
                "  僅進 raw、未進結構化統計："
                f"{len(summary['coverage']['skipped_files'])}"
            ),
            "",
            "── 資料筆數 ──",
        ] + [f"  {k}: {v}" for k, v in counts.items() if v > 0] + [
            "",
            f"── 驗證：{'✅ 通過' if not errors else f'❌ {total_errors} 個錯誤'} "
            f"{'⚠️ ' + str(total_warnings) + ' 個警告' if total_warnings else ''}──",
        ]
        if errors:
            lines += [""] + [f"  ❌ {e['message']}" for e in errors[:5]]
        if warnings:
            lines += ["", "── 警告內容 ──"]
            lines += [f"  ⚠️ {item['message']}" for item in warnings[:10]]
        if skipped:
            lines += ["", "── 僅進 raw、未進結構化統計 ──"]
            lines += [
                f"  • {name}：{reason}"
                for name, reason in list(skipped.items())[:15]
            ]

        root = tk.Tk()
        root.withdraw()
        if errors:
            messagebox.showerror("上傳結果", "\n".join(lines))
        else:
            messagebox.showinfo("上傳結果", "\n".join(lines))
        root.destroy()
    except Exception:
        print(json.dumps(summary, ensure_ascii=False, indent=2))


# ── 主流程 ────────────────────────────────────────────────────────────────────
def run(
    source_dir: Path,
    config: ClinicConfig,
    batch_id: str,
    dry_run: bool = False,
    output: Optional[Path] = None,
) -> dict:
    print(f"[1/5] 保存原始來源列索引…")
    raw_result = collect_raw_sources(source_dir)

    print(f"[2/5] 解析 {source_dir.name}（{config.source_system}）…")
    parser = get_parser(config.source_system)
    parse_result = parser.parse(source_dir, config, batch_id)
    parse_result.bundle.raw_source_files = raw_result.source_files
    parse_result.bundle.raw_source_rows = raw_result.rows
    raw_issues = [
        ValidationIssue(
            severity="warning",
            dataset="raw_source_rows",
            code="raw_rows_not_expanded",
            message=f"{file_name}：{reason}",
            source_file=file_name,
        )
        for file_name, reason in raw_result.unreadable_files.items()
    ]

    print(f"[3/5] 驗證資料…")
    validation = validate_bundle(parse_result.bundle)
    all_issues = parse_result.issues + raw_issues + validation.issues
    is_valid = not any(i.severity == "error" for i in all_issues)

    counts = parse_result.bundle.counts()
    summary = {
        "batch_id":     batch_id,
        "clinic_code":  config.clinic_code,
        "clinic_name":  config.clinic_name,
        "source_system": config.source_system,
        "source_dir":   str(source_dir),
        "dry_run":      dry_run,
        "dataset_counts": counts,
        "coverage": {
            "discovered_files": parse_result.coverage.discovered_files,
            "parsed_files":     parse_result.coverage.parsed_files,
            "skipped_files":    parse_result.coverage.skipped_files,
            "parsed_rows":      parse_result.coverage.parsed_rows,
            "unmatched_rows":   parse_result.coverage.unmatched_rows,
            "unlinked_rows":    parse_result.coverage.unlinked_rows,
            "raw_unreadable_files": raw_result.unreadable_files,
        },
        "validation": {
            "is_valid":      is_valid,
            "issue_count":   len(all_issues),
            "error_count":   sum(1 for i in all_issues if i.severity == "error"),
            "warning_count": sum(1 for i in all_issues if i.severity == "warning"),
            "issues": [
                {
                    "severity":     i.severity,
                    "dataset":      i.dataset,
                    "code":         i.code,
                    "message":      i.message,
                    "source_file":  i.source_file,
                    "source_row":   i.source_row,
                }
                for i in all_issues[:50]
            ],
        },
        "staged": False,
    }

    if not dry_run and is_valid:
        print(f"[4/5] 寫入 raw 與 staging…")
        writer = PostgresStagingWriter()
        clinic_id = writer.get_clinic_id(config.clinic_code)
        # clinic_name 空白時從 DB 補
        if not config.clinic_name:
            from db_pipeline.storage import _run_query
            db_name = _run_query(
                f"SELECT clinic_name FROM meta.clinics "
                f"WHERE clinic_code='{config.clinic_code}' LIMIT 1;"
            )
            if db_name:
                import dataclasses
                config = dataclasses.replace(config, clinic_name=db_name)
            summary["clinic_name"] = config.clinic_name
        # 合併所有 issues 傳給 writer，讓驗證紀錄完整寫入 meta.validation_errors
        from db_pipeline.validation.models import ValidationReport
        full_report = ValidationReport(issues=all_issues, dataset_counts=counts)
        stage_result = writer.stage(
            clinic_id=clinic_id,
            batch_id=batch_id,
            bundle=parse_result.bundle,
            validation_report=full_report,
            source_system=config.source_system,
            source_root=str(source_dir),
            requested_by=socket.gethostname(),
        )
        summary["staged"] = True
        summary["staged_counts"] = stage_result.staged_counts
        print(f"[5/5] 完成！batch_id={batch_id}")
    elif dry_run:
        print(f"[4/5] 試跑模式，跳過寫入。")
        print(f"[5/5] 完成！")
    else:
        print(f"[4/5] 驗證失敗，跳過寫入。")
        print(f"[5/5] 完成（未寫入）。")

    if output:
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(summary, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    return summary


# ── CLI 入口 ──────────────────────────────────────────────────────────────────
def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="解析診所資料並寫入 staging（parse → validate → stage）")
    parser.add_argument("--source-dir", type=Path, help="診所來源資料夾")
    parser.add_argument("--config",     type=Path, help="診所 JSON 設定檔")
    parser.add_argument("--batch-id",   default="", help="批次 UUID（空字串 = 自動產生）")
    parser.add_argument("--dry-run",    action="store_true", help="只解析不寫入 DB")
    parser.add_argument("--output",     type=Path, help="摘要 JSON 輸出路徑")
    args = parser.parse_args(argv)

    # 來源資料夾
    source_dir = args.source_dir
    if source_dir is None:
        source_dir = _pick_directory("選擇診所來源資料夾")
    if not source_dir or not source_dir.is_dir():
        print("未選擇來源資料夾，程式已取消。")
        return 1
    source_dir = source_dir.resolve()

    # 設定檔
    config_path = args.config
    if config_path is None:
        # 先找同名資料夾名稱對應的 JSON
        candidates = list((SCRIPT_DIR / "config" / "examples").glob("*.json"))
        config_path = _pick_file(
            "選擇診所設定 JSON 檔",
            initial=SCRIPT_DIR / "config" / "examples",
        )
    if not config_path or not config_path.is_file():
        print("未選擇設定檔，程式已取消。")
        return 1
    config = load_clinic_config(config_path)

    # batch_id：必須是合法 UUID；若傳入非 UUID 字串則轉成 uuid5（可重現）
    raw_bid = args.batch_id.strip()
    if not raw_bid:
        batch_id = str(uuid.uuid4())
    else:
        try:
            batch_id = str(uuid.UUID(raw_bid))  # 驗證格式
        except ValueError:
            # 非 UUID 字串 → 轉成確定性 UUID（uuid5）
            # 加入 clinic_code，防止不同診所同名批次產生相同 UUID
            batch_id = str(uuid.uuid5(uuid.NAMESPACE_DNS, f"{config.clinic_code}/{raw_bid}"))

    summary = run(
        source_dir=source_dir,
        config=config,
        batch_id=batch_id,
        dry_run=args.dry_run,
        output=args.output,
    )

    _show_result(summary, dry_run=args.dry_run)
    return 0 if summary["validation"]["is_valid"] else 2


if __name__ == "__main__":
    raise SystemExit(main())
