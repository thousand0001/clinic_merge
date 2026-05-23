#!/usr/bin/env python3
from __future__ import annotations

import csv
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext
from typing import Iterable


ENCODING_CANDIDATES = (
    "utf-16",
    "utf-16le",
    "utf-8-sig",
    "utf-8",
    "cp950",
    "big5",
    "gbk",
    "gb18030",
)


def detect_encoding(
    csv_path: str | Path,
    encodings: Iterable[str] = ENCODING_CANDIDATES,
) -> str:
    path = Path(csv_path)
    raw = path.read_bytes()

    if raw.startswith(b"\xff\xfe") or raw.startswith(b"\xfe\xff"):
        return "utf-16"
    if raw.startswith(b"\xef\xbb\xbf"):
        return "utf-8-sig"

    last_error: Exception | None = None
    for encoding in encodings:
        try:
            sample = raw.decode(encoding)
            if "\ufffd" in sample:
                continue
            return encoding
        except UnicodeDecodeError as exc:
            last_error = exc
    raise RuntimeError(f"無法判斷 {path.name} 的編碼: {last_error}") from last_error


def preview_csv(csv_path: Path, encoding: str, preview_rows: int) -> list[list[str]]:
    rows: list[list[str]] = []
    with csv_path.open("r", encoding=encoding, newline="") as f:
        reader = csv.reader(f)
        for idx, row in enumerate(reader, start=1):
            rows.append(row)
            if idx >= preview_rows:
                break
    return rows


def preview_csv_safe(csv_path: Path, encoding: str, preview_rows: int) -> tuple[bool, str]:
    try:
        rows = preview_csv(csv_path, encoding, preview_rows)
        return True, format_preview(rows)
    except Exception as exc:
        return False, f"讀取失敗：{exc}"


def format_preview(rows: list[list[str]]) -> str:
    lines: list[str] = []
    for idx, row in enumerate(rows, start=1):
        lines.append(f"{idx}: {row}")
    return "\n".join(lines) if lines else "(沒有可預覽內容)"


def main() -> int:
    root = tk.Tk()
    root.withdraw()

    csv_path_str = filedialog.askopenfilename(
        title="選擇要測試編碼的 CSV 檔案",
        filetypes=[("CSV files", "*.csv *.CSV"), ("All files", "*.*")],
    )
    if not csv_path_str:
        return 0

    csv_path = Path(csv_path_str).expanduser().resolve()
    if not csv_path.exists():
        messagebox.showerror("找不到檔案", f"找不到檔案：\n{csv_path}")
        return 1

    try:
        encoding = detect_encoding(csv_path)
    except Exception as exc:
        messagebox.showerror("測試失敗", str(exc))
        return 1

    preview_sections: list[str] = []
    seen: set[str] = set()
    ordered_candidates = [encoding, *ENCODING_CANDIDATES]
    for candidate in ordered_candidates:
        if candidate in seen:
            continue
        seen.add(candidate)
        ok, content = preview_csv_safe(csv_path, candidate, 8)
        status = "可讀" if ok else "失敗"
        preview_sections.append(
            f"===== {candidate} ({status}) =====\n{content}"
        )

    viewer = tk.Toplevel()
    viewer.title("CSV 編碼測試結果")
    viewer.geometry("1100x700")

    header = (
        f"檔案：{csv_path}\n"
        f"偵測編碼：{encoding}\n"
        "下方會列出多種編碼預覽，請直接看哪一段是正常中文。\n"
    )
    tk.Label(viewer, text=header, justify="left", anchor="w").pack(
        fill="x", padx=12, pady=(12, 6)
    )

    text = scrolledtext.ScrolledText(viewer, wrap="word")
    text.pack(fill="both", expand=True, padx=12, pady=(0, 12))
    text.insert("1.0", "\n\n".join(preview_sections))
    text.configure(state="disabled")

    viewer.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
