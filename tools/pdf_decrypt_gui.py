#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from pypdf import PdfReader, PdfWriter


def open_path(target: Path) -> None:
    if sys.platform.startswith("darwin"):
        subprocess.Popen(["open", str(target)])
        return
    if os.name == "nt":
        os.startfile(str(target))  # type: ignore[attr-defined]
        return
    subprocess.Popen(["xdg-open", str(target)])


def decrypt_pdf_file(pdf_path: Path, password: str, output_path: Path) -> Path:
    reader = PdfReader(str(pdf_path))
    if not reader.is_encrypted:
        raise ValueError(f"{pdf_path.name} 不是加密 PDF，不需要解密。")

    result = reader.decrypt(password)
    if result == 0:
        raise ValueError(f"{pdf_path.name} 密碼錯誤，無法解密。")

    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("wb") as fh:
        writer.write(fh)
    return output_path


def collect_pdf_files(folder_path: Path) -> list[Path]:
    return sorted(
        path
        for path in folder_path.rglob("*.pdf")
        if path.is_file() and "解碼" not in path.parts
    )


class PDFDecryptApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("PDF 解密工具")
        self.root.geometry("620x300")
        self.root.resizable(False, False)

        self.selection_type = tk.StringVar(value="file")
        self.selected_path = tk.StringVar()
        self.password = tk.StringVar()
        self.status = tk.StringVar(value="請先選擇 PDF 檔案或資料夾。")

        self._build_ui()

    def _build_ui(self) -> None:
        frame = ttk.Frame(self.root, padding=18)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="選擇方式").grid(row=0, column=0, sticky="w")
        radio_frame = ttk.Frame(frame)
        radio_frame.grid(row=0, column=1, sticky="w", pady=(0, 12))
        ttk.Radiobutton(
            radio_frame,
            text="單一 PDF 檔案",
            variable=self.selection_type,
            value="file",
            command=self.clear_selection,
        ).pack(side="left", padx=(0, 16))
        ttk.Radiobutton(
            radio_frame,
            text="整個資料夾",
            variable=self.selection_type,
            value="folder",
            command=self.clear_selection,
        ).pack(side="left")

        ttk.Label(frame, text="已選路徑").grid(row=1, column=0, sticky="nw")
        path_frame = ttk.Frame(frame)
        path_frame.grid(row=1, column=1, sticky="ew")
        self.path_label = ttk.Label(
            path_frame,
            textvariable=self.selected_path,
            relief="sunken",
            anchor="w",
            width=46,
        )
        self.path_label.grid(row=0, column=0, sticky="ew")
        path_frame.columnconfigure(0, weight=1)
        ttk.Button(frame, text="選取檔案或資料夾", command=self.browse).grid(
            row=2, column=1, sticky="w", pady=(10, 0)
        )

        ttk.Label(frame, text="PDF 密碼").grid(row=3, column=0, sticky="w", pady=(16, 0))
        ttk.Entry(frame, textvariable=self.password, show="*", width=34).grid(
            row=3, column=1, sticky="w", pady=(16, 0)
        )

        ttk.Button(frame, text="開始解密", command=self.run_decrypt).grid(
            row=4, column=1, sticky="w", pady=(22, 0)
        )

        ttk.Label(
            frame,
            text="路徑請用右側「選取」按鈕開啟視窗挑選。輸出會建立在「解碼」子資料夾。",
            foreground="#555555",
        ).grid(row=5, column=0, columnspan=2, sticky="w", pady=(18, 6))

        status_label = ttk.Label(
            frame,
            textvariable=self.status,
            wraplength=560,
            foreground="#0b5f2a",
            justify="left",
        )
        status_label.grid(row=6, column=0, columnspan=2, sticky="w")

        frame.columnconfigure(1, weight=1)

    def clear_selection(self) -> None:
        self.selected_path.set("")
        if self.selection_type.get() == "file":
            self.status.set("已切換為單一檔案模式，請選擇 PDF。")
        else:
            self.status.set("已切換為資料夾模式，會批次處理資料夾內所有 PDF。")

    def browse(self) -> None:
        if self.selection_type.get() == "file":
            path = filedialog.askopenfilename(
                title="選擇 PDF 檔案",
                filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")],
            )
        else:
            path = filedialog.askdirectory(title="選擇資料夾")

        if path:
            self.selected_path.set(path)
            self.status.set(f"已選取：{path}")

    def run_decrypt(self) -> None:
        selected = self.selected_path.get().strip()
        password = self.password.get()

        if not selected:
            messagebox.showwarning("尚未選取", "請先選擇 PDF 檔案或資料夾。")
            return
        if not password:
            messagebox.showwarning("缺少密碼", "請輸入 PDF 密碼。")
            return

        target_path = Path(selected)
        try:
            if self.selection_type.get() == "file":
                output_file = decrypt_pdf_file(
                    target_path,
                    password,
                    target_path.parent / "解碼" / target_path.name,
                )
                self.status.set(f"解密完成：{output_file}")
                messagebox.showinfo("完成", f"已解密完成：\n{output_file}")
                open_path(output_file)
                return

            pdf_files = collect_pdf_files(target_path)
            if not pdf_files:
                raise ValueError("選取的資料夾內找不到 PDF 檔案。")

            output_dir = target_path / "解碼"
            success_count = 0
            failed: list[str] = []

            for pdf_file in pdf_files:
                try:
                    relative_path = pdf_file.relative_to(target_path)
                    decrypt_pdf_file(pdf_file, password, output_dir / relative_path)
                    success_count += 1
                except Exception as exc:
                    failed.append(f"{pdf_file.name}: {exc}")

            if success_count == 0:
                raise ValueError("沒有任何 PDF 成功解密。\n" + "\n".join(failed[:10]))

            self.status.set(f"批次完成：成功 {success_count} 個，輸出資料夾：{output_dir}")
            summary = [f"成功解密 {success_count} 個 PDF。", f"輸出資料夾：\n{output_dir}"]
            if failed:
                summary.append("\n以下檔案失敗：")
                summary.extend(failed[:10])
                if len(failed) > 10:
                    summary.append(f"... 其餘 {len(failed) - 10} 個失敗檔案略過顯示")
            messagebox.showinfo("批次完成", "\n".join(summary))
            open_path(output_dir)
        except Exception as exc:
            self.status.set(f"解密失敗：{exc}")
            messagebox.showerror("解密失敗", str(exc))


def main() -> None:
    root = tk.Tk()
    style = ttk.Style(root)
    if "clam" in style.theme_names():
        style.theme_use("clam")
    PDFDecryptApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
