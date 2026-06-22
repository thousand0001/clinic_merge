#!/usr/bin/env python3
"""
Big5/CP950 文字編輯器 v2
- 雙 Tab：文字編輯 / 表格檢視（逗號自動對齊）
- 字元數統計（總數＋選取數）
- 表格支援搜尋、欄位排序
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import csv
import io
import sys

ENCODINGS = ["cp950", "utf-8", "utf-8-sig", "big5", "utf-16"]


class TextEditor:
    def __init__(self, root):
        self.root = root
        self.root.title("文字編輯器（CP950/Big5）")
        self.root.geometry("1100x750")

        self.current_file = None
        self.current_encoding = tk.StringVar(value="cp950")
        self.delimiter_var = tk.StringVar(value=",")
        self.modified = False
        self._all_rows = []
        self._sort_reverse = {}

        self._build_menu()
        self._build_toolbar()
        self._build_notebook()
        self._build_statusbar()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ── menu ───────────────────────────────────────────────────

    def _build_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="檔案", menu=file_menu)
        file_menu.add_command(label="開新檔案", accelerator="Ctrl+N", command=self._new_file)
        file_menu.add_command(label="開啟...", accelerator="Ctrl+O", command=self._open_file)
        file_menu.add_command(label="儲存", accelerator="Ctrl+S", command=self._save_file)
        file_menu.add_command(label="另存新檔...", accelerator="Ctrl+Shift+S", command=self._save_as)
        file_menu.add_separator()
        file_menu.add_command(label="以指定編碼重新開啟", command=self._reload_with_encoding)
        file_menu.add_separator()
        file_menu.add_command(label="結束", command=self._on_close)

        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="編輯", menu=edit_menu)
        edit_menu.add_command(label="復原", accelerator="Ctrl+Z",
                              command=lambda: self.text.event_generate("<<Undo>>"))
        edit_menu.add_command(label="重做", accelerator="Ctrl+Y",
                              command=lambda: self.text.event_generate("<<Redo>>"))
        edit_menu.add_separator()
        edit_menu.add_command(label="剪下", accelerator="Ctrl+X",
                              command=lambda: self.text.event_generate("<<Cut>>"))
        edit_menu.add_command(label="複製", accelerator="Ctrl+C",
                              command=lambda: self.text.event_generate("<<Copy>>"))
        edit_menu.add_command(label="貼上", accelerator="Ctrl+V",
                              command=lambda: self.text.event_generate("<<Paste>>"))
        edit_menu.add_separator()
        edit_menu.add_command(label="全選", accelerator="Ctrl+A", command=self._select_all)

        self.root.bind("<Control-n>", lambda e: self._new_file())
        self.root.bind("<Control-o>", lambda e: self._open_file())
        self.root.bind("<Control-s>", lambda e: self._save_file())
        self.root.bind("<Control-S>", lambda e: self._save_as())
        self.root.bind("<Control-a>", lambda e: self._select_all())

    # ── toolbar ────────────────────────────────────────────────

    def _build_toolbar(self):
        toolbar = ttk.Frame(self.root, relief="raised")
        toolbar.pack(side="top", fill="x", padx=2, pady=2)

        ttk.Button(toolbar, text="開啟", command=self._open_file, width=6).pack(side="left", padx=2)
        ttk.Button(toolbar, text="儲存", command=self._save_file, width=6).pack(side="left", padx=2)
        ttk.Button(toolbar, text="另存", command=self._save_as, width=6).pack(side="left", padx=2)

        ttk.Separator(toolbar, orient="vertical").pack(side="left", fill="y", padx=6)

        ttk.Label(toolbar, text="編碼：").pack(side="left")
        ttk.Combobox(toolbar, textvariable=self.current_encoding,
                     values=ENCODINGS, width=10, state="readonly").pack(side="left", padx=2)
        ttk.Button(toolbar, text="重新開啟", command=self._reload_with_encoding).pack(side="left", padx=4)

        ttk.Separator(toolbar, orient="vertical").pack(side="left", fill="y", padx=6)

        ttk.Label(toolbar, text="分隔符：").pack(side="left")
        ttk.Entry(toolbar, textvariable=self.delimiter_var, width=4).pack(side="left", padx=2)
        ttk.Button(toolbar, text="重整表格", command=self._refresh_table).pack(side="left", padx=2)

    # ── notebook ───────────────────────────────────────────────

    def _build_notebook(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True)
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        text_frame = ttk.Frame(self.notebook)
        self.notebook.add(text_frame, text="  文字編輯  ")
        self._build_text_tab(text_frame)

        table_frame = ttk.Frame(self.notebook)
        self.notebook.add(table_frame, text="  表格檢視  ")
        self._build_table_tab(table_frame)

    def _build_text_tab(self, parent):
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True)

        self.line_numbers = tk.Text(
            frame, width=5, padx=4, state="disabled",
            bg="#f0f0f0", fg="#888888", relief="flat",
            font=("Courier New", 11), cursor="arrow",
        )
        self.line_numbers.pack(side="left", fill="y")

        vscroll = ttk.Scrollbar(frame, orient="vertical")
        vscroll.pack(side="right", fill="y")
        hscroll = ttk.Scrollbar(parent, orient="horizontal")
        hscroll.pack(side="bottom", fill="x")

        self.text = tk.Text(
            frame, wrap="none", undo=True,
            font=("Courier New", 11),
            yscrollcommand=lambda *a: self._sync_yscroll(vscroll, *a),
            xscrollcommand=hscroll.set,
        )
        self.text.pack(side="left", fill="both", expand=True)

        vscroll.config(command=self._on_vscroll)
        hscroll.config(command=self.text.xview)

        self.text.bind("<<Modified>>", self._on_text_modified)
        self.text.bind("<KeyRelease>", lambda e: self._on_text_key())
        self.text.bind("<ButtonRelease>", lambda e: self._update_status())
        self.text.bind("<<Selection>>", lambda e: self._update_status())

    def _build_table_tab(self, parent):
        # 搜尋列
        search_bar = ttk.Frame(parent)
        search_bar.pack(side="top", fill="x", padx=6, pady=4)
        ttk.Label(search_bar, text="搜尋：").pack(side="left")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._filter_table())
        ttk.Entry(search_bar, textvariable=self.search_var, width=32).pack(side="left", padx=4)
        ttk.Button(search_bar, text="清除", command=lambda: self.search_var.set("")).pack(side="left")
        self.row_count_var = tk.StringVar(value="")
        ttk.Label(search_bar, textvariable=self.row_count_var, foreground="#555").pack(side="left", padx=12)

        # Treeview
        tv_frame = ttk.Frame(parent)
        tv_frame.pack(fill="both", expand=True)

        vsb = ttk.Scrollbar(tv_frame, orient="vertical")
        hsb = ttk.Scrollbar(tv_frame, orient="horizontal")
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")

        self.tree = ttk.Treeview(tv_frame, show="headings",
                                 yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.pack(fill="both", expand=True)
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)

        # 交替行顏色
        style = ttk.Style()
        style.configure("Treeview", rowheight=22, font=("Courier New", 10))
        style.configure("Treeview.Heading", font=("", 10, "bold"))

    # ── statusbar ──────────────────────────────────────────────

    def _build_statusbar(self):
        bar = ttk.Frame(self.root, relief="sunken")
        bar.pack(side="bottom", fill="x")

        self.status_var = tk.StringVar(value="就緒")
        ttk.Label(bar, textvariable=self.status_var, anchor="w").pack(side="left", padx=6)

        self.char_var = tk.StringVar(value="字元數: 0")
        ttk.Label(bar, textvariable=self.char_var).pack(side="right", padx=12)

        self.pos_var = tk.StringVar(value="行 1, 欄 1")
        ttk.Label(bar, textvariable=self.pos_var).pack(side="right", padx=6)

        ttk.Separator(bar, orient="vertical").pack(side="right", fill="y", pady=2)

    # ── scrolling ──────────────────────────────────────────────

    def _sync_yscroll(self, scrollbar, *args):
        scrollbar.set(*args)
        self.line_numbers.yview_moveto(args[0])

    def _on_vscroll(self, *args):
        self.text.yview(*args)
        self.line_numbers.yview(*args)

    # ── text tab helpers ───────────────────────────────────────

    def _on_text_key(self):
        self._update_line_numbers()
        self._update_status()

    def _update_line_numbers(self):
        content = self.text.get("1.0", "end-1c")
        lines = content.count("\n") + 1
        self.line_numbers.config(state="normal")
        self.line_numbers.delete("1.0", "end")
        self.line_numbers.insert("1.0", "\n".join(str(i) for i in range(1, lines + 1)))
        self.line_numbers.config(state="disabled")

    def _update_status(self):
        # 游標位置
        try:
            row, col = self.text.index("insert").split(".")
            self.pos_var.set(f"行 {row}, 欄 {int(col)+1}")
        except Exception:
            pass

        # 字元數
        total = len(self.text.get("1.0", "end-1c"))
        try:
            sel = self.text.get("sel.first", "sel.last")
            self.char_var.set(f"字元數: {total:,}  │  選取: {len(sel):,}")
        except tk.TclError:
            self.char_var.set(f"字元數: {total:,}")

        # 標題列＋狀態列
        enc = self.current_encoding.get()
        fname = os.path.basename(self.current_file) if self.current_file else "未命名"
        mod = " *" if self.modified else ""
        self.status_var.set(f"{fname}{mod}  │  編碼: {enc}")
        self.root.title(f"{'*' if self.modified else ''}{fname} — 文字編輯器")

    def _on_text_modified(self, event=None):
        if self.text.edit_modified():
            self.modified = True
            self._update_status()
            self.text.edit_modified(False)

    # ── tab switch ─────────────────────────────────────────────

    def _on_tab_changed(self, event):
        if self.notebook.index("current") == 1:
            self._refresh_table()

    # ── table view ─────────────────────────────────────────────

    def _refresh_table(self):
        content = self.text.get("1.0", "end-1c").strip()
        if not content:
            return

        delim = self.delimiter_var.get() or ","
        if delim in ("\\t", "tab"):
            delim = "\t"

        try:
            rows = list(csv.reader(io.StringIO(content), delimiter=delim))
        except Exception as e:
            messagebox.showerror("CSV 解析錯誤", str(e))
            return

        if not rows:
            return

        # 跳過前導標題行（不含分隔符的行，如「費用年月:11401」）
        self._title_lines = []
        header_idx = 0
        for i, row in enumerate(rows):
            if len(row) > 1:
                header_idx = i
                break
            self._title_lines.append(rows[i][0] if rows[i] else "")
        else:
            # 全部都只有一欄，直接用第一行
            header_idx = 0
            self._title_lines = []

        headers = [h.strip() for h in rows[header_idx]]
        data_rows = rows[header_idx + 1:]
        # 去除完全空白的尾行
        data_rows = [r for r in data_rows if any(v.strip() for v in r)]

        # 重建欄位
        self._sort_reverse = {}
        self.tree["columns"] = headers
        for h in headers:
            col_idx = headers.index(h)
            # 計算此欄最大寬度（取前 200 列估算）
            sample = [str(r[col_idx]) if col_idx < len(r) else "" for r in ([headers] + data_rows[:200])]
            max_chars = max((len(v) for v in sample), default=4)
            width = max(min(max_chars * 8 + 16, 320), 50)
            self.tree.heading(h, text=h,
                              command=lambda c=h: self._sort_column(c))
            self.tree.column(h, width=width, minwidth=40, anchor="w")

        self._all_rows = data_rows
        self.search_var.set("")
        # 顯示前導標題（如「費用年月:11401」）於搜尋列右側
        title_text = "  ".join(t.strip() for t in self._title_lines if t.strip())
        self.row_count_var.set(f"{title_text}　" if title_text else "")
        self._populate_tree(data_rows)

    def _populate_tree(self, rows):
        self.tree.delete(*self.tree.get_children())
        for i, row in enumerate(rows):
            self.tree.insert("", "end", values=row,
                             tags=("even" if i % 2 == 0 else "odd",))
        self.tree.tag_configure("even", background="#ffffff")
        self.tree.tag_configure("odd", background="#f2f6fb")
        title_text = "  ".join(t.strip() for t in getattr(self, "_title_lines", []) if t.strip())
        prefix = f"{title_text}　" if title_text else ""
        self.row_count_var.set(f"{prefix}共 {len(rows):,} 筆")

    def _filter_table(self):
        keyword = self.search_var.get().lower()
        if not keyword:
            self._populate_tree(self._all_rows)
        else:
            filtered = [r for r in self._all_rows
                        if any(keyword in str(v).lower() for v in r)]
            self._populate_tree(filtered)

    def _sort_column(self, col):
        rev = self._sort_reverse.get(col, False)
        data = [(self.tree.set(k, col), k) for k in self.tree.get_children("")]
        try:
            data.sort(key=lambda t: float(t[0]) if t[0] else float("-inf"), reverse=rev)
        except ValueError:
            data.sort(key=lambda t: t[0], reverse=rev)
        for i, (_, k) in enumerate(data):
            self.tree.move(k, "", i)
            self.tree.item(k, tags=("even" if i % 2 == 0 else "odd",))
        self._sort_reverse[col] = not rev
        arrow = " ↑" if rev else " ↓"
        # 重設所有欄標題，只在當前欄加箭頭
        for h in self.tree["columns"]:
            label = h + (arrow if h == col else "")
            self.tree.heading(h, text=label,
                              command=lambda c=h: self._sort_column(c))

    # ── file operations ────────────────────────────────────────

    def _new_file(self):
        if not self._confirm_discard():
            return
        self.text.delete("1.0", "end")
        self.current_file = None
        self.modified = False
        self._update_line_numbers()
        self._update_status()

    def _open_file(self):
        if not self._confirm_discard():
            return
        path = filedialog.askopenfilename(
            filetypes=[("文字/CSV", "*.txt *.csv"), ("所有檔案", "*.*")]
        )
        if path:
            self._load_file(path)

    def _load_file(self, path):
        enc = self.current_encoding.get()
        try:
            with open(path, "r", encoding=enc, errors="replace") as f:
                content = f.read()
            self.text.delete("1.0", "end")
            self.text.insert("1.0", content)
            self.current_file = path
            self.modified = False
            self.text.edit_modified(False)
            self._update_line_numbers()
            self._update_status()
        except Exception as e:
            messagebox.showerror("開啟失敗", f"無法以 {enc} 開啟：\n{e}")

    def _reload_with_encoding(self):
        if not self.current_file:
            messagebox.showinfo("提示", "請先開啟一個檔案")
            return
        if self.modified and not messagebox.askyesno("確認", "有未儲存的變更，重新開啟會遺失，繼續？"):
            return
        self._load_file(self.current_file)

    def _save_file(self):
        if self.current_file:
            self._write_file(self.current_file)
        else:
            self._save_as()

    def _save_as(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("文字檔", "*.txt"), ("CSV", "*.csv"), ("所有檔案", "*.*")]
        )
        if path:
            self._write_file(path)
            self.current_file = path

    def _write_file(self, path):
        enc = self.current_encoding.get()
        try:
            content = self.text.get("1.0", "end-1c")
            with open(path, "w", encoding=enc, errors="replace") as f:
                f.write(content)
            self.modified = False
            self._update_status()
        except Exception as e:
            messagebox.showerror("儲存失敗", f"無法以 {enc} 儲存：\n{e}")

    def _select_all(self):
        self.text.tag_add("sel", "1.0", "end")
        self._update_status()

    def _confirm_discard(self):
        if self.modified:
            return messagebox.askyesno("確認", "有未儲存的變更，繼續會遺失，確定？")
        return True

    def _on_close(self):
        if self._confirm_discard():
            self.root.destroy()


def main():
    root = tk.Tk()
    editor = TextEditor(root)
    if len(sys.argv) > 1:
        editor._load_file(sys.argv[1])
    root.mainloop()


if __name__ == "__main__":
    main()
