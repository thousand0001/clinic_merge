#!/usr/bin/env python3
"""
FM.txt 檢視 / 編輯器
健保署 MHB5 規格：208 bytes/筆，CP950，CRLF

驗證顏色：
  紅   (error)      行長 ≠ 208 bytes
  橙   (warn_value) 固定欄位值錯誤（A/PLAN_NO/BRANCH）或 CLOSE 欄位非空白
  黃   (warn_blank) 必填欄位空白（PID/NAME/BIRTHDAY 等）
  綠   (ok)         全部正確
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import sys

ENCODING   = "cp950"
RECORD_LEN = 208

FM_FIELDS = [
    ("A",          1,  "固定 A"),
    ("PLAN_NO",    2,  "計畫代號"),
    ("BRANCH",     1,  "分支"),
    ("HOSP_ID",   10,  "院所代碼"),
    ("PID",       10,  "身份證號"),
    ("BIRTHDAY",   8,  "生日"),
    ("NAME",      12,  "姓名"),
    ("SEX",        1,  "性別"),
    ("ADDR",     120,  "地址"),
    ("TEL",       15,  "電話"),
    ("PRSN_ID",   10,  "醫事人員"),
    ("CASE_TYPE",  1,  "案件類型"),
    ("CASE_DATE",  8,  "案件日期"),
    ("CLOSE_DATE", 8,  "終止日"),
    ("CLOSE_RSN",  1,  "終止原因"),
]

assert sum(f[1] for f in FM_FIELDS) == RECORD_LEN

# 固定值欄位（應等於指定值）
FIXED_VALUES  = [("A", "A"), ("PLAN_NO", "17"), ("BRANCH", "1")]
# 必填欄位（不可全空白）
REQUIRED      = ("HOSP_ID", "PID", "BIRTHDAY", "NAME", "SEX", "CASE_TYPE", "CASE_DATE")
# 必須空白的欄位
MUST_BLANK    = ("CLOSE_DATE", "CLOSE_RSN")

# 嚴重性標籤（同時作為 Treeview tag 名稱）
SEV_OK    = "ok"
SEV_BLANK = "warn_blank"
SEV_VALUE = "warn_value"
SEV_ERROR = "error"

_SEV_RANK = {SEV_OK: 0, SEV_BLANK: 1, SEV_VALUE: 2, SEV_ERROR: 3}


def _worst(a: str, b: str) -> str:
    return a if _SEV_RANK[a] >= _SEV_RANK[b] else b


# ── 工具函式 ───────────────────────────────────────────────────────────────────

def _field_offset(name: str) -> tuple[int, int]:
    off = 0
    for n, w, _ in FM_FIELDS:
        if n == name:
            return off, w
        off += w
    raise KeyError(name)


def parse_record(raw: bytes) -> dict:
    fields, off = {}, 0
    for name, width, _ in FM_FIELDS:
        chunk = raw[off: off + width]
        if len(chunk) < width:
            chunk = chunk.ljust(width, b" ")
        fields[name] = chunk.decode(ENCODING, errors="replace")
        off += width
    return fields


def encode_field(value: str, width: int) -> bytes:
    b = value.encode(ENCODING, errors="replace")
    if len(b) > width:
        b = b[:width]
        if b and b[-1] >= 0x81:   # 不在 CP950 lead byte 中間截斷
            b = b[:-1] + b" "
    return b.ljust(width, b" ")


def validate_record(raw: bytes, fields: dict) -> tuple[str, list[str]]:
    """Returns (severity_tag, [human-readable messages])."""
    if len(raw) != RECORD_LEN:
        return SEV_ERROR, [f"字元數錯誤：行長 {len(raw)} bytes（應為 {RECORD_LEN}）"]

    msgs, sev = [], SEV_OK

    # 固定值錯誤 → 橙
    for fname, expected in FIXED_VALUES:
        val = fields.get(fname, "").strip()
        if not val:
            msgs.append(f"{fname}：空白（應為 '{expected}'）")
            sev = _worst(sev, SEV_VALUE)
        elif val != expected:
            msgs.append(f"{fname}='{val}'（應為 '{expected}'）")
            sev = _worst(sev, SEV_VALUE)

    # 應為空白卻有值 → 橙
    for fname in MUST_BLANK:
        if fields.get(fname, "").strip():
            msgs.append(f"{fname}：應為空白，實為 '{fields[fname].strip()}'")
            sev = _worst(sev, SEV_VALUE)

    # 必填欄位空白 → 黃
    for fname in REQUIRED:
        if not fields.get(fname, "").strip():
            msgs.append(f"{fname}：空白")
            sev = _worst(sev, SEV_BLANK)

    return sev, msgs


# ── 主視窗 ─────────────────────────────────────────────────────────────────────

class FMEditor:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("FM.txt 編輯器")
        self.root.geometry("1300x760")

        self.current_file: str | None = None
        self.raw_lines: list[bytes]   = []
        self.modified                 = False
        self._table_data: list[tuple[int, dict, bytes]] = []
        self._undo_stack: list[list[bytes]] = []
        self._show_sep = tk.BooleanVar(value=True)

        # pack 順序：bottom 先 pack 的最靠下
        self._build_toolbar()
        self._build_statusbar()     # 最底
        self._build_delete_panel()  # 狀態列正上方
        self._build_notebook()      # 填滿剩餘空間

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ── 工具列 ─────────────────────────────────────────────────

    def _build_toolbar(self):
        bar = ttk.Frame(self.root, relief="raised")
        bar.pack(side="top", fill="x", padx=2, pady=2)

        ttk.Button(bar, text="開啟", command=self._open_file, width=6).pack(side="left", padx=2)
        ttk.Button(bar, text="儲存", command=self._save_file, width=6).pack(side="left", padx=2)
        ttk.Button(bar, text="另存", command=self._save_as,  width=6).pack(side="left", padx=2)

        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=6)
        ttk.Checkbutton(bar, text="原始檢視顯示欄位分隔線",
                        variable=self._show_sep,
                        command=self._refresh_raw).pack(side="left", padx=4)
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=6)

        self.err_btn = ttk.Button(bar, text="⚠ 檢查中…", command=self._jump_next_error)
        self.err_btn.pack(side="left", padx=4)

    # ── 刪除列面板 ──────────────────────────────────────────────

    def _build_delete_panel(self):
        bar = ttk.Frame(self.root, relief="groove")
        bar.pack(side="bottom", fill="x", padx=2, pady=(0, 1))

        ttk.Label(bar, text="刪除列：").pack(side="left", padx=(8, 2), pady=5)
        self.del_entry = ttk.Entry(bar, width=26, font=("Courier New", 10))
        self.del_entry.pack(side="left", padx=2)
        ttk.Label(bar, text="（如：2,4,7 或 3-5）", foreground="#888").pack(side="left", padx=(2, 8))

        ttk.Button(bar, text="刪除", command=self._delete_rows, width=6).pack(side="left", padx=2)
        self.undo_btn = ttk.Button(bar, text="復原", command=self._undo,
                                   state="disabled", width=6)
        self.undo_btn.pack(side="left", padx=4)

        self.del_info = tk.StringVar(value="")
        ttk.Label(bar, textvariable=self.del_info, foreground="#555").pack(side="left", padx=8)

        self.del_entry.bind("<Return>", lambda e: self._delete_rows())

    def _parse_row_input(self, text: str) -> list[int]:
        rows: set[int] = set()
        for part in text.replace(" ", "").split(","):
            if not part:
                continue
            if "-" in part:
                a, b = part.split("-", 1)
                rows.update(range(int(a), int(b) + 1))
            else:
                rows.add(int(part))
        return sorted(rows)

    def _delete_rows(self):
        text = self.del_entry.get().strip()
        if not text:
            return
        try:
            wanted = self._parse_row_input(text)
        except ValueError:
            messagebox.showerror("輸入錯誤",
                "格式不正確，請輸入數字或範圍，如：2,4,7 或 3-5", parent=self.root)
            return

        valid = [r for r in wanted if 1 <= r <= len(self.raw_lines)]
        if not valid:
            messagebox.showwarning("無效列號",
                f"列號超出範圍（目前共 {len(self.raw_lines)} 筆）", parent=self.root)
            return

        if not messagebox.askyesno("確認刪除",
                f"確定刪除第 {valid} 列，共 {len(valid)} 筆？", parent=self.root):
            return

        # 存入 undo stack
        self._undo_stack.append(list(self.raw_lines))
        if len(self._undo_stack) > 20:
            self._undo_stack.pop(0)

        for idx in sorted((r - 1 for r in valid), reverse=True):
            del self.raw_lines[idx]

        self.modified = True
        self.del_entry.delete(0, "end")
        self.del_info.set(f"已刪除 {len(valid)} 筆（可復原 {len(self._undo_stack)} 步）")
        self.undo_btn.config(state="normal")
        self._rebuild_table_data()
        self._refresh_raw()
        self._filter_table()
        self._update_status()

    def _undo(self):
        if not self._undo_stack:
            return
        self.raw_lines = self._undo_stack.pop()
        self.modified  = True
        remaining = len(self._undo_stack)
        self.del_info.set(f"已復原（還可復原 {remaining} 步）")
        self.undo_btn.config(state="normal" if remaining else "disabled")
        self._rebuild_table_data()
        self._refresh_raw()
        self._filter_table()
        self._update_status()

    # ── Notebook ────────────────────────────────────────────────

    def _build_notebook(self):
        self.nb = ttk.Notebook(self.root)
        self.nb.pack(fill="both", expand=True)

        raw_f   = ttk.Frame(self.nb)
        table_f = ttk.Frame(self.nb)
        self.nb.add(raw_f,   text="  原始資料  ")
        self.nb.add(table_f, text="  表格解析  ")

        self._build_raw_tab(raw_f)
        self._build_table_tab(table_f)

    # ── 原始資料 Tab ─────────────────────────────────────────────

    def _build_raw_tab(self, parent):
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True)

        self.raw_sidebar = tk.Text(
            frame, width=9, padx=4, state="disabled",
            bg="#f5f5f5", fg="#777", relief="flat",
            font=("Courier New", 10), cursor="arrow",
        )
        self.raw_sidebar.pack(side="left", fill="y")
        self.raw_sidebar.tag_configure("ok",  foreground="#4CAF50")
        self.raw_sidebar.tag_configure("err", foreground="red")

        vscroll = ttk.Scrollbar(frame, orient="vertical")
        vscroll.pack(side="right", fill="y")
        hscroll = ttk.Scrollbar(parent, orient="horizontal")
        hscroll.pack(side="bottom", fill="x")

        self.raw_text = tk.Text(
            frame, wrap="none", state="disabled",
            font=("Courier New", 10),
            yscrollcommand=lambda *a: self._raw_yscroll_cb(vscroll, *a),
            xscrollcommand=hscroll.set,
        )
        self.raw_text.pack(side="left", fill="both", expand=True)
        self.raw_text.tag_configure("err_row", background="#ffe0e0")

        vscroll.config(command=self._raw_vscroll)
        hscroll.config(command=self.raw_text.xview)
        self.raw_text.bind("<<Selection>>", self._on_raw_selection)

    def _on_raw_selection(self, event=None):
        try:
            sel = self.raw_text.get("sel.first", "sel.last")
            b   = len(sel.encode(ENCODING, errors="replace"))
            self.sel_var.set(f"選取: {b} bytes")
        except tk.TclError:
            self.sel_var.set("")

    def _raw_yscroll_cb(self, scrollbar, *args):
        scrollbar.set(*args)
        self.raw_sidebar.yview_moveto(args[0])

    def _raw_vscroll(self, *args):
        self.raw_text.yview(*args)
        self.raw_sidebar.yview(*args)

    def _refresh_raw(self):
        use_sep = self._show_sep.get()

        self.raw_text.config(state="normal")
        self.raw_text.delete("1.0", "end")
        self.raw_sidebar.config(state="normal")
        self.raw_sidebar.delete("1.0", "end")

        for i, raw in enumerate(self.raw_lines):
            blen  = len(raw)
            is_ok = blen == RECORD_LEN

            self.raw_sidebar.insert("end",
                f"{i+1:>4} {blen:>3}\n", "ok" if is_ok else "err")

            if use_sep and is_ok:
                off, parts = 0, []
                for _, w, _ in FM_FIELDS:
                    parts.append(raw[off: off + w].decode(ENCODING, errors="replace"))
                    off += w
                line_str = "|".join(parts)
            else:
                line_str = raw.decode(ENCODING, errors="replace")

            self.raw_text.insert("end", line_str + "\n")
            if not is_ok:
                self.raw_text.tag_add("err_row", f"{i+1}.0", f"{i+1}.end")

        self.raw_text.config(state="disabled")
        self.raw_sidebar.config(state="disabled")

    # ── 表格解析 Tab ─────────────────────────────────────────────

    def _build_table_tab(self, parent):
        top = ttk.Frame(parent)
        top.pack(side="top", fill="x", padx=6, pady=4)

        ttk.Label(top, text="搜尋：").pack(side="left")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._filter_table())
        ttk.Entry(top, textvariable=self.search_var, width=28).pack(side="left", padx=4)
        ttk.Button(top, text="清除", command=lambda: self.search_var.set("")).pack(side="left")

        self.table_info = tk.StringVar(value="")
        ttk.Label(top, textvariable=self.table_info, foreground="#555").pack(side="left", padx=10)

        # 圖例（右側，從右往左 pack）
        for bg, label in [
            ("#e8ffe8", "正確"),
            ("#fff3cd", "必填空白"),
            ("#ffe8c4", "欄位值錯誤"),
            ("#ffe0e0", "格式錯誤"),
        ]:
            ttk.Label(top, text=label).pack(side="right")
            tk.Frame(top, bg=bg, width=14, height=14,
                     relief="solid", bd=1).pack(side="right", padx=(4, 1))

        # Treeview
        tv_f = ttk.Frame(parent)
        tv_f.pack(fill="both", expand=True)

        vsb = ttk.Scrollbar(tv_f, orient="vertical")
        hsb = ttk.Scrollbar(tv_f, orient="horizontal")
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")

        col_ids = ["row"] + [f[0] for f in FM_FIELDS] + ["bytes", "狀態"]
        self.tree = ttk.Treeview(tv_f, columns=col_ids, show="headings",
                                  yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.pack(fill="both", expand=True)
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)

        widths = {
            "row": 40,  "A": 28,    "PLAN_NO": 58,  "BRANCH": 50,
            "HOSP_ID": 92, "PID": 100, "BIRTHDAY": 76, "NAME": 94,
            "SEX": 36,  "ADDR": 200, "TEL": 120, "PRSN_ID": 94,
            "CASE_TYPE": 70, "CASE_DATE": 76, "CLOSE_DATE": 76, "CLOSE_RSN": 66,
            "bytes": 52, "狀態": 240,
        }
        self.tree.column("row", width=widths["row"], anchor="e")
        self.tree.heading("row", text="列")
        for name, wb, _ in FM_FIELDS:
            self.tree.column(name, width=widths.get(name, 80), minwidth=28, anchor="w")
            self.tree.heading(name, text=f"{name}\n({wb}B)", anchor="w")
        self.tree.column("bytes", width=widths["bytes"], anchor="e")
        self.tree.heading("bytes", text="Bytes")
        self.tree.column("狀態", width=widths["狀態"], anchor="w")
        self.tree.heading("狀態", text="狀態")

        style = ttk.Style()
        style.configure("Treeview", rowheight=22, font=("Courier New", 10))
        style.configure("Treeview.Heading", font=("", 10, "bold"))

        self._apply_tree_tags()
        self.tree.bind("<Double-1>", self._on_double_click)

    def _apply_tree_tags(self):
        self.tree.tag_configure(SEV_OK,    background="#e8ffe8")
        self.tree.tag_configure(SEV_BLANK, background="#fff3cd")
        self.tree.tag_configure(SEV_VALUE, background="#ffe8c4")
        self.tree.tag_configure(SEV_ERROR, background="#ffe0e0")

    def _rebuild_table_data(self):
        self._table_data = [(i, parse_record(r), r) for i, r in enumerate(self.raw_lines)]

    def _filter_table(self):
        kw = self.search_var.get().lower()
        self.tree.delete(*self.tree.get_children())

        warn_n, shown = 0, 0
        for i, fields, raw in self._table_data:
            if kw and not any(kw in v.lower() for v in fields.values()):
                continue
            sev, msgs = validate_record(raw, fields)
            if sev != SEV_OK:
                warn_n += 1
            vals = ([i + 1]
                    + [fields.get(f[0], "") for f in FM_FIELDS]
                    + [len(raw), "; ".join(msgs) if msgs else "✓"])
            self.tree.insert("", "end", iid=str(i), values=vals, tags=(sev,))
            shown += 1

        total = len(self._table_data)
        self.table_info.set(f"顯示 {shown}/{total} 筆，{warn_n} 筆異常")
        self.err_btn.config(text=f"⚠ {warn_n} 筆異常" if warn_n else "✓ 全部正確")
        self._apply_tree_tags()

    # ── 儲存格編輯 ──────────────────────────────────────────────

    def _on_double_click(self, event):
        if self.tree.identify("region", event.x, event.y) != "cell":
            return
        col_id  = self.tree.identify_column(event.x)
        row_id  = self.tree.identify_row(event.y)
        if not row_id:
            return

        col_idx   = int(col_id.lstrip("#")) - 1
        col_names = ["row"] + [f[0] for f in FM_FIELDS] + ["bytes", "狀態"]
        editable  = {f[0] for f in FM_FIELDS}

        if col_idx == 0 or col_names[col_idx] not in editable:
            return

        field_name = col_names[col_idx]
        line_idx   = int(row_id)
        current    = self._table_data[line_idx][1].get(field_name, "")
        self._open_edit_dialog(line_idx, field_name, current)

    def _open_edit_dialog(self, line_idx: int, field_name: str, current: str):
        offset, width = _field_offset(field_name)
        desc = next(f[2] for f in FM_FIELDS if f[0] == field_name)

        win = tk.Toplevel(self.root)
        win.title(f"編輯 {field_name}")
        win.geometry("500x200")
        win.resizable(False, False)
        win.grab_set()

        ttk.Label(win, text=f"{field_name} — {desc}   （上限 {width} bytes）",
                  font=("", 11, "bold")).pack(padx=16, pady=(14, 4), anchor="w")

        entry_var = tk.StringVar(value=current.strip())
        entry = ttk.Entry(win, textvariable=entry_var, font=("Courier New", 12), width=44)
        entry.pack(padx=16, fill="x")
        entry.select_range(0, "end")
        entry.focus_set()

        info_var = tk.StringVar()
        info_lbl = ttk.Label(win, textvariable=info_var)
        info_lbl.pack(padx=16, pady=4, anchor="w")

        def on_change(*_):
            b = len(entry_var.get().encode(ENCODING, errors="replace"))
            info_var.set(f"目前 {b} / {width} bytes")
            info_lbl.config(foreground="red" if b > width else "#4CAF50")

        entry_var.trace_add("write", on_change)
        on_change()

        def save():
            val  = entry_var.get()
            blen = len(val.encode(ENCODING, errors="replace"))
            if blen > width and not messagebox.askyesno(
                "超出長度",
                f"'{val}' 編碼後 {blen} bytes，超過 {width} bytes 上限。\n截斷後儲存？",
                parent=win,
            ):
                return
            raw = self.raw_lines[line_idx]
            self.raw_lines[line_idx] = (raw[:offset]
                                        + encode_field(val, width)
                                        + raw[offset + width:])
            new_raw = self.raw_lines[line_idx]
            self._table_data[line_idx] = (line_idx, parse_record(new_raw), new_raw)
            self.modified = True
            self._refresh_raw()
            self._filter_table()
            self._update_status()
            if self.tree.exists(str(line_idx)):
                self.tree.selection_set(str(line_idx))
                self.tree.see(str(line_idx))
            win.destroy()

        bf = ttk.Frame(win)
        bf.pack(pady=10)
        ttk.Button(bf, text="儲存", command=save,        width=10).pack(side="left", padx=6)
        ttk.Button(bf, text="取消", command=win.destroy, width=10).pack(side="left", padx=6)
        win.bind("<Return>", lambda e: save())
        win.bind("<Escape>", lambda e: win.destroy())

    # ── 跳至錯誤 ────────────────────────────────────────────────

    def _jump_next_error(self):
        self.nb.select(1)
        items = self.tree.get_children()
        sel   = self.tree.selection()
        start = list(items).index(sel[0]) if sel and sel[0] in items else -1
        candidates = items[start + 1:] + items[:start + 1]
        for item in candidates:
            if any(t in (SEV_ERROR, SEV_VALUE, SEV_BLANK)
                   for t in self.tree.item(item, "tags")):
                self.tree.selection_set(item)
                self.tree.see(item)
                return
        messagebox.showinfo("完成", "沒有找到異常筆數", parent=self.root)

    # ── 檔案 I/O ────────────────────────────────────────────────

    def _open_file(self):
        path = filedialog.askopenfilename(
            filetypes=[("FM.txt", "FM*.txt *.txt"), ("所有檔案", "*.*")]
        )
        if path:
            self._load_file(path)

    def _load_file(self, path: str):
        try:
            with open(path, "rb") as f:
                data = f.read()
            lines = data.replace(b"\r\n", b"\n").split(b"\n")
            if lines and lines[-1] == b"":
                lines = lines[:-1]
            self.raw_lines    = lines
            self.current_file = path
            self.modified     = False
            self._undo_stack.clear()
            self.undo_btn.config(state="disabled")
            self.del_info.set("")
            self._rebuild_table_data()
            self._refresh_raw()
            self._filter_table()
            self._update_status()
        except Exception as e:
            messagebox.showerror("開啟失敗", str(e))

    def _save_file(self):
        if self.current_file:
            self._write_file(self.current_file)
        else:
            self._save_as()

    def _save_as(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("FM.txt", "*.txt"), ("所有檔案", "*.*")],
        )
        if path:
            self._write_file(path)

    def _write_file(self, path: str):
        try:
            data = b"\r\n".join(self.raw_lines) + b"\r\n"
            with open(path, "wb") as f:
                f.write(data)
            self.current_file = path
            self.modified     = False
            self._update_status()
        except Exception as e:
            messagebox.showerror("儲存失敗", str(e))

    # ── 狀態列 ──────────────────────────────────────────────────

    def _build_statusbar(self):
        bar = ttk.Frame(self.root, relief="sunken")
        bar.pack(side="bottom", fill="x")
        self.status_var = tk.StringVar(value="就緒")
        self.record_var = tk.StringVar(value="")
        self.sel_var    = tk.StringVar(value="")
        ttk.Label(bar, textvariable=self.status_var, anchor="w").pack(side="left",  padx=6)
        ttk.Label(bar, textvariable=self.record_var, anchor="e").pack(side="right", padx=12)
        ttk.Separator(bar, orient="vertical").pack(side="right", fill="y", pady=2)
        ttk.Label(bar, textvariable=self.sel_var,
                  foreground="#1976D2").pack(side="right", padx=8)

    def _update_status(self):
        fname = os.path.basename(self.current_file) if self.current_file else "未開啟"
        mod   = " *" if self.modified else ""
        total = len(self.raw_lines)
        errs  = sum(1 for r in self.raw_lines if len(r) != RECORD_LEN)
        self.status_var.set(f"{fname}{mod}  │  CP950  │  {RECORD_LEN} bytes/筆")
        self.record_var.set(f"共 {total} 筆  │  行長錯誤 {errs} 筆")
        self.root.title(f"{'*' if self.modified else ''}{fname} — FM.txt 編輯器")

    def _on_close(self):
        if self.modified and not messagebox.askyesno("確認", "有未儲存的變更，確定離開？"):
            return
        self.root.destroy()


def main():
    root = tk.Tk()
    editor = FMEditor(root)
    if len(sys.argv) > 1:
        editor._load_file(sys.argv[1])
    root.mainloop()


if __name__ == "__main__":
    main()
