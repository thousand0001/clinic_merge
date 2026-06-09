# -*- coding: utf-8 -*-
"""
產生家庭醫師整合性照護計畫名單上傳檔（FM.txt）

格式規範：MHB5 系統，定長 208 BYTES，CP950 編碼
檔名規則：業務組別(1)+醫事機構代號(10)+上傳月份(2)+流水號(2)+FM.txt

用法（CLI）：
    python tools/generate_fm_txt.py \\
        --hosp-id 3501013059 \\
        --template "路徑/親親家-醫聖 115指定會員模板.xlsx" \\
        --excel   "路徑/親親家庭診所選會員_0609_1548.xlsx" \\
        --dest    "輸出資料夾/" \\
        [--plan-no 01] [--serial 01] [--prsn-id A123456789]

    # 書田（無指定會員模板，用自選名單 sheet）：
    python tools/generate_fm_txt.py \\
        --hosp-id 4001020028 \\
        --excel   "路徑/書田泌尿科眼科診所選會員_0609_1548.xlsx" \\
        --mode    shuda \\
        --dest    "輸出資料夾/"

GUI 模式（不帶任何參數直接執行）：
    python tools/generate_fm_txt.py
"""
from __future__ import annotations

import argparse
import datetime
import sys
from pathlib import Path

import openpyxl

# db_pipeline 在上一層
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from db_pipeline.normalization import branch_from_address
from db_pipeline.storage import _run_query


# ── 常數 ──────────────────────────────────────────────────────────────────────
DEFAULT_PLAN_NO = "01"
DEFAULT_PRSN_ID = "A123456789"
TODAY = datetime.date.today().strftime("%Y%m%d")


# ── 工具函式 ──────────────────────────────────────────────────────────────────
def sex_from_id(pid: str) -> str:
    """台灣身分證/居留證第 2 碼推性別 → 1=男 2=女"""
    if not pid or len(pid) < 2:
        return " "
    g = str(pid)[1].upper()
    if g in ("1", "8", "A", "C"):
        return "1"
    if g in ("2", "9", "B", "D"):
        return "2"
    return " "


def to_yyyymmdd(v) -> str:
    if v is None:
        return "        "
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%Y%m%d")
    s = str(v).strip()
    if len(s) == 8 and s.isdigit():
        return s
    # 支援 YYYY/MM/DD 或 YYYY-MM-DD
    for sep in ("/", "-"):
        parts = s.split(sep)
        if len(parts) == 3 and all(p.isdigit() for p in parts):
            return f"{parts[0]}{parts[1].zfill(2)}{parts[2].zfill(2)}"
    return "        "


def case_abc(v) -> str:
    """個案類別數值 → FM.txt A/B/C"""
    if v is None:
        return "A"
    s = str(v).strip().lower()
    if s == "6":
        return "C"
    if "b" in s:
        return "B"
    return "A"


def fb(value, byte_len: int, encoding: str = "cp950") -> bytes:
    """固定長度欄位，左靠補半型空白"""
    b = ("" if value is None else str(value)).encode(encoding, errors="replace")
    while len(b) > byte_len:
        b = b[:-1]
    return b + b" " * (byte_len - len(b))


def make_record(
    plan_no: str, branch: int, hosp_id: str,
    pid: str, birthday: str, name: str, sex: str,
    addr: str, tel: str, prsn_id: str,
    case_type: str, case_date: str,
) -> bytes:
    rec = (
        b"A"
        + plan_no.encode()
        + str(branch).encode()
        + fb(hosp_id, 10)
        + fb(pid, 10)
        + fb(birthday, 8)
        + fb(name, 12)
        + sex.encode()
        + fb(addr, 120)
        + fb(tel, 15)
        + fb(prsn_id, 10)
        + case_type.encode()
        + fb(case_date, 8)
        + fb("", 8)   # CLOSE_DATE 空白
        + b" "        # CLOSE_RSN 空白
    )
    assert len(rec) == 208, f"記錄長度錯誤：{len(rec)}"
    return rec


# ── 從選會員 Excel 建 ID lookup ───────────────────────────────────────────────
def build_lookup(excel_path: Path) -> dict:
    wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    ws = wb["醫生看(從會員指標內容Key過來)"]
    lookup = {}
    for row in ws.iter_rows(min_row=4, values_only=True):
        pid = str(row[0]).strip() if row[0] else None
        if not pid or pid == "None":
            continue
        tel = str(row[4]).strip() if row[4] else (str(row[5]).strip() if row[5] else "")
        lookup[pid] = {
            "name":      str(row[1]).strip() if row[1] else "",
            "birthday":  to_yyyymmdd(row[2]),
            "tel":       tel,
            "case_type": case_abc(row[48]),   # AW 欄
            "case_date": to_yyyymmdd(row[28]) if row[28] else TODAY,  # AC 欄
        }
    return lookup


# ── 主邏輯 ────────────────────────────────────────────────────────────────────
def generate(
    hosp_id: str,
    excel_path: Path,
    dest_dir: Path,
    template_path: Path | None = None,
    mode: str = "template",
    plan_no: str = DEFAULT_PLAN_NO,
    serial: str = "01",
    prsn_id: str = DEFAULT_PRSN_ID,
    upload_month: str | None = None,
) -> list[Path]:
    # 從 DB 取診所資料
    r = _run_query(
        f"SELECT clinic_name, institution_address, institution_phone "
        f"FROM meta.clinics WHERE clinic_code='{hosp_id}' LIMIT 1;"
    )
    if not r:
        raise ValueError(f"找不到診所代碼：{hosp_id}")
    parts = r.split("|")
    clinic_addr  = parts[1] if len(parts) > 1 else ""
    clinic_phone = parts[2] if len(parts) > 2 else ""

    branch = branch_from_address(clinic_addr) or 1
    month  = upload_month or datetime.date.today().strftime("%m")
    lookup = build_lookup(excel_path)
    records = []

    def _rec(pid, birthday, case_type_raw):
        m = lookup.get(pid, {})
        tel = m.get("tel", "") or clinic_phone
        return make_record(
            plan_no   = plan_no,
            branch    = branch,
            hosp_id   = hosp_id,
            pid       = pid,
            birthday  = birthday or m.get("birthday", "        "),
            name      = m.get("name", ""),
            sex       = sex_from_id(pid),
            addr      = clinic_addr,
            tel       = tel,
            prsn_id   = prsn_id,
            case_type = case_abc(case_type_raw),
            case_date = m.get("case_date", TODAY),
        )

    if mode == "template" and template_path:
        wb = openpyxl.load_workbook(template_path, read_only=True, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            pid = str(row[1]).strip() if row[1] else None
            if not pid or pid == "None":
                continue
            records.append(_rec(pid, to_yyyymmdd(row[2]), row[3]))
    else:  # shuda：從自選名單 sheet
        wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
        ws = wb["自選名單(從會員指標內容Key過來)"]
        for row in ws.iter_rows(min_row=3, values_only=True):
            pid = str(row[1]).strip() if row[1] else None
            if not pid or pid == "None":
                continue
            m = lookup.get(pid, {})
            records.append(_rec(pid, m.get("birthday", "        "), m.get("case_type", "A")))

    # 超過 9999 自動拆檔（serial 遞增）
    MAX_PER_FILE = 9999
    chunks = [records[i:i+MAX_PER_FILE] for i in range(0, len(records), MAX_PER_FILE)]
    out_paths = []
    serial_int = int(serial)
    for chunk in chunks:
        sn = f"{serial_int:02d}"
        fname = f"{branch}{hosp_id}{month}{sn}FM.txt"
        out_path = dest_dir / fname
        with open(out_path, "wb") as f:
            for rec in chunk:
                f.write(rec + b"\r\n")
        print(f"✓ {fname}：{len(chunk)} 筆，業務組別={branch}")
        out_paths.append(out_path)
        serial_int += 1

    return out_paths


# ── 診所搜尋 ──────────────────────────────────────────────────────────────────
def search_clinics(keyword: str) -> list[dict]:
    """依診所名稱模糊搜尋，回傳 [{'code':..., 'name':...}, ...]"""
    kw = keyword.strip().replace("'", "''")
    r = _run_query(
        f"SELECT clinic_code, clinic_name FROM meta.clinics "
        f"WHERE clinic_name LIKE '%{kw}%' OR official_name LIKE '%{kw}%' "
        f"ORDER BY clinic_name LIMIT 20;"
    )
    if not r:
        return []
    results = []
    for line in r.splitlines():
        parts = line.split("|")
        if len(parts) >= 2:
            results.append({"code": parts[0], "name": parts[1]})
    return results


def resolve_hosp_id(hosp_id: str | None, hosp_name: str | None) -> str:
    """從代碼或名稱解析出唯一的 clinic_code，多筆時 CLI 互動選擇。"""
    if hosp_id:
        return hosp_id.strip()
    if not hosp_name:
        raise ValueError("請提供 --hosp-id 或 --hosp-name")
    matches = search_clinics(hosp_name)
    if not matches:
        raise ValueError(f"找不到診所：{hosp_name}")
    if len(matches) == 1:
        print(f"找到：{matches[0]['code']} {matches[0]['name']}")
        return matches[0]["code"]
    print("找到多筆，請選擇：")
    for i, m in enumerate(matches, 1):
        print(f"  {i}. {m['code']}  {m['name']}")
    while True:
        try:
            choice = int(input("輸入編號："))
            if 1 <= choice <= len(matches):
                return matches[choice - 1]["code"]
        except (ValueError, KeyboardInterrupt):
            pass
        print("無效輸入，請重試。")


# ── GUI 模式 ──────────────────────────────────────────────────────────────────
def _gui():
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox, simpledialog
    except ImportError:
        print("無法使用 GUI，請用 CLI 參數模式。")
        return

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    # 支援輸入代碼或名稱
    keyword = simpledialog.askstring(
        "搜尋診所",
        "輸入醫事機構代碼（10碼）或診所名稱（模糊搜尋）：",
        parent=root,
    )
    if not keyword:
        return

    keyword = keyword.strip()
    # 判斷是代碼還是名稱
    if keyword.isdigit() or (len(keyword) == 10 and keyword[0].isdigit()):
        hosp_id = keyword
    else:
        matches = search_clinics(keyword)
        if not matches:
            messagebox.showerror("找不到", f"找不到診所：{keyword}")
            root.destroy()
            return
        if len(matches) == 1:
            hosp_id = matches[0]["code"]
        else:
            # 多筆 → 跳選單
            win = tk.Toplevel(root)
            win.title("選擇診所")
            win.attributes("-topmost", True)
            lb = tk.Listbox(win, width=50, height=min(len(matches), 15))
            for m in matches:
                lb.insert(tk.END, f"{m['code']}  {m['name']}")
            lb.pack(padx=10, pady=10)
            lb.selection_set(0)
            selected = [None]

            def confirm():
                idx = lb.curselection()
                if idx:
                    selected[0] = matches[idx[0]]["code"]
                win.destroy()

            tk.Button(win, text="確認", command=confirm).pack(pady=5)
            win.grab_set()
            root.wait_window(win)
            if not selected[0]:
                root.destroy()
                return
            hosp_id = selected[0]

    excel = filedialog.askopenfilename(title="選擇「選會員」Excel", filetypes=[("Excel","*.xlsx")])
    if not excel:
        return

    template = filedialog.askopenfilename(
        title="選擇「指定會員模板」Excel（書田可略過，直接取消）",
        filetypes=[("Excel","*.xlsx"), ("所有檔案","*.*")]
    )

    dest = filedialog.askdirectory(title="選擇輸出資料夾")
    if not dest:
        return

    mode = "template" if template else "shuda"
    try:
        out = generate(
            hosp_id       = hosp_id.strip(),
            excel_path    = Path(excel),
            dest_dir      = Path(dest),
            template_path = Path(template) if template else None,
            mode          = mode,
        )
        messagebox.showinfo("完成", "已產生：\n" + "\n".join(str(p) for p in out))
    except Exception as e:
        messagebox.showerror("錯誤", str(e))

    root.destroy()


# ── CLI 入口 ──────────────────────────────────────────────────────────────────
def main():
    if len(sys.argv) == 1:
        _gui()
        return

    parser = argparse.ArgumentParser(description="產生家醫計畫名單上傳 FM.txt")
    id_group = parser.add_mutually_exclusive_group(required=True)
    id_group.add_argument("--hosp-id",   help="醫事機構代碼（10碼）")
    id_group.add_argument("--hosp-name", help="診所名稱（模糊搜尋）")
    parser.add_argument("--excel",     required=True, type=Path, help="選會員輸出 Excel 路徑")
    parser.add_argument("--template",  type=Path, default=None, help="指定會員模板 xlsx（書田可省略）")
    parser.add_argument("--dest",      required=True, type=Path, help="輸出資料夾")
    parser.add_argument("--mode",      default="template", choices=["template","shuda"])
    parser.add_argument("--plan-no",   default=DEFAULT_PLAN_NO, help="計畫期別（預設 01）")
    parser.add_argument("--serial",    default="01", help="流水號（預設 01）")
    parser.add_argument("--prsn-id",   default=DEFAULT_PRSN_ID, help="醫事人員身分證")
    parser.add_argument("--month",     default=None, help="上傳月份 MM（預設本月）")
    args = parser.parse_args()

    hosp_id = resolve_hosp_id(args.hosp_id, args.hosp_name)
    generate(
        hosp_id       = hosp_id,
        excel_path    = args.excel,
        dest_dir      = args.dest,
        template_path = args.template,
        mode          = args.mode,
        plan_no       = args.plan_no,
        serial        = args.serial,
        prsn_id       = args.prsn_id,
        upload_month  = args.month,
    )


if __name__ == "__main__":
    main()
