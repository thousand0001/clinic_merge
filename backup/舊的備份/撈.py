# -*- coding: utf-8 -*-
import os
import pandas as pd
from tkinter import Tk, filedialog


def select_file():
    """開啟檔案選擇視窗"""
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="請選擇輸入檔案",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    return file_path


def process_excel(file_path):
    print(f"讀取檔案: {file_path}")

    # 讀取所有 sheet
    xls = pd.ExcelFile(file_path)
    all_sheets = xls.sheet_names

    print(f"發現 sheets: {all_sheets}")

    result_list = []

    for sheet in all_sheets:
        # ❌ 略過「工作表19」
        if sheet == "工作表19":
            print(f"跳過 sheet: {sheet}")
            continue

        print(f"處理 sheet: {sheet}")

        try:
            df = pd.read_excel(file_path, sheet_name=sheet)

            # 基本欄位檢查
            if not {"身分證號", "姓名", "件數", "申請金額"}.issubset(df.columns):
                print(f"  ⚠ 欄位不符，跳過: {sheet}")
                continue

            # 分組統計
            grouped = df.groupby(["身分證號", "姓名"], as_index=False).agg({
                "件數": "sum",
                "申請金額": "sum"
            })

            result_list.append(grouped)

        except Exception as e:
            print(f"  ❌ 發生錯誤: {e}")

    if not result_list:
        print("❌ 沒有任何有效資料")
        return

    # 合併全部
    final_df = pd.concat(result_list, ignore_index=True)

    # 再總結一次（跨 sheet）
    final_df = final_df.groupby(["身分證號", "姓名"], as_index=False).agg({
        "件數": "sum",
        "申請金額": "sum"
    })

    # 輸出檔名（前6字）
    base_name = os.path.basename(file_path)
    prefix = base_name[:6]
    output_path = os.path.join(
        os.path.dirname(file_path),
        f"{prefix}_撈出.xlsx"
    )

    final_df.to_excel(output_path, index=False)

    print(f"\n✅ 完成！輸出檔案: {output_path}")

    # 自動開檔（Mac / Windows）
    try:
        if os.name == "nt":
            os.startfile(output_path)
        else:
            os.system(f'open "{output_path}"')
    except:
        pass


if __name__ == "__main__":
    file_path = select_file()

    if not file_path:
        print("❌ 未選擇檔案")
    else:
        process_excel(file_path)