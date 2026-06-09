# 資料上傳流程規則

> 根據 2026-06-10 稽核結論制定。目前僅為規則，尚未全部實作。

---

## 核心原則

1. **全欄位**：資料夾內所有檔案的欄位內容全部上傳
2. **去重**：相同自然鍵只保留一筆
3. **新蓋舊**：新批次驗證成功後取代舊批次，舊批次標記 superseded

---

## raw 層規則

- `.xlsx / .xlsm / .csv / .txt` 的所有非空列必須存入 `raw.uploaded_rows`
- `.xls / .pdf / .ini` 目前只存檔案資訊，**待實作**：至少轉換 `.xls` → `.xlsx` 後展開
- raw 保留來源原貌，**允許重複列**（供稽核用），不強制去重
- 欄位值需保留表頭對應（目前只存陣列，**待改善**）

## staging 層規則

### 去重
- 每個資料集需定義自然鍵（natural key），資料庫層用 `ON CONFLICT` 或 UPSERT 防止重複
  - `staging.members`：自然鍵 = `clinic_id + patient_id_normalized`
  - `staging.monthly_claims`：自然鍵 = `clinic_id + patient_id + roc_year + month`
  - 目前只有主鍵，**無自然鍵約束**，待補
- 同一會員多筆時，依「來源檔案修改時間較新者優先」，不得用「第一個非空值」

### 新批次蓋舊批次
- 新批次 `status = validated` 後，以 transaction 執行：
  1. 舊批次 `status` 更新為 `superseded`
  2. 新批次寫入 `current` 層（或設為 active）
  3. 舊批次資料保留 history，不刪除（供回復用）
- 目前 writer 只刪除相同 `batch_id` 的資料，**不會標記舊批次 superseded**，待修

---

## Parser 規則

- 每個 parser 必須達到：`discovered_files = parsed_files + 明確排除檔案`
- 不得只寫「v1 尚未實作此來源類型」跳過，每個跳過的檔案必須有明確原因
- 待完成的 parser：
  - **方鼎**：P4P、篩檢、檢驗仍略過
  - **杏翔**：sm_* 的 P4P、篩檢、檢驗仍略過
  - **耀聖、展望、自行系統**：仍有未實作類型

### name/phone enrichment 規則（已實作：自行系統）
- 資料夾內任何 xlsx 若有 ID + 姓名欄，自動補回 members 中的空白姓名
- 優先順序：來源檔越新越優先

---

## 已知問題（待修）

| 問題 | 位置 | 嚴重度 |
|------|------|--------|
| raw 7,616 組重複，多 8,852 列 | raw.uploaded_rows | 中（稽核用可接受） |
| staging.members 最新批次多 2,211 筆（鈞安 1,283、崇恩 928） | staging | 高 |
| 新批次不標記舊批次 superseded | storage/__init__.py L466 | 高 |
| 同會員多筆採第一非空值，不分新舊 | member_builder.py L59 | 中 |
| .xls/.pdf 未展開欄位內容 | collector.py L102 | 低 |

---

## 測試規則

- 現有 30 項測試通過，但未覆蓋：
  - 跨檔去重
  - 新批次蓋舊批次
  - 完整欄位上傳
- 上述三項每條規則實作後，必須有對應測試
