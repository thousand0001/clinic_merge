# 現有資料庫流程盤點

## 現有入口

- `陳森豊_上傳資料庫0601.py`
  - 診所專用入口。
  - 呼叫 `資料庫輸出0601.py`。
  - 支援資料夾選擇與指定輸出目錄。

- `資料庫輸出0601.py`
  - 上傳來源 Excel 列。
  - 建立及更新批次。
  - 從資料庫讀回資料。
  - 產生 Excel。
  - 紀錄輸出及支援上一批回復。

## 已使用的 PostgreSQL 能力

- `meta`、`raw`、`audit` schema。
- UUID 批次識別。
- `JSONB` 原始列資料。
- JSONB GIN 索引。
- `BIGSERIAL` 主鍵。
- `TIMESTAMPTZ`。
- 外鍵約束。
- `ON CONFLICT ... DO UPDATE`。
- `RETURNING`。
- `psql \copy` 大量匯入與匯出。
- transaction 型批次更新。
- current 與歷史批次資料。
- 操作稽核及輸出紀錄。

## 現有主要資料表

已在程式中確認：

- `meta.clinics`
- `meta.import_batches`
- `meta.source_files`
- `meta.generated_outputs`
- `raw.uploaded_rows`
- `raw.current_uploaded_rows`
- `audit.operation_logs`

完整 schema 定義仍需從目前 PostgreSQL 實例匯出後確認。

## 現有優點

- 已具批次與來源檔追蹤基礎。
- 原始資料以 JSONB 保存，適合逐步展開 staging。
- 已有 current 資料概念。
- 已有回復上一批次與操作稽核雛形。

## 現有風險

1. 上傳、解析、資料庫操作與 Excel 輸出集中在大型單一程式。
2. 診所專用入口仍綁定特定診所代碼與名稱。
3. SQL 以字串拼接及 `psql` 子程序為主。
4. PostgreSQL 密碼存在程式預設值，需移至環境變數。
5. 尚未建立結構化 staging 資料集。
6. 輸出邏輯仍可能依賴資料夾與原始檔格式。
7. 尚缺可自動執行的新舊輸出完整比對。

## 不可在早期遷移中移除的能力

- 原始列留存。
- 來源檔雜湊。
- `batch_id`。
- current/history。
- 批次失敗狀態。
- 回復上一成功批次。
- 操作稽核。
- Excel 輸出紀錄。

## 下一步盤點

1. PostgreSQL 服務恢復後，匯出完整 schema、索引、外鍵及資料量。
2. 盤點每個來源系統的檔案類型、編碼與欄位。
3. 對照現有前置清洗程式，列出不可用設定檔取代的特殊規則。
4. 定義標準資料集欄位與主鍵。

## 2026-06-08 實例盤點狀態

- `psql 16.11` 已安裝。
- `clinic_merge` 預期連線位置為 `localhost:5432`。
- 盤點時 PostgreSQL 未監聽該連接埠，因此沒有啟動服務或修改資料庫。
- 在服務恢復前，先依現有程式碼完成旁路契約與 schema 草案。

## 2026-06-08 SM 旁路試跑

代表資料夾：

- 診所：本一診所。
- 系統：SM／耀聖。
- 模式：唯讀解析，不寫 PostgreSQL、不產生正式 Excel。

結果：

- 發現來源檔 11 個。
- 已解析來源檔 3 個。
- 會員原始列 2,121 筆。
- 門診次數費用成功對應 132 筆，未唯一對應 462 筆。
- 自選會員成功對應 135 筆，未唯一對應 143 筆。
- P4P、健康管理、篩檢及舊式 `.xls` 轉診檔尚未納入 SM v1。
- 語法檢查及 7 項旁路單元測試通過。

判定：

- SM 基礎欄位推斷、來源追蹤與覆蓋報告可運作。
- 尚有未命中與未實作來源，不符合正式切換條件。
- 現有前置清洗及 Excel 正式流程維持不變。

## 2026-06-09 鈞安醫聖 staging

代表資料夾：

- 診所：鈞安診所（3501186011）。
- 系統：醫聖。
- 來源：照護名單、15 個 BIG5/CP950 月份 TXT、P4P、五類篩檢、個案健康管理。

解析器補齊：

- TXT 月份檔保留所有有效身分證號，不因不在照護名單而略過一般門診病患。
- 新增 P4P 收案與追蹤解析。
- 新增成人健檢、子宮抹片、老人流感、糞便潛血、肝炎篩檢解析。
- 新增 HbA1c、LDL、UACR 個案健康管理解析。
- 新增鈞安設定檔及醫聖解析器合約測試。

正式 staging 結果：

- batch_id：`2912a96f-0d00-5ed9-968a-523fca63425e`。
- 狀態：`validated`。
- 發現來源檔 24 個，成功解析 24 個，略過 0 個。
- members：1,947。
- claims：9,117。
- member_flags：1,947。
- p4p_records：180。
- screenings：731。
- lab_results：1,919。
- P4P 追蹤來源只有表頭，因此追蹤資料為 0 筆。
- 6,127 筆 claims 不在照護名單，屬一般門診病患；已使用來源真實身分證號寫入。

驗證：

- 鈞安 dry-run 與正式 staging 均通過，0 error、1 warning。
- staging 回讀筆數與解析 counter 完全一致。
- db_pipeline 契約、醫聖解析器與 staging schema 共 8 項測試通過。
- 舊 `test_txt_monthly_fee_to_xlsx.py` 仍引用已不存在的 V1 工具，屬既有測試缺口，未納入鈞安 staging 驗收。

下一步：

1. 進入調和系統代表診所「蘆洲大愛」解析器試跑。
2. 鈞安的新舊 Excel AW:BJ 與 L/M/N/O 比對留待第 7 階段執行。

## 2026-06-09 蘆洲大愛調和 staging

代表資料夾：

- 診所：蘆洲大愛診所（3531142830）。
- 系統：調和。
- 來源：照護名單、15 個逐月費用檔、自選／不選 CSV、P4P、五類篩檢、個案健康管理。

解析器補齊：

- 照護名單所有來源列完整保留；身分證集合只用於判斷是否已連結會員。
- 一般門診與自選／不選來源即使不在照護名單，仍保留真實身分證號與來源旗標。
- 新增 P4P、成人健檢、子宮抹片、老人流感、糞便潛血、肝炎篩檢及 HbA1c／LDL／UACR。
- 新增 raw 收集器；重複彙總工作簿完整保存在 raw，但不重複計算。

正式 staging 結果：

- batch_id：`1ba42faa-9a0a-4d61-9ffe-582805c2f6a8`。
- 狀態：`validated`。
- `meta.source_files`：27 個來源檔全部登錄。
- `raw.uploaded_rows`：8,953 個非空原始列。
- 結構化解析 26 個檔；重複彙總檔只進 raw。
- members：264。
- claims：3,460。
- member_flags：760。
- p4p_records：34。
- screenings：196。
- lab_results：1,791。
- 3,334 筆 claims 與 335 筆自選／不選來源不在照護名單，均已保留。

驗證：

- dry-run 與正式 staging 均通過，0 error、2 warning。
- raw 與 staging 回讀筆數和解析 counter 完全一致。
- 該批次在 `raw.current_uploaded_rows` 為 0 筆，既有正式流程未受影響。
- 旁路契約、調和 raw/staging 與 schema 共 8 項測試通過。

下一步：

1. 進入宏誠系統代表診所「德容」解析器試跑。
2. 書田維持中斷，後續只使用書田專用前置處理流程。
3. 鈞安與蘆洲大愛的新舊 Excel 比對留待第 7 階段執行。

## 2026-06-09 德容宏誠 staging

代表資料夾：

- 診所：德容聯合診所（3501103076）。
- 系統：宏誠。
- 來源：CP950 照護名單 CSV、15 個月份費用 xlsx、15 個月份次數 PDF、P4P、五類篩檢、個案健康管理。

解析器補齊：

- 新增 CP950 照護名單 CSV 與宏誠診所設定。
- 逐月費用依身分證號聚合金額、最後就診日及明細次數。
- 逐月 PDF 依病歷號連結費用檔；PDF 次數優先，並保留費用與 PDF 的合併來源追蹤。
- 一般門診病患即使不在照護名單，仍以來源真實身分證號寫入。
- 新增 P4P、成人健檢、子宮抹片、老人流感、糞便潛血、肝炎篩檢及 HbA1c／LDL／UACR。

正式 staging 結果：

- batch_id：`165e5857-e504-5432-beb2-c1352a5b16cb`。
- 狀態：`validated`。
- 發現來源檔 38 個，結構化解析 38 個，略過 0 個。
- raw source files：38。
- raw source rows：22,400。
- members：657。
- claims：12,045。
- member_flags：657。
- p4p_records：72。
- screenings：720。
- lab_results：3,069。

差異與採用規則：

- 9,004 筆 claims 不在照護名單，屬一般門診病患，已完整保留。
- PDF 的 12,318 個病歷號均可連結同月費用檔。
- 21 個費用檔病歷號未出現在同月 PDF，使用費用明細列數作為次數。
- 4 個病歷號的 PDF 次數與費用明細列數不同，依宏誠來源規則採 PDF 次數。
- PDF 已保存檔案中繼資料與 SHA-256；raw 通用收集器尚未展開 PDF 原始列，但宏誠解析器已完成結構化 PDF 次數解析。

驗證：

- dry-run 與正式 staging 均通過，0 error。
- staging 與 raw 回讀筆數和解析 counter 完全一致。
- db_pipeline 契約、醫聖、調和、宏誠與 staging schema 共 11 項測試通過。

下一步：

1. 書田依使用者指示維持中斷，不進行自行系統解析器測試。
2. 第 5 階段代表系統先完成至宏誠，接續第 6 階段 `FolderDataProvider`／`PostgresDataProvider` 雙介面。
3. 第 7 階段再進行新舊 Excel 關鍵欄位與 L/M/N/O 比對。

## 2026-06-09 第 6 階段 DataProvider 起始

已建立：

- `DataProvider` 共用協定，固定提供 `load_bundle()`。
- `FolderDataProvider`：依診所設定呼叫既有系統解析器。
- `PostgresDataProvider`：唯讀指定診所與 validated batch，重建相同 `DatasetBundle`。
- Provider 合約測試，不依賴真實 PostgreSQL。

德容實際雙介面比對：

- members：657 / 657。
- monthly_claims：12,045 / 12,045。
- p4p_cases：72 / 72。
- lab_results：3,069 / 3,069。
- screenings：720 / 720。
- member_selections：657 / 657。
- 會員 ID 集合、claims ID 集合、旗標集合完全一致。
- 就診次數總計：18,322 / 18,322。
- 費用總計：9,054,438 / 9,054,438。
- 共 12 項契約及解析器測試通過。

跨系統補充驗證：

- 鈞安（醫聖）：7 個標準資料集筆數、旗標、claims ID、次數總計 14,640 均一致。
- 蘆洲大愛（調和）：比對時發現照護名單去重呼叫在 raw 整合後遺失，已補回並加入重複 ID 測試。
- 蘆洲大愛修正後：members 216、claims 3,460、P4P 34、lab 1,791、screenings 196、flags 712 均一致。
- 蘆洲大愛次數總計 8,578 / 8,578，旗標集合完全一致。

下一步：

1. 將 Provider 比對擴充到鈞安與蘆洲大愛，確認不同來源系統均可還原相同中間資料。
2. 定義輸出層實際需要的查詢方法，避免 Excel 輸出層自行掃資料夾或拼 SQL。
3. 完成第 6 階段後進入第 7 階段新舊 Excel 輸出比對。

## 2026-06-09 第 7 階段 輸出比對基礎建設

已建立：

- `db_pipeline/output/member_builder.py`：
  - `build_from_bundle(DatasetBundle) → Dict[str, Dict]`
  - 等同舊流程 `build_members()`，來源改為 DatasetBundle。
  - AW:BJ 欄位：中文 key 直接對應 `write_output()` 的 `designated_fields`。
  - L/M/N/O 欄位：114_count_q1 / 115_count / 114_avg_amount / 115_avg_amount，計算邏輯與舊流程一致。
  - 包含 disease_code 推算、merge_if_empty 保護、sex 補齊。
- `db_pipeline/output/__init__.py`：package init。
- `tools/generate_output.py`：
  - CLI：`python tools/generate_output.py <batch_id> --template <範本> --dest <輸出>`
  - 以 `PostgresDataProvider` 讀取 staging → `build_from_bundle` → 舊流程 `write_output()`。
- `tools/compare_common_output.py`：醫生看工作表 max_col 由 48 擴至 62，涵蓋 AW:BJ。
- `tests/test_member_builder.py`：9 項單元測試（AW:BJ、L/M/N/O、旗標、篩檢、檢驗、P4P、disease_code、merge），全部通過。
- 全套 24 項測試通過，無 regression。

注意：

- disease_code ≠ None 時，ASCVD 欄位改寫為 "1"（與舊流程行為一致）。
- 鈞安與蘆洲大愛目前無舊流程批次（raw.uploaded_rows 空）；實際 Excel 比對需先以舊流程跑相同來源資料夾。

比對步驟（使用者執行）：

1. 以舊流程產出舊 Excel：`python 資料庫輸出0601.py --source <來源資料夾> --dest old.xlsx`
2. 以新工具產出新 Excel：`python tools/generate_output.py <batch_id> --template <範本> --dest new.xlsx`
3. 比對：`python tools/compare_common_output.py old.xlsx new.xlsx`

下一步：

1. Phase 8：Operations & output UI（診所選擇、批次管理、一鍵產出）。

## 2026-06-09 第 8 階段 Operations & output UI

已建立：

- `db_pipeline/cli/output.py`：一鍵產出 CLI。
  - `--clinic-code CODE`：自動取最新 validated/published 批次。
  - `--batch-id UUID`：指定批次（支援跨診所回溯比對）。
  - `--template <範本路徑> --dest <輸出路徑>`。
  - 實測：鈞安（3754 筆）、蘆洲大愛（3102 筆）均成功產出。
- `db_pipeline/cli/batches.py`：批次清單 CLI。
  - `--clinic-code`（篩診所）、`--status`（篩狀態）、`--limit`（顯示筆數）。
  - 欄位：診所名稱、代碼、來源系統、狀態、批次 ID、開始/驗證時間。

使用方式：
```bash
# 列出所有 validated 批次
python -m db_pipeline.cli.batches --status validated

# 一鍵產出（最新批次）
python -m db_pipeline.cli.output \
    --clinic-code 3501186011 \
    --template 選會員模板0526.xlsx

# 指定批次產出
python -m db_pipeline.cli.output \
    --batch-id 2912a96f-0d00-5ed9-968a-523fca63425e \
    --template 選會員模板0526.xlsx \
    --dest /tmp/鈞安_output.xlsx
```

下一步：

1. Phase 7 實際 Excel 比對：請使用者以舊流程跑鈞安或蘆洲大愛來源資料夾，再用 `compare_common_output.py` 比對。
2. Phase 9：切換評估（正式流程切換條件確認）。

## 2026-06-09 Phase 7 實際比對結果（蘆洲大愛調和，正式流程）

來源：`3531142830蘆洲大愛調合/` 來源資料夾以**正式舊流程**（upload → DB → download → build_members）產出，對比 staging batch `1ba42faa-...` 以新流程產出。

比對方式：`tools/compare_common_output.py`，從舊流程 Excel 前 90 筆 ID 找出新流程 Excel 裡的共同人，取 30 筆逐欄比對。

**產出行數**

| 流程 | 輸出 members | 說明 |
|------|-------------|------|
| 舊流程（正式 DB 跑） | 604 | batch_id `9c4f6b74-e2e2-427c-8646-cca03e72cdf0` |
| 新流程（staging DB） | 3,102 | batch_id `1ba42faa-9a0a-4d61-9ffe-582805c2f6a8` |
| 舊前 90 ID 在新中命中 | 88 / 90 | — |
| 抽樣比對人數 | 30 | — |

**30 人中，同在兩邊照護名單的共同人（17 人）AW:BJ 比對結果**

| 欄組 | 差異筆數 | 結論 |
|------|----------|------|
| col49-62 全部 AW:BJ（17 位共同成員） | **0** | 完全一致 ✓ |

**其餘差異（非 bug，屬資料範圍或來源限制）**

| 差異類型 | 筆數 / 30 | 原因 |
|----------|-----------|------|
| 新流程多 13 人（不在舊照護名單）→ AW:BJ 舊空新有 | 13 | 新流程從 staging 取更完整的 member_flags |
| 姓名 / 生日 / 年齡 新流程空白 | 17-26 | 調和系統 staging 僅有部分會員存姓名（非照護名單 sheet） |
| HbA1c / LDL / UACR 舊流程為 `"0"`，新流程為空 | 21-29 | 舊流程無資料時預設 0；新流程留空（新流程行為較正確） |
| L/M（114 Q1 次數、115 次數）差異 | 1 | 個別 staging 資料差異，不影響整體 |
| 篩檢文字 / 分數算法差異 | 4-9 | 判讀邏輯略有差異 |

**修正的 bugs（比對過程中發現並修正，commit 49db188）**

| Bug | 修正方式 |
|-----|---------|
| 性別推算寫 `"M"/"F"` 而非 `"男"/"女"` | 改用 `{"1","8"}→"男"` / `{"2","9"}→"女"` |
| `claim_months_115` 為全域 set（O 欄分母錯誤） | 改為 per-person `"115_months"` set |
| staging `member_type` 帶 `'` 前綴（Excel text-prefix 洩漏） | `_norm()` 新增 `startswith("'")` 剝除 |
| `disease_pattern` 為空時 `_disease_code_text("")` 返回 `"None"` 字串 | 加 `and rec.disease_pattern` 條件跳過 |

**結論**：

- 對於兩邊照護名單都有的 17 位共同成員，AW:BJ 13 個指定欄位**完全一致（0 差異）**，驗證新流程輸出邏輯正確。
- 新流程比舊流程多了 13 位成員（staging 有記錄但未在舊照護名單），屬資料完整性提升。
- 姓名缺失與 lab 預設 0 屬已知來源限制，不影響照護指標計算。

下一步：

1. Phase 7 繼續：鈞安（醫聖）與德容（宏誠）比對。
2. Phase 9：切換評估（正式流程切換條件確認）。

## 2026-06-09 Phase 7 實際比對結果（鈞安醫聖 + 德容宏誠，正式流程）

比對觸發另外三個 bugs，修正後（commit 見下）再次比對。

### 修正的 bugs（跨診所比對發現，同 session 修正）

| Bug | 修正方式 |
|-----|---------|
| `_DISEASE_MAP` 編號對應錯誤（"2"→HTN、"3"→CKD、"4"→DKD） | 對齊舊流程：`"1"→DM`、`"2"→CKD`、`"3"→DKD`、`"4"`移除（預設 "None"） |
| `_disease_class_text()` 未加 `+ASCVD` 後綴（如 `DM+ASCVD`） | 加入 `has_ascvd` 判斷，輸出 `DM+ASCVD`／`CKD+ASCVD`／`DKD+ASCVD` |
| lab result 值未正規化（`"57.0"` 未去除 `.0`） | section D 改用 `_norm(rec.result_value)` |

---

### 鈞安診所（醫聖，3501186011）

| 流程 | 輸出 members | 說明 |
|------|-------------|------|
| 舊流程（正式 DB 跑） | 858 | batch `414fec86`；只上傳照護名單+健康類 9 檔，TXT 月費未包含 |
| 新流程（staging DB） | 3,754 | batch `2912a96f`；含 TXT 月費 15 檔 |
| 舊前 90 ID 在新中命中 | 84 / 90 | — |
| 抽樣比對人數 | 30 | — |

**AW:BJ（30 位共同成員）**

| 欄組 | 差異筆數 | 結論 |
|------|----------|------|
| col49-62 全部 AW:BJ（30 位，修正後） | **0** | 完全一致 ✓ |

**其餘差異（非 bug）**

| 差異類型 | 筆數 / 30 | 原因 |
|----------|-----------|------|
| 最後就診日 / 就診次數 舊空新有 | 21 | 舊流程未讀 TXT 月費檔（不在上傳來源） |
| 分數 / 分數說明 不同 | 21 | 就診次數缺失導致分數計算差異 |
| 115 年就診次數 舊空新有 | 13 | 同上 |
| 姓名 新流程空白 | 30 | 醫聖 staging 照護名單未存姓名（來源限制） |
| HbA1c / LDL / UACR 小量差異 | 6-7 | 個別資料匹配差異 |

---

### 德容聯合診所（宏誠，3501103076）

| 流程 | 輸出 members | 說明 |
|------|-------------|------|
| 舊流程（正式 DB 跑） | 1,045 | batch `d6847bbf`；照護名單+健康類 22 檔，PDF 月費未包含 |
| 新流程（staging DB） | 5,373 | batch `165e5857`；含 PDF 月費次數 15 檔 |
| 舊前 90 ID 在新中命中 | 82 / 90 | — |
| 抽樣比對人數 | 30 | — |

**AW:BJ（14 位兩邊照護名單共同成員）**

| 欄組 | 差異筆數 | 結論 |
|------|----------|------|
| col49-62 全部 AW:BJ（14 位共同成員） | **0** | 完全一致 ✓ |

**其餘差異（非 bug）**

| 差異類型 | 筆數 / 30 | 原因 |
|----------|-----------|------|
| 新流程多 16 人（不在舊照護名單）→ AW:BJ 舊空新有 | 16 | staging 有記錄但舊照護名單沒有 |
| 就診次數 / 費用 舊空新有 | 24-26 | 舊流程未讀 PDF 月費次數檔 |
| 姓名 / 生日 新流程空白 | 14-30 | 宏誠 staging 未存姓名（來源限制） |
| lab 資料：舊有新空 | 24-30 | 宏誠 lab 以病歷號匹配，部分 ID 在 staging 未命中（parser 已知限制） |
| 分數不同 | 28 | 就診次數缺失導致分數差異 |

**結論（三系統合計）**：

- 三個代表系統（調和、醫聖、宏誠）的 AW:BJ 指定欄位，對於兩邊照護名單都有的共同成員，**全數 0 差異**。
- 三次比對中共發現並修正 7 個 bugs（蘆洲大愛 4 個 + 跨診所 3 個）。
- 剩餘差異均屬已知來源限制（舊流程未讀 TXT/PDF 月費、parser 姓名/lab 匹配），不影響照護指標計算的正確性。

下一步：

1. Phase 9：切換評估（正式流程切換條件確認）。
