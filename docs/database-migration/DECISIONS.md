# 資料庫遷移設計決策

## D001：保留 PostgreSQL

狀態：採用。

原因：

- 現有流程已使用 schema、UUID、JSONB、GIN、批次回復與稽核。
- 後續目標是多診所、歷史版本及資料庫輸出，不適合直接降為 SQLite。

## D002：旁路建置

狀態：採用。

新資料庫流程不得直接取代現有正式流程。必須先建立獨立模組、schema、
輸出目錄及比對報告。

## D003：解析器不直接寫資料庫

狀態：採用。

解析器輸出標準資料集與驗證結果；storage 層負責 transaction、
staging、current、history 與 rollback。

## D004：雙 DataProvider

狀態：採用。

保留 `FolderDataProvider`，新增 `PostgresDataProvider`。
兩者提供相同資料契約，避免輸出層綁死資料來源。

## D005：設定檔只描述差異

狀態：採用。

JSON/YAML 可設定系統類型、編碼、工作表別名、欄位別名及少量開關。
複雜清洗與計算規則保留在可測試的 Python 程式。

## D006：所有來源先進 raw，非會員仍進 staging

狀態：採用。

- 每個來源檔都登錄於 `meta.source_files`。
- 每個可讀的非空來源列都寫入 `raw.uploaded_rows`。
- 是否存在照護名單只影響會員欄位是否可補齊，不影響費用、名單旗標、
  P4P、篩檢或檢驗資料是否進 staging。
- 非會員資料保留真實身分證號；沒有來源的姓名、生日、電話等欄位留空。
- 重複彙總檔保留於 raw，但不重複進入結構化統計。
- 遷移驗證期間不更新 `raw.current_uploaded_rows`，避免影響既有正式流程。
