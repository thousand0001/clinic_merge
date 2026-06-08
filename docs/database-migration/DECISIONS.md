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

