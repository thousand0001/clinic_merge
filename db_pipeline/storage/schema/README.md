# staging_v1 migration

`001_staging_v1.sql` 是旁路 staging schema 草案。

目前狀態：

- 尚未在正式 PostgreSQL 執行。
- 不修改既有 `meta`、`raw`、`audit` 資料。
- 只建立新的 `staging_v1` schema。
- 所有標準資料列都保留來源檔、工作表、列號及雜湊。

執行前必要檢查：

1. PostgreSQL 服務可連線。
2. 匯出現有 schema 備份。
3. 確認 `meta.clinics`、`meta.import_batches` 的主鍵型別。
4. 使用 transaction 在測試資料庫驗證。
5. 執行後核對所有索引與外鍵。

目前不可直接執行於正式資料庫。

