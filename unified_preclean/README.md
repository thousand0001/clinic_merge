# unified_preclean

旁路版統一前置清洗原型。第一階段只做來源盤點與格式辨識，不改原始檔、不產正式 Excel、不取代現有前置清洗。

## 目標分工

- `unified_preclean`：讀懂不同診所原始資料，盤點檔案、判斷資料類型、找表頭與欄位 alias、產生覆蓋報告。
- `選會員_共用核心_*.py`：維持只吃標準資料、合併、計分與輸出 Excel。

## 使用

```bash
.venv/bin/python -m unified_preclean.cli <來源資料夾> -o /private/tmp/report.xlsx --json-summary
```

## 目前階段

- 支援讀取 `.xlsx/.xlsm/.xls/.ods/.csv/.txt/.pdf`。
- 依檔名、sheet 名與表頭推定資料類型。
- 偵測表頭列、核心欄位 alias、有效 ID 數、日期/金額/次數欄。
- 後續才會把各 parser 輸出成標準中間資料，不在第一階段直接接管正式流程。

