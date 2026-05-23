# AGENTS.md

本專案是診所會員名單、疾病分流、檢驗資料、預防保健資料、分析表統計的 Excel 自動化工具。

請 Codex / AI coding agent 修改本專案時遵守以下規則。

---

# 一、基本原則

1. 不要破壞既有 Excel 樣板格式。
2. 不要任意改變既有欄位順序。
3. 不要任意刪除既有功能。
4. 優先做最小幅度修改，除非使用者明確要求重構。
5. Python 需相容 Python 3.9。
6. 輸出檔名需加時間戳，避免覆蓋舊檔。
7. 輸出檔預設存到輸入檔同資料夾或使用者指定資料夾。
8. 不要自動刪除原始檔案、樣板檔案或輸出檔案。
9. 中文訊息請使用繁體中文。

---

# 二、openpyxl 效能規則

處理大型 Excel / ODS / XLSX 時，不要用 openpyxl 完整載入整本 workbook 後才做統計。

禁止以下流程：

1. 產生輸出 Excel 後，再用 openpyxl `load_workbook()` 重新打開輸出檔統計。
2. 掃描輸出 workbook 的所有 sheet。
3. 掃描所有列、所有欄來統計「分析表」。
4. 對大型 xlsx 做完整 workbook 驗證。
5. openpyxl 卡住後反覆重跑同一個驗證流程。
6. 為整張工作表所有空白格套樣式。
7. 對整本 workbook 做無必要的深度格式複製。

正確做法：

1. 在資料讀取、分類、分流階段同步建立統計 counter。
2. 每筆資料分到哪個 sheet 時，立即統計備註代碼 0~5。
3. 最後直接把 counter 寫入「分析表」。
4. 不要再從輸出 workbook 回頭掃分頁統計。
5. openpyxl 主要用於寫入 Excel、保留格式、套用樣板，不要用來反覆掃描大型輸出檔。

---

# 三、分析表統計規則

分析表統計必須在資料分類階段完成。

建議使用：

```python
from collections import defaultdict, Counter

analysis_counter = defaultdict(Counter)
```

範例：

```python
sheet_name = "DM"
remark_code = "3"
analysis_counter[sheet_name][remark_code] += 1
```

常見 sheet：

```text
DM
CKD
DKD
ASCVD
CKD+ASCVD
DM+ASCVD
DKD+ASCVD
```

備註代碼通常是：

```text
0
1
2
3
4
5
```

最後直接把 `analysis_counter` 寫入「分析表」。

不要為了統計分析表而重新讀取輸出 workbook。

---

# 四、資料比對基本規則

身分證號比對前要正規化：

```python
def normalize_id(value):
    if value is None:
        return ""
    s = str(value).strip().upper()
    if s.endswith(".0"):
        s = s[:-2]
    s = s.replace(" ", "")
    return s
```

日期需支援：

```text
2026/03/20
2026-03-20
115/03/20
1150320
```

缺失值視為空：

```text
-
—
–
空白
None
NaN
```

不要用未正規化的 ID 直接比對。

---

# 五、Excel 格式規則

輸出 workbook 需盡量保留：

1. 樣板格式。
2. 合併儲存格。
3. 欄寬。
4. 列高。
5. 框線。
6. 底色。
7. 字型。
8. 對齊方式。
9. 分析表固定版面。

不要因為改效能而破壞樣板外觀。

不要把格式套到整張表、整欄、整列，避免檔案變大或卡住。

---

# 六、卡住程序處理規則

如果 openpyxl / Python 驗證程序卡住：

1. 先詢問使用者是否允許停止。
2. 可以使用 `ps` 查詢 PID。
3. 不要自行刪除檔案。
4. 不要停止不相關程序。
5. 停止後不要再重跑同一個慢速驗證。
6. 應改程式邏輯，避免再次卡住。

可使用：

```bash
ps -axo pid,command | grep "檔名關鍵字"
```

停止前需讓使用者確認。

---

# 七、執行修正後程式規則

當 Codex 完成程式修改，需要執行修正後的 Python 程式產生正式 Excel 時，需先詢問使用者。

允許執行的條件：

1. 執行的是本專案內的 Python 程式。
2. 目的是產生正式 Excel 輸出檔。
3. 不會刪除原始檔案。
4. 不會覆蓋既有輸出檔，檔名需加時間戳。
5. 不會執行慢速 openpyxl 完整 workbook 驗證。

執行完成後，只回報：

1. 是否成功。
2. 輸出檔路徑。
3. 是否有錯誤訊息。

不要再重新打開大型輸出 workbook 做完整驗證。

---

# 八、修改風格

修改程式時：

1. 優先保留原本函式名稱。
2. 優先新增小函式，不要大幅破壞架構。
3. 不要刪除使用者註解。
4. 不要改掉使用者指定的檔名規則。
5. 不要把繁體中文訊息改成簡體中文。
6. 不要把醫療欄位名稱翻成英文。
7. 不要改變 Excel sheet 名稱，除非使用者要求。
8. 若不確定，先做最小修改。

---

# 九、目前最重要的要求

若任務涉及「分析表」統計，務必遵守：

```text
不要重新讀取輸出 workbook 掃描分頁。
不要用 openpyxl 對大型輸出檔做完整驗證。
分析表統計必須在資料分類階段同步累計。
最後直接把 counter 寫入分析表。
```

---

# 十、專案結構說明