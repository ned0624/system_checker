# 週檢小幫手

依每週系統檢核產出的 CSV，自動完成兩件事：

1. 在 **BOS_SystemCheck.xlsx** 填寫每台伺服器的「警告 / 錯誤」O/X 與備註
2. 把 CSV 全部內容貼到 **系統檢核問題追蹤.xlsx** 留存紀錄

另附 **可忽略事件ID 搜尋工具**，可快速查詢事件 ID 是否在可忽略清單。

## 檔案說明

| 檔案 | 說明 |
|---|---|
| `fill_check.py` | 主程式：填檢核表 + 貼追蹤表 |
| `search_ignore.py` | 可忽略事件ID 搜尋工具（互動式 CLI） |
| `config.json` | 設定檔（檔名、工作表、欄位標題） |
| `csv/` | 放每週的 CSV 檢核結果 |
| `BOS_SystemCheck.xlsx` | 系統檢核表（要填 O/X 的表） |
| `系統檢核問題追蹤.xlsx` | 問題追蹤表，內含「可忽略事件ID」清單與「事件檢視器_BOS」 |

## 環境需求

```
pip install pandas openpyxl
```

## 每週作業流程

### 1. 放入 CSV

把本週的 CSV 放到 `csv/` 資料夾，檔名格式：

```
日期_伺服器名.csv      例: 20261015_P3-WFA-AP01V.csv
```

- 伺服器名要與 BOS_SystemCheck.xlsx 第一欄的名稱相同（不分大小寫）
- 日期用來計算週別；省略日期時以執行當天計算
- CSV 需有 `ID`、`事件分類`、`事件說明` 欄位（欄名可在 config.json 修改），編碼 UTF-8 或 Big5 皆可
- 只讀等級為「錯誤」「警告」的事件（「重大 / Critical」算錯誤），「資訊」等其他等級一律略過，不填檢核表也不貼追蹤表

### 2. 更新可忽略清單（視需要）

在「系統檢核問題追蹤.xlsx」的「可忽略事件ID」工作表 `ID` 欄加入可忽略的事件 ID。
一格可放多個 ID，用 `,`、`、` 或換行分隔。

### 3. 執行

先**關閉**兩個 Excel 檔，再執行：

```
python fill_check.py
python fill_check.py <csv資料夾> <檢核表xlsx> <追蹤表xlsx>   # 暫時指定其他路徑
```

三個路徑平常設定在 config.json 的 `paths`（見下方「設定檔」）。路徑不存在時程式會直接提示並停止。

### 4. 程式處理內容

對每一個 CSV：

**(A) 填 BOS_SystemCheck.xlsx**

| 情況 | 警告欄 | 錯誤欄 | 備註 |
|---|---|---|---|
| 所有 ID 都在可忽略清單 | O | O | 空白 |
| 有不在清單的「警告」事件 | X | — | 列出事件 |
| 有不在清單的「錯誤」事件 | — | X | 列出事件 |

沒有問題的那一欄填 O。備註格式：`[等級] ID 事件說明`，同一事件只記一次。
檢核表找不到該伺服器時會顯示「略過」，但仍會貼到追蹤表。

**(B) 貼到 系統檢核問題追蹤.xlsx 的「事件檢視器_BOS」工作表**

工作表要先有標題列，程式依標題找欄位，接在最後一筆資料下面：

| 標題 | 內容 |
|---|---|
| Server Name | 檔名中的伺服器名 |
| 周 | 依檔名日期算 ISO 週次，`W` + 年後兩碼 + 週次，例：20261015 → `W2642` |
| 錯誤/警告 | CSV 的 `LevelDisplayName` |
| 事件識別碼 | CSV 的 `Id` |
| 記錄檔名稱 | CSV 的 `LogName` |
| 事件內容 | CSV 的 `Message` |
| 可能造成的影響、Action | 不動，人工填寫 |
| 可忽略 | 程式填公式 `=IF(COUNTIF('可忽略事件ID'!$A$2:$A$25555,"="&D列)>0,"V","")`，在可忽略清單就顯示 V |

- 只貼錯誤/警告事件（含可忽略的事件），資訊不貼
- 同一伺服器、同一週別的舊資料會先刪除再貼上，**重跑不會重複**（那幾列人工填的影響、Action 也會一起被刪）
- 標題與 CSV 欄位的對應可在 config.json 的 `track.columns` 修改

### 5. 確認結果

畫面會列出每台伺服器的結果，例如：

```
可忽略事件ID 共 25 筆
P3-WFA-AP01V
  檢核表: 警告=O 錯誤=X
    [錯誤] 2002 服務停止
  追蹤表: W2642 貼上 2 筆
已儲存: BOS_SystemCheck.xlsx
已儲存: 系統檢核問題追蹤.xlsx
```

出現「無法寫入…」表示 Excel 檔仍開啟中，關閉後重跑即可。

## 可忽略事件ID 搜尋工具

```
python search_ignore.py
```

執行後輸入 ID 查詢，多筆用 `,`、`，` 或 `、` 隔開：

```
ID> 1001,8
  [可忽略] 1001
           第 2 列  說明: 磁碟警告可忽略
  [不在清單] 8
共 2 筆：可忽略 1 筆，不在清單 1 筆
```

| 指令 | 作用 |
|---|---|
| `r` / `reload` | 重新讀取 Excel（清單更新後使用） |
| `q` / `quit` / `exit` / `離開` | 結束程式，也可直接關閉視窗 |

## 設定檔 config.json

### 路徑設定 `paths`

```json
"paths": {
    "csv_dir": "D:/週檢/csv",
    "check_xlsx": "D:/週檢/BOS_SystemCheck.xlsx",
    "track_xlsx": "//fileserver/IT/週檢/系統檢核問題追蹤.xlsx"
}
```

| 欄位 | 說明 |
|---|---|
| `csv_dir` | CSV 資料夾 |
| `check_xlsx` | 檢核表（填 O/X） |
| `track_xlsx` | 問題追蹤表（貼 CSV），**可忽略事件ID 清單也從這個檔案讀** |

- 可用絕對路徑、網路磁碟或共用資料夾路徑
- Windows 路徑請用正斜線 `/`，或把反斜線寫兩次：`"D:\\週檢\\csv"`
- 相對路徑以 config.json 所在資料夾為基準（包成 exe 時為 exe 所在資料夾）

### 其他設定

| 區段 | 欄位 | 說明 |
|---|---|---|
| `ignore_list` | `sheet` / `id_column` | 可忽略事件ID 的工作表與欄位 |
| `csv` | `id_column` / `level_column` / `desc_column` | CSV 的 ID、等級（警告/錯誤）、說明欄名 |
| `check_sheet` | `sheet` | 工作表名稱，`null` = 第一個工作表 |
| | `header_row` | 標題列在第幾列 |
| | `warn_header` / `error_header` / `note_header` | 警告、錯誤、備註欄標題 |
| | `output` | 另存路徑，`null` = 覆寫原檔 |
| `track` | `sheet` | 追蹤表的工作表（例：`事件檢視器_BOS`） |
| | `header_row` | 標題列在第幾列 |
| | `host_header` / `week_header` | 伺服器、週別欄的標題（`Server Name` / `周`） |
| | `columns` | 追蹤表標題 → CSV 欄位，例：`"事件識別碼": "Id"` |
| | `ignore_header` | 可忽略欄的標題，`null` = 不填 |
| | `output` | 另存路徑，`null` = 覆寫原檔 |

> 注意：`track.sheet` 不要設成「可忽略事件ID」，否則貼上的 ID 下次都會被當成可忽略。

## 打包成執行檔（Windows）

```
pip install pyinstaller
pyinstaller --onefile search_ignore.py
pyinstaller --onefile fill_check.py
```

產出在 `dist\`。exe 旁邊放 `config.json`；Excel 與 CSV 資料夾放在 `paths` 設定的位置（用相對路徑時就放在 exe 旁邊）。
