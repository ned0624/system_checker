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

## 一鍵執行（撈 CSV + 填表）

> 公司網路看不到 GitHub 上的 `.bat` 檔時，請看同名的 `.txt`：[run_weekly.bat.txt](run_weekly.bat.txt)、[build_exe.bat.txt](build_exe.bat.txt)。
> 內容與 `.bat` 相同（UTF-8 編碼方便在網頁上看）。要從 txt 還原成 bat 時，用記事本「另存新檔」→ 編碼選 **ANSI**，檔名改成 `.bat`，否則中文訊息會變亂碼。

```
D:\SystemCheck\
├─ run_weekly.bat          ← 每週點這個
├─ EventLogCheck.ps1       撈事件檢視器 → EventLog\日期_伺服器名.csv
├─ serverlist.txt          伺服器清單（一行一台）
├─ fill_check.py          填檢核表、貼追蹤表（或 fill_check\ 資料夾內的 exe）
├─ search_ignore.py / .exe 可忽略事件ID 搜尋
├─ config.json
├─ EventLog\               CSV 輸出（每週累積）
└─ logs\                   執行紀錄（依日期命名）
```

### 撈取腳本（兩個版本）

| 腳本 | 在哪裡執行 | 撈誰 | 輸出 |
|---|---|---|---|
| `EventLogCheck.ps1` | 你的電腦（`run_weekly.bat` 用這支） | `serverlist.txt` 裡的每台伺服器（遠端） | `EventLog\日期_伺服器名.csv` |
| `EventLogCheck_Local.ps1` | 每台伺服器上 | 本機 | `EventLog\日期_電腦名稱.csv` |

共同行為：
- 只抓 **重大 / 錯誤 / 警告**（Level 1~3），不抓資訊；預設近 7 天
- 同一來源、同一 ID 只留一筆，`TimeCreated` 為最近一次發生時間，`Count` 為期間內次數
- 某個記錄檔讀取失敗會寫進 log；整台連不上時**不產生 CSV**，log 最後列出「撈取失敗」的伺服器
- 沒有任何錯誤/警告的伺服器會產生**只有標題的 CSV**，檢核表會填 O O
- log 與 fill_check 共用 `logs\日期.log`

參數（都可省略）：
```
powershell -ExecutionPolicy Bypass -File EventLogCheck.ps1 -ServerList D:\SystemCheck\serverlist.txt -OutDir D:\SystemCheck\EventLog -LogDir D:\SystemCheck\logs -Days 7
powershell -ExecutionPolicy Bypass -File EventLogCheck_Local.ps1 -OutDir \\fileserver\SystemCheck\EventLog -Days 7
```
本機版的 `-OutDir` 可設成共用資料夾，讓各台伺服器的 CSV 集中到同一處給 fill_check 讀。

`serverlist.txt` 一行一台，空行與 `#` 開頭的行會略過。

CSV 多了 `TimeCreated`、`Count` 欄位；追蹤表想顯示的話，在追蹤表加標題（例如「發生時間」「次數」），
再於 config.json 的 `track.columns` 加上 `"發生時間": "TimeCreated", "次數": "Count"`。

`run_weekly.bat` 可直接雙擊執行，依序：
1. `EventLogCheck.ps1`：讀 `serverlist.txt`，撈每台伺服器近 7 天的事件，輸出 `EventLog\20261016_伺服器名.csv`
2. 檢查有沒有產生**今天日期**的 CSV；沒有就停止（避免拿上週的 CSV 重跑、蓋掉上週追蹤表裡填的影響 / Action）
3. 有 `fill_check.exe` 就執行 exe，沒有就用 Python 執行 `fill_check.py`；兩者都沒有會顯示錯誤並停止
4. 依結果顯示「完成」或「填表沒有完成」，停在視窗等按任意鍵；詳細結果記在 `logs\`

也可以排進 Windows 工作排程器，每週固定時間執行 `run_weekly.bat`。

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

### 6. Log 檔

log 寫在 `logs\` 資料夾，依執行日期命名（例：`20261016.log`），同一天執行多次會接在同一個檔案後面，
每次以 `===== 開始執行 日期 時間 =====` 分隔。內容與畫面輸出相同，最後有彙總：

```
===== 完成：CSV 共 4 個，有狀況 2 個 =====
  20261015_BAD.csv: 錯誤：CSV 無法讀取，整個略過（找不到欄位「LevelDisplayName」…）
  20261015_NOHOST.csv: 略過：檢核表找不到此伺服器，未填 O/X（追蹤表照貼）
```

| 狀況 | 處理 |
|---|---|
| 單一 CSV 讀不到、缺欄位 | 略過該 CSV，其他照常，彙總列出 |
| 檢核表找不到伺服器 | 不填 O/X，追蹤表照貼，彙總列出 |
| 設定檔、路徑、工作表、標題有誤 | 程式中止，**兩個 Excel 都不會存檔**；畫面顯示一行原因，log 內有完整錯誤內容 |
| Excel 開啟中無法存檔 | 程式中止，關閉 Excel 後重跑（重跑不會重複貼） |

log 資料夾可在 config.json 加 `"log_dir": "D:/週檢/logs"` 指定。包成 exe 時視窗會直接關閉，請看 log 確認結果。

## 可忽略事件ID 搜尋工具

`search_ignore.py` 是獨立程式，不需要 config.json，單獨複製 exe 就能用。

```
python search_ignore.py
```

1. 先輸入可忽略清單的 Excel 路徑（可把檔案直接拖進視窗；直接按 Enter 沿用上次的檔案）
2. 再輸入 ID 查詢，多筆用 `,`、`，` 或 `、` 隔開：

```
可忽略清單 Excel 路徑> D:\週檢\系統檢核問題追蹤.xlsx
已載入可忽略事件ID 共 25 筆（系統檢核問題追蹤.xlsx / 可忽略事件ID / 欄位: 可忽略事件ID）

ID> 200,1801
  [可忽略] 200
           第 2 列  說明: WU
  [不在清單] 1801
共 2 筆：可忽略 1 筆，不在清單 1 筆
```

- 讀「可忽略事件ID」工作表；沒有這個工作表時會列出所有工作表讓你選
- ID 欄依標題「可忽略事件ID」或「ID」尋找，都沒有時用第一欄
- 上次的路徑記在程式旁的 `search_ignore_last.txt`

| 指令 | 作用 |
|---|---|
| `r` / `reload` | 重新讀取 Excel（清單更新後使用） |
| `f` / `file` | 換一個 Excel 檔 |
| `q` / `quit` / `exit` / `離開` | 結束程式，也可直接關閉視窗 |

## 設定檔 config.json

### 路徑設定 `paths`

三個位置可以各自不同：

```json
"paths": {
    "csv_dir": "D:/SystemCheck/EventLog",
    "check_xlsx": "E:/健檢/BOS_系統穩定健檢狀態記錄表_{week}.xlsx",
    "track_xlsx": "//fileserver/IT/週檢/系統檢核問題追蹤.xlsx"
}
```

路徑中的 `{week}` 會自動換成這批 CSV 的週別（例：`_{week}.xlsx` → `_W2642.xlsx`），
檢核表每週換檔名時不用改設定。執行時畫面與 log 會印出實際使用的三個路徑。

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

## 下載執行檔

到 [Releases](https://github.com/ned0624/system_checker/releases/latest) 下載（程式一更新推到 master，GitHub Actions 會自動在 Windows 上重新打包並發布新版本）：

| 檔案 | 內容 |
|---|---|
| `SystemCheck.zip` | 解壓到 `D:\` 會得到 `D:\SystemCheck\`：`fill_check\`（含 `fill_check.exe`、`config.json`）、`search_ignore.exe`、`run_weekly.bat`、`EventLogCheck.ps1`、空的 `EventLog\`、`logs\` |
| `search_ignore.exe` | 只要搜尋工具時單獨下載 |

使用前：
1. 修改 `D:\SystemCheck\fill_check\config.json` 的 `paths`（CSV、檢核表、追蹤表位置）
2. 準備 `D:\SystemCheck\serverlist.txt`（一行一台伺服器）
3. 雙擊 `run_weekly.bat`

## 打包成執行檔（Windows）

```
pip install pyinstaller pandas openpyxl
build_exe.bat
```

| 程式 | 打包方式 | 產出 | 說明 |
|---|---|---|---|
| search_ignore | `--onefile`，排除 pandas/numpy/PIL/lxml | `dist\search_ignore.exe`（約 10 MB） | 單一檔案，複製到哪都能用 |
| fill_check | `--onedir` | `dist\fill_check\` 資料夾 | 整個資料夾一起複製，`config.json` 放在 `fill_check.exe` 旁邊 |

**為什麼不全部用 `--onefile`？** `--onefile` 每次啟動都要把整包程式解壓縮到暫存資料夾，
公司防毒會逐一掃描解出來的檔案；含 pandas 的程式可能要等好幾分鐘才出現第一行文字。
`--onedir` 不用解壓縮，啟動快很多。

啟動仍然很慢時：
- exe 放在網路磁碟（例如 P:）會更慢，請複製到本機（例如 D:\SystemCheck）執行
- 請 IT 把 exe 所在資料夾加入防毒的排除清單

`run_weekly.bat` 會先找 `D:\SystemCheck\fill_check\fill_check.exe`（把 `dist\fill_check\` 整個資料夾複製到 `D:\SystemCheck\fill_check\`），找不到才用 Python 執行 `fill_check.py`。
