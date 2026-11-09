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
| `monitor_check.py` | 監控數據分析：Kettle 匯出的 CSV 有沒有超過 spec、使用量是否持續上升 |
| `config.json` | **共用設定檔**：Kettle、monitor_check、fill_check 都讀這一份（路徑、template、工作表、欄位標題） |
| `csv/` | 放每週的 CSV 檢核結果 |
| `BOS_SystemCheck.xlsx` | 系統檢核表（要填 O/X 的表） |
| `系統檢核問題追蹤.xlsx` | 問題追蹤表，內含「可忽略事件ID」清單與「事件檢視器_BOS」 |

## 環境需求

```
pip install pandas openpyxl
```

## 一鍵執行（撈 event log → Kettle + monitor_check → 填表）

> 公司網路看不到 GitHub 上的 `.bat` 檔時，請看同名的 `.txt`：[run_weekly.bat.txt](run_weekly.bat.txt)、[build_exe.bat.txt](build_exe.bat.txt)。
> 內容與 `.bat` 完全相同（全英文），把副檔名 `.txt` 去掉就是 bat。

```
D:\SystemCheck\
├─ run_weekly.bat          ← 每週點這個
├─ EventLogCheck.ps1       撈事件檢視器 → EventLog\伺服器名.csv（每次覆蓋）
├─ serverlist.txt          伺服器清單（一行一台）
├─ excluded_ids.txt        撈取時直接排除的事件 ID
├─ fill_check.py          填檢核表、貼追蹤表（或 fill_check\ 資料夾內的 exe）
├─ monitor_check.exe       監控數據分析報表（Kettle 匯出的 CSV）
├─ search_ignore.py / .exe 可忽略事件ID 搜尋
├─ config.json             共用設定檔（Kettle、monitor_check、fill_check）
├─ 2026-W45\               Kettle 每週匯出的監控 CSV（主機名.csv）
├─ EventLog\               CSV 輸出（每週累積）
└─ logs\                   執行紀錄（日期.log；本機版為 日期_電腦名稱.log）
```

### 撈取腳本（兩個版本）

| 腳本 | 在哪裡執行 | 撈誰 | 輸出 |
|---|---|---|---|
| `EventLogCheck.ps1` | 你的電腦（`run_weekly.bat` 用這支） | `serverlist.txt` 裡的每台伺服器（遠端） | `EventLog\伺服器名.csv` |
| `EventLogCheck_Local.ps1` | 每台伺服器上 | 本機 | `EventLog\電腦名稱.csv` |

共同行為：
- CSV 輸出到**固定資料夾**（SETTINGS 的 `OutDir`，例 `D:\SystemCheck\EventLog`），檔名 `<伺服器>.csv`，**每次執行直接覆蓋**；
  config.json 的 `paths.csv_dir` 要指同一個資料夾
- 只抓 **重大 / 錯誤 / 警告**（Level 1~3），不抓資訊；預設近 7 天
- 同一來源、同一 ID 只留一筆，`TimeCreated` 為最近一次發生時間，`Count` 為期間內次數
- 某個記錄檔讀取失敗會寫進 log；整台連不上時**不產生 CSV**，log 最後列出「撈取失敗」的伺服器
- 排除後沒有任何錯誤/警告的伺服器**不產生 CSV**，只在 log 記錄
- 這次沒產生 CSV（沒有事件或撈取失敗）時，會**刪掉上次留下的 `<伺服器>.csv`**（log：`removed old CSV`），fill_check 不會誤用舊資料，該台填 O O
- 結束碼：`0` 有寫出 CSV、`1` 錯誤（資料夾無法寫入／全部伺服器失敗）、`2` 沒有任何事件可報告（bat 會略過填表）
- 開始時會測試輸出與 log 資料夾能否寫入；不行就顯示執行帳號並停止（共用資料夾需給該帳號「修改」權限）
- log：`EventLogCheck.ps1` 與 fill_check 共用 `logs\日期.log`；`EventLogCheck_Local.ps1` 寫 `logs\日期_電腦名稱.log`（例如 `20261029_T1-WFA-AP02V.log`），多台寫到同一個共用資料夾也分得開
- **排除 ID**：`excluded_ids.txt`（預設 `D:\SystemCheck\excluded_ids.txt`，可用 `-ExcludeFile` 指定）裡的事件 ID 直接不寫進 CSV，
  所以不會出現在檢核表備註與追蹤表。檔案不存在就不排除。一行一個或用逗號分隔，`#` 之後是註解。
  log 會記錄載入了幾個 ID、每台排除了幾筆事件。

**路徑設定**：直接改 ps1 開頭的 `SETTINGS` 區塊（`$Settings = @'` 到 `'@` 之間），例如：
```
$Settings = @'
OutDir      = \\xintec.com.tw\Supporting_Org\中文資料夾\BOS\Z0_Others\SystemCheck\EventLog
LogDir      = \\xintec.com.tw\Supporting_Org\中文資料夾\BOS\Z0_Others\SystemCheck\logs
ExcludeFile = D:\SystemCheck\excluded_ids.txt
Days        = 7
'@
```
可以寫中文資料夾與共用資料夾路徑：程式會自己判斷 ps1 的存檔編碼（UTF-8 有無 BOM、ANSI 都可以）重新讀這一段，
不會再變成亂碼。輸出或 log 資料夾無法建立/寫入、或所有伺服器都失敗時，ps1 會以結束碼 1 結束，`run_weekly.bat` 會停下不填表。
`fill_check` 讀 CSV 的資料夾是 config.json 的 `paths.csv_dir`，要和 ps1 的 `OutDir` 指同一個位置。

**撈取時間**：SETTINGS 的 `StartTime` / `EndTime`（格式 `yyyy-MM-dd` 或 `yyyy-MM-dd HH:mm`，只寫日期 = 當天 00:00）：

| 設定 | 撈取區間 |
|---|---|
| 都不設定（預設） | 現在往前 7 天 |
| 只設 `Days` | 現在往前 `Days` 天 |
| 只設 `StartTime` | `StartTime` ～ 現在 |
| `StartTime` + `EndTime` | 指定區間，例：`2026-10-15` ～ `2026-10-22` |

CSV 檔名的日期與路徑中的 `{week}` 依**結束時間**那天決定，補撈過去某週後用 `fill_check --week` 指定同一週處理：
```
powershell -ExecutionPolicy Bypass -File EventLogCheck.ps1 -StartTime 2026-10-15 -EndTime 2026-10-22
fill_check.exe --week W2643
```

命令列參數（可省略，有給會覆蓋 SETTINGS）：
```
powershell -ExecutionPolicy Bypass -File EventLogCheck.ps1 -ServerList D:\SystemCheck\serverlist.txt -OutDir D:\SystemCheck\EventLog -LogDir D:\SystemCheck\logs -Days 7
powershell -ExecutionPolicy Bypass -File EventLogCheck_Local.ps1 -OutDir \\fileserver\SystemCheck\EventLog -Days 7
```
本機版的 `-OutDir` 可設成共用資料夾，讓各台伺服器的 CSV 集中到同一處給 fill_check 讀。

`serverlist.txt` 一行一台，空行與 `#` 開頭的行會略過。

CSV 多了 `TimeCreated`、`Count` 欄位；追蹤表想顯示的話，在追蹤表加標題（例如「發生時間」「次數」），
再於 config.json 的 `track.columns` 加上 `"發生時間": "TimeCreated", "次數": "Count"`。

`run_weekly.bat` 可直接雙擊執行，依序：

| 步驟 | 內容 | 失敗時 |
|---|---|---|
| [1/4] 撈 event log | `COLLECTOR` 指定的撈取腳本（預設 `EventLogCheck.ps1`；只撈本機時把 `set COLLECTOR=` 改成 `EventLogCheck_Local.ps1`） | 結束碼 1（資料夾無法使用或全部伺服器失敗）→ **整個停止**，不跑 Kettle、不填表；結束碼 2（沒有事件）照常往下 |
| [2/4] Kettle | `call %KITCHEN% /file:%KETTLE_JOB%`（Job 讀 config.json，見「Kettle 讀 config.json」） | Kitchen / Job 找不到或結束碼非 0 → 跳過 [3/4]，**繼續填表但監控欄清空** |
| [3/4] monitor_check | 有 `monitor_check.exe` 就執行 exe，沒有就用 Python 執行 `monitor_check.py`，產生監控報表 | 結束碼非 0 → **繼續填表但監控欄清空** |
| [4/4] fill_check | 有 `fill_check\fill_check.exe` 就執行 exe，沒有就用 Python；[2]、[3] 失敗時加 `--clear-monitor` | 依結果顯示 `Done.` 或錯誤訊息 |

最後列出失敗的步驟（`[WARN] Failed steps: Kettle`），停在視窗等按任意鍵。

bat 開頭的設定（依實際安裝位置修改）：

```
set BASE=D:\SystemCheck
set COLLECTOR=EventLogCheck.ps1
set KITCHEN=D:\data-integration\Kitchen.bat
set KETTLE_JOB=%BASE%\kettle\monitor.kjb
```

監控欄清空 = 範本中 CPU / Memory / … 打 O 的格子變空白、備註不寫 `[監控]`，事件檢視器欄照常填寫。

也可以排進 Windows 工作排程器，每週固定時間執行 `run_weekly.bat`。

## 每週作業流程

### 1. 放入 CSV

把本週的 CSV 放到 `csv/` 資料夾，檔名格式：

```
伺服器名.csv           例: P3-WFA-AP01V.csv（撈取腳本的輸出，每次覆蓋）
日期_伺服器名.csv      例: 20261015_P3-WFA-AP01V.csv（舊格式，仍可讀）
```

- 伺服器名要與 BOS_SystemCheck.xlsx 第一欄的名稱相同（不分大小寫）
- 沒有日期的檔名一律算**目標週**（預設本週，`--week` 指定的週也一樣），追蹤表的「周」填目標週；
  有日期的舊檔名照舊依日期計算週別，只處理目標週的；同一台兩種都有時用沒有日期的
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

### 檢核表範本模式（每週產生完整檔案）

在 config.json 的 `check_sheet` 設定 `template` 後，每次執行會產生當週的完整檢核表：

```json
"check_sheet": {
    "template": "D:/SystemCheck/template/BOS_系統穩定健檢狀態記錄表_template.xlsx",
    "output": "//xintec.com.tw/tsm/Supporting_Org/資訊及通訊技術處/80 品質管理/System Check/{week}/BOS_系統穩定健檢狀態記錄表_{week}.xlsx",
    "antivirus_header": "防毒軟體",
    "link_header": "連結",
    "link_target": "//xintec.com.tw/tsm/Supporting_Org/資訊及通訊技術處/80 品質管理/System Check/{week}/Reference/BOS",
    "link_text": "Evidence_{week}",
    ...
}
```

- `output` 路徑可用 `{week}`（W2642）、`{yy}`（26）、`{ww}`（42），週別資料夾不存在會自動建立
- CSV 也可以放在每週的資料夾：ps1 SETTINGS 的 `OutDir` 與 config.json 的 `paths.csv_dir` 都可以寫 `{week}`，例如
  `Z:\Supporting_Org\資訊及通訊技術處\80 品質管理\System Check\{week}\Reference\BOS`。兩邊用同一套 ISO 週別（週一到週日），
  ps1 依執行當天、fill_check 依今天（或 `--week`）換成 `W2643`
- 網路磁碟機代號（例如 `Z:`）只在使用者登入時存在；工作排程器設成「不論使用者是否登入都執行」時，請改用 `\\伺服器\共用` 的 UNC 路徑
- 當週檔已存在（重跑）就以它為底更新，手動填的 CPU、Memory 等欄位會保留；不存在才從範本產生
- **有底色的格子**（灰、黑、黃色分隔列…）與合併儲存格一律不填
- 有本週 CSV 的伺服器：依事件填 O/X 與備註；**沒有本週 CSV 的伺服器：事件檢視器兩欄填 O O**（全部被排除或撈取失敗都會是 O，請看撈取 log 確認失敗的伺服器）
- 防毒軟體欄全部填 O
- 連結欄：設定 `link_target` 時，每一列的超連結由程式產生，例如只連到當週的 BOS 資料夾：
  ```json
  "link_target": "//xintec.com.tw/tsm/Supporting_Org/資訊及通訊技術處/80 品質管理/System Check/{week}/Reference/BOS",
  "link_text": "Evidence_{week}"
  ```
  可用 `{week}`（W2642）、`{yy}`、`{ww}`、`{host}`（該列伺服器名稱）。
  `link_text` 是格子顯示的文字：`"Evidence_{week}"` → `Evidence_W2643`、`"Evidence_W{ww}"` → `Evidence_W43`、
  `"BOS - {host}_Evidence_W{ww}"` → `BOS - erpdis02_Evidence_W43`；`null` = 沿用範本原本的文字並換週別，空白格顯示 `BOS`。
  有底色的連結格不會加連結。不設定 `link_target` 時，只把範本既有連結的 `W2640` → `W2642`、`_W40` → `_W42`
- 只處理**本週**的 CSV（今天所在的 ISO 週）；補跑其他週用 `python fill_check.py --week W2641`
- 不設定 `template` 時維持原本行為（直接改 `paths.check_xlsx`），但一樣只處理本週的 CSV

### Excel 被別人開著時

Excel 開著檔案時會鎖住，任何程式都寫不進去。fill_check 會：

- 先把 CSV 結果算好，再**等檔案沒人開著**才寫入：每 5 分鐘檢查一次，畫面與 log 顯示是誰開著
  （`[WAIT] Check sheet is open by 王小明, retry in 300 s (waited 10/120 min)`）
- 等到可以寫入時**重新讀取最新的檔案**再套用、立刻存檔，不會蓋掉別人在等待期間存的修改
- 超過 2 小時仍被開著：不存檔，顯示 `NOT saved`，以結束碼 3 結束（`run_weekly.bat` 會提示稍後再跑，重跑是安全的）

設定（config.json，可省略）：
```json
"wait_if_open": { "retry_seconds": 300, "max_wait_minutes": 120 }
```
`max_wait_minutes` 設 `0` 表示不等待。建議用 Windows 工作排程器把 `run_weekly.bat` 排在**下班後或清晨**，
大家都沒開檔案的時段，最不容易遇到被鎖住。

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

log 資料夾可在 config.json 加 `"log_dir": "D:/週檢/logs"` 指定，路徑可用 `{week}`（例: `.../System Check/{week}/Reference/BOS`）。包成 exe 時視窗會直接關閉，請看 log 確認結果。

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

## 監控數據分析 monitor_check

分工：**Kettle 只負責撈資料，檢核由 monitor_check（Python）做**。超標用 SQL 也做得到，
但「每天平均、從最高點回落」的趨勢判定在 SQL/Kettle 很難寫、難測、難調門檻。
結果會由 fill_check 填進檢核表（見下方「填進檢核表」），monitor_check.exe 另外產生明細報表。

分析 Kettle（`monitor.ktr`）每週匯出的每台主機 CSV（`D:\SystemCheck\2026-W45\<主機>.csv`，**最近 7 天**、每分鐘一筆，
只有 template 打 O 的欄位）。天數由 config.json 的 `monitor.days` 控制（SQL 用 `interval ${DAYS} day`）。
Kettle 與 monitor_check 已包含在 `run_weekly.bat` 的 [2/4]、[3/4]，不用另外排程：

```
time_min,cpu,memory,swap,disk1,disk2,disk3
upper_spec,60.00,25.60,10.00,40.00,420.00,625.00
2026-11-03 16:42,4.62,22.10,1.20,29.95,377.02,600.60
```

- 檔名 = 主機名稱；第一列可以是表頭，也可以直接是 `upper_spec`（沒表頭時欄位依上面順序）
- `time_min` 以外的欄位都會分析，只打 O 部分欄位或 Kettle 之後加欄位都不用改程式；空白值略過
- 編碼 UTF-8（含 BOM）或 MS950 都可以；讀不了的檔案顯示 `[WARN]` 後繼續下一台
- **只有時間、每一欄都沒有數值**的檔案顯示 `[WARN] <主機>: no values in any column`，
  通常是 Kettle SQL 的 `monitor_item` 對不到（例如 Linux 掛載點名稱不同），見下方「只有時間沒有數據」

### 檢查項目

| 項目 | 判定 |
|---|---|
| 超過 spec | 數值 > `upper_spec`；記錄超標分鐘數、首次/最後時間、最長連續分鐘數、最大值與時間。spec 空白則不檢查 |
| 持續上升沒降下來 | 每天取平均（從最後一筆往回每 24 小時一段，最舊不完整的一天捨棄），一週升幅 ≥ spec 的 5%，且過程中從最高點回落 ≤ spec 的 1%，**而且最後沒有降下來**：最後 60 分鐘平均 − 最前 60 分鐘平均（`edge_minutes`）也要 ≥ spec 的 5%，否則不算持續上升（例：spec 4.8，開始 60 分鐘平均 3.6、最後 60 分鐘平均 3.7，差 0.1 < 0.24 → 已降下來）。報表的 `rise` 就是這個頭尾平均差。日夜起伏、上下跳動的不會誤判；判定上升時估算照目前速度約幾小時（天）後達 spec。不足 3 天的資料不判定（`rising` 顯示 `-`）；沒有 spec 時以第一天平均為基準 |

設定在共用 `config.json` 的 `monitor` 段（Kettle、monitor_check、fill_check 都讀這段）：

```json
"monitor": {
    "csv_dir": "D:/SystemCheck/{yyyy}-W{ww}",
    "days": 7,
    "report_dir": "D:/SystemCheck/monitor_report",
    "x_over_minutes": 5,
    "columns": {"CPU": "cpu", "Memory": "memory", "SWAP": "swap", "Disk 1": "disk1", "Disk 2": "disk2", "Disk 3": "disk3"},
    "trend": null
}
```

| 欄位 | 誰用 | 說明 |
|---|---|---|
| `csv_dir` | Kettle 輸出、monitor_check / fill_check 讀取 | 每週監控 CSV 資料夾；`{week}` / `{yyyy}` / `{yy}` / `{ww}` 三者都支援（例: `.../System Check/{week}/Monitor`） |
| `days` | Kettle | 撈最近幾天（SQL 的 `interval ${DAYS} day`） |
| `report_dir` | monitor_check | 報表輸出資料夾 |
| `x_over_minutes` | fill_check | 連續超過 spec 幾分鐘以上才在檢核表填 X（預設 5）；報表的 `OVER SPEC` 仍列出所有超標 |
| `columns` | fill_check | 檢核表 / template 的監控欄標題 → CSV 欄位 |
| `trend` | monitor_check、fill_check | 持續上升門檻，null = 預設 `{"bucket_minutes": 1440, "min_buckets": 3, "min_rise_pct_of_spec": 5, "max_drop_pct_of_spec": 1, "edge_minutes": 60}` |

log 寫到頂層的 `log_dir`。路徑可用 `{week}` / `{yyyy}` / `{yy}` / `{ww}`（今天的 ISO 週，例: W2645 / 2026 / 26 / 45），
Kettle 的週資料夾 `%x-W%v`（2026-W45）就寫成 `{yyyy}-W{ww}`；相對路徑以 config.json 所在資料夾為基準。
舊版的 `monitor_check.json` 仍可讀（會顯示 `[WARN] ... deprecated`），請把內容併進 `monitor` 段後刪除。

### 輸出

- 報表 `report_dir\yyyyMMdd_monitor_report.csv`（Excel 可直接開），一台主機一個指標一列：
  `host, result, metric, spec, max, max_time, over_count, over_first, over_last, longest_over_min, rise, rising, hours_to_spec`

  | 欄位 | 意思 |
  |---|---|
  | spec | upper_spec |
  | max / max_time | 一週最大值與時間 |
  | over_count / over_first / over_last | 超過 spec 的分鐘數、第一次與最後一次時間 |
  | longest_over_min | 最長連續超標分鐘數 |
  | rise | 升幅（最後 60 分鐘平均 − 最前 60 分鐘平均） |
  | rising | Y = 持續上升、N = 沒有、`-` = 資料不足 |
  | hours_to_spec | 持續上升時，照目前速度約幾小時到 spec |
  - `result`：`OK` / `OVER SPEC` / `RISING` / `OVER SPEC+RISING` / `NO DATA`
- 畫面與 log `log_dir\yyyyMMdd_monitor_check.log`：每台一行摘要，最後列出所有超標與持續上升的項目
- 結束碼：0 = 完成（不論有沒有發現問題），1 = 錯誤（設定檔、CSV 資料夾、報表或 log 寫不進去）

### 執行

```
monitor_check.exe              （讀 config.json 的 monitor.csv_dir）
monitor_check.exe D:\SystemCheck\2026-W44    （臨時指定 CSV 資料夾，例如補看上週）
```

`monitor_check.exe` 是單一檔案、只用 Python 內建模組（不含 pandas），放在 `D:\SystemCheck\`（和 config.json 同一層）。
`run_weekly.bat` 會在 Kettle 之後自動執行它；要單獨重看報表時直接雙擊或用上面的指令。

### Kettle 讀 config.json

Kettle 不寫死路徑，template、輸出資料夾、天數都從 `D:\SystemCheck\config.json` 讀（和 fill_check 同一份 template）：

1. 新增 Transformation「read_config.ktr」：
   - **JSON input**：File 頁加入 `D:\SystemCheck\config.json`（Encoding UTF-8）；Fields 頁：

     | Name | Path | Type |
     |---|---|---|
     | TEMPLATE | `$.check_sheet.template` | String |
     | MONITOR_DIR | `$.monitor.csv_dir` | String |
     | DAYS | `$.monitor.days` | Integer |

   - **Set Variables**：三個欄位設成同名變數，Scope = **Valid in the root job**
2. 原本的 monitor.ktr：
   - **Microsoft Excel input**：File 改成 `${TEMPLATE}`（工作表照原本設定）
   - **Table input**：勾 **Replace variables in script**，SQL 改兩處：
     ```sql
     -- 輸出資料夾：原本 concat('D:/SystemCheck/', date_format(now(), '%x-W%v'), '/') as dir
     -- {week} / {yyyy} / {yy} / {ww} 都換成本週（%x = ISO 年、%v = ISO 週），順序不能對調
     concat(
       replace(replace(replace(replace('${MONITOR_DIR}',
         '{week}', concat('W', right(date_format(now(), '%x'), 2), date_format(now(), '%v'))),
         '{yyyy}', date_format(now(), '%x')),
         '{yy}',   right(date_format(now(), '%x'), 2)),
         '{ww}',   date_format(now(), '%v')),
       '/') as dir
     -- 天數：原本 interval 7 day（monitor_time 與 lm_time 兩處）
     interval ${DAYS} day
     ```
3. 新增 Job「monitor.kjb」：Start → Transformation（read_config.ktr）→ Transformation（monitor.ktr）
4. Job 存成 `D:\SystemCheck\kettle\monitor.kjb`（`run_weekly.bat` 的 `KETTLE_JOB`），每週由 `run_weekly.bat` 執行，不用另外排程

之後要換 template、輸出資料夾或天數，只改 config.json。

### 填進檢核表（fill_check）

`config.json` 有 `monitor` 段時，fill_check 會分析本週（或 `--week` 指定週）的 Kettle 資料夾，
填進檢核表的 CPU / Memory / SWAP / Disk 1 / Disk 2 / Disk 3 欄：

```json
"monitor": {
    "csv_dir": "D:/SystemCheck/{yyyy}-W{ww}",
    "columns": {"CPU": "cpu", "Memory": "memory", "SWAP": "swap", "Disk 1": "disk1", "Disk 2": "disk2", "Disk 3": "disk3"},
    "trend": null
}
```

| 狀況 | 格子 | 備註 |
|---|---|---|
| 範本**沒打 O** | 不動 | — |
| 有打 O、超過 spec 且**連續 ≥ 5 分鐘** | X | `[監控] CPU 超過 spec 60.00 共 30 分鐘，最長連續 30 分鐘，最高 75.00（10/13 11:00）` |
| 有打 O、超過 spec 但最長連續 < 5 分鐘 | O | `[監控] CPU 超過 spec 60.00 共 7 分鐘，最長連續 4 分鐘，最高 75.00（10/11 10:40）` |
| 有打 O、持續上升（沒超標） | O | `[監控] Memory 持續上升 15.43 → 20.62，約 5.8 天達 spec` |
| 有打 O、正常 | O | — |
| 有打 O、沒有 CSV 或該欄沒數值 | 空白 | `[監控] Memory 無監控資料` |

- 哪些格子有打 O **一律從範本讀**（重跑時當週檔的 O 已換成結果），有底色的格子不寫
- 備註和事件檢視器的內容並存；`[監控]` 開頭的行重跑時會先移除再寫，不會重複
- Kettle 資料夾不存在：顯示 `[WARN] Monitor CSV folder not found`，監控欄清空
- `--clear-monitor`（run_weekly.bat 在 Kettle 或 monitor_check 失敗時加上）：不讀監控資料，監控欄清空
- `columns` 的左邊是檢核表的欄位標題（要和第 `header_row` 列的文字一樣），右邊是 Kettle CSV 的欄位
- `trend` 為 null 時用 monitor_check 的預設（每天平均、5% / 1%）

### 只有時間沒有數據

Kettle SQL 的 CTE `d` 會留下「對不到類別（`item` 為 NULL）」或「沒打 O」的監控項目（例如 Traffic In/Out、
名稱不同的 Linux 掛載點）。標題列有 `where exists`，但 upper_spec 列和數據列沒有過濾，
所以只要主機有任何監控資料，就會照分鐘輸出「時間 + 空值」。

1. 先查這台實際的 `monitor_item`，補齊 SQL 裡 `case ... in (...)` 的清單：
   ```sql
   select hostname, monitor_item, count(*) n, min(monitor_time), max(monitor_time)
   from STM.monitor_system_history
   where hostname in ('P3-APX-AP02V', 'T1-GUI-AP01P') and monitor_time > now() - interval 7 day
   group by hostname, monitor_item order by hostname, monitor_item;
   ```
2. CTE `d` 只留「有打 O 且對得到」的項目，三段 union 都跟著正確，對不到的主機也不會產生檔案：
   ```sql
   d as (
     select tm, item, monitor_data, upper_spec from (
       select date_format(t.monitor_time, '%Y-%m-%d %H:%i') as tm,
              case ... end as item,          -- 原本的 case 不變
              t.monitor_data, t.upper_spec,
              p.c_cpu, p.c_mem, p.c_swap, p.c_d1, p.c_d2, p.c_d3
       from STM.monitor_system_history t join p on t.hostname = p.h
       where t.monitor_time > now() - interval 7 day
         and t.lm_time > now() - interval 7 day
     ) x
     where (item = 'cpu' and c_cpu) or (item = 'mem' and c_mem) or (item = 'swap' and c_swap)
        or (item = 'd1' and c_d1) or (item = 'd2' and c_d2) or (item = 'd3' and c_d3)
   )
   ```
   標題列的 `where exists (select 1 from d where d.item is not null)` 可改成 `where exists (select 1 from d)`。

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
- 相對路徑以 config.json 所在資料夾為基準
- 設定檔只有一份 `D:\SystemCheck\config.json`，三個程式共用：fill_check 依序找 exe 同資料夾、上一層
  （`fill_check\fill_check.exe` → `D:\SystemCheck\config.json`）；monitor_check 讀同一層；Kettle 用 JSON input 讀。
  執行時畫面與 log 第一行印出 `Config: <實際讀到的路徑>`

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
| `SystemCheck.zip` | 解壓到 `D:\` 會得到 `D:\SystemCheck\`：`config.json`（共用）、`fill_check\`（`fill_check.exe` + `_internal\`）、`monitor_check.exe`、`search_ignore.exe`、`run_weekly.bat`、`EventLogCheck.ps1`、空的 `EventLog\`、`logs\` |
| `search_ignore.exe` | 只要搜尋工具時單獨下載 |

使用前：
1. 修改 `D:\SystemCheck\config.json` 的 `paths`、`check_sheet.template`、`monitor`（CSV、檢核表、追蹤表、template、監控資料夾）
   - 從舊版升級：把 `fill_check\config.json` 移到 `D:\SystemCheck\`（取代新版的範例前先比對內容），
     舊的 `monitor_check.json` 內容併進 `monitor` 段後刪除
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
| monitor_check | `--onefile`，排除 pandas/numpy/openpyxl/PIL/lxml | `dist\monitor_check.exe` | 單一檔案，放在 `config.json` 同一層 |
| fill_check | `--onedir` | `dist\fill_check\` 資料夾 | 整個資料夾一起複製到 `D:\SystemCheck\fill_check\`，讀上一層的 `config.json` |

**為什麼不全部用 `--onefile`？** `--onefile` 每次啟動都要把整包程式解壓縮到暫存資料夾，
公司防毒會逐一掃描解出來的檔案；含 pandas 的程式可能要等好幾分鐘才出現第一行文字。
`--onedir` 不用解壓縮，啟動快很多。

啟動仍然很慢時：
- exe 放在網路磁碟（例如 P:）會更慢，請複製到本機（例如 D:\SystemCheck）執行
- 請 IT 把 exe 所在資料夾加入防毒的排除清單

`run_weekly.bat` 會先找 `D:\SystemCheck\fill_check\fill_check.exe`（把 `dist\fill_check\` 整個資料夾複製到 `D:\SystemCheck\fill_check\`），找不到才用 Python 執行 `fill_check.py`。
