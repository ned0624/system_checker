# -*- coding: utf-8 -*-
"""
週檢小幫手：依 CSV 檢核結果填寫系統檢核表，並把 CSV 內容貼到問題追蹤表

流程:
1. 讀取「系統檢核問題追蹤.xlsx」的「可忽略事件ID」清單
2. 讀取指定資料夾內所有 CSV，檔名 = [日期_]伺服器名（例: 20261015_P3-WFA-AP01V.csv）
3. 檢核表（BOS_SystemCheck.xlsx）：到第一欄找到同名伺服器那一列
   - CSV 裡的 ID 全部都在可忽略清單 → 「警告」「錯誤」兩欄都寫 O
   - 有 ID 不在清單 → 依 CSV 的等級欄位（警告/錯誤）在對應欄打 X，
     另一欄若沒有問題則寫 O；不在清單的事件說明寫進「備註」
4. 問題追蹤表（系統檢核問題追蹤.xlsx）：CSV 全部內容（含可忽略的事件）貼到最後面
   - 第一欄：伺服器名
   - 第二欄：週別，依檔名日期算 ISO 週次，格式 W+年後兩碼+週次（20261015 → W2642）
             檔名沒有日期時用今天日期
   - 第三欄起：CSV 內容（依標題對應欄位，表中沒有的標題會自動新增在最右邊）
   - 同一台伺服器、同一週別的舊資料會先刪除再貼上，重跑不會重複

需求套件: pip install pandas openpyxl
用法:
    python fill_check.py
    python fill_check.py <csv資料夾> <檢核表xlsx> <追蹤表xlsx>   （有給的參數會取代設定檔的路徑）

設定檔: 與本程式同資料夾的 config.json
    paths.csv_dir                             要讀取的 CSV 資料夾
    paths.check_xlsx                          要填寫 O/X 的檢核表（BOS_SystemCheck.xlsx）
    paths.track_xlsx                          問題追蹤表（系統檢核問題追蹤.xlsx），可忽略清單也從這裡讀
    ignore_list.sheet / id_column             可忽略事件ID 的工作表與欄位
    csv.id_column / level_column / desc_column  CSV 的 ID、警告/錯誤分類、說明欄位
    check_sheet.sheet                         工作表名稱，null = 第一個工作表
    check_sheet.header_row                    標題列在第幾列
    check_sheet.warn/error/note_header        警告、錯誤、備註欄的標題
    check_sheet.output                        另存路徑，null = 直接覆寫檢核表
    track.sheet                               工作表名稱，不存在會自動建立；null = 第一個工作表
    track.host_header / week_header           第一欄、第二欄的標題
    track.output                              另存路徑，null = 直接覆寫追蹤表
    路徑可用絕對路徑（例: D:/週檢/BOS_SystemCheck.xlsx 或 \\\\server\\share\\週檢\\csv），
    JSON 裡的反斜線要寫兩次（\\），或改用正斜線 /
    相對路徑以 config.json 所在資料夾為基準
"""
import sys
import glob
import json
import os
import re
from datetime import date, datetime
import pandas as pd
import openpyxl

# ===== 讀取設定檔 =====
# 包成執行檔（PyInstaller）時，config.json 放在 exe 旁邊
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")


def _path(p):
    return os.path.join(BASE_DIR, p) if p else p


with open(CONFIG_PATH, encoding="utf-8-sig") as f:
    _cfg = json.load(f)

CSV_DIR    = _path(_cfg["paths"]["csv_dir"])
CHECK_XLSX = _path(_cfg["paths"]["check_xlsx"])
TRACK_XLSX = _path(_cfg["paths"]["track_xlsx"])

IGNORE_XLSX   = TRACK_XLSX
IGNORE_SHEET  = _cfg["ignore_list"]["sheet"]
IGNORE_COLUMN = _cfg["ignore_list"]["id_column"]

CSV_ID_COLUMN    = _cfg["csv"]["id_column"]
CSV_LEVEL_COLUMN = _cfg["csv"]["level_column"]
CSV_DESC_COLUMN  = _cfg["csv"]["desc_column"]

CHECK_SHEET  = _cfg["check_sheet"]["sheet"]
HEADER_ROW   = _cfg["check_sheet"]["header_row"]
WARN_HEADER  = _cfg["check_sheet"]["warn_header"]
ERROR_HEADER = _cfg["check_sheet"]["error_header"]
NOTE_HEADER  = _cfg["check_sheet"]["note_header"]
OUTPUT_XLSX  = _path(_cfg["check_sheet"]["output"])

TRACK_SHEET  = _cfg["track"]["sheet"]
HOST_HEADER  = _cfg["track"]["host_header"]
WEEK_HEADER  = _cfg["track"]["week_header"]
TRACK_OUTPUT = _path(_cfg["track"]["output"])
# ======================


def norm(v):
    """統一 ID 格式：去空白、去掉 Excel 讀成 float 的 .0"""
    if pd.isna(v):
        return None
    s = str(v).strip()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s or None


def read_csv_any(path):
    """依序嘗試常見編碼（含 Big5/CP950）"""
    for enc in ("utf-8-sig", "cp950", "big5", "utf-8"):
        try:
            df = pd.read_csv(path, dtype=str, encoding=enc)
            df.columns = df.columns.str.strip()
            return df
        except UnicodeDecodeError:
            continue
    raise ValueError(f"無法判斷編碼: {path}")


def load_ignore_ids(xlsx_path):
    df = pd.read_excel(xlsx_path, sheet_name=IGNORE_SHEET, dtype=str)
    df.columns = df.columns.str.strip()
    if IGNORE_COLUMN not in df.columns:
        raise KeyError(f"Excel 找不到欄位「{IGNORE_COLUMN}」，現有欄位: {list(df.columns)}")
    ids = set()
    for v in df[IGNORE_COLUMN]:
        # 若一格內有多個 ID（以逗號/換行分隔）也一併拆開
        if pd.isna(v):
            continue
        for part in str(v).replace("\n", ",").replace("、", ",").split(","):
            n = norm(part)
            if n:
                ids.add(n)
    return ids


# ---------- 檢核表 O/X ----------

def find_columns(ws):
    """依標題文字找出 警告/錯誤/備註 的欄號"""
    headers = {str(c.value).strip(): c.column for c in ws[HEADER_ROW] if c.value is not None}
    cols = {}
    for h in (WARN_HEADER, ERROR_HEADER, NOTE_HEADER):
        if h not in headers:
            raise KeyError(f"檢核表找不到標題「{h}」，現有標題: {list(headers)}")
        cols[h] = headers[h]
    return cols


def find_host_rows(ws):
    """第一欄的伺服器名稱 → 列號（不分大小寫）"""
    rows = {}
    for r in range(HEADER_ROW + 1, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if v is not None and str(v).strip():
            rows[str(v).strip().upper()] = r
    return rows


def check_csv(df, ignore_ids):
    """回傳 (有警告, 有錯誤, 備註文字)；只計算不在可忽略清單的事件"""
    for col in (CSV_ID_COLUMN, CSV_LEVEL_COLUMN, CSV_DESC_COLUMN):
        if col not in df.columns:
            raise KeyError(f"找不到欄位「{col}」，現有欄位: {list(df.columns)}")

    has_warn = has_error = False
    notes = []
    seen = set()
    for _, row in df.iterrows():
        eid = norm(row[CSV_ID_COLUMN])
        if not eid or eid in ignore_ids:
            continue
        level = str(row[CSV_LEVEL_COLUMN]).strip() if pd.notna(row[CSV_LEVEL_COLUMN]) else ""
        desc = str(row[CSV_DESC_COLUMN]).strip() if pd.notna(row[CSV_DESC_COLUMN]) else ""
        if "錯誤" in level or level.lower() == "error":
            has_error = True
        elif "警告" in level or level.lower() == "warning":
            has_warn = True
        else:
            print(f"  [注意] ID {eid} 等級「{level}」無法辨識，未計入警告/錯誤")
        key = (level, eid, desc)
        if key not in seen:  # 同一事件重複出現只記一次
            seen.add(key)
            notes.append(f"[{level}] {eid} {desc}".strip())
    return has_warn, has_error, "\n".join(notes)


def fill_check_row(ws, cols, r, df, ignore_ids, host):
    has_warn, has_error, note = check_csv(df, ignore_ids)
    ws.cell(row=r, column=cols[WARN_HEADER]).value = "X" if has_warn else "O"
    ws.cell(row=r, column=cols[ERROR_HEADER]).value = "X" if has_error else "O"
    note_cell = ws.cell(row=r, column=cols[NOTE_HEADER])
    note_cell.value = note or None
    if note:
        note_cell.alignment = openpyxl.styles.Alignment(wrap_text=True, vertical="top")
    print(f"  檢核表: 警告={'X' if has_warn else 'O'} 錯誤={'X' if has_error else 'O'}"
          + (f"\n    {note.replace(chr(10), chr(10) + '    ')}" if note else ""))


# ---------- 問題追蹤表 ----------

def parse_filename(path):
    """檔名拆出 (伺服器名, 日期)，例: 20261015_P3-WFA-AP01V.csv → (P3-WFA-AP01V, 2026-10-15)"""
    name = os.path.splitext(os.path.basename(path))[0].strip()
    m = re.match(r"^(\d{8})_(.+)$", name)
    if m:
        try:
            return m.group(2).strip(), datetime.strptime(m.group(1), "%Y%m%d").date()
        except ValueError:
            pass
    return re.sub(r"^\d+_", "", name), date.today()


def week_label(d):
    """ISO 週次，例: 2026-10-15 → W2642"""
    year, week, _ = d.isocalendar()
    return f"W{year % 100:02d}{week:02d}"


def get_track_sheet(wb):
    if not TRACK_SHEET:
        return wb.worksheets[0]
    if TRACK_SHEET in wb.sheetnames:
        return wb[TRACK_SHEET]
    return wb.create_sheet(TRACK_SHEET)


def ensure_headers(ws, csv_columns):
    """確保第一列有 伺服器/週別/CSV 欄位標題，回傳 標題 → 欄號"""
    if ws.max_row == 1 and ws.cell(row=1, column=1).value is None:
        ws.cell(row=1, column=1).value = HOST_HEADER
        ws.cell(row=1, column=2).value = WEEK_HEADER
    headers = {str(c.value).strip(): c.column for c in ws[1] if c.value is not None}
    next_col = max(headers.values(), default=0) + 1
    for col in csv_columns:
        if col not in headers:
            ws.cell(row=1, column=next_col).value = col
            headers[col] = next_col
            next_col += 1
    return headers


def delete_existing(ws, host, week):
    """刪除同一伺服器、同一週別的舊資料（由下往上刪才不會錯位）"""
    removed = 0
    for r in range(ws.max_row, 1, -1):
        h = ws.cell(row=r, column=1).value
        w = ws.cell(row=r, column=2).value
        if h is not None and str(h).strip().upper() == host.upper() and str(w).strip() == week:
            ws.delete_rows(r)
            removed += 1
    return removed


def last_data_row(ws):
    """最後一列有資料的列號（避免 max_row 被空白格式列撐大）"""
    for r in range(ws.max_row, 0, -1):
        if any(c.value is not None for c in ws[r]):
            return r
    return 0


def paste_csv(ws, df, host, week):
    headers = ensure_headers(ws, list(df.columns))
    removed = delete_existing(ws, host, week)
    r = last_data_row(ws) + 1
    for _, row in df.iterrows():
        ws.cell(row=r, column=1).value = host
        ws.cell(row=r, column=2).value = week
        for col in df.columns:
            v = row[col]
            ws.cell(row=r, column=headers[col]).value = None if pd.isna(v) else v
        r += 1
    print(f"  追蹤表: {week} 貼上 {len(df)} 筆" + (f"（取代舊資料 {removed} 筆）" if removed else ""))


# ---------- 主程式 ----------

def save(wb, path):
    try:
        wb.save(path)
    except PermissionError:
        print(f"無法寫入 {path}，請先關閉 Excel 中開啟的檔案後再執行")
        sys.exit(1)
    print(f"已儲存: {path}")


def main():
    csv_dir = sys.argv[1] if len(sys.argv) > 1 else CSV_DIR
    check_xlsx = sys.argv[2] if len(sys.argv) > 2 else CHECK_XLSX
    track_xlsx = sys.argv[3] if len(sys.argv) > 3 else TRACK_XLSX
    check_output = OUTPUT_XLSX or check_xlsx
    track_output = TRACK_OUTPUT or track_xlsx

    for label, p in (("CSV 資料夾", csv_dir), ("檢核表", check_xlsx), ("追蹤表", track_xlsx)):
        if not os.path.exists(p):
            print(f"[錯誤] 找不到{label}: {p}（請確認 config.json 的 paths）")
            sys.exit(1)

    ignore_ids = load_ignore_ids(track_xlsx)
    print(f"可忽略事件ID 共 {len(ignore_ids)} 筆")

    csv_files = sorted(glob.glob(os.path.join(csv_dir, "*.csv")))
    if not csv_files:
        print(f"資料夾內沒有 CSV: {csv_dir}")
        return

    check_wb = openpyxl.load_workbook(check_xlsx)
    check_ws = check_wb[CHECK_SHEET] if CHECK_SHEET else check_wb.worksheets[0]
    cols = find_columns(check_ws)
    host_rows = find_host_rows(check_ws)

    track_wb = openpyxl.load_workbook(track_xlsx)
    track_ws = get_track_sheet(track_wb)

    for path in csv_files:
        host, d = parse_filename(path)
        print(host)
        try:
            df = read_csv_any(path)
        except Exception as e:
            print(f"  [錯誤] {e}")
            continue

        r = host_rows.get(host.upper())
        if r is None:
            print("  [略過] 檢核表第一欄找不到此伺服器")
        else:
            try:
                fill_check_row(check_ws, cols, r, df, ignore_ids, host)
            except Exception as e:
                print(f"  [錯誤] 檢核表: {e}")

        paste_csv(track_ws, df, host, week_label(d))

    save(check_wb, check_output)
    save(track_wb, track_output)


if __name__ == "__main__":
    main()
