# -*- coding: utf-8 -*-
"""
週檢小幫手：把 CSV 檢核結果貼到「系統檢核問題追蹤.xlsx」

流程:
1. 讀取指定資料夾內所有 CSV，檔名 = [日期_]伺服器名（例: 20261015_P3-WFA-AP01V.csv）
2. 每一筆 CSV 資料貼到追蹤表的工作表最後面：
   - 第一欄：伺服器名
   - 第二欄：週別，依檔名日期算 ISO 週次，格式 W+年後兩碼+週次（20261015 → W2642）
             檔名沒有日期時用今天日期
   - 第三欄起：CSV 內容（依標題對應欄位，表中沒有的標題會自動新增在最右邊）
3. 同一台伺服器、同一週別的舊資料會先刪除再貼上，重跑不會重複

需求套件: pip install pandas openpyxl
用法:
    python fill_check.py
    python fill_check.py <csv資料夾> <追蹤表xlsx>

設定檔: 與本程式同資料夾的 config.json
    track.xlsx         要貼上的追蹤表
    track.sheet        工作表名稱，不存在會自動建立；null = 第一個工作表
    track.host_header  第一欄標題（伺服器名）
    track.week_header  第二欄標題（週別）
    track.output       另存路徑，null = 直接覆寫追蹤表
    csv.dir            要讀取的 CSV 資料夾
    設定檔中的相對路徑以 config.json 所在資料夾為基準
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
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")


def _path(p):
    return os.path.join(BASE_DIR, p) if p else p


with open(CONFIG_PATH, encoding="utf-8-sig") as f:
    _cfg = json.load(f)

TRACK_XLSX  = _path(_cfg["track"]["xlsx"])
TRACK_SHEET = _cfg["track"]["sheet"]
HOST_HEADER = _cfg["track"]["host_header"]
WEEK_HEADER = _cfg["track"]["week_header"]
OUTPUT_XLSX = _path(_cfg["track"]["output"])

CSV_DIR = _path(_cfg["csv"]["dir"])
# ======================


def read_csv_any(path):
    """依序嘗試常見編碼（含 Big5/CP950）"""
    for enc in ("utf-8-sig", "cp950", "big5", "utf-8"):
        try:
            return pd.read_csv(path, dtype=str, encoding=enc)
        except UnicodeDecodeError:
            continue
    raise ValueError(f"無法判斷編碼: {path}")


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


def get_sheet(wb):
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


def main():
    csv_dir = sys.argv[1] if len(sys.argv) > 1 else CSV_DIR
    track_xlsx = sys.argv[2] if len(sys.argv) > 2 else TRACK_XLSX
    output = OUTPUT_XLSX or track_xlsx

    csv_files = sorted(glob.glob(os.path.join(csv_dir, "*.csv")))
    if not csv_files:
        print(f"資料夾內沒有 CSV: {csv_dir}")
        return

    wb = openpyxl.load_workbook(track_xlsx)
    ws = get_sheet(wb)

    for path in csv_files:
        host, d = parse_filename(path)
        week = week_label(d)
        try:
            df = read_csv_any(path)
        except Exception as e:
            print(f"[錯誤] {host}: {e}")
            continue
        df.columns = df.columns.str.strip()

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
        print(f"{host} {week}: 貼上 {len(df)} 筆" + (f"（取代舊資料 {removed} 筆）" if removed else ""))

    try:
        wb.save(output)
    except PermissionError:
        print(f"無法寫入 {output}，請先關閉 Excel 中開啟的檔案後再執行")
        sys.exit(1)
    print(f"已儲存: {output}")


if __name__ == "__main__":
    main()
