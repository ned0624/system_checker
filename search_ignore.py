# -*- coding: utf-8 -*-
"""
可忽略事件ID 搜尋小工具（CLI，獨立程式，不需要 config.json）

1. 執行後輸入可忽略清單的 Excel 路徑（可直接把檔案拖進視窗；直接按 Enter 沿用上次的檔案）
2. 輸入 ID 查詢，一次可輸入多筆，用逗號隔開（, ，、 都可以），
   逐筆顯示是否在可忽略清單，在清單內的會一併列出該列的其他欄位（例如說明）

指令:
    r / reload        重新讀取 Excel（清單有更新時用）
    f / file          換一個 Excel 檔
    q / quit / exit   結束程式（也可以直接關閉視窗）

清單位置: 工作表「可忽略事件ID」（找不到時列出工作表讓你選），
          ID 欄依標題「可忽略事件ID」或「ID」尋找，都沒有時用第一欄

需求套件: pip install pandas openpyxl
打包: pip install pyinstaller
      pyinstaller --onefile search_ignore.py
"""
import os
import re
import sys
import pandas as pd

DEFAULT_SHEET = "可忽略事件ID"
ID_HEADERS = ("可忽略事件ID", "ID", "Id", "事件識別碼")
QUIT_CMDS = {"q", "quit", "exit", "離開"}
RELOAD_CMDS = {"r", "reload"}
FILE_CMDS = {"f", "file"}

def _base_dir():
    """程式所在資料夾；包成 exe 時是 exe 所在資料夾"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


# 記住上次的 Excel 路徑（放在程式旁邊）
LAST_PATH_FILE = os.path.join(_base_dir(), "search_ignore_last.txt")


class Quit(Exception):
    pass


def ask(prompt):
    try:
        text = input(prompt).strip()
    except (EOFError, KeyboardInterrupt):
        print()
        raise Quit
    if text.lower() in QUIT_CMDS:
        raise Quit
    return text


def norm(v):
    """統一 ID 格式：去空白、去掉 Excel 讀成 float 的 .0"""
    if pd.isna(v):
        return None
    s = str(v).strip()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s or None


def read_last_path():
    try:
        with open(LAST_PATH_FILE, encoding="utf-8") as f:
            return f.read().strip()
    except OSError:
        return ""


def save_last_path(path):
    try:
        with open(LAST_PATH_FILE, "w", encoding="utf-8") as f:
            f.write(path)
    except OSError:
        pass


def ask_excel():
    """詢問 Excel 路徑，直到找到檔案為止"""
    last = read_last_path()
    while True:
        hint = f"（Enter 沿用 {last}）" if last else ""
        path = ask(f"\n可忽略清單 Excel 路徑{hint}> ").strip('"').strip("'")
        if not path and last:
            path = last
        if not path:
            continue
        if os.path.isfile(path):
            return path
        print(f"[錯誤] 找不到檔案: {path}")


def pick_sheet(path):
    sheets = pd.ExcelFile(path).sheet_names
    if DEFAULT_SHEET in sheets:
        return DEFAULT_SHEET
    print(f"找不到工作表「{DEFAULT_SHEET}」，請選擇：")
    for i, name in enumerate(sheets, 1):
        print(f"  {i}. {name}")
    while True:
        text = ask("工作表編號> ")
        if text.isdigit() and 1 <= int(text) <= len(sheets):
            return sheets[int(text) - 1]


def load_ignore_table(path, sheet):
    """回傳 ID → 該列其他欄位的說明文字（一格多個 ID 時每個 ID 都對應同一列）"""
    df = pd.read_excel(path, sheet_name=sheet, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    id_col = next((h for h in ID_HEADERS if h in df.columns), df.columns[0])
    others = [c for c in df.columns if c != id_col and not c.startswith("Unnamed")]
    table = {}
    for idx, row in df.iterrows():
        v = row[id_col]
        if pd.isna(v):
            continue
        info = "  ".join(f"{c}: {str(row[c]).strip()}" for c in others if pd.notna(row[c]))
        info = f"第 {idx + 2} 列" + (f"  {info}" if info else "")
        for part in str(v).replace("\n", ",").replace("、", ",").split(","):
            n = norm(part)
            if n:
                table.setdefault(n, []).append(info)
    print(f"已載入可忽略事件ID 共 {len(table)} 筆（{os.path.basename(path)} / {sheet} / 欄位: {id_col}）")
    return table


def split_ids(text):
    seen, ids = set(), []
    for part in re.split(r"[,，、;\s]+", text):
        n = norm(part)
        if n and n not in seen:
            seen.add(n)
            ids.append(n)
    return ids


def search(table, text):
    ids = split_ids(text)
    if not ids:
        return
    hit = 0
    for eid in ids:
        if eid in table:
            hit += 1
            print(f"  [可忽略] {eid}")
            for info in table[eid]:
                print(f"           {info}")
        else:
            print(f"  [不在清單] {eid}")
    print(f"共 {len(ids)} 筆：可忽略 {hit} 筆，不在清單 {len(ids) - hit} 筆")


def open_excel():
    """選檔並載入，失敗就重新選，回傳 (path, sheet, table)"""
    while True:
        path = ask_excel()
        try:
            sheet = pick_sheet(path)
            table = load_ignore_table(path, sheet)
        except Quit:
            raise
        except Exception as e:
            print(f"[錯誤] 讀取失敗（檔案若在 Excel 中開啟仍可讀；請確認是 xlsx 檔）: {e}")
            continue
        save_last_path(path)
        return path, sheet, table


def main():
    print("可忽略事件ID 搜尋")
    print("輸入 ID 查詢，多筆用逗號隔開；r 重新讀取；f 換檔案；q 結束（或直接關閉視窗）")
    try:
        path, sheet, table = open_excel()
        while True:
            text = ask("\nID> ")
            cmd = text.lower()
            if cmd in RELOAD_CMDS:
                try:
                    table = load_ignore_table(path, sheet)
                except Exception as e:
                    print(f"[錯誤] 讀取失敗，沿用舊清單: {e}")
            elif cmd in FILE_CMDS:
                path, sheet, table = open_excel()
            else:
                search(table, text)
    except Quit:
        pass


if __name__ == "__main__":
    main()
