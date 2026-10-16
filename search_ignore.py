# -*- coding: utf-8 -*-
"""
可忽略事件ID 搜尋小工具（CLI）

執行後輸入 ID 查詢，一次可輸入多筆，用逗號隔開（, ，、 都可以），
逐筆顯示是否在可忽略清單，在清單內的會一併列出該列的其他欄位（例如說明）。

指令:
    r / reload        重新讀取 Excel（清單有更新時用）
    q / quit / exit   結束程式（也可以直接關閉視窗）

設定沿用 config.json 的 ignore_list（xlsx / sheet / id_column），
包成執行檔時 config.json 與 Excel 放在 exe 同一個資料夾。

打包: pip install pyinstaller
      pyinstaller --onefile search_ignore.py
"""
import re
import pandas as pd
from fill_check import norm, IGNORE_XLSX, IGNORE_SHEET, IGNORE_COLUMN

QUIT_CMDS = {"q", "quit", "exit", "離開"}
RELOAD_CMDS = {"r", "reload"}


def load_ignore_table():
    """回傳 ID → 該列其他欄位的說明文字（一格多個 ID 時每個 ID 都對應同一列）"""
    df = pd.read_excel(IGNORE_XLSX, sheet_name=IGNORE_SHEET, dtype=str)
    df.columns = df.columns.str.strip()
    if IGNORE_COLUMN not in df.columns:
        raise KeyError(f"Excel 找不到欄位「{IGNORE_COLUMN}」，現有欄位: {list(df.columns)}")
    others = [c for c in df.columns if c != IGNORE_COLUMN and not c.startswith("Unnamed")]
    table = {}
    for idx, row in df.iterrows():
        v = row[IGNORE_COLUMN]
        if pd.isna(v):
            continue
        info = "  ".join(f"{c}: {str(row[c]).strip()}" for c in others if pd.notna(row[c]))
        info = f"第 {idx + 2} 列" + (f"  {info}" if info else "")
        for part in str(v).replace("\n", ",").replace("、", ",").split(","):
            n = norm(part)
            if n:
                table.setdefault(n, []).append(info)
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


def load():
    table = load_ignore_table()
    print(f"已載入可忽略事件ID 共 {len(table)} 筆（{IGNORE_XLSX} / {IGNORE_SHEET}）")
    return table


def main():
    print("可忽略事件ID 搜尋")
    print("輸入 ID 查詢，多筆用逗號隔開；r 重新讀取清單；q 結束（或直接關閉視窗）")
    try:
        table = load()
    except Exception as e:
        print(f"[錯誤] 讀取清單失敗: {e}")
        input("按 Enter 結束...")
        return

    while True:
        try:
            text = input("\nID> ").strip()
        except (EOFError, KeyboardInterrupt):
            print()
            break
        cmd = text.lower()
        if cmd in QUIT_CMDS:
            break
        if cmd in RELOAD_CMDS:
            try:
                table = load()
            except Exception as e:
                print(f"[錯誤] 讀取清單失敗，沿用舊清單: {e}")
            continue
        search(table, text)


if __name__ == "__main__":
    main()
