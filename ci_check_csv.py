# -*- coding: utf-8 -*-
"""CI 用：檢查 EventLogCheck*.ps1 產生的 CSV 能被 fill_check 讀取"""
import glob
import os
import sys
import pandas as pd

ev_dir, computer, log_dir = sys.argv[1:4]
files = {os.path.basename(p).split("_", 1)[1][:-4].upper(): p for p in glob.glob(os.path.join(ev_dir, "*.csv"))}
print("CSV:", sorted(files))
assert computer.upper() in files, "local version did not produce CSV"
assert "LOCALHOST" in files, "server-list version did not produce CSV for localhost"
assert "NO-SUCH-HOST-XYZ" not in files, "unreachable host should not produce CSV"

for name, path in files.items():
    df = pd.read_csv(path, dtype=str, encoding="utf-8-sig")
    assert list(df.columns) == ["LevelDisplayName", "Id", "LogName", "Message", "TimeCreated", "Count"], df.columns
    levels = set(df["LevelDisplayName"].dropna())
    assert levels <= {"重大", "錯誤", "警告"}, f"{name}: unexpected levels {levels}"
    assert df["LevelDisplayName"].notna().all(), f"{name}: empty level"
    print(f"{name}: {len(df)} rows, levels={sorted(levels)}")
    print(df[["LevelDisplayName", "Id", "LogName", "TimeCreated", "Count"]].head(5).to_string())

log = open(glob.glob(os.path.join(log_dir, "*.log"))[0], encoding="utf-8-sig").read()
print(log)
assert "撈取失敗: no-such-host-xyz" in log, "unreachable host not reported in log"
print("OK")
