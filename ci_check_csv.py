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

# server-list run used an exclude file with 11, 46
ids = set(pd.read_csv(files["LOCALHOST"], dtype=str, encoding="utf-8-sig")["Id"])
assert not ids & {"11", "46"}, f"excluded IDs present in localhost CSV: {ids & {'11', '46'}}"
local_ids = set(pd.read_csv(files[computer.upper()], dtype=str, encoding="utf-8-sig")["Id"])
print("excluded IDs present in local (unfiltered) CSV:", sorted(local_ids & {"11", "46"}))

log = open(glob.glob(os.path.join(log_dir, "*.log"))[0], encoding="utf-8-sig").read()
print(log)
assert "Excluded event IDs: 2" in log, "exclude file not loaded"
assert "No exclude file" in log, "local run should report no exclude file"
assert "FAILED: no-such-host-xyz" in log, "unreachable host not reported in log"
assert "cannot list event logs: \n" not in log and "cannot list event logs: \r" not in log, "empty failure reason"
print("OK")
