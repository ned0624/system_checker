# -*- coding: utf-8 -*-
"""CI: SETTINGS block of EventLogCheck_Local.ps1 must work with non-ASCII paths in any file encoding.

make  <tmp>             write copies of the script whose SETTINGS point to non-ASCII folders
check <tmp> <computer>  verify each copy wrote its CSV and log into its own folder
"""
import glob
import os
import re
import sys
from datetime import date

import pandas as pd

# (variant, file encoding, folder name). The runner's ANSI code page is cp1252, so the ANSI
# variant uses a cp1252 character; on zh-TW Windows the same code path decodes Big5.
VARIANTS = [
    ("utf8_nobom", "utf-8", "中文資料夾_utf8"),
    ("utf8_bom", "utf-8-sig", "中文資料夾_bom"),
    ("ansi", "cp1252", "Café_ansi"),
]

mode, tmp = sys.argv[1], sys.argv[2]
if mode == "make":
    src = open("EventLogCheck_Local.ps1", encoding="ascii").read()
    for variant, enc, folder in VARIANTS:
        base = os.path.join(tmp, folder)
        block = (f"$Settings = @'\nOutDir      = {base}\\EventLog\nLogDir      = {base}\\logs\n"
                 f"ExcludeFile = {tmp}\\no-such-exclude.txt\nDays        = 7\n'@")
        out = re.sub(r"(?ms)^\$Settings = @'\n.*?\n'@", lambda _: block, src.replace("\r\n", "\n"), count=1)
        assert out != src
        os.makedirs(os.path.join(tmp, variant), exist_ok=True)
        with open(os.path.join(tmp, variant, "EventLogCheck_Local.ps1"), "w", encoding=enc, newline="\r\n") as f:
            f.write(out)
        print("made", variant, enc, base)
else:
    computer = sys.argv[3]
    today = date.today().strftime("%Y%m%d")
    for variant, enc, folder in VARIANTS:
        base = os.path.join(tmp, folder)
        csv = os.path.join(base, "EventLog", f"{today}_{computer}.csv")
        csvs = glob.glob(os.path.join(base, "EventLog", "*.csv"))
        assert csvs, f"{variant}: no CSV in {base}\\EventLog"
        df = pd.read_csv(csvs[0], dtype=str, encoding="utf-8-sig")
        assert set(df["LevelDisplayName"].dropna()) <= {"重大", "錯誤", "警告"}, df["LevelDisplayName"].unique()
        assert glob.glob(os.path.join(base, "logs", "*.log")), f"{variant}: no log"
        print(f"{variant}: OK {os.path.basename(csvs[0])} ({len(df)} rows) in {base}")
    print("OK")
