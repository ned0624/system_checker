# -*- coding: utf-8 -*-
"""CI: CSVs written for an explicit period - file date = end date, all TimeCreated inside the period.
Usage: python ci_period_check.py <csv folder> "<start yyyy-MM-dd HH:mm>" "<end yyyy-MM-dd HH:mm>"
"""
import glob
import os
import sys
from datetime import datetime

import pandas as pd

folder, start, end = sys.argv[1], *(datetime.strptime(s, "%Y-%m-%d %H:%M") for s in sys.argv[2:4])
files = glob.glob(os.path.join(folder, "*.csv"))
print("CSV:", [os.path.basename(f) for f in files])
for f in files:
    assert os.path.basename(f).startswith(end.strftime("%Y%m%d") + "_"), f"file date must be the end date: {f}"
    df = pd.read_csv(f, dtype=str, encoding="utf-8-sig")
    times = pd.to_datetime(df["TimeCreated"])
    print(f"{os.path.basename(f)}: {len(df)} rows, {times.min()} .. {times.max()}")
    assert (times >= start).all() and (times <= end).all(), "TimeCreated outside the period"
print("PERIOD OK")
