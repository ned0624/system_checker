# -*- coding: utf-8 -*-
"""CI 用：產生假的 Kettle 匯出 CSV，驗證 monitor_check 的超標與持續上升判定

用法: python ci_monitor_test.py <tmp資料夾>
"""
import csv
import glob
import json
import math
import os
import shutil
import subprocess
import sys
from datetime import date, datetime, timedelta

tmp = os.path.abspath(sys.argv[1])
shutil.rmtree(tmp, ignore_errors=True)
app = os.path.join(tmp, "app")
os.makedirs(app)
shutil.copy("monitor_check.py", app)
y, w, _ = date.today().isocalendar()
week = f"W{y % 100:02d}{w:02d}"
export = os.path.join(tmp, f"{y}-W{w:02d}")   # Kettle 週資料夾 %x-W%v
os.makedirs(export)
# 共用的 config.json 放在 app 的上一層（D:\\SystemCheck\\config.json 的位置），相對路徑以它為基準
json.dump({"log_dir": "logs", "monitor": {"csv_dir": "{yyyy}-W{ww}", "report_dir": os.path.join("{week}", "report")}},
          open(os.path.join(tmp, "config.json"), "w"))

START = datetime(2026, 10, 28, 0, 0)
N = 7 * 24 * 60   # Kettle 撈 7 天、每分鐘一筆
DAY = 24 * 60
HEADER = "time_min,cpu,memory,disk1,disk2,disk3"


def write(name, spec, rows, header=True, enc="utf-8-sig", columns=HEADER, start=START):
    lines = ([columns] if header else []) + ["upper_spec," + spec]
    lines += [f"{(start + timedelta(minutes=i)):%Y-%m-%d %H:%M}," + ",".join("" if v is None else f"{v:.2f}" for v in r)
              for i, r in enumerate(rows)]
    open(os.path.join(export, name), "w", encoding=enc, newline="").write("\r\n".join(lines) + "\r\n")


# HOST-A: cpu 日夜起伏 + 三段超標（20 / 5 / 1 分鐘）；memory 每天穩定上升（也有日夜起伏）；
#         disk1 持平；disk2 前 4 天上升後清掉；disk3 空白
rows = []
for i in range(N):
    cpu = 20 + 15 * math.sin(2 * math.pi * i / DAY) + 3 * math.sin(i / 37)
    if 100 <= i < 120:
        cpu = 70
    elif 200 <= i < 205:
        cpu = 65
    elif i == 300:
        cpu = 61
    mem = 15 + 0.0005 * i + 0.5 * math.sin(2 * math.pi * i / DAY)
    disk2 = 370 + 8 * i / DAY if i < 4 * DAY else 372
    rows.append((cpu, mem, 30.0, disk2, None))
write("host-a.csv", "60.00,25.60,40.00,420.00,625.00", rows)

# HOST-B: 無表頭（time_min,cpu,memory,swap,...）、MS950；memory 上升但第 4 天整天回落 2 → 不算持續上升；cpu 沒有 spec
rows = [(10 + 0.0001 * i, 15 + 0.0005 * i - (2 if 3 * DAY <= i < 4 * DAY else 0), 5.0, 20.0, 100.0, 50.0)
        for i in range(N)]
write("host-b.csv", ",25.60,10.00,40.00,420.00,625.00", rows, header=False, enc="cp950")

# HOST-C: 壞檔 → WARN 但不中斷
open(os.path.join(export, "host-c.csv"), "w").write("garbage\n1,2,3\n")

# HOST-E: 只有時間沒有數值（Kettle SQL 的 monitor_item 對不到）→ WARN
write("host-e.csv", ",", [(None, None)] * 120, columns="time_min,cpu,disk2")

# HOST-F: 只打 O 部分欄位（cpu、disk2）→ 只分析這兩欄
write("host-f.csv", "60.00,420.00", [(10.0, 300.0)] * 120, columns="time_min,cpu,disk2")

# HOST-G: 從 08:37 開始的 7 天（Kettle 依執行時間往回撈），cpu 只有日夜起伏 → 頭尾半天不能造成誤判
rows = [(30 + 25 * math.sin(2 * math.pi * (i + 517) / DAY - 1.2),) for i in range(N)]
write("host-g.csv", "60.00", rows, columns="time_min,cpu", start=START + timedelta(minutes=517))

# HOST-D: 只有 3 小時 → 不判定上升，超標照算
rows = [(10.0, 30.0 if i == 10 else 20.0, 30.0, 300.0, 500.0) for i in range(180)]
write("host-d.csv", "60.00,25.60,40.00,420.00,625.00", rows)


def run(*args):
    r = subprocess.run([sys.executable, os.path.join(app, "monitor_check.py"), *args],
                       capture_output=True, text=True, encoding="utf-8", errors="replace",
                       env=dict(os.environ, PYTHONIOENCODING="utf-8"))
    print(r.stdout, r.stderr)
    return r


r = run()
assert r.returncode == 0, r.returncode
reports = glob.glob(os.path.join(tmp, week, "report", "*_monitor_report.csv"))
assert len(reports) == 1, f"report not in week folder: {reports}"
rep = {(x["host"], x["metric"]): x for x in csv.DictReader(open(reports[0], encoding="utf-8-sig"))}
header = open(reports[0], encoding="utf-8-sig").readline()
assert "first_hour_avg" not in header and "last_hour_avg" not in header and "max_drop" not in header, header

a = rep["HOST-A", "cpu"]
assert a["over_count"] == "26" and a["longest_over_min"] == "20", a
assert a["over_first"] == "2026-10-28 01:40" and a["over_last"] == "2026-10-28 05:00", a
assert a["max"] == "70.00" and a["rising"] == "N" and a["result"] == "OVER SPEC", a
m = rep["HOST-A", "memory"]
assert m["rising"] == "Y" and m["result"] == "RISING" and float(m["hours_to_spec"]) > 0, m
assert rep["HOST-A", "disk1"]["result"] == "OK"
assert rep["HOST-A", "disk2"]["rising"] == "N", rep["HOST-A", "disk2"]
assert rep["HOST-A", "disk3"]["result"] == "NO DATA"
assert rep["HOST-B", "swap"]["spec"] == "10.00" and rep["HOST-B", "disk3"]["result"] == "OK", rep["HOST-B", "disk3"]
assert rep["HOST-E", "cpu"]["result"] == "NO DATA" and rep["HOST-E", "disk2"]["result"] == "NO DATA"
assert sorted(m for h, m in rep if h == "HOST-F") == ["cpu", "disk2"]
assert rep["HOST-F", "disk2"]["result"] == "OK" and rep["HOST-F", "disk2"]["spec"] == "420.00"
assert rep["HOST-B", "memory"]["rising"] == "N", rep["HOST-B", "memory"]
assert rep["HOST-G", "cpu"]["rising"] == "N" and rep["HOST-G", "cpu"]["result"] == "OK", rep["HOST-G", "cpu"]
assert rep["HOST-B", "cpu"]["spec"] == "" and rep["HOST-B", "cpu"]["over_count"] == "0", rep["HOST-B", "cpu"]
assert not any(h == "HOST-C" for h, _ in rep), "broken file should not be in report"
d = rep["HOST-D", "memory"]
assert d["over_count"] == "1" and d["rising"] == "-" and d["result"] == "OVER SPEC", d
log = open(glob.glob(os.path.join(tmp, "logs", "*_monitor_check.log"))[0], encoding="utf-8-sig").read()
assert "[WARN] HOST-E: 120 rows, no values in any column" in log, log
assert "[WARN] HOST-C" in log and "[RISING] HOST-A memory" in log and "[OVER SPEC] HOST-A cpu" in log, log
assert not os.path.exists(os.path.join(tmp, "{week}")), "literal {week} folder created"

r = run(os.path.join(tmp, "no_such_dir"))
assert r.returncode == 1 and "CSV folder not found" in r.stdout, r.returncode

# 舊版 monitor_check.json（沒有 config.json）仍可用，並提示改用 config.json
os.remove(os.path.join(tmp, "config.json"))
json.dump({"csv_dir": os.path.join(tmp, "{yyyy}-W{ww}"), "report_dir": os.path.join(tmp, "legacy_report"),
           "log_dir": os.path.join(tmp, "logs")}, open(os.path.join(app, "monitor_check.json"), "w"))
r = run()
assert r.returncode == 0 and "monitor_check.json is deprecated" in r.stdout, r.returncode
assert glob.glob(os.path.join(tmp, "legacy_report", "*_monitor_report.csv")), "legacy config not used"
os.remove(os.path.join(app, "monitor_check.json"))
r = run()
assert r.returncode == 1 and "no config.json" in r.stdout, r.returncode
print("MONITOR TEST OK")
