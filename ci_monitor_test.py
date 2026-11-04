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
export = os.path.join(tmp, week, "export")
os.makedirs(export)
json.dump({"csv_dir": os.path.join(tmp, "{week}", "export"), "report_dir": os.path.join(tmp, "{week}", "report"),
           "log_dir": os.path.join(tmp, "logs")}, open(os.path.join(app, "monitor_check.json"), "w"))

START = datetime(2026, 11, 3, 0, 0)
N = 24 * 60
HEADER = "time_min,cpu,memory,disk_c,disk_d,disk_e"


def write(name, spec, rows, header=True, enc="utf-8-sig"):
    lines = ([HEADER] if header else []) + ["upper_spec," + spec]
    lines += [f"{(START + timedelta(minutes=i)):%Y-%m-%d %H:%M}," + ",".join("" if v is None else f"{v:.2f}" for v in r)
              for i, r in enumerate(rows)]
    open(os.path.join(export, name), "w", encoding=enc, newline="").write("\r\n".join(lines) + "\r\n")


# HOST-A: cpu 上下跳動 + 三段超標（20 / 5 / 1 分鐘）；memory 穩定上升；disk_c 持平；disk_d 先升後清掉；disk_e 空白
rows = []
for i in range(N):
    cpu = 20 + 15 * math.sin(i / 37) + 0.002 * i
    if 100 <= i < 120:
        cpu = 70
    elif 200 <= i < 205:
        cpu = 65
    elif i == 300:
        cpu = 61
    mem = 15 + 0.004 * i + 0.05 * math.sin(i * 1.7)
    disk_d = 370 + 30 * i / 720 if i < 720 else 372
    rows.append((cpu, mem, 30.0, disk_d, None))
write("host-a.csv", "60.00,25.60,40.00,420.00,625.00", rows)

# HOST-B: 無表頭、MS950；memory 上升但中途回落 2 → 不算持續上升；cpu 沒有 spec
rows = [(10 + 0.001 * i, 15 + 0.005 * i - (2 if 600 <= i < 700 else 0), 20.0, 100.0, 50.0) for i in range(N)]
write("host-b.csv", ",25.60,40.00,420.00,625.00", rows, header=False, enc="cp950")

# HOST-C: 壞檔 → WARN 但不中斷
open(os.path.join(export, "host-c.csv"), "w").write("garbage\n1,2,3\n")

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

a = rep["HOST-A", "cpu"]
assert a["over_count"] == "26" and a["longest_over_min"] == "20", a
assert a["over_first"] == "2026-11-03 01:40" and a["over_last"] == "2026-11-03 05:00", a
assert a["max"] == "70.00" and a["rising"] == "N" and a["result"] == "OVER SPEC", a
m = rep["HOST-A", "memory"]
assert m["rising"] == "Y" and m["result"] == "RISING" and float(m["hours_to_spec"]) > 0, m
assert rep["HOST-A", "disk_c"]["result"] == "OK"
assert rep["HOST-A", "disk_d"]["rising"] == "N", rep["HOST-A", "disk_d"]
assert rep["HOST-A", "disk_e"]["result"] == "NO DATA"
assert rep["HOST-B", "memory"]["rising"] == "N", rep["HOST-B", "memory"]
assert rep["HOST-B", "cpu"]["spec"] == "" and rep["HOST-B", "cpu"]["over_count"] == "0", rep["HOST-B", "cpu"]
assert not any(h == "HOST-C" for h, _ in rep), "broken file should not be in report"
d = rep["HOST-D", "memory"]
assert d["over_count"] == "1" and d["rising"] == "-" and d["result"] == "OVER SPEC", d
log = open(glob.glob(os.path.join(tmp, "logs", "*_monitor_check.log"))[0], encoding="utf-8-sig").read()
assert "[WARN] HOST-C" in log and "[RISING] HOST-A memory" in log and "[OVER SPEC] HOST-A cpu" in log, log
assert not os.path.exists(os.path.join(tmp, "{week}")), "literal {week} folder created"

r = run(os.path.join(tmp, "no_such_dir"))
assert r.returncode == 1 and "CSV folder not found" in r.stdout, r.returncode
print("MONITOR TEST OK")
