# -*- coding: utf-8 -*-
"""CI（Windows）：用假的撈取腳本、Kitchen、monitor_check、fill_check 測 run_weekly.bat → run_weekly.ps1 的流程

情境：全部成功 / Kettle 失敗 / monitor_check 失敗 / 沒有 config.json / 沒有事件 / 撈 event log 失敗，
另外用真的 EventLogCheck.ps1 / EventLogCheck_Local.ps1 跑一次（config.json 的中文 + {week} 路徑）
用法: python ci_run_weekly_test.py <tmp資料夾>
"""
import json
import os
import shutil
import subprocess
import sys
from datetime import date

tmp = os.path.abspath(sys.argv[1])


def scenario(name, collector=0, kitchen=0, monitor=0, config=True):
    base = os.path.join(tmp, name)
    shutil.rmtree(base, ignore_errors=True)
    os.makedirs(os.path.join(base, "kettle"))
    log = os.path.join(base, "calls.txt")
    shutil.copy("run_weekly.bat", base)   # bat 用 %~dp0 找同資料夾的 run_weekly.ps1
    shutil.copy("run_weekly.ps1", base)
    if config:   # log_dir 相對路徑 + {week} → <base>\W2645\logs；路徑從 config.json 傳給撈取腳本
        json.dump({"log_dir": "{week}/logs",
                   "weekly": {"collector": "EventLogCheck.ps1", "kitchen": "Kitchen.bat",
                              "kettle_job": "kettle/monitor.kjb"},
                   "paths": {"csv_dir": "//fileserver/中文資料夾/{week}/BOS"},
                   "eventlog": {"server_list": "serverlist.txt", "exclude_file": "D:/SystemCheck/excluded_ids.txt",
                                "days": 7, "start_time": "", "end_time": ""}},
                  open(os.path.join(base, "config.json"), "w", encoding="utf-8"), ensure_ascii=False)
    # 假的撈取腳本：把收到的參數寫成 UTF-8 JSON（中文路徑不會因主控台編碼變成 ?）
    open(os.path.join(base, "EventLogCheck.ps1"), "w").write(
        "param([string]$ServerList, [string]$OutDir, [string]$LogDir, [string]$StartTime, [string]$EndTime,"
        " [string]$Days, [string]$ExcludeFile)\n"
        f'Add-Content "{log}" "collector"\n'
        f'[IO.File]::WriteAllText("{base}\\collector_args.json", ($PSBoundParameters | ConvertTo-Json), '
        "(New-Object Text.UTF8Encoding($false)))\n"
        f"exit {collector}\n")
    open(os.path.join(base, "Kitchen.bat"), "w", newline="\r\n").write(
        f'@echo off\necho kitchen %*>>"{log}"\nexit /b {kitchen}\n')
    open(os.path.join(base, "kettle", "monitor.kjb"), "w").write("<job/>")
    stub = 'import sys\nopen(r"{log}", "a").write("{name} " + " ".join(sys.argv[1:]) + "\\n")\nsys.exit({code})\n'
    open(os.path.join(base, "monitor_check.py"), "w").write(stub.format(log=log, name="monitor", code=monitor))
    open(os.path.join(base, "fill_check.py"), "w").write(stub.format(log=log, name="fill", code=0))
    r = subprocess.run(["cmd", "/c", os.path.join(base, "run_weekly.bat")], stdin=subprocess.DEVNULL,
                       capture_output=True, text=True, errors="replace")
    calls = open(log).read().split("\n") if os.path.exists(log) else []
    calls = [c.strip() for c in calls if c.strip()]
    print(f"--- {name}\n{r.stdout}\ncalls: {calls}")
    return r.stdout, calls


out, calls = scenario("all_ok")
assert calls[0] == "collector" and calls[1].startswith("kitchen /file:") and calls[2] == "monitor", calls
args = json.load(open(os.path.join(tmp, "all_ok", "collector_args.json"), encoding="utf-8"))
base_ok = os.path.join(tmp, "all_ok")
assert args == {"OutDir": r"\\fileserver\中文資料夾\{week}\BOS",           # {week} 由撈取腳本換成週別
                "LogDir": os.path.join(base_ok, "{week}", "logs"),
                "ServerList": os.path.join(base_ok, "serverlist.txt"),
                "ExcludeFile": r"D:\SystemCheck\excluded_ids.txt",
                "Days": "7"}, args                                            # 空的 start_time / end_time 不傳
y, w, _ = date.today().isocalendar()
log_dir = os.path.join(tmp, "all_ok", f"W{y % 100:02d}{w:02d}", "logs")
kettle_log = os.path.join(log_dir, f"{date.today():%Y%m%d}_kettle.log")
assert f"/logfile:{kettle_log}" in calls[1], (calls[1], kettle_log)   # Kettle log → config.json 的 log_dir
assert os.path.isdir(log_dir), "log_dir not created"
assert "Kettle log: " + kettle_log in out, out
assert calls[3] == "fill", calls
assert "Failed steps" not in out and "Done." in out, out

out, calls = scenario("kettle_failed", kitchen=1)
assert [c.split()[0] for c in calls] == ["collector", "kitchen", "fill"], calls   # monitor_check skipped
assert calls[2] == "fill --clear-monitor", calls
assert "Failed steps: Kettle" in out, out

out, calls = scenario("monitor_failed", monitor=1)
assert [c.split()[0] for c in calls] == ["collector", "kitchen", "monitor", "fill"], calls
assert calls[3] == "fill --clear-monitor", calls
assert "Failed steps: monitor_check" in out, out

out, calls = scenario("no_config", config=False)   # config.json 讀不到 → [1/4] 失敗，整個停止
assert calls == [], calls
assert "Cannot read" in out and "[1/4]" not in out, out

out, calls = scenario("no_events", collector=2)   # exit 2 = no CSV, the rest still runs
assert [c.split()[0] for c in calls] == ["collector", "kitchen", "monitor", "fill"], calls

out, calls = scenario("collector_failed", collector=1)
assert calls == ["collector"], calls
assert "Event log collection failed" in out, out
# ===== 真的撈取腳本經 run_weekly.ps1：config.json 的中文 + {week} 路徑、相對路徑 =====
real = os.path.join(tmp, "real")
shutil.rmtree(real, ignore_errors=True)
os.makedirs(real)
for f in ("EventLogCheck.ps1", "EventLogCheck_Local.ps1", "run_weekly.ps1"):
    shutil.copy(f, real)
open(os.path.join(real, "fill_check.py"), "w").write("import sys\nsys.exit(0)\n")   # Kitchen 不存在 → Kettle 失敗，直接填表
open(os.path.join(real, "servers.txt"), "w").write("localhost\n")
open(os.path.join(real, "ex.txt"), "w").write("# test\n11\n")
week = f"W{y % 100:02d}{w:02d}"
for collector, host in (("EventLogCheck.ps1", "LOCALHOST"), ("EventLogCheck_Local.ps1", os.environ["COMPUTERNAME"].upper())):
    json.dump({"log_dir": "記錄檔", "paths": {"csv_dir": "中文資料夾/{week}/BOS"},
               "weekly": {"collector": collector, "kitchen": "no_kitchen.bat", "kettle_job": "no.kjb"},
               "eventlog": {"server_list": "servers.txt", "exclude_file": "ex.txt", "days": 7}},
              open(os.path.join(real, "config.json"), "w", encoding="utf-8"), ensure_ascii=False)
    r = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
                        os.path.join(real, "run_weekly.ps1")], capture_output=True, text=True, errors="replace")
    print(f"--- real {collector} (exit {r.returncode})\n{r.stdout}\n{r.stderr}")
    assert r.returncode == 0 and "Event log collection failed" not in r.stdout, r.returncode
    assert "Failed steps: Kettle" in r.stdout, r.stdout
    csv_dir = os.path.join(real, "中文資料夾", week, "BOS")
    if "No error/warning events after exclusion" not in r.stdout:   # 撈到事件才有 CSV
        names = [n.upper() for n in os.listdir(csv_dir)]
        assert any(n.endswith(f"_{host}.CSV") for n in names), (collector, names)
logs = os.listdir(os.path.join(real, "記錄檔"))
assert any(n.endswith(".log") for n in logs), logs
log_text = "".join(open(os.path.join(real, "記錄檔", n), encoding="utf-8-sig").read() for n in logs if n.endswith(".log"))
assert "Excluded event IDs: 1" in log_text, log_text
print("RUN_WEEKLY TEST OK")
