# -*- coding: utf-8 -*-
"""CI（Windows）：用假的撈取腳本、Kitchen、monitor_check、fill_check 測 run_weekly.bat 的流程

情境：全部成功 / Kettle 失敗 / monitor_check 失敗 / 撈 event log 失敗
用法: python ci_run_weekly_test.py <tmp資料夾>
"""
import os
import shutil
import subprocess
import sys

tmp = os.path.abspath(sys.argv[1])


def scenario(name, collector=0, kitchen=0, monitor=0):
    base = os.path.join(tmp, name)
    shutil.rmtree(base, ignore_errors=True)
    os.makedirs(os.path.join(base, "kettle"))
    log = os.path.join(base, "calls.txt")
    bat = open("run_weekly.bat", encoding="ascii").read()
    bat = bat.replace(r"set BASE=D:\SystemCheck", f"set BASE={base}")
    bat = bat.replace(r"set KITCHEN=D:\data-integration\Kitchen.bat", rf"set KITCHEN={base}\Kitchen.bat")
    assert base in bat and rf"{base}\Kitchen.bat" in bat, "bat variables not found"
    open(os.path.join(base, "run_weekly.bat"), "w", encoding="ascii", newline="\r\n").write(bat)
    open(os.path.join(base, "EventLogCheck.ps1"), "w").write(f'Add-Content "{log}" "collector"\nexit {collector}\n')
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

out, calls = scenario("no_events", collector=2)   # exit 2 = no CSV, the rest still runs
assert [c.split()[0] for c in calls] == ["collector", "kitchen", "monitor", "fill"], calls

out, calls = scenario("collector_failed", collector=1)
assert calls == ["collector"], calls
assert "Event log collection failed" in out, out
print("RUN_WEEKLY TEST OK")
