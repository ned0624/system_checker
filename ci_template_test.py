# -*- coding: utf-8 -*-
"""CI: fill_check template mode end to end.

Builds a template check sheet + tracking file + CSVs in a temp folder, runs fill_check.py --week W2642
twice (second run must keep manual edits) and checks the generated weekly file.
Usage: python ci_template_test.py <empty temp folder>
"""
import json
import math
import glob
import os
import shutil
import subprocess
import sys

from datetime import datetime, timedelta

import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.styles.colors import Color
from openpyxl.worksheet.hyperlink import Hyperlink

tmp = os.path.abspath(sys.argv[1])
# mode: weekswap (no link_target) / rule (link_target, link_text null) / ruletext (link_target + link_text)
#       locked (file open by someone) / weekdir (paths.csv_dir contains {week})
#       monitor (Kettle monitor CSVs -> CPU/Memory/Disk 2 columns marked O in the template + notes)
MODE = sys.argv[2] if len(sys.argv) > 2 else "weekswap"
os.makedirs(tmp, exist_ok=True)
app = os.path.join(tmp, "app")
os.makedirs(app, exist_ok=True)
shutil.copy("fill_check.py", app)
shutil.copy("monitor_check.py", app)

GRAY = PatternFill("solid", fgColor="FF808080")
BLACK = PatternFill("solid", fgColor="FF000000")
YELLOW = PatternFill("solid", fgColor="FFFFFF00")
THEME_GRAY = PatternFill("solid", fgColor=Color(theme=0, tint=-0.25))
WHITE = PatternFill("solid", fgColor="FFFFFFFF")
SHARE = r"\\xintec.com.tw\tsm\Supporting_Org\Dept\System Check"

# ---- template ----
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Server"
ws.append(["Host Name", "Checker", "事件檢視器\n系統錯誤", "事件檢視器\n警告事件", "防毒軟體", "CPU", "連結", "備註"]
          + (["Memory", "Disk 2"] if MODE == "monitor" else []))
for host in ("SRV-A", "SRV-B", "SRV-C", "SRV-D"):
    ws.append([host, "Ned"])
ws.append(["(季檢,下次 10/1, 2027/1/7)"])   # row 6: yellow merged separator
ws.append(["SRV-E", "Ned"])                  # row 7
for c in "CDE":
    ws[f"{c}4"].fill = GRAY                  # SRV-C: event + antivirus gray
ws["E5"].fill = THEME_GRAY                   # SRV-D: antivirus theme gray
ws["G5"].fill = GRAY                         # SRV-D: link cell gray -> no link
ws["C7"].fill = BLACK                        # SRV-E: error column black
ws["D2"].fill = WHITE                        # explicit white is still fillable
for c in "ABCDEFGH":
    ws[f"{c}6"].fill = YELLOW
ws.merge_cells("A6:H6")
ws["G2"] = "SRV-A_Evidence_W40"
ws["G2"].hyperlink = Hyperlink(ref="G2", target=f"file:///{SHARE}\\W2640\\Reference\\BOS - SRV-A_Evidence_W40.docx")
ws["G3"] = f'=HYPERLINK("{SHARE}\\W2640\\Reference\\BOS - SRV-B_Evidence_W40.docx","SRV-B_Evidence_W40")'
if MODE == "monitor":
    # 打 O = 要檢查的監控項目：SRV-A CPU/Memory/Disk 2、SRV-B CPU/Memory、SRV-D CPU；SRV-E 沒打 O
    for ref in ("F2", "I2", "J2", "F3", "I3", "F5"):
        ws[ref] = "O"
tpl = os.path.join(tmp, "template.xlsx")
wb.save(tpl)

# ---- tracking file ----
wb = openpyxl.Workbook()
ig = wb.active
ig.title = "可忽略事件ID"
ig.append(["可忽略事件ID"])
ig.append([10016])
tr = wb.create_sheet("事件檢視器_BOS")
tr.append(["Server Name", "周", "錯誤/警告", "事件識別碼", "記錄檔名稱", "事件內容", "可能造成的影響", "Action", "可忽略"])
track = os.path.join(tmp, "track.xlsx")
wb.save(track)

# ---- CSVs: SRV-A this week (W2642 = 2026-10-12..18), SRV-B only last week ----
csv_dir = os.path.join(tmp, "csv")
csv_setting = csv_dir
if MODE == "weekdir":   # csv_dir with {week}: CSVs live in the week folder
    csv_setting = os.path.join(tmp, "{week}", "csv")
    csv_dir = os.path.join(tmp, "W2642", "csv")
os.makedirs(csv_dir, exist_ok=True)
hdr = "LevelDisplayName,Id,LogName,Message,TimeCreated,Count\n"
open(os.path.join(csv_dir, "20261015_SRV-A.csv"), "w", encoding="utf-8-sig").write(
    hdr + "錯誤,1801,System,boot cert,2026-10-15 10:00:00,1\n警告,10016,System,DCOM,2026-10-15 10:00:00,3\n")
open(os.path.join(csv_dir, "20261008_SRV-B.csv"), "w", encoding="utf-8-sig").write(
    hdr + "錯誤,999,System,last week error,2026-10-08 10:00:00,1\n")

out_pattern = os.path.join(tmp, "out", "{week}", "BOS_{week}.xlsx")
cfg = {
    "paths": {"csv_dir": csv_setting, "check_xlsx": tpl, "track_xlsx": track},
    "ignore_list": {"sheet": "可忽略事件ID", "id_column": "可忽略事件ID"},
    "csv": {"id_column": "Id", "level_column": "LevelDisplayName", "desc_column": "Message"},
    "check_sheet": {"sheet": None, "header_row": 1, "warn_header": "事件檢視器\n警告事件",
                    "error_header": "事件檢視器\n系統錯誤", "note_header": "備註", "output": out_pattern,
                    "template": tpl, "antivirus_header": "防毒軟體", "link_header": "連結"},
    "track": {"sheet": "事件檢視器_BOS", "header_row": 1, "host_header": "Server Name", "week_header": "周",
              "ignore_header": "可忽略", "output": None},
    # weekdir: config.json 放在 app 的上一層（共用設定檔位置），log_dir 用相對路徑 → 以設定檔資料夾為基準
    "log_dir": os.path.join("{week}", "logs") if MODE == "weekdir" else os.path.join(tmp, "logs"),
}
LINK_ROOT = os.path.join(tmp, "share")
if MODE in ("rule", "ruletext"):
    cfg["check_sheet"]["link_target"] = LINK_ROOT.replace("\\", "/") + "/{week}/Reference/BOS"
    cfg["check_sheet"]["link_text"] = "BOS - {host}_Evidence_W{ww}" if MODE == "ruletext" else None
KETTLE = os.path.join(tmp, "kettle", "2026-W42")   # Kettle 週資料夾 %x-W%v
if MODE == "monitor":
    cfg["monitor"] = {"csv_dir": os.path.join(tmp, "kettle", "{yyyy}-W{ww}")}
    os.makedirs(KETTLE)
    start = datetime(2026, 10, 11, 9, 0)
    DAY = 1440

    def kettle_csv(host, header, spec, rows):
        lines = [header, "upper_spec," + spec] + [
            f"{start + timedelta(minutes=i):%Y-%m-%d %H:%M}," + ",".join("" if x is None else f"{x:.2f}" for x in r)
            for i, r in enumerate(rows)]
        open(os.path.join(KETTLE, host + ".csv"), "w", encoding="utf-8").write("\n".join(lines) + "\n")

    # SRV-A: cpu 超標 30 分鐘 → X；memory 每天上升 → O + 備註；disk2 正常 → O
    kettle_csv("SRV-A", "time_min,cpu,memory,disk2", "60.00,25.60,420.00",
               [(75.0 if 3000 <= i < 3030 else 20 + 10 * math.sin(2 * math.pi * i / DAY), 15 + 0.0006 * i, 300.0)
                for i in range(7 * DAY)])
    # SRV-B: cpu 正常 → O；memory 打 O 但沒有數值 → 空白 + 備註
    kettle_csv("SRV-B", "time_min,cpu,memory", "60.00,25.60", [(20.0, None)] * (7 * DAY))
    # SRV-D: 沒有 CSV → CPU 空白 + 備註
json.dump(cfg, open(os.path.join(tmp if MODE == "weekdir" else app, "config.json"), "w", encoding="utf-8"),
          ensure_ascii=False, indent=2)


def run(*extra):
    r = subprocess.run([sys.executable, os.path.join(app, "fill_check.py"), "--week", "W2642", *extra],
                       capture_output=True, text=True, encoding="utf-8", errors="replace",
                       env=dict(os.environ, PYTHONIOENCODING="utf-8"))
    print(r.stdout, r.stderr)
    assert r.returncode == 0, r.returncode
    if MODE == "weekdir":
        assert "Config: " + os.path.join(tmp, "config.json") in r.stdout, "config.json in parent folder not used"
    return r.stdout


run()
if MODE == "weekdir":   # log_dir with {week}: log goes to the week folder, no literal {week} folder
    assert glob.glob(os.path.join(tmp, "W2642", "logs", "*.log")), "log not written to week folder"
    assert not os.path.exists(os.path.join(tmp, "{week}")), "literal {week} folder created"
out = os.path.join(tmp, "out", "W2642", "BOS_W2642.xlsx")
assert os.path.isfile(out), "weekly file not created in W2642 folder"
ws = openpyxl.load_workbook(out).active
v = lambda ref: ws[ref].value

# SRV-A (row 2): error X (1801), warning O (10016 ignorable), note has 1801 only
assert (v("C2"), v("D2")) == ("X", "O"), (v("C2"), v("D2"))
assert "1801" in v("H2") and "10016" not in v("H2"), v("H2")
# SRV-B (row 3): only last week's CSV -> treated as no CSV -> O O
assert (v("C3"), v("D3")) == ("O", "O") and (MODE == "monitor" or v("H3") is None), (v("C3"), v("D3"), v("H3"))
# SRV-C (row 4): gray cells untouched
assert (v("C4"), v("D4"), v("E4")) == (None, None, None), (v("C4"), v("D4"), v("E4"))
# SRV-D (row 5): theme-gray antivirus untouched, events O O
assert v("E5") is None and (v("C5"), v("D5")) == ("O", "O"), (v("E5"), v("C5"), v("D5"))
# separator row 6 untouched
assert all(v(f"{c}6") is None for c in "BCDEFGH"), [v(f"{c}6") for c in "BCDEFGH"]
# SRV-E (row 7): black error cell untouched, warning O, antivirus O
assert v("C7") is None and v("D7") == "O" and v("E7") == "O", (v("C7"), v("D7"), v("E7"))
# antivirus O on normal rows
assert [v(f"E{r}") for r in (2, 3)] == ["O", "O"]
# links
if MODE in ("weekswap", "locked", "weekdir", "monitor"):
    assert v("G2") == "SRV-A_Evidence_W42", v("G2")
    assert "\\W2642\\" in ws["G2"].hyperlink.target and "_W42.docx" in ws["G2"].hyperlink.target, ws["G2"].hyperlink.target
    assert "\\W2642\\" in v("G3") and "SRV-B_Evidence_W42" in v("G3") and "W2640" not in v("G3"), v("G3")
else:
    folder = (LINK_ROOT.replace("/", "\\") + "\\W2642\\Reference\\BOS")
    for ref, host in (("G2", "SRV-A"), ("G3", "SRV-B"), ("G4", "SRV-C"), ("G7", "SRV-E")):
        assert ws[ref].hyperlink is not None and ws[ref].hyperlink.target == folder, (ref, ws[ref].hyperlink and ws[ref].hyperlink.target)
    assert ws["G5"].hyperlink is None and v("G5") is None, "gray link cell must stay empty"
    if MODE == "rule":
        assert v("G2") == "SRV-A_Evidence_W42", v("G2")     # existing text, week updated
        assert v("G3") == "SRV-B_Evidence_W42", v("G3")     # HYPERLINK formula -> its text, week updated
        assert v("G4") == "BOS" and v("G7") == "BOS", (v("G4"), v("G7"))  # blank -> folder name
    else:
        assert v("G2") == "BOS - SRV-A_Evidence_W42" and v("G7") == "BOS - SRV-E_Evidence_W42", (v("G2"), v("G7"))
    assert ws["G4"].font.underline == "single", ws["G4"].font.underline
if MODE == "monitor":
    assert (v("F2"), v("I2"), v("J2")) == ("X", "O", "O"), (v("F2"), v("I2"), v("J2"))
    assert "1801" in v("H2") and "[監控] CPU 超過 spec 60.00 共 30 分鐘" in v("H2"), v("H2")
    assert "[監控] Memory 持續上升" in v("H2") and "Disk 2" not in v("H2"), v("H2")
    assert (v("F3"), v("I3"), v("J3")) == ("O", None, None), (v("F3"), v("I3"), v("J3"))
    assert v("H3") == "[監控] Memory 無監控資料", v("H3")
    assert v("F5") is None and v("H5") == "[監控] CPU 無監控資料", (v("F5"), v("H5"))
    assert v("F7") is None and v("I7") is None, (v("F7"), v("I7"))
    first_notes = (v("H2"), v("H3"), v("H5"))
# tracking sheet: only this week's CSV pasted
trs = openpyxl.load_workbook(track)["事件檢視器_BOS"]
pasted = [(r[0], r[1]) for r in trs.iter_rows(min_row=2, values_only=True) if r[0]]
assert pasted and all(p == ("SRV-A", "W2642") for p in pasted), pasted

# second run keeps manual edits made in the weekly file
wb = openpyxl.load_workbook(out)
MANUAL = "B2" if MODE == "monitor" else "F2"   # monitor 模式的 F2 是 CPU 結果
wb.active[MANUAL] = "manual 30%"
wb.save(out)
run()
ws = openpyxl.load_workbook(out).active
assert ws[MANUAL].value == "manual 30%", ws[MANUAL].value
if MODE == "monitor":   # 重跑：打 O 從範本讀，[監控] 備註不重複
    assert (ws["F2"].value, ws["I3"].value, ws["F5"].value) == ("X", None, None)
    assert (ws["H2"].value, ws["H3"].value, ws["H5"].value) == first_notes, (ws["H2"].value, first_notes)
    # --clear-monitor（run_weekly.bat：Kettle / monitor_check 失敗）：有 CSV 也清空，[監控] 備註移除，事件欄照常
    o = run("--clear-monitor")
    assert "monitor columns cleared" in o, o
    ws = openpyxl.load_workbook(out).active
    assert (ws["F2"].value, ws["I2"].value, ws["J2"].value, ws["F3"].value, ws["F5"].value) == (None,) * 5
    assert "[監控]" not in ws["H2"].value and "1801" in ws["H2"].value, ws["H2"].value
    assert ws["H3"].value is None and ws["C2"].value == "X", (ws["H3"].value, ws["C2"].value)
    # Kettle 資料夾不存在：WARN，監控欄同樣清空
    os.remove(out)
    shutil.rmtree(KETTLE)
    o = run()
    assert "Monitor CSV folder not found" in o and "monitor columns cleared" in o, o
    ws = openpyxl.load_workbook(out).active
    assert (ws["F2"].value, ws["I2"].value, ws["F5"].value) == (None, None, None), (ws["F2"].value, ws["F5"].value)
    assert "[監控]" not in ws["H2"].value, ws["H2"].value
assert ws["C2"].value == "X"
assert ws["G2"].value == ("BOS - SRV-A_Evidence_W42" if MODE == "ruletext" else "SRV-A_Evidence_W42"), ws["G2"].value
print(f"TEMPLATE TEST OK ({MODE})")

if MODE != "locked":
    sys.exit(0)

# ===== locked: the weekly file is open by someone =====
import contextlib
import io
import time

owner = os.path.join(os.path.dirname(out), "~$" + os.path.basename(out))
open(owner, "wb").write(bytes([8]) + b"TestUser" + b" " * 46)   # Excel owner file format

# --- A. simulated lock (any OS): locked twice, another user saves an edit meanwhile -> edit kept ---
sys.path.insert(0, app)
import fill_check as fc   # noqa: E402  (reads app/config.json)
real_is_locked = fc.is_locked
calls = {"n": 0}


def fake_is_locked(path):
    if os.path.abspath(path) == os.path.abspath(out) and calls["n"] < 2:
        calls["n"] += 1
        if calls["n"] == 1:   # someone edits and saves while we wait
            wb2 = openpyxl.load_workbook(out)
            wb2.active["F4"] = "edited while waiting"
            wb2.save(out)
        return True
    return real_is_locked(path)


fc.is_locked = fake_is_locked
fc.time.sleep = lambda s: None
fc.RETRY_SECONDS, fc.MAX_WAIT_MINUTES = 1, 1
buf = io.StringIO()
sys.argv = ["fill_check.py", "--week", "W2642"]
with contextlib.redirect_stdout(buf):
    try:
        fc.main()
        code = 0
    except SystemExit as e:
        code = e.code or 0
log = buf.getvalue()
print(log)
assert code == 0, code
assert log.count("[WAIT] Check sheet is open by TestUser") == 2, log
ws = openpyxl.load_workbook(out).active
assert ws["F4"].value == "edited while waiting", "edit saved during the wait was overwritten"
assert ws["F2"].value == "manual 30%" and ws["C2"].value == "X"

# --- B. simulated timeout: still locked -> exit 3, file untouched ---
fc.is_locked = lambda path: os.path.abspath(path) == os.path.abspath(out) or real_is_locked(path)
fc.MAX_WAIT_MINUTES = 0
before = open(out, "rb").read()
buf = io.StringIO()
with contextlib.redirect_stdout(buf):
    try:
        fc.main()
        code = 0
    except SystemExit as e:
        code = e.code or 0
print(buf.getvalue())
assert code == 3, code
assert "still open by TestUser" in buf.getvalue() and "NOT saved" in buf.getvalue()
assert open(out, "rb").read() == before, "locked file must not be modified"
print("SIMULATED LOCK OK")

# --- C. real exclusive lock held by another process (Windows only) ---
if os.name == "nt":
    import ctypes
    import threading
    from ctypes import wintypes

    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    k32.CreateFileW.restype = wintypes.HANDLE
    k32.CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, wintypes.LPVOID,
                                wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
    k32.CloseHandle.argtypes = [wintypes.HANDLE]

    def hold_lock(seconds):
        """Open the file like Excel does (others may read, nobody else may write); release after N s"""
        GENERIC_RW, FILE_SHARE_READ, OPEN_EXISTING = 0xC0000000, 1, 3
        h = k32.CreateFileW(out, GENERIC_RW, FILE_SHARE_READ, None, OPEN_EXISTING, 0x80, None)
        assert h not in (None, wintypes.HANDLE(-1).value), f"CreateFileW failed: {ctypes.get_last_error()}"
        assert real_is_locked(out), "file is held exclusively but is_locked() returned False"
        print("lock held, is_locked() = True")
        timer = threading.Timer(seconds, k32.CloseHandle, args=(h,))
        timer.start()
        return timer

    def run_with_wait(retry, max_min):
        cfg["wait_if_open"] = {"retry_seconds": retry, "max_wait_minutes": max_min}
        json.dump(cfg, open(os.path.join(app, "config.json"), "w", encoding="utf-8"), ensure_ascii=False)
        r = subprocess.run([sys.executable, os.path.join(app, "fill_check.py"), "--week", "W2642"],
                           capture_output=True, text=True, encoding="utf-8", errors="replace",
                           env=dict(os.environ, PYTHONIOENCODING="utf-8"))
        print(r.stdout, r.stderr)
        return r

    timer = hold_lock(12)
    r = run_with_wait(3, 1)                       # waits ~12 s, then saves
    timer.join()
    assert r.returncode == 0 and "[WAIT] Check sheet is open by TestUser" in r.stdout and "Saved:" in r.stdout, r.returncode

    before = open(out, "rb").read()
    timer = hold_lock(15)
    r = run_with_wait(2, 0.1)                     # gives up after 6 s
    timer.join()
    assert open(out, "rb").read() == before, "locked file must not be modified"
    assert r.returncode == 3 and "NOT saved" in r.stdout, r.returncode
    print("REAL LOCK OK")

os.remove(owner)
print("LOCK TEST OK")
