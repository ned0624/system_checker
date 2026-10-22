# -*- coding: utf-8 -*-
"""CI: fill_check template mode end to end.

Builds a template check sheet + tracking file + CSVs in a temp folder, runs fill_check.py --week W2642
twice (second run must keep manual edits) and checks the generated weekly file.
Usage: python ci_template_test.py <empty temp folder>
"""
import json
import os
import shutil
import subprocess
import sys

import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.styles.colors import Color
from openpyxl.worksheet.hyperlink import Hyperlink

tmp = os.path.abspath(sys.argv[1])
# mode: weekswap (no link_target) / rule (link_target, link_text null) / ruletext (link_target + link_text)
MODE = sys.argv[2] if len(sys.argv) > 2 else "weekswap"
os.makedirs(tmp, exist_ok=True)
app = os.path.join(tmp, "app")
os.makedirs(app, exist_ok=True)
shutil.copy("fill_check.py", app)

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
ws.append(["Host Name", "Checker", "事件檢視器\n系統錯誤", "事件檢視器\n警告事件", "防毒軟體", "CPU", "連結", "備註"])
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
os.makedirs(csv_dir, exist_ok=True)
hdr = "LevelDisplayName,Id,LogName,Message,TimeCreated,Count\n"
open(os.path.join(csv_dir, "20261015_SRV-A.csv"), "w", encoding="utf-8-sig").write(
    hdr + "錯誤,1801,System,boot cert,2026-10-15 10:00:00,1\n警告,10016,System,DCOM,2026-10-15 10:00:00,3\n")
open(os.path.join(csv_dir, "20261008_SRV-B.csv"), "w", encoding="utf-8-sig").write(
    hdr + "錯誤,999,System,last week error,2026-10-08 10:00:00,1\n")

out_pattern = os.path.join(tmp, "out", "{week}", "BOS_{week}.xlsx")
cfg = {
    "paths": {"csv_dir": csv_dir, "check_xlsx": tpl, "track_xlsx": track},
    "ignore_list": {"sheet": "可忽略事件ID", "id_column": "可忽略事件ID"},
    "csv": {"id_column": "Id", "level_column": "LevelDisplayName", "desc_column": "Message"},
    "check_sheet": {"sheet": None, "header_row": 1, "warn_header": "事件檢視器\n警告事件",
                    "error_header": "事件檢視器\n系統錯誤", "note_header": "備註", "output": out_pattern,
                    "template": tpl, "antivirus_header": "防毒軟體", "link_header": "連結"},
    "track": {"sheet": "事件檢視器_BOS", "header_row": 1, "host_header": "Server Name", "week_header": "周",
              "ignore_header": "可忽略", "output": None},
    "log_dir": os.path.join(tmp, "logs"),
}
LINK_ROOT = os.path.join(tmp, "share")
if MODE in ("rule", "ruletext"):
    cfg["check_sheet"]["link_target"] = LINK_ROOT.replace("\\", "/") + "/{week}/Reference/BOS"
    cfg["check_sheet"]["link_text"] = "BOS - {host}_Evidence_W{ww}" if MODE == "ruletext" else None
json.dump(cfg, open(os.path.join(app, "config.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=2)


def run():
    r = subprocess.run([sys.executable, os.path.join(app, "fill_check.py"), "--week", "W2642"],
                       capture_output=True, text=True, encoding="utf-8", errors="replace",
                       env=dict(os.environ, PYTHONIOENCODING="utf-8"))
    print(r.stdout, r.stderr)
    assert r.returncode == 0, r.returncode


run()
out = os.path.join(tmp, "out", "W2642", "BOS_W2642.xlsx")
assert os.path.isfile(out), "weekly file not created in W2642 folder"
ws = openpyxl.load_workbook(out).active
v = lambda ref: ws[ref].value

# SRV-A (row 2): error X (1801), warning O (10016 ignorable), note has 1801 only
assert (v("C2"), v("D2")) == ("X", "O"), (v("C2"), v("D2"))
assert "1801" in v("H2") and "10016" not in v("H2"), v("H2")
# SRV-B (row 3): only last week's CSV -> treated as no CSV -> O O
assert (v("C3"), v("D3"), v("H3")) == ("O", "O", None), (v("C3"), v("D3"), v("H3"))
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
if MODE in ("weekswap", "locked"):
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
# tracking sheet: only this week's CSV pasted
trs = openpyxl.load_workbook(track)["事件檢視器_BOS"]
pasted = [(r[0], r[1]) for r in trs.iter_rows(min_row=2, values_only=True) if r[0]]
assert pasted and all(p == ("SRV-A", "W2642") for p in pasted), pasted

# second run keeps manual edits made in the weekly file
wb = openpyxl.load_workbook(out)
wb.active["F2"] = "manual 30%"
wb.save(out)
run()
ws = openpyxl.load_workbook(out).active
assert ws["F2"].value == "manual 30%", ws["F2"].value
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
    marker = out + ".locked"

    def hold_lock(seconds):
        if os.path.exists(marker):
            os.remove(marker)
        # another process opens the file without sharing (like Excel), then writes a marker
        ps = (f"$f=[IO.File]::Open('{out}','Open','ReadWrite','None'); "
              f"Set-Content -LiteralPath '{marker}' 'held'; "
              f"Start-Sleep -Seconds {seconds}; $f.Close()")
        p = subprocess.Popen(["powershell", "-NoProfile", "-Command", ps])
        for _ in range(300):          # PowerShell may start slowly
            if os.path.exists(marker) or p.poll() is not None:
                break
            time.sleep(0.2)
        assert os.path.exists(marker), f"lock holder did not start (exit {p.poll()})"
        assert real_is_locked(out), "file is held exclusively but is_locked() returned False"
        print("lock held by another process, is_locked() = True")
        return p

    def run_with_wait(retry, max_min):
        cfg["wait_if_open"] = {"retry_seconds": retry, "max_wait_minutes": max_min}
        json.dump(cfg, open(os.path.join(app, "config.json"), "w", encoding="utf-8"), ensure_ascii=False)
        r = subprocess.run([sys.executable, os.path.join(app, "fill_check.py"), "--week", "W2642"],
                           capture_output=True, text=True, encoding="utf-8", errors="replace",
                           env=dict(os.environ, PYTHONIOENCODING="utf-8"))
        print(r.stdout, r.stderr)
        return r

    p = hold_lock(12)
    r = run_with_wait(3, 1)                       # waits ~12 s, then saves
    p.wait()
    assert r.returncode == 0 and "[WAIT] Check sheet is open by TestUser" in r.stdout and "Saved:" in r.stdout, r.returncode

    p = hold_lock(30)
    r = run_with_wait(2, 0.1)                     # gives up after 6 s
    p.kill()
    assert r.returncode == 3 and "NOT saved" in r.stdout, r.returncode
    print("REAL LOCK OK")

os.remove(owner)
print("LOCK TEST OK")
