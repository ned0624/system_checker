# -*- coding: utf-8 -*-
"""
監控數據分析：檢查 Kettle（monitor.ktr）匯出的每台主機 CSV

CSV 格式（檔名 = 主機名稱，例: D:/export/p3-wfa-ap01v.csv）:
    time_min,cpu,memory,disk_c,disk_d,disk_e
    upper_spec,60.00,25.60,40.00,420.00,625.00
    2026-11-03 16:42,4.62,22.10,29.95,377.02,600.60
    ...（每分鐘一筆）
    - 沒有表頭時（第一列就是 upper_spec），欄位依 time_min,cpu,memory,disk_c,disk_d,disk_e
    - time_min 以外的欄位都當成指標分析；空白值略過（例: 主機沒有 E 槽）
    - 編碼 UTF-8（含 BOM）或 MS950 都可以

檢查項目（每台主機、每個指標）:
1. 超過 spec：數值 > upper_spec 的筆數、首次/最後時間、最長連續分鐘數、最大值
2. 持續上升沒降下來：每 trend.bucket_minutes 分鐘取平均，
   整段升幅 ≥ spec × min_rise_pct_of_spec%，且過程中從最高點回落 ≤ spec × max_drop_pct_of_spec%
   （沒有 spec 時以第一段平均為基準）；判定上升時估算照目前速度約幾小時後達 spec

輸出:
    report_dir/yyyyMMdd_monitor_report.csv   一台主機一個指標一列（UTF-8 BOM，Excel 可直接開）
    log_dir/yyyyMMdd_monitor_check.log       畫面訊息（同一天多次執行附加在同一檔）

用法:
    monitor_check.exe [csv資料夾]          （有給參數會取代設定檔的 csv_dir）

設定檔: 與本程式同資料夾的 monitor_check.json
    csv_dir / report_dir / log_dir   路徑可用 {week}/{yy}/{ww}（今天的 ISO 週，例: W2645 / 26 / 45）
                                     相對路徑以設定檔所在資料夾為基準
    trend.bucket_minutes             幾分鐘取一次平均（預設 60）
    trend.min_buckets                至少幾段才判定上升（預設 6）
    trend.min_rise_pct_of_spec       升幅門檻，spec 的百分比（預設 5）
    trend.max_drop_pct_of_spec       回落容許值，spec 的百分比（預設 1）

結束碼: 0 = 完成（不論有沒有發現問題），1 = 錯誤（設定檔、CSV 資料夾、報表或 log 寫入失敗）
只用 Python 內建模組，打包成單一 exe（不需要 pandas）
"""
import csv
import glob
import json
import os
import sys
import traceback
from datetime import date, datetime, timedelta

DEFAULT_COLUMNS = ["time_min", "cpu", "memory", "disk_c", "disk_d", "disk_e"]
SPEC_ROW = "upper_spec"
TIME_FORMATS = ("%Y-%m-%d %H:%M", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M", "%Y/%m/%d %H:%M:%S")
REPORT_HEADER = ["host", "metric", "spec", "max", "max_time", "over_count", "over_first", "over_last",
                 "longest_over_min", "first_hour_avg", "last_hour_avg", "rise", "max_drop", "rising",
                 "hours_to_spec", "result"]


def _base_dir():
    """程式所在資料夾；包成執行檔（PyInstaller）時是 exe 所在資料夾，設定檔放在旁邊"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


BASE_DIR = _base_dir()
CONFIG_PATH = os.path.join(BASE_DIR, "monitor_check.json")


def week_label(d):
    """ISO 週次，例: 2026-11-04 → W2645"""
    year, week, _ = d.isocalendar()
    return f"W{year % 100:02d}{week:02d}"


def week_path(path, week):
    """路徑中的 {week}/{yy}/{ww} 換成週別，例: W2645 / 26 / 45"""
    return path.replace("{week}", week).replace("{yy}", week[1:3]).replace("{ww}", week[3:5])


def resolve(p, week):
    return os.path.join(BASE_DIR, week_path(p, week))


def to_float(s):
    s = (s or "").strip()
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def to_time(s):
    s = (s or "").strip()
    for fmt in TIME_FORMATS:
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def read_rows(path):
    """依序試 UTF-8（含 BOM）、MS950"""
    for enc in ("utf-8-sig", "cp950"):
        try:
            with open(path, encoding=enc, newline="") as f:
                return [r for r in csv.reader(f) if any(c.strip() for c in r)]
        except UnicodeDecodeError:
            continue
    raise ValueError("file is neither UTF-8 nor MS950")


def load_host(path):
    """讀一台主機的 CSV → (指標名稱 list, {指標: spec}, [(時間, {指標: 值})], 略過的列數)"""
    rows = read_rows(path)
    if not rows:
        raise ValueError("empty file")
    first = rows[0][0].strip().lower()
    if first == SPEC_ROW:
        width = max(len(r) for r in rows)
        names = DEFAULT_COLUMNS[:width] + [f"col{i + 1}" for i in range(len(DEFAULT_COLUMNS), width)]
    elif first == DEFAULT_COLUMNS[0]:
        names = [c.strip() for c in rows[0]]
        rows = rows[1:]
    else:
        raise ValueError(f"first row is neither a header ({DEFAULT_COLUMNS[0]}) nor {SPEC_ROW}: {rows[0][:3]}")
    metrics = names[1:]
    spec, data, skipped = None, [], 0
    for r in rows:
        values = {m: to_float(r[i + 1]) if i + 1 < len(r) else None for i, m in enumerate(metrics)}
        if r[0].strip().lower() == SPEC_ROW:
            spec = values
            continue
        t = to_time(r[0])
        if t is None:
            skipped += 1
            continue
        data.append((t, values))
    if spec is None:
        raise ValueError(f"no {SPEC_ROW} row")
    if not data:
        raise ValueError("no data rows")
    data.sort(key=lambda x: x[0])
    return metrics, spec, data, skipped


def check_over(points, spec):
    """超過 spec 的筆數、首次/最後時間、最長連續分鐘數（相鄰分鐘才算連續）"""
    res = {"over_count": 0, "over_first": None, "over_last": None, "longest_over_min": 0}
    if spec is None:
        return res
    run, prev = 0, None
    for t, v in points:
        if v > spec:
            res["over_count"] += 1
            res["over_first"] = res["over_first"] or t
            res["over_last"] = t
            run = run + 1 if prev is not None and t - prev == timedelta(minutes=1) else 1
            prev = t
            res["longest_over_min"] = max(res["longest_over_min"], run)
        else:
            prev = None
    return res


def check_trend(points, spec, cfg):
    """每段平均 → 升幅、最大回落、是否持續上升、約幾小時後達 spec"""
    bucket = cfg["bucket_minutes"]
    start = points[0][0]
    sums = {}
    for t, v in points:
        k = int((t - start).total_seconds() // 60 // bucket)
        s = sums.setdefault(k, [0.0, 0])
        s[0] += v
        s[1] += 1
    keys = sorted(sums)
    avgs = [sums[k][0] / sums[k][1] for k in keys]
    res = {"first_hour_avg": avgs[0], "last_hour_avg": avgs[-1], "rise": None, "max_drop": None,
           "rising": "-", "hours_to_spec": None}
    if len(avgs) < cfg["min_buckets"]:
        return res
    peak, drop = avgs[0], 0.0
    for a in avgs:
        peak = max(peak, a)
        drop = max(drop, peak - a)
    rise = avgs[-1] - avgs[0]
    res.update(rise=rise, max_drop=drop)
    base = spec if spec else avgs[0]
    if not base or base <= 0:
        return res
    rising = rise >= base * cfg["min_rise_pct_of_spec"] / 100 and drop <= base * cfg["max_drop_pct_of_spec"] / 100
    res["rising"] = "Y" if rising else "N"
    hours = (keys[-1] - keys[0]) * bucket / 60
    if rising and spec and hours > 0 and avgs[-1] < spec:
        res["hours_to_spec"] = (spec - avgs[-1]) / (rise / hours)
    return res


def analyze(host, metrics, spec, data, trend_cfg):
    """一台主機 → 每個指標一列結果"""
    out = []
    for m in metrics:
        points = [(t, v[m]) for t, v in data if v[m] is not None]
        row = {"host": host, "metric": m, "spec": spec.get(m)}
        if not points:
            row["result"] = "NO DATA"
            out.append(row)
            continue
        t_max, v_max = max(points, key=lambda p: p[1])
        row.update(max=v_max, max_time=t_max)
        row.update(check_over(points, row["spec"]))
        row.update(check_trend(points, row["spec"], trend_cfg))
        flags = [f for f, hit in (("OVER SPEC", row["over_count"]), ("RISING", row["rising"] == "Y")) if hit]
        row["result"] = "+".join(flags) or "OK"
        out.append(row)
    return out


def fmt(v):
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d %H:%M")
    if isinstance(v, float):
        return f"{v:.2f}"
    return str(v)


def describe(r):
    """畫面/log 的一句說明"""
    parts = []
    if r["over_count"]:
        parts.append(f"over spec {fmt(r['spec'])}: {r['over_count']} min, longest {r['longest_over_min']} min, "
                     f"max {fmt(r['max'])} at {fmt(r['max_time'])}")
    if r.get("rising") == "Y":
        s = f"rising {fmt(r['first_hour_avg'])} -> {fmt(r['last_hour_avg'])}"
        if r.get("hours_to_spec") is not None:
            s += f", reaches spec in ~{r['hours_to_spec']:.1f} h"
        parts.append(s)
    return f"{r['metric']} {'; '.join(parts)}"


def main(cfg, week, csv_dir):
    trend = {"bucket_minutes": 60, "min_buckets": 6, "min_rise_pct_of_spec": 5, "max_drop_pct_of_spec": 1}
    trend.update(cfg.get("trend") or {})
    if not os.path.isdir(csv_dir):
        print(f"[ERROR] CSV folder not found: {csv_dir} (check csv_dir in monitor_check.json)")
        return 1
    files = sorted(glob.glob(os.path.join(csv_dir, "*.csv")))
    print(f"CSV folder : {csv_dir} ({len(files)} files)")
    rows, findings, warns = [], [], []
    for p in files:
        host = os.path.splitext(os.path.basename(p))[0].upper()
        try:
            metrics, spec, data, skipped = load_host(p)
        except Exception as e:
            print(f"  [WARN] {host}: cannot read {p}: {e}")
            warns.append(host)
            continue
        res = analyze(host, metrics, spec, data, trend)
        rows += res
        bad = [r for r in res if r["result"] not in ("OK", "NO DATA")]
        findings += bad
        note = f", skipped {skipped} rows with bad time" if skipped else ""
        print(f"  {host}: {len(data)} rows {fmt(data[0][0])} ~ {fmt(data[-1][0])}{note}, "
              + (", ".join(f"{r['metric']} {r['result']}" for r in bad) if bad else "OK"))

    report_dir = resolve(cfg.get("report_dir") or "monitor_report", week)
    report = os.path.join(report_dir, f"{datetime.now():%Y%m%d}_monitor_report.csv")
    try:
        os.makedirs(report_dir, exist_ok=True)
        with open(report, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(REPORT_HEADER)
            for r in rows:
                w.writerow([fmt(r.get(k)) for k in REPORT_HEADER])
    except OSError as e:
        print(f"[ERROR] Cannot write report {report}: {e}")
        return 1

    print(f"\n===== Done: {len(files)} files, {len(findings)} findings, {len(warns)} files not read =====")
    for r in findings:
        print(f"  [{r['result']}] {r['host']} {describe(r)}")
    for h in warns:
        print(f"  [WARN] {h}: file not read, see above")
    print(f"Report: {report}")
    return 0


class _Tee:
    """畫面輸出同時寫進 log 檔"""
    def __init__(self, *streams):
        self.streams = streams

    def write(self, s):
        for st in self.streams:
            try:
                st.write(s)
            except (UnicodeEncodeError, OSError, ValueError):
                pass

    def flush(self):
        for st in self.streams:
            try:
                st.flush()
            except (OSError, ValueError):
                pass


def run():
    week = week_label(date.today())
    try:
        with open(CONFIG_PATH, encoding="utf-8-sig") as f:
            cfg = json.load(f)
    except Exception as e:
        print(f"[ERROR] Cannot read config file {CONFIG_PATH}: {type(e).__name__}: {e}")
        sys.exit(1)
    log_dir = resolve(cfg.get("log_dir") or "logs", week)
    log_path = os.path.join(log_dir, f"{datetime.now():%Y%m%d}_monitor_check.log")
    try:
        os.makedirs(log_dir, exist_ok=True)
        enc = "utf-8" if os.path.exists(log_path) else "utf-8-sig"
        log = open(log_path, "a", encoding=enc)
    except OSError as e:
        print(f"[ERROR] Cannot write log {log_path}: {e}")
        sys.exit(1)
    console, code = sys.stdout, 1
    with log:
        sys.stdout = _Tee(console, log)
        try:
            print(f"\n===== monitor_check start {datetime.now():%Y-%m-%d %H:%M:%S} =====")
            csv_dir = sys.argv[1] if len(sys.argv) > 1 else resolve(cfg.get("csv_dir") or "export", week)
            code = main(cfg, week, csv_dir)
        except Exception as e:
            print(f"[ERROR] Aborted: {type(e).__name__}: {e}")
            log.write("\n" + traceback.format_exc())
        finally:
            print(f"End {datetime.now():%Y-%m-%d %H:%M:%S}, log: {log_path}")
            sys.stdout = console
    sys.exit(code)


if __name__ == "__main__":
    run()
