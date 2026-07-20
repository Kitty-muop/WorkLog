import openpyxl
import pandas as pd
from pathlib import Path

WORKLOG = Path("D:/WorkLog/worklog.xlsx")

def load_ws():
    return openpyxl.load_workbook(WORKLOG, data_only=True)['Time Entries']

def empty_rows(ws):
    count = 0
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=10):
        if all(c.value is None for c in row):
            count += 1
    return count

def missing_descriptions(ws):
    out = []
    for r in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        date, dur, desc = r[0], r[8], r[9]
        if date and dur and dur > 0 and (desc is None or str(desc).strip() in ("", "0")):
            out.append({"date": date, "duration": dur})
    return out

def missing_categories(ws):
    out = []
    for r in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        date, cat, dur = r[0], r[5], r[8]
        if date and dur and dur > 0 and (cat is None or str(cat).strip() == ""):
            out.append({"date": date, "duration": dur})
    return out

def utilization_gaps(ws):
    daily = {}
    for r in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        date, dur = r[0], r[8]
        if date and dur and dur > 0:
            d = pd.Timestamp(date).date()
            daily[d] = daily.get(d, 0) + dur
    if not daily:
        return []
    mn, mx = min(daily), max(daily)
    gaps = []
    for d in pd.date_range(mn, mx, freq='D'):
        if d.weekday() < 5 and d.date() not in daily:
            gaps.append(d.date())
    return gaps

def run():
    ws = load_ws()
    issues = []

    e = empty_rows(ws)
    if e:
        issues.append({"type": "EMPTY_ROWS", "severity": "medium", "detail": f"{e} empty rows"})

    for m in missing_descriptions(ws):
        issues.append({"type": "MISSING_DESC", "severity": "low", "detail": f"{m['date']}: {m['duration']}h no description"})

    for m in missing_categories(ws):
        issues.append({"type": "MISSING_CATEGORY", "severity": "low", "detail": f"{m['date']}: {m['duration']}h no category"})

    for g in utilization_gaps(ws):
        issues.append({"type": "UTILIZATION_GAP", "severity": "medium", "detail": f"No hours on {g} (weekday)"})

    return issues

if __name__ == '__main__':
    for i in run():
        print(f"[{i['severity'].upper()}] {i['type']}: {i['detail']}")
