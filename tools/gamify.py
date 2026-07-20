import json
import openpyxl
import pandas as pd
from datetime import date, datetime
from pathlib import Path

WORKLOG = Path("D:/WorkLog/worklog.xlsx")
SCORE_FILE = Path("D:/WorkLog/.gamify_data.json")

SCORE_THRESHOLDS = [0, 10, 25, 50, 100, 200, 350, 500]
LEVEL_NAMES = [
    "Rookie", "Bronze", "Silver", "Gold",
    "Platinum", "Diamond", "Master", "Grandmaster"
]
DAILY_TARGET_HOURS = 7.5

def load_gamify_data():
    if SCORE_FILE.exists():
        return json.loads(SCORE_FILE.read_text())
    return {
        "total_score": 0,
        "streak": 0,
        "max_streak": 0,
        "last_logged_date": None,
        "daily_history": {},
        "level": 0,
    }

def save_gamify_data(data):
    SCORE_FILE.write_text(json.dumps(data, indent=2, default=str))

def get_time_entries(ws):
    entries = []
    for r in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        date_raw, _, project, task, _, category, _, _, dur, desc = (
            r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9]
        )
        if date_raw and dur and dur > 0:
            entries.append({
                "date": pd.Timestamp(date_raw).date(),
                "project": str(project or ""),
                "task": str(task or ""),
                "category": str(category or ""),
                "duration": dur,
                "description": str(desc or ""),
            })
    return entries

def compute_daily_summary(entries):
    daily = {}
    for e in entries:
        d = e["date"]
        if d not in daily:
            daily[d] = {"total_hours": 0, "task_count": 0, "categories": set()}
        daily[d]["total_hours"] += e["duration"]
        daily[d]["task_count"] += 1
        daily[d]["categories"].add(e["category"])
    return daily

def update_streak(data, daily_summary, today):
    sorted_dates = sorted(daily_summary.keys())
    if not sorted_dates:
        data["streak"] = 0
        data["last_logged_date"] = None
        return

    # Check if logged today
    last = sorted_dates[-1]
    data["last_logged_date"] = str(last)

    # Recalculate streak from today backwards
    streak = 0
    check = today
    while check in daily_summary or check.weekday() >= 5:
        if check.weekday() < 5 and check not in daily_summary:
            break
        if check in daily_summary:
            streak += 1
        check -= pd.Timedelta(days=1)
    data["streak"] = streak
    if streak > data["max_streak"]:
        data["max_streak"] = streak

def compute_scores(entries, daily_summary, data, today):
    today_str = str(today)
    today_entry = daily_summary.get(today, None)
    today_score = 0

    if today_entry:
        util_pct = min(today_entry["total_hours"] / DAILY_TARGET_HOURS, 1.0)
        today_score += round(util_pct * 50)

        task_bonus = min(today_entry["task_count"] * 5, 15)
        today_score += task_bonus

        category_count = len(today_entry["categories"])
        if category_count >= 2:
            today_score += 5

    streak_bonus = min(data["streak"] * 2, 20)
    today_score += streak_bonus

    daily_history = data.get("daily_history", {})
    if today_str not in daily_history:
        daily_history[today_str] = []
    daily_history[today_str].append({
        "date": today_str,
        "score": today_score,
        "hours": today_entry["total_hours"] if today_entry else 0,
        "tasks": today_entry["task_count"] if today_entry else 0,
    })
    data["daily_history"] = daily_history

    data["total_score"] += today_score

    # Determine level
    for i, threshold in reversed(list(enumerate(SCORE_THRESHOLDS))):
        if data["total_score"] >= threshold:
            data["level"] = i
            break

    return today_score

def compute_consistency(entries):
    if not entries:
        return 0.0, 0, 0
    dates = sorted(set(e["date"] for e in entries))
    if not dates:
        return 0.0, 0, 0
    mn, mx = dates[0], dates[-1]
    total_weekdays = sum(
        1 for d in pd.date_range(mn, mx, freq='D')
        if d.weekday() < 5
    )
    logged_weekdays = sum(1 for d in dates if d.weekday() < 5)
    pct = (logged_weekdays / total_weekdays * 100) if total_weekdays > 0 else 0
    return round(pct, 1), logged_weekdays, total_weekdays

def compute_task_performance(ws, entries):
    ws_kpi = openpyxl.load_workbook(WORKLOG, data_only=True)['KPIs']
    tasks = {}
    for r in ws_kpi.iter_rows(min_row=2, max_row=ws_kpi.max_row, values_only=True):
        name, project, _, days, hours_deadline, status = r[0], r[1], r[2], r[3], r[4], r[5]
        if name:
            deadline_h = hours_deadline if hours_deadline else (days * 7.5 if days else None)
            logged = sum(e["duration"] for e in entries if e["task"] == str(name))
            pct = (logged / deadline_h * 100) if deadline_h and deadline_h > 0 else 0
            tasks[str(name)] = {
                "project": str(project),
                "logged_hours": round(logged, 2),
                "deadline_hours": deadline_h,
                "performance_pct": round(pct, 1),
                "status": str(status),
            }
    return tasks

def run():
    wb = openpyxl.load_workbook(WORKLOG, data_only=True)
    ws = wb['Time Entries']
    today = date.today()

    entries = get_time_entries(ws)
    daily = compute_daily_summary(entries)

    data = load_gamify_data()
    update_streak(data, daily, today)
    today_score = compute_scores(entries, daily, data, today)
    consistency, logged_days, total_days = compute_consistency(entries)
    task_perf = compute_task_performance(ws, entries)

    save_gamify_data(data)

    return {
        "today_score": today_score,
        "total_score": data["total_score"],
        "level": data["level"],
        "level_name": LEVEL_NAMES[data["level"]],
        "streak": data["streak"],
        "max_streak": data["max_streak"],
        "consistency_pct": consistency,
        "logged_weekdays": logged_days,
        "total_weekdays": total_days,
        "tasks": task_perf,
        "daily_summary": {
            str(k): {
                "hours": round(v["total_hours"], 2),
                "tasks": v["task_count"],
                "categories": list(v["categories"]),
            }
            for k, v in daily.items()
        },
    }

if __name__ == '__main__':
    result = run()
    print(json.dumps(result, indent=2, default=str))
