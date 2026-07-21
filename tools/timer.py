#!/usr/bin/env python3
"""
Work Log Timer CLI — for WPS Office users
Usage:
  python timer.py start                  Start timer
  python timer.py stop [options]         Stop timer, write to worklog.xlsx
  python timer.py pomo [--work 25] [--rest 5]   Pomodoro session
  python timer.py today                  Show today's hours
  python timer.py week                   Show this week's hours
  python timer.py cancel                 Cancel running timer (discard)
  python timer.py status                 Show timer status
"""

import argparse
import json
import os
import sys
import time
import subprocess
import datetime
from pathlib import Path

import openpyxl
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).parent))
import db
import webhook

STATE_FILE = Path.home() / '.worklog_timer.json'
WORKLOG_FILE = 'worklog.xlsx'

DAY_NAMES = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']


# ─── State ───────────────────────────────────────────────────────────

def load_state():
    if STATE_FILE.exists():
        return json.loads(STATE_FILE.read_text(encoding='utf-8'))
    return {}


def save_state(data):
    STATE_FILE.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding='utf-8')


def clear_state():
    if STATE_FILE.exists():
        STATE_FILE.unlink()


def get_elapsed(state):
    """Get total elapsed seconds from state. Works with paused or running state."""
    acc = state.get('accumulated_seconds', 0)
    seg = state.get('segment_start')
    if seg and not state.get('paused', False):
        sd = datetime.datetime.fromisoformat(seg)
        acc += (datetime.datetime.now() - sd).total_seconds()
    return acc


# ─── Git Auto-detect ─────────────────────────────────────────────────

def detect_project():
    try:
        top = subprocess.check_output(
            ['git', 'rev-parse', '--show-toplevel'],
            stderr=subprocess.DEVNULL
        ).decode().strip()
        return Path(top).name
    except Exception:
        return None


def detect_branch():
    try:
        return subprocess.check_output(
            ['git', 'branch', '--show-current'],
            stderr=subprocess.DEVNULL
        ).decode().strip()
    except Exception:
        return None


def detect_commit_msg():
    try:
        return subprocess.check_output(
            ['git', 'log', '-1', '--format=%s'],
            stderr=subprocess.DEVNULL
        ).decode().strip()
    except Exception:
        return None


# ─── XLSX helpers ────────────────────────────────────────────────────

def get_day_name(d):
    return DAY_NAMES[d.weekday()]


def find_next_row(ws):
    """Find first empty row in Time Entries (col A check)."""
    for r in range(2, ws.max_row + 2):
        if ws.cell(r, 1).value is None:
            return r
    return ws.max_row + 1


def get_next_subtask_code(wb, father_task_code):
    """Generate next subtask code like KPI-1-ST-1, KPI-1-ST-2..."""
    ws = wb['Time Entries']
    max_st = 0
    for r2 in range(2, ws.max_row + 1):
        sc = ws.cell(r2, 6).value
        if sc and str(sc).startswith(f'{father_task_code}-ST-'):
            try:
                num = int(str(sc).rsplit('-', 1)[1])
                max_st = max(max_st, num)
            except:
                pass
    return f'{father_task_code}-ST-{max_st + 1}'


def append_entry(project, task, subtask, subtask_code, category, description,
                 start_dt, end_dt, break_info=None):
    """Append one row to Time Entries sheet and save."""
    wb = load_workbook(WORKLOG_FILE)
    ws = wb['Time Entries']
    r = find_next_row(ws)

    date_val = start_dt.date()
    day_str = get_day_name(date_val)

    ws.cell(r, 1).value = date_val
    ws.cell(r, 1).number_format = 'dd-mm-yyyy'
    ws.cell(r, 2).value = day_str
    ws.cell(r, 3).value = project
    ws.cell(r, 4).value = task
    ws.cell(r, 5).value = subtask or None
    ws.cell(r, 6).value = subtask_code or None
    ws.cell(r, 7).value = category or 'Work'
    ws.cell(r, 8).value = start_dt.time().replace(microsecond=0)
    ws.cell(r, 8).number_format = 'hh:mm'
    ws.cell(r, 9).value = end_dt.time().replace(microsecond=0)
    ws.cell(r, 9).number_format = 'hh:mm'
    ws.cell(r, 10).value = f'=IF(AND(H{r}<>"",I{r}<>""),(I{r}-H{r})*24,"")'
    ws.cell(r, 11).value = description or ''
    ws.cell(r, 12).value = break_info or ''
    # Helper cols M-R: let formulas auto-propagate (leave blank, formulas fill)

    db.insert_entry(project, task, subtask, subtask_code, category, description,
                    start_dt, end_dt, break_info or '')

    wb.save(WORKLOG_FILE)
    print(f"  Written to {WORKLOG_FILE} (row {r})")
    return r


# ─── Commands ────────────────────────────────────────────────────────

def cmd_start(args):
    state = load_state()
    if 'segment_start' in state and not state.get('paused', False):
        sd = datetime.datetime.fromisoformat(state['segment_start'])
        elapsed = get_elapsed(state)
        print(f"Timer already running (started {sd.strftime('%H:%M:%S')}, "
              f"{int(elapsed)//60}m elapsed)")
        print("  Use 'stop' to save, 'pause' to pause, or 'cancel' to discard")
        return 1
    if state.get('paused', False):
        print("Timer is paused. Use 'continue' to resume or 'cancel' to discard.")
        return 1

    now = datetime.datetime.now()
    info = {
        'accumulated_seconds': 0,
        'segment_start': now.isoformat(),
        'paused': False,
        'project': args.project or detect_project(),
        'task': args.task or detect_branch(),
        'subtask': args.subtask,
        'subtask_code': args.subtask_code,
        'category': args.category or 'Development',
        'description': args.description or detect_commit_msg(),
    }
    save_state(info)
    print(f"Timer started at {now.strftime('%H:%M:%S')}")
    if info['project']:
        print(f"  Project: {info['project']}")
    if info['task']:
        print(f"  Task:    {info['task']}")
    if info['description']:
        print(f"  Desc:    {info['description']}")
    if not info['project'] or not info['task']:
        print("  (Use -p PROJECT -t TASK to override auto-detect)")
    return 0


def cmd_stop(args):
    state = load_state()
    if 'segment_start' not in state and state.get('accumulated_seconds', 0) <= 0:
        print("No timer running. Use 'start' first.")
        return 1

    now = datetime.datetime.now()
    acc = get_elapsed(state)
    end_dt = now

    seg = state.get('segment_start')
    if seg and not state.get('paused', False):
        start_dt = datetime.datetime.fromisoformat(seg)
    else:
        start_dt = now - datetime.timedelta(seconds=acc)

    # Close any open pause
    pause_log = state.get('pause_log', [])
    if state.get('paused', False):
        pause_start = datetime.datetime.fromisoformat(state['segment_start'])
        pause_duration = (now - pause_start).total_seconds() / 60
        reason = state.get('pause_reason', 'break')
        pause_log.append({
            'reason': reason,
            'start': pause_start.isoformat(),
            'end': now.isoformat(),
            'duration_min': int(pause_duration)
        })

    # Build break info string
    break_info = ''
    total_break_min = 0
    if pause_log:
        parts = []
        for p in pause_log:
            parts.append(f"{p['reason']}({p['duration_min']}m)")
            total_break_min += p['duration_min']
        break_info = '; '.join(parts)

    project = args.project or state.get('project') or detect_project() or input("Project: ")
    task = args.task or state.get('task') or detect_branch() or input("Father Task: ")
    subtask = args.subtask or state.get('subtask')
    subtask_code = args.subtask_code or state.get('subtask_code')
    category = args.category or state.get('category') or 'Development'
    description = args.description or state.get('description') or input("Description: ")

    if subtask and not subtask_code:
        wb_sc = load_workbook(WORKLOG_FILE)
        subtask_code = get_next_subtask_code(wb_sc, task)

    duration = acc / 3600
    print(f"Stopped at {end_dt.strftime('%H:%M:%S')}")
    print(f"  Duration: {duration:.2f}h ({start_dt.strftime('%H:%M')} - {end_dt.strftime('%H:%M')})")
    print(f"  Project:  {project}")
    print(f"  Task:     {task}")
    if subtask:
        print(f"  Sub Task: {subtask}")
    print(f"  Category: {category}")
    print(f"  Desc:     {description}")
    if break_info:
        print(f"  Breaks:   {break_info}")
        if total_break_min >= 60:
            print(f"  Break total: {total_break_min}m ({total_break_min/60:.1f}h) off-task")

    append_entry(project, task, subtask, subtask_code, category, description, start_dt, end_dt, break_info)
    webhook.send(project=project, task=task, duration_h=duration,
                 description=description, category=category, subtask=subtask,
                 break_info=break_info,
                 start_time=start_dt.strftime('%H:%M'), end_time=end_dt.strftime('%H:%M'))
    clear_state()
    return 0


def cmd_cancel(args):
    state = load_state()
    if 'segment_start' not in state and state.get('accumulated_seconds', 0) <= 0:
        print("No timer running.")
        return 1
    acc = get_elapsed(state)
    print(f"Cancelled timer ({int(acc)//60}m discarded)")
    clear_state()
    return 0


def cmd_status(args):
    state = load_state()
    if 'segment_start' not in state and state.get('accumulated_seconds', 0) <= 0:
        print("No timer running.")
        return 0

    acc = get_elapsed(state)
    mins = int(acc) // 60
    hours = mins / 60

    if state.get('paused', False):
        reason = state.get('pause_reason', 'break')
        print(f"Timer PAUSED ({mins}m / {hours:.2f}h accumulated)")
        print(f"  Reason: {reason}")
    else:
        sd = datetime.datetime.fromisoformat(state['segment_start'])
        print(f"Timer running: {sd.strftime('%H:%M:%S')} ({mins}m / {hours:.2f}h elapsed)")

    if state.get('project'):
        print(f"  Project: {state['project']}")
    if state.get('task'):
        print(f"  Task:    {state['task']}")
    return 0


def cmd_pause(args):
    state = load_state()
    if 'segment_start' not in state:
        print("No timer running.")
        return 1
    if state.get('paused', False):
        print("Timer is already paused. Use 'continue' to resume.")
        return 1

    now = datetime.datetime.now()
    seg = datetime.datetime.fromisoformat(state['segment_start'])
    elapsed = (now - seg).total_seconds()
    state['accumulated_seconds'] = state.get('accumulated_seconds', 0) + elapsed
    state['segment_start'] = now.isoformat()
    state['paused'] = True
    reason = args.reason if args.reason else 'break'
    state['pause_reason'] = reason
    save_state(state)

    total = int(state['accumulated_seconds']) // 60
    print(f"Timer paused at {now.strftime('%H:%M:%S')} ({total}m accumulated)")
    print(f"  Reason: {reason}")
    return 0


def cmd_continue(args):
    state = load_state()
    if not state.get('paused', False):
        print("Timer is not paused.")
        return 1

    now = datetime.datetime.now()
    pause_start = datetime.datetime.fromisoformat(state['segment_start'])
    pause_duration = (now - pause_start).total_seconds() / 60
    reason = state.get('pause_reason', 'break')

    pause_log = state.get('pause_log', [])
    pause_log.append({
        'reason': reason,
        'start': pause_start.isoformat(),
        'end': now.isoformat(),
        'duration_min': int(pause_duration)
    })
    state['pause_log'] = pause_log
    state['pause_reason'] = ''

    state['segment_start'] = now.isoformat()
    state['paused'] = False
    save_state(state)

    print(f"Timer resumed at {now.strftime('%H:%M:%S')} ({int(pause_duration)}m {reason})")
    if state.get('project'):
        print(f"  Project: {state['project']}")
    if state.get('task'):
        print(f"  Task:    {state['task']}")
    return 0


def get_duration(ws, r):
    """Get numeric duration from row. Computes from start/end if formula not evaluated."""
    dur = ws.cell(r, 10).value
    if isinstance(dur, (int, float)):
        return float(dur)
    start = ws.cell(r, 8).value
    end = ws.cell(r, 9).value
    if start and end:
        if isinstance(start, datetime.datetime):
            s = start
        elif isinstance(start, datetime.time):
            s = datetime.datetime.combine(datetime.date.today(), start)
        else:
            return 0.0
        if isinstance(end, datetime.datetime):
            e = end
        elif isinstance(end, datetime.time):
            e = datetime.datetime.combine(datetime.date.today(), end)
        else:
            return 0.0
        return max(0, (e - s).total_seconds() / 3600)
    return 0.0


def cmd_today(args):
    if not os.path.exists(WORKLOG_FILE):
        print(f"No {WORKLOG_FILE} found.")
        return 1
    wb = load_workbook(WORKLOG_FILE)
    ws = wb['Time Entries']
    today = datetime.date.today()
    total = 0.0
    count = 0
    print(f"Today ({today.strftime('%d-%m-%Y')}):")
    print(f"{'Project':<15} {'Task':<20} {'Hours':>6} {'Desc':<30}")
    print("-" * 75)
    for r in range(2, ws.max_row + 1):
        d = ws.cell(r, 1).value
        if d is None:
            break
        if isinstance(d, datetime.datetime):
            d = d.date()
        if d != today:
            continue
        dur = get_duration(ws, r)
        total += dur
        count += 1
        proj = ws.cell(r, 3).value or ''
        task = ws.cell(r, 4).value or ''
        desc = ws.cell(r, 11).value or ''
        print(f"{str(proj):<15} {str(task):<20} {dur:>6.2f} {str(desc):<30}")
    print("-" * 75)
    print(f"{'TOTAL':<15} {'':<20} {total:>6.2f}")
    if total >= 7.5:
        print("  Full day!")
    else:
        rem = 7.5 - total
        print(f"  Partial day: {rem:.2f}h remaining")
    return 0


def cmd_week(args):
    if not os.path.exists(WORKLOG_FILE):
        print(f"No {WORKLOG_FILE} found.")
        return 1
    wb = load_workbook(WORKLOG_FILE)
    ws = wb['Time Entries']

    today = datetime.date.today()
    monday = today - datetime.timedelta(days=today.weekday())
    sunday = monday + datetime.timedelta(days=6)

    print(f"Week: {monday.strftime('%d-%m-%Y')} - {sunday.strftime('%d-%m-%Y')}")
    print()

    entries = {}
    total = 0.0
    for r in range(2, ws.max_row + 1):
        d = ws.cell(r, 1).value
        if d is None:
            break
        if isinstance(d, datetime.datetime):
            d = d.date()
        if not (monday <= d <= sunday):
            continue
        dur = get_duration(ws, r)
        total += dur
        proj = str(ws.cell(r, 3).value or '?')
        task = str(ws.cell(r, 4).value or '?')
        key = f"{proj}: {task}"
        entries.setdefault(key, 0)
        entries[key] += dur

    print(f"{'Task':<35} {'Hours':>6}")
    print("-" * 45)
    for k, v in sorted(entries.items(), key=lambda x: -x[1]):
        print(f"{k:<35} {v:>6.2f}")
    print("-" * 45)
    print(f"{'TOTAL':<35} {total:>6.2f}")
    if total:
        days_worked = len({ws.cell(r, 1).value
                          for r in range(2, ws.max_row + 1)
                          if ws.cell(r, 1).value
                          and (isinstance(ws.cell(r, 1).value, datetime.datetime) and
                               monday <= ws.cell(r, 1).value.date() <= sunday) or
                          (isinstance(ws.cell(r, 1).value, datetime.date) and
                           monday <= ws.cell(r, 1).value <= sunday)})
        print(f"  Days active: {days_worked}, Avg: {total / max(days_worked, 1):.2f}h/day")
    return 0


def cmd_pomo(args):
    work_mins = args.work
    rest_mins = args.rest

    print(f"Pomodoro: {work_mins}min work + {rest_mins}min rest")
    print("  Press Ctrl+C to stop")

    cycles = 0
    while True:
        cycles += 1
        print(f"\n--- Cycle {cycles}: WORK {work_mins}min ---")
        try:
            countdown(work_mins * 60)
        except KeyboardInterrupt:
            print("\nPomodoro stopped.")
            break

        print(f"Work done! Rest {rest_mins}min...")
        try:
            countdown(rest_mins * 60, rest=True)
        except KeyboardInterrupt:
            print("\nPomodoro stopped.")
            break

    print(f"\nCompleted {cycles} pomodoro cycles.")
    return 0


# ─── KPI commands ────────────────────────────────────────────────────

def get_kpis():
    """Read KPIs sheet, return list of dicts."""
    if not os.path.exists(WORKLOG_FILE):
        return []
    wb = load_workbook(WORKLOG_FILE)
    ws = wb['KPIs']
    kpis = []
    for r in range(2, ws.max_row + 1):
        name = ws.cell(r, 2).value
        if name is None or name == '' or name == '[Enter father task name]':
            break
        kpis.append({
            'row': r,
            'code': str(ws.cell(r, 1).value or ''),
            'name': str(name),
            'project': str(ws.cell(r, 3).value or ''),
            'date': ws.cell(r, 4).value,
            'deadline_days': ws.cell(r, 5).value,
            'status': str(ws.cell(r, 7).value or ''),
            'completed': ws.cell(r, 8).value,
        })
    return kpis


def get_total_hours_for_task(task_name):
    """Get total logged hours for a father task from Time Entries."""
    if not os.path.exists(WORKLOG_FILE):
        return 0.0
    wb = load_workbook(WORKLOG_FILE)
    ws = wb['Time Entries']
    total = 0.0
    for r in range(2, ws.max_row + 1):
        t = ws.cell(r, 4).value
        if t is None:
            break
        if str(t) == task_name:
            total += get_duration(ws, r)
    return total


def cmd_kpi(args):
    if not args.kpi_cmd:
        print("Usage: kpi {list|add|edit|rename|delete|done|status}")
        print()
        print("  list       List all father tasks with status")
        print("  add        Add a new father task")
        print("  edit       Edit project/deadline")
        print("  rename     Rename father task (updates KPIs + Time Entries)")
        print("  delete     Delete a father task")
        print("  done       Mark a father task as completed")
        print("  status     Show overall task progress")
        return 1
    return args.kpi_func(args)


def cmd_kpi_list(args):
    kpis = get_kpis()
    if not kpis:
        print("No KPIs found. Add one with: kpi add -t NAME -p PROJECT -d DAYS")
        return 0
    print(f"{'Code':<10} {'Father Task':<25} {'Project':<15} {'Deadline':<10} {'Hours':>8} {'Status':<15}")
    print("-" * 85)
    for k in kpis:
        hours = get_total_hours_for_task(k['name'])
        dl = k['deadline_days']
        dl_str = f"{dl}d" if dl else '—'
        print(f"{k['code']:<10} {k['name']:<25} {k['project']:<15} {dl_str:<10} {hours:>8.2f} {k['status']:<15}")
    return 0


def cmd_kpi_add(args):
    if not args.name:
        print("Error: -t NAME is required")
        return 1
    wb = load_workbook(WORKLOG_FILE)
    ws = wb['KPIs']
    r = 2
    while ws.cell(r, 2).value and str(ws.cell(r, 2).value).strip() not in ('', '[Enter father task name]'):
        if str(ws.cell(r, 2).value).strip() == args.name.strip():
            print(f"KPI '{args.name}' already exists at row {r}.")
            return 1
        r += 1

    max_num = 0
    for r_check in range(2, ws.max_row + 1):
        code_val = ws.cell(r_check, 1).value
        if code_val and str(code_val).startswith('KPI-'):
            try:
                num = int(str(code_val).split('-')[1])
                max_num = max(max_num, num)
            except:
                pass
    new_code = f'KPI-{max_num + 1}'

    ws.cell(r, 1).value = new_code
    ws.cell(r, 2).value = args.name
    ws.cell(r, 3).value = args.project or ''
    ws.cell(r, 4).value = datetime.date.today()
    ws.cell(r, 4).number_format = 'dd-mm-yyyy'
    ws.cell(r, 5).value = args.deadline
    ws.cell(r, 6).value = f'=IF(E{r}<>"",E{r}*7.5,"")'
    ws.cell(r, 7).value = 'In Progress'
    ws.cell(r, 8).value = None
    wb.save(WORKLOG_FILE)
    print(f"KPI added: {args.name} (code={new_code}, project={args.project or '?'}, deadline={args.deadline}d)")
    return 0


def cmd_kpi_done(args):
    if not args.name:
        print("Error: -t NAME is required")
        return 1
    kpis = get_kpis()
    found = [k for k in kpis if k['name'] == args.name]
    if not found:
        print(f"KPI '{args.name}' not found.")
        return 1
    wb = load_workbook(WORKLOG_FILE)
    ws = wb['KPIs']
    r = found[0]['row']
    ws.cell(r, 7).value = 'Done'
    ws.cell(r, 8).value = datetime.date.today()
    ws.cell(r, 8).number_format = 'dd-mm-yyyy'
    wb.save(WORKLOG_FILE)
    print(f"KPI '{args.name}' marked as Done.")
    return 0


def cmd_kpi_edit(args):
    kpis = get_kpis()
    found = [k for k in kpis if k['name'] == args.name]
    if not found:
        print(f"KPI '{args.name}' not found.")
        return 1
    wb = load_workbook(WORKLOG_FILE)
    ws = wb['KPIs']
    r = found[0]['row']
    changed = []
    if args.project is not None:
        ws.cell(r, 3).value = args.project
        changed.append(f"project={args.project}")
    if args.deadline is not None:
        ws.cell(r, 5).value = args.deadline
        changed.append(f"deadline={args.deadline}d")
    wb.save(WORKLOG_FILE)
    if changed:
        print(f"KPI '{args.name}' updated: {', '.join(changed)}")
    else:
        print("No changes made. Use -p PROJECT and/or -d DEADLINE.")
    return 0


def cmd_kpi_rename(args):
    kpis = get_kpis()
    found = [k for k in kpis if k['name'] == args.name]
    if not found:
        print(f"KPI '{args.name}' not found.")
        return 1
    new_name = args.new_name or args.name
    if new_name == args.name:
        print("New name is same as old name.")
        return 0
    wb = load_workbook(WORKLOG_FILE)
    ws_kpi = wb['KPIs']
    ws_te = wb['Time Entries']
    r = found[0]['row']
    ws_kpi.cell(r, 2).value = new_name
    for r2 in range(2, ws_te.max_row + 1):
        old_val = ws_te.cell(r2, 4).value
        if old_val and str(old_val).strip() == args.name:
            ws_te.cell(r2, 4).value = new_name
    wb.save(WORKLOG_FILE)
    print(f"Renamed '{args.name}' -> '{new_name}' (updated in KPIs + Time Entries)")
    return 0


def cmd_kpi_delete(args):
    kpis = get_kpis()
    found = [k for k in kpis if k['name'] == args.name]
    if not found:
        print(f"KPI '{args.name}' not found.")
        return 1
    wb = load_workbook(WORKLOG_FILE)
    ws = wb['KPIs']
    r = found[0]['row']
    ws.delete_rows(r)
    wb.save(WORKLOG_FILE)
    print(f"KPI '{args.name}' deleted.")
    return 0


def cmd_kpi_status(args):
    kpis = get_kpis()
    if not kpis:
        print("No KPIs found.")
        return 0
    done_count = sum(1 for k in kpis if k['status'] == 'Done')
    in_progress = sum(1 for k in kpis if k['status'] != 'Done')
    total_hours = sum(get_total_hours_for_task(k['name']) for k in kpis)
    print(f"{'Code':<10} {'Task':<25} {'Hours':>8} {'Deadline':<10} {'Status':<12} {'Completed':<15}")
    print("-" * 82)
    for k in kpis:
        hours = get_total_hours_for_task(k['name'])
        dl = k['deadline_days']
        dl_str = f"{dl}d" if dl else '—'
        comp = ''
        if k['completed']:
            comp = k['completed'].strftime('%d-%m-%Y') if hasattr(k['completed'], 'strftime') else str(k['completed'])
        print(f"{k['code']:<10} {k['name']:<25} {hours:>8.2f} {dl_str:<10} {k['status']:<12} {comp:<15}")
    print("-" * 82)
    print(f"Done: {done_count} | In Progress: {in_progress} | Total hours: {total_hours:.2f}")
    return 0


# ─── Tasks command ───────────────────────────────────────────────────

def cmd_tasks(args):
    if not os.path.exists(WORKLOG_FILE):
        print(f"No {WORKLOG_FILE} found.")
        return 1
    wb = load_workbook(WORKLOG_FILE)
    ws = wb['Time Entries']

    groups = {}
    for r in range(2, ws.max_row + 1):
        ft = ws.cell(r, 4).value
        if ft is None or str(ft).strip() == '':
            break
        ft = str(ft).strip()
        st = ws.cell(r, 5).value
        sc = ws.cell(r, 6).value
        cat = ws.cell(r, 7).value or ''
        desc = ws.cell(r, 11).value or ''
        dur = get_duration(ws, r)
        groups.setdefault(ft, []).append({'sub': str(st or ''), 'sc': str(sc or ''), 'cat': str(cat), 'dur': dur, 'desc': str(desc)})

    if args.task:
        ft = args.task
        entries = groups.get(ft, [])
        if not entries:
            print(f"No entries for father task '{ft}'.")
            return 0
        total = sum(e['dur'] for e in entries)
        print(f"Father Task: {ft}")
        print(f"{'Sub Task':<20} {'Sub Task Code':<15} {'Category':<15} {'Hours':>8} {'Description':<30}")
        print("-" * 90)
        for e in entries:
            print(f"{e['sub']:<20} {e['sc']:<15} {e['cat']:<15} {e['dur']:>8.2f} {e['desc']:<30}")
        print("-" * 90)
        print(f"{'TOTAL':<20} {'':<15} {'':<15} {total:>8.2f}")
    else:
        if not groups:
            print("No time entries found.")
            return 0
        print(f"{'Father Task':<25} {'Entries':>8} {'Hours':>8} {'Last Date':<15}")
        print("-" * 58)
        for ft, entries in groups.items():
            total = sum(e['dur'] for e in entries)
            print(f"{ft:<25} {len(entries):>8} {total:>8.2f}")
        print("-" * 58)
        total_all = sum(sum(e['dur'] for e in entries) for entries in groups.values())
        print(f"{'ALL':<25} {'':>8} {total_all:>8.2f}")
    return 0


# ─── Subtask ─────────────────────────────────────────────────────────

def cmd_subtask(args):
    if not args.sub_cmd:
        print("subtask: list | rename | delete")
        print("  list              Show all subtask names")
        print("  rename -s OLD -n NEW  Rename subtask in all entries")
        print("  delete -s NAME    Clear subtask from matching entries")
        return
    return args.sub_func(args)


def cmd_sub_list(args):
    if not os.path.exists(WORKLOG_FILE):
        print(f"No {WORKLOG_FILE} found.")
        return 1
    wb = load_workbook(WORKLOG_FILE)
    ws = wb['Time Entries']
    seen = {}
    for r in range(2, ws.max_row + 1):
        ft = ws.cell(r, 4).value
        st = ws.cell(r, 5).value
        if not ft or not st:
            continue
        ft = str(ft).strip()
        st = str(st).strip()
        if args.task and args.task.lower() not in ft.lower():
            continue
        dur = get_duration(ws, r)
        seen.setdefault(st, {'father_tasks': set(), 'total_hours': 0.0})
        seen[st]['father_tasks'].add(ft)
        seen[st]['total_hours'] += dur
    if not seen:
        print("No subtasks found.")
        return 0
    print(f"{'Subtask':<25} {'Father Tasks':<35} {'Hours':>8}")
    print('-' * 70)
    for name, info in sorted(seen.items()):
        fts = ', '.join(sorted(info['father_tasks']))
        print(f"{name:<25} {fts:<35} {info['total_hours']:>8.2f}")
    return 0


def cmd_sub_rename(args):
    if not os.path.exists(WORKLOG_FILE):
        print(f"No {WORKLOG_FILE} found.")
        return 1
    wb = load_workbook(WORKLOG_FILE)
    ws = wb['Time Entries']
    count = 0
    for r in range(2, ws.max_row + 1):
        st = ws.cell(r, 5).value
        if st and str(st).strip() == args.subtask:
            ws.cell(r, 5).value = args.new_name
            count += 1
    wb.save(WORKLOG_FILE)
    print(f"Renamed {count} subtask entries '{args.subtask}' -> '{args.new_name}'.")
    return 0


def cmd_sub_delete(args):
    if not os.path.exists(WORKLOG_FILE):
        print(f"No {WORKLOG_FILE} found.")
        return 1
    wb = load_workbook(WORKLOG_FILE)
    ws = wb['Time Entries']
    count = 0
    for r in range(2, ws.max_row + 1):
        st = ws.cell(r, 5).value
        if st and str(st).strip() == args.subtask:
            ws.cell(r, 5).value = None
            count += 1
    wb.save(WORKLOG_FILE)
    print(f"Cleared {count} subtask entries matching '{args.subtask}'.")
    return 0


def countdown(seconds, rest=False):
    """Countdown timer with progress display."""
    bar_width = 30
    start = time.time()
    end = start + seconds
    symbol = ' ' if rest else '\u2588'
    while True:
        remaining = max(0, end - time.time())
        elapsed = seconds - remaining
        pct = elapsed / seconds if seconds > 0 else 1
        filled = int(pct * bar_width)
        bar = symbol * filled + '\u250a' + '\u00b7' * (bar_width - filled)
        mins = int(remaining) // 60
        secs = int(remaining) % 60
        label = 'REST' if rest else 'WORK'
        print(f"\r  {label} [{bar}] {mins:02d}:{secs:02d}  ", end='', flush=True)
        if remaining <= 0:
            print("\a")  # beep
            print()
            return
        time.sleep(0.5)


# ─── Main ────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description='Work Log Timer — CLI tool for WPS Office users',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__)
    sub = parser.add_subparsers(dest='cmd')

    p_start = sub.add_parser('start', help='Start timer')
    p_start.add_argument('-p', '--project', help='Project name (auto: git folder)')
    p_start.add_argument('-t', '--task', help='Father task (auto: git branch)')
    p_start.add_argument('-s', '--subtask', help='Sub task')
    p_start.add_argument('--subtask-code', help='Sub task code')
    p_start.add_argument('-c', '--category', help='Category (default: Development)')
    p_start.add_argument('-d', '--description', help='Description (auto: git commit msg)')

    p_stop = sub.add_parser('stop', help='Stop timer and save')
    p_stop.add_argument('-p', '--project', help='Override project')
    p_stop.add_argument('-t', '--task', help='Override father task')
    p_stop.add_argument('-s', '--subtask', help='Sub task')
    p_stop.add_argument('--subtask-code', help='Sub task code')
    p_stop.add_argument('-c', '--category', help='Override category')
    p_stop.add_argument('-d', '--description', help='Override description')

    sub.add_parser('cancel', help='Cancel running timer')
    sub.add_parser('status', help='Show timer status')
    p_pause = sub.add_parser('pause', help='Pause timer (break, meeting, lunch)')
    p_pause.add_argument('-r', '--reason', help='Pause reason: meeting, lunch, break, review, context-switch, etc.')
    sub.add_parser('continue', help='Resume paused timer')
    sub.add_parser('webhook-test', help='Send test webhook notification')
    sub.add_parser('today', help='Show today summary')
    sub.add_parser('week', help='Show this week summary')

    p_pomo = sub.add_parser('pomo', help='Pomodoro timer')
    p_pomo.add_argument('--work', type=int, default=25, help='Work minutes (default: 25)')
    p_pomo.add_argument('--rest', type=int, default=5, help='Rest minutes (default: 5)')

    # KPI subcommands
    p_kpi = sub.add_parser('kpi', help='Manage father tasks (KPIs)')
    ksub = p_kpi.add_subparsers(dest='kpi_cmd')
    p_kl = ksub.add_parser('list', help='List all father tasks with status')
    p_kl.set_defaults(kpi_func=cmd_kpi_list)
    p_ka = ksub.add_parser('add', help='Add a new father task')
    p_ka.add_argument('-t', '--name', required=True, help='Father task name')
    p_ka.add_argument('-p', '--project', help='Project name')
    p_ka.add_argument('-d', '--deadline', type=int, default=7, help='Deadline in days (default: 7)')
    p_ka.set_defaults(kpi_func=cmd_kpi_add)
    p_kd = ksub.add_parser('done', help='Mark a father task as done')
    p_kd.add_argument('-t', '--name', required=True, help='Father task name')
    p_kd.set_defaults(kpi_func=cmd_kpi_done)
    p_ke = ksub.add_parser('edit', help='Edit father task project/deadline')
    p_ke.add_argument('-t', '--name', required=True, help='Father task name')
    p_ke.add_argument('-p', '--project', help='New project name')
    p_ke.add_argument('-d', '--deadline', type=int, help='New deadline in days')
    p_ke.set_defaults(kpi_func=cmd_kpi_edit)
    p_kr = ksub.add_parser('rename', help='Rename father task')
    p_kr.add_argument('-t', '--name', required=True, help='Current father task name')
    p_kr.add_argument('-n', '--new-name', help='New father task name')
    p_kr.set_defaults(kpi_func=cmd_kpi_rename)
    p_kdel = ksub.add_parser('delete', help='Delete a father task')
    p_kdel.add_argument('-t', '--name', required=True, help='Father task name to delete')
    p_kdel.set_defaults(kpi_func=cmd_kpi_delete)
    p_ks = ksub.add_parser('status', help='Show overall task progress')
    p_ks.set_defaults(kpi_func=cmd_kpi_status)

    # Subtask subcommands
    p_sub = sub.add_parser('subtask', help='Manage subtasks in Time Entries')
    ssub = p_sub.add_subparsers(dest='sub_cmd')
    p_sl = ssub.add_parser('list', help='List all subtasks')
    p_sl.add_argument('-t', '--task', help='Filter by father task')
    p_sl.set_defaults(sub_func=cmd_sub_list)
    p_sr = ssub.add_parser('rename', help='Rename subtask in all entries')
    p_sr.add_argument('-s', '--subtask', required=True, help='Current subtask name')
    p_sr.add_argument('-n', '--new-name', required=True, help='New subtask name')
    p_sr.set_defaults(sub_func=cmd_sub_rename)
    p_sdel = ssub.add_parser('delete', help='Clear subtask from all matching entries')
    p_sdel.add_argument('-s', '--subtask', required=True, help='Subtask name to clear')
    p_sdel.set_defaults(sub_func=cmd_sub_delete)

    p_tasks = sub.add_parser('tasks', help='Show tasks grouped by father task')
    p_tasks.add_argument('-t', '--task', help='Filter by father task name')

    args = parser.parse_args()

    if not args.cmd:
        parser.print_help()
        return 1

    cmds = {
        'start': cmd_start,
        'stop': cmd_stop,
        'cancel': cmd_cancel,
        'status': cmd_status,
        'pause': cmd_pause,
        'continue': cmd_continue,
        'today': cmd_today,
        'week': cmd_week,
        'pomo': cmd_pomo,
        'kpi': cmd_kpi,
        'tasks': cmd_tasks,
        'subtask': cmd_subtask,
        'webhook-test': webhook.send_test,
    }
    return cmds[args.cmd](args)


if __name__ == '__main__':
    sys.exit(main())
