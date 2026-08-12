#!/usr/bin/env python3
"""Google Sheets sync module for WorkLog."""

import json
import os
import sys
from pathlib import Path

SHEET_ID_VAR = 'WORKLOG_GSHEETS_ID'
KEY_PATH_VAR = 'WORKLOG_GSHEETS_KEY'
_worksheet_cache = {}


def _get_service():
    try:
        from google.oauth2.service_account import Credentials
        import gspread
    except ImportError:
        return None
    key_path = os.environ.get(KEY_PATH_VAR)
    sheet_id = os.environ.get(SHEET_ID_VAR)
    if not key_path or not sheet_id:
        return None
    key_path = Path(key_path).expanduser()
    if not key_path.exists():
        return None
    creds = Credentials.from_service_account_file(
        str(key_path),
        scopes=['https://www.googleapis.com/auth/spreadsheets']
    )
    gc = gspread.authorize(creds)
    return gc.open_by_key(sheet_id)


def _ws(name):
    if name not in _worksheet_cache:
        svc = _get_service()
        if not svc:
            return None
        _worksheet_cache[name] = svc.worksheet(name)
    return _worksheet_cache[name]


def is_enabled():
    return bool(os.environ.get(SHEET_ID_VAR) and os.environ.get(KEY_PATH_VAR))


def sync_timer_start(project, task, subtask, subtask_code, category, description):
    """Log timer start to Google Sheets with zero duration"""
    if not is_enabled():
        return False
    from datetime import datetime, date
    ws = _ws('Time Entries')
    if not ws:
        return False
    now = datetime.now()
    d = date.today()
    start_time = now.strftime('%H:%M')
    # End time same as start for zero duration entry
    end_time = start_time
    duration_h = 0.0
    row = [d.strftime('%d-%m-%Y'), d.strftime('%a'), project, task,
           subtask or '', subtask_code or '', category,
           start_time, end_time, f'{duration_h:.2f}', description or '', '']
    ws.append_row(row, value_input_option='USER_ENTERED')
    return True


def sync_time_entry(project, task, subtask, subtask_code, category, description,
                    start_time, end_time, duration_h, break_info='', date_val=None):
    if not is_enabled():
        return False
    from datetime import date
    ws = _ws('Time Entries')
    if not ws:
        return False
    d = date_val or date.today()
    row = [d.strftime('%d-%m-%Y'), d.strftime('%a'), project, task,
           subtask or '', subtask_code or '', category,
           start_time, end_time, f'{duration_h:.2f}', description or '', break_info or '']
    ws.append_row(row, value_input_option='USER_ENTERED')
    return True


def sync_kpi_add(code, name, project, deadline_days, status='In Progress'):
    if not is_enabled():
        return False
    ws = _ws('KPIs')
    if not ws:
        return False
    from datetime import date
    today = date.today()
    row = [code, name, project, today.strftime('%d-%m-%Y'),
           str(deadline_days) if deadline_days else '', f'={deadline_days}*7.5' if deadline_days else '',
           status, '', '']
    ws.append_row(row, value_input_option='USER_ENTERED')
    return True


def sync_kpi_update(name, field, value):
    """Update a KPI field by matching name in column B."""
    if not is_enabled():
        return False
    ws = _ws('KPIs')
    if not ws:
        return False
    col_map = {'status': 7, 'project': 3, 'deadline_days': 5, 'completed_date': 8}
    col = col_map.get(field)
    if not col:
        return False
    cell = ws.find(name, in_column=2)
    if not cell:
        return False
    ws.update_cell(cell.row, col, value)
    return True


def sync_kpi_delete(name):
    """Delete a KPI row by matching name."""
    if not is_enabled():
        return False
    ws = _ws('KPIs')
    if not ws:
        return False
    cell = ws.find(name, in_column=2)
    if not cell:
        return False
    ws.delete_rows(cell.row)
    return True


def sync_kpi_rename(old_name, new_name):
    """Rename KPI in all rows that reference it."""
    if not is_enabled():
        return False
    ws = _ws('KPIs')
    if not ws:
        return False
    cell = ws.find(old_name, in_column=2)
    if cell:
        ws.update_cell(cell.row, 2, new_name)
    return True


def sync_project_add(code, name, description='', status='Active'):
    if not is_enabled():
        return False
    ws = _ws('Projects')
    if not ws:
        return False
    from datetime import date
    row = [code, name, description or '', status, date.today().strftime('%d-%m-%Y')]
    ws.append_row(row, value_input_option='USER_ENTERED')
    return True


def sync_project_rename_propagate(old_name, new_name):
    """Update project name references in KPIs (col 3) and Time Entries (col 3)."""
    if not is_enabled():
        return False
    ws_kpi = _ws('KPIs')
    if ws_kpi:
        cells = ws_kpi.findall(old_name, in_column=3)
        for cell in cells:
            ws_kpi.update_cell(cell.row, 3, new_name)
    ws_te = _ws('Time Entries')
    if ws_te:
        cells = ws_te.findall(old_name, in_column=3)
        for cell in cells:
            ws_te.update_cell(cell.row, 3, new_name)
    return True


def sync_kpi_rename_propagate(old_name, new_name):
    """Update KPI name references in Time Entries (col 4)."""
    if not is_enabled():
        return False
    ws = _ws('Time Entries')
    if not ws:
        return False
    cells = ws.findall(old_name, in_column=4)
    for cell in cells:
        ws.update_cell(cell.row, 4, new_name)
    return True


def sync_subtask_rename(old_name, new_name):
    """Rename subtask name in all Time Entries rows (col 5)."""
    if not is_enabled():
        return False
    ws = _ws('Time Entries')
    if not ws:
        return False
    cells = ws.findall(old_name, in_column=5)
    for cell in cells:
        ws.update_cell(cell.row, 5, new_name)
    return True


def sync_subtask_delete(name):
    """Clear subtask name in all Time Entries rows (col 5)."""
    if not is_enabled():
        return False
    ws = _ws('Time Entries')
    if not ws:
        return False
    cells = ws.findall(name, in_column=5)
    for cell in cells:
        ws.update_cell(cell.row, 5, '')
    return True


def sync_subtask_add(code, name, father_task, project, status='In Progress'):
    if not is_enabled():
        return False
    ws = _ws('SubTasks')
    if not ws:
        return False
    from datetime import date
    today = date.today()
    row = [code, name, father_task, project, status,
           today.strftime('%d-%m-%Y'), '', '']
    ws.append_row(row, value_input_option='USER_ENTERED')
    return True


def sync_subtask_update(name, field, value):
    if not is_enabled():
        return False
    ws = _ws('SubTasks')
    if not ws:
        return False
    col_map = {'status': 5, 'father_task': 3, 'project': 4, 'completed_date': 7}
    col = col_map.get(field)
    if not col:
        return False
    cell = ws.find(name, in_column=2)
    if not cell:
        return False
    ws.update_cell(cell.row, col, value)
    return True


def sync_project_update(name, field, value):
    if not is_enabled():
        return False
    ws = _ws('Projects')
    if not ws:
        return False
    col_map = {'name': 2, 'status': 4, 'description': 3}
    col = col_map.get(field)
    if not col:
        return False
    if field == 'name':
        cell = ws.find(name, in_column=2)
    else:
        cell = ws.find(name, in_column=2)
    if not cell:
        return False
    ws.update_cell(cell.row, col, value)
    return True