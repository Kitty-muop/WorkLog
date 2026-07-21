"""SQLite backend for WorkLog. Mirrors Excel Time Entries and KPIs."""
import os
import sqlite3
import datetime
from pathlib import Path

DB_PATH = Path.home() / '.worklog.db'
ENABLED = os.environ.get('WORKLOG_DB', '').lower() == 'sqlite'


def get_conn():
    return sqlite3.connect(str(DB_PATH))


def init():
    conn = get_conn()
    conn.execute('''
        CREATE TABLE IF NOT EXISTS time_entries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            day TEXT NOT NULL,
            project TEXT,
            father_task TEXT,
            sub_task TEXT,
            sub_task_code TEXT,
            category TEXT DEFAULT 'Work',
            start_time TEXT NOT NULL,
            end_time TEXT NOT NULL,
            duration REAL,
            description TEXT,
            break_info TEXT DEFAULT '',
            created_at TEXT DEFAULT (datetime('now'))
        )
    ''')
    conn.commit()
    conn.close()


def insert_entry(project, task, subtask, subtask_code, category, description,
                 start_dt, end_dt, break_info=''):
    if not ENABLED:
        return
    init()
    conn = get_conn()
    duration = max(0, (end_dt - start_dt).total_seconds() / 3600)
    day_names = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
    conn.execute('''
        INSERT INTO time_entries
            (date, day, project, father_task, sub_task, sub_task_code,
             category, start_time, end_time, duration, description, break_info)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        start_dt.date().isoformat(),
        day_names[start_dt.weekday()],
        project or '', task or '', subtask or '', subtask_code or '',
        category or 'Work',
        start_dt.time().strftime('%H:%M'),
        end_dt.time().strftime('%H:%M'),
        round(duration, 2),
        description or '',
        break_info or ''
    ))
    conn.commit()
    conn.close()


def get_entries(date_from=None, date_to=None):
    if not ENABLED:
        return []
    init()
    conn = get_conn()
    query = 'SELECT * FROM time_entries'
    params = []
    if date_from and date_to:
        query += ' WHERE date >= ? AND date <= ?'
        params = [date_from.isoformat(), date_to.isoformat()]
    query += ' ORDER BY date, start_time'
    rows = conn.execute(query, params).fetchall()
    conn.close()
    return rows
