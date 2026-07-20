# Worklog Auto-formulas Design

## Overview
Transform the existing worklog.xlsx so users only edit in "Time Entries", while other sheets auto-populate using Excel 365 dynamic array formulas. Add KPI tracking and monthly performance analysis.

## Sheets

### 0. Data Normalization (helper columns in Time Entries)
Because merged cells only hold value in the top cell, formulas in other sheets can't read them.
**Hidden columns J-K** auto-propagate values downward for reliable formula references.

| Column (hidden) | Header | Formula |
|-----------------|--------|---------|
| J | Project (norm) | `=IF(C{n}<>"", C{n}, J{n-1})` — carries forward last non-blank Project |
| K | Task (norm) | `=IF(D{n}<>"", D{n}, K{n-1})` — carries forward last non-blank Task |

All other sheets reference J and K instead of C and D.

### 1. Time Entries (minimal change)
- Columns: A-I same as before (Date, Day, Project, Task, Category, Start Time, End Time, Duration, Description)
- J-K: hidden helper columns (auto-filled, user never edits)
- Duration formula: =(G-F)*24 (existing)
- **Father tasks**: any task name the user adds to the KPIs sheet

### 2. KPIs (new sheet)
Manual entry for tracking father task deadlines. User adds any "big task" here.

| Column | Header | Type | Notes |
|--------|--------|------|-------|
| A | Father Task | Text | Task name (must match Time Entries Task column) |
| B | Project | Text | Project name |
| C | Date | Date | Date of father task |
| D | Deadline (days) | Number | Working days allowed (e.g. 5 = 5 working days) |
| E | Deadline (h) | Formula | =D*7.5 (auto-calc, 7.5h per day) |
| F | Status | Dropdown | Not Started / In Progress / Done / Late |
| G | Completed Date | Date | Actual completion date |
| H | Notes | Text | Optional |

### 3. Weekly Summary (revamped — fully auto-populated)
**Auto-detects latest week** from Time Entries via MAX date.
**Grouped by Project with subtotals** using REDUCE/VSTACK/LAMBDA.

Layout:
```
Row 1: Week label (e.g. "Week 27 (Jun 29 - Jul 05, 2026)")
Row 2: Headers: Task / Project | Category | Mon DD | Tue DD | Wed DD | Thu DD | Fri DD | Sat DD | Sun DD | Weekly Total
Rows 3+: Auto-generated from Time Entries

Project Header (bold, no data)
  Task 1               Cat     2.0   0     1.5   ...
  Task 2               Cat     1.0   0.5   0     ...
Project Subtotal               3.0   0.5   1.5   ...
                              ...next project...
GRAND TOTAL                    xx    xx    xx    xx
```

Key formulas:
- Latest date: `=MAX('Time Entries'!A:A)`
- Week start (Mon): `=latestDate - WEEKDAY(latestDate, 2) + 1`
- Unique tasks in week: `=UNIQUE(FILTER('Time Entries'!K:K, (dateCol>=ws)*(dateCol<=we)*(taskCol<>"")))`  (uses normalized Task column K)
- Daily hours: `=SUMIFS('Time Entries'!H:H, 'Time Entries'!K:K, taskName, 'Time Entries'!A:A, date)` (uses normalized Task column K)

### 4. Daily Detail (revamped — fully auto-populated)
**Auto-shows latest day** or user can change date by editing cell A1.

Layout:
```
Cell A1: date input (default: =MAX('Time Entries'!A:A) — latest day)
Row 2: "Daily Detail — [formatted date from A1]"
Row 3: Headers: Time | Project | Task | Category | Duration | Description
Rows 4+: =FILTER('Time Entries'!F:K, 'Time Entries'!A:A = A1) — shows Start-End time, Project, Task, Category, Duration, Description
...
Summary row:
  Total Hours: =SUM( filteredDuration )
  % Working Day: =TotalHours / 7.5 * 100
```

### 5. Monthly (new — performance sheet)
**Auto-populated** from Time Entries + KPIs.
**Logic**: Only shows tasks that exist in the KPIs sheet (those are "father tasks" by definition).
**Month**: Auto-detected from latest Time Entry date: `=MAX('Time Entries'!A:A)`. Extracts year/month via `YEAR()` and `MONTH()`.
**Layout**: Cell A1 = month label, A2+ = table via dynamic array formulas.

```
Cell A1: Month label (e.g. "July 2026") derived from latest Time Entry date.
Cell A3: =LET(...) dynamic array that filters father tasks from KPIs matching this month.
```

| Father Task | Project | Total Hours | Deadline (h) | % Performance | Status | Completed Date |
|-------------|---------|------------|--------------|---------------|--------|----------------|
| Task A | Proj X | 20 | 30 | 150% | ✓ Done | Jul 10 |

- **Father Task list**: `=FILTER(KPIs!A:A, (YEAR('Time Entries' date)=monthYear)*(MONTH('Time Entries' date)=monthMonth))` — but simplified: filter KPIs rows where the father task exists in Time Entries within the month
- **Total Hours**: `=SUMIFS('Time Entries'!H:H, 'Time Entries'!K:K, taskName, 'Time Entries'!A:A, ">="&monthStart, 'Time Entries'!A:A, "<="&monthEnd)` — sums all entries matching the task name in the selected month (uses normalized Task column K + date range)
- **Deadline (h)**: `=XLOOKUP(taskName, KPIs!A:A, KPIs!E:E)` from KPIs sheet
- **% Performance**: `=IF(TotalHours=0, "N/A", DeadlineHours/TotalHours*100)` (>100% = ahead; exactly 100% = on time; <100% = behind)
- **Status**: =IF(deadlineMet, "✓ On Time", "✗ Late") based on Completed Date vs deadline
- **Bottom section**: Monthly summary rows (total hours, avg performance %, count of on-time vs late tasks)

## Performance Calculation Rules
- 1 working day = 7.5 hours (7h30)
- 1 calendar day = 24h
- Performance = Assigned Deadline (h) ÷ Actual Hours × 100%
- If Performance ≥ 100% → On Time / Ahead
- If Performance < 100% → Behind / Late

## Auto-population Logic
- All sheets reference Time Entries only (read-only for auto-filled cells)
- Cells that require input are clearly marked (white background)
- Auto-filled cells use formulas (light gray background)
- Excel 365 dynamic arrays (FILTER, UNIQUE, SUMIFS, XLOOKUP, LET, VSTACK, REDUCE)
