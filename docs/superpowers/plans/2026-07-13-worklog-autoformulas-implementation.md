# Worklog Auto-formulas Implementation Plan

> **Note:** This is an Excel workbook modification, not a code project. Implementation uses openpyxl to write formulas into the xlsx.

**Goal:** Transform worklog.xlsx so only "Time Entries" is editable; all other sheets auto-populate with Excel 365 formulas.

**Approach:** Python script with openpyxl builds all formulas. No VBA. Uses Excel 365 dynamic array functions (UNIQUE, FILTER, SUMIFS, XLOOKUP, LET, VSTACK).

**Tech Stack:** Python 3.10+, openpyxl, Excel 365

---

### Task 1: Add helper columns J-K to Time Entries

**Files:**
- Modify: `D:\WorkLog\worklog.xlsx`

- [ ] **Step 1: Write Python script stub**

Create script `D:\WorkLog\build_workbook.py` to open and modify the workbook.

```python
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from copy import copy

wb = openpyxl.load_workbook(r'D:\WorkLog\worklog.xlsx')
```

- [ ] **Step 2: Add helper column formulas to Time Entries**

J2 (Project, propagated): `=IF(C2<>"",C2,J1)`
K2 (Task, propagated): `=IF(D2<>"",D2,K1)`

Fill down to row 100 (pre-filled for future entries).

Hide columns J-K (set hidden=True on column dimension).

Style: light gray fill, narrow italic font to distinguish from user input.

- [ ] **Step 3: Verify helper columns**

Run script, open xlsx, check that J and K propagate values correctly.

---

### Task 2: Create KPIs sheet

**Files:**
- Modify: `D:\WorkLog\worklog.xlsx`

- [ ] **Step 4: Create KPIs sheet**

Headers in Row 1:
- A1: "Father Task"
- B1: "Project"
- C1: "Date"
- D1: "Deadline (days)"
- E1: "Deadline (h)" — formula: `=IF(D2<>"",D2*7.5,"")`
- F1: "Status"
- G1: "Completed Date"
- H1: "Notes"

Style: Blue header fill (#4472C4), white bold font.

Set column widths: A=30, B=20, C=12, D=14, E=14, F=14, G=14, H=30

Add data validation for Status column (F): dropdown with "Not Started", "In Progress", "Done", "Late"

Add a sample instruction row:
- A2: "Enter father tasks here. Task name must match Time Entries exactly."

---

### Task 3: Rewrite Weekly Summary with auto-formulas

**Files:**
- Modify: `D:\WorkLog\worklog.xlsx`

- [ ] **Step 5: Clear old data and set up week detection**

Clear all existing data in Weekly Summary (keep headers as template, replace with formulas).

Row 1: Week info
- A1: "Week {weeknum} ({start_date} - {end_date}, {year})" — manually set by formula or label
- B1: hidden reference cell for latest date: `=MAX('Time Entries'!A:A)`
- C1: week start (Mon): `=B1-WEEKDAY(B1,2)+1`

Row 2: Day headers with dates
- A2: "Task"
- B2: "Category"
- C2: formula for Monday date: `=$C$1` formatted as "ddd mm/dd"
- D2: formula: `=$C$1+1`
- E2: `=$C$1+2`
- F2: `=$C$1+3`
- G2: `=$C$1+4`
- H2: `=$C$1+5`
- I2: `=$C$1+6`
- J2: "Weekly Total"

- [ ] **Step 6: Auto-populate task list**

A4 (spills down): 
```
=SORT(UNIQUE(FILTER('Time Entries'!J:J&": "&'Time Entries'!K:K,
  ('Time Entries'!A:A>=$C$1)*('Time Entries'!A:A<=$C$1+6)*
  ('Time Entries'!K:K<>""))))
```

This creates "Project: Task" list sorted alphabetically.

- [ ] **Step 7: Hours per day formulas**

B4 (Category): `=XLOOKUP(TRIM(MID(A4,FIND(": ",A4)+2,99)), 'Time Entries'!K:K, 'Time Entries'!E:E,,0)`

C4 (Mon hours): `=SUMIFS('Time Entries'!H:H, 'Time Entries'!J:J&": "&'Time Entries'!K:K, A4#, 'Time Entries'!A:A, $C$1)`

D4 (Tue): `=SUMIFS('Time Entries'!H:H, 'Time Entries'!J:J&": "&'Time Entries'!K:K, A4#, 'Time Entries'!A:A, $C$1+1)`

...repeat for each day (Wed-Sun: E4, F4, G4, H4, I4)...

J4 (Weekly Total): `=SUM(C4#:I4#)` — sums across all days for each task row

Fill C4:J4 down enough rows.

- [ ] **Step 8: Create project subtotals section**

Below the task list, add a project subtotals area. Use formulas to reference unique projects:

A{last+2}: "PROJECT TOTALS" (bold header)

For each unique project from the week:
```
=SORT(UNIQUE(FILTER('Time Entries'!J:J,
  ('Time Entries'!A:A>=$C$1)*('Time Entries'!A:A<=$C$1+6)*
  ('Time Entries'!J:J<>""))))
```

Project total hours per day: `=SUMIFS('Time Entries'!H:H, 'Time Entries'!J:J, projectName, 'Time Entries'!A:A, date)`

- [ ] **Step 9: Grand total row**

Last row: "GRAND TOTAL" with SUM of all subtotals.

---

### Task 4: Rewrite Daily Detail with formulas

**Files:**
- Modify: `D:\WorkLog\worklog.xlsx`

- [ ] **Step 10: Set up date selector and FILTER formula**

Clear old data.

A1: "Daily Detail"
C1: Date selector — `=MAX('Time Entries'!A:A)` (latest day). User can override.

Row 2: "Showing entries for: [date from C1]"

Row 3: Headers
- A3: "Time Range"
- B3: "Project"
- C3: "Task"
- D3: "Category"
- E3: "Duration (h)"
- F3: "Description"

Row 4 (spills down):
```
=FILTER('Time Entries'!F:G&"-"&'Time Entries'!G:G, 'Time Entries'!A:A=$C$1, "No entries")
```

Wait, that doesn't work cleanly. Better to use multiple FILTER columns:

A4: `=FILTER('Time Entries'!F:F&" - "&'Time Entries'!G:G, 'Time Entries'!A:A=$C$1, "")` — Time Range as "HH:MM - HH:MM"
B4: `=FILTER('Time Entries'!J:J, 'Time Entries'!A:A=$C$1, "")` — Project (normalized)
C4: `=FILTER('Time Entries'!K:K, 'Time Entries'!A:A=$C$1, "")` — Task (normalized)
D4: `=FILTER('Time Entries'!E:E, 'Time Entries'!A:A=$C$1, "")` — Category
E4: `=FILTER('Time Entries'!H:H, 'Time Entries'!A:A=$C$1, "")` — Duration
F4: `=FILTER('Time Entries'!I:I, 'Time Entries'!A:A=$C$1, "")` — Description

- [ ] **Step 11: Add day summary section**

Below the filtered data, add summary:

```
Total Hours: =SUM(E4#)
% Working Day: =totalHours / 7.5 * 100  (formatted as %)
Status: =IF(% >= 100%, "✓ Full day", "✗ Under: " & (7.5 - totalHours) & "h remaining")
```

---

### Task 5: Create Monthly performance sheet

**Files:**
- Modify: `D:\WorkLog\worklog.xlsx`

- [ ] **Step 12: Set up month detection**

A1: "Month: [auto-detected]"
B1: Latest date reference: `=MAX('Time Entries'!A:A)`
C1: Month start: `=DATE(YEAR(B1),MONTH(B1),1)`
D1: Month end: `=EOMONTH(B1,0)`

- [ ] **Step 13: Father task list from KPIs**

Row 2: Headers
- A2: "Father Task"
- B2: "Project"
- C2: "Total Hours"
- D2: "Deadline (h)"
- E2: "% Performance"
- F2: "Status"
- G2: "Completed Date"

A3 (spills down): 
```
=FILTER(KPIs!A:A, 
  (KPIs!A:A<>"")*(KPIs!A:A<>"Father Task"),
  "No KPIs set")
```
Lists all father tasks from KPIs sheet.

- [ ] **Step 14: Performance calculation formulas**

B3 (Project): `=XLOOKUP(A3#, KPIs!A:A, KPIs!B:B)`
C3 (Total Hours): `=SUMIFS('Time Entries'!H:H, 'Time Entries'!K:K, A3#, 'Time Entries'!A:A, ">="&$C$1, 'Time Entries'!A:A, "<="&$D$1)`
D3 (Deadline h): `=XLOOKUP(A3#, KPIs!A:A, KPIs!E:E)`
E3 (% Performance): `=IF(C3#=0, "N/A", D3#/C3#*100)` — formatted as "0.0%"
F3 (Status): `=XLOOKUP(A3#, KPIs!A:A, KPIs!F:F)`
G3 (Completed Date): `=XLOOKUP(A3#, KPIs!A:A, KPIs!G:G)`

- [ ] **Step 15: Monthly summary section**

Below the task table:

```
Total hours this month: =SUM(C3#)
Avg Performance: =AVERAGE(E3#)  (excluding N/A)
Tasks On Time: =COUNTIF(F3#, "Done")
Tasks Late: =COUNTIF(F3#, "Late")
Overall Status: =IF(COUNTIF(F3#,"Late")>0, "⚠ Needs attention", "✓ All on track")
```

---

### Task 6: Formatting and polish

**Files:**
- Modify: `D:\WorkLog\worklog.xlsx`

- [ ] **Step 16: Apply formatting**

- Weekly Summary: Blue header (#4472C4), alternating row colors for readability
- Daily Detail: Same blue header style
- Monthly: Blue header, green/yellow/red conditional formatting for Performance column
- KPIs: Blue header
- All formula cells: light gray fill to distinguish from input cells
- Hide gridlines on non-input sheets where appropriate
- Freeze panes on all sheets for column headers

- [ ] **Step 17: Final verification**

Run script. Open the xlsx in Excel 365. Verify:
1. Helper columns J-K propagate correctly
2. Weekly Summary shows tasks from the latest week
3. Daily Detail shows entries for the selected date
4. Monthly shows KPIs and calculates performance
5. All formulas resolve correctly (no #REF!, #N/A, etc.)
