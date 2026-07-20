# Graph Report - WorkLog  (2026-07-17)

## Corpus Check
- 13 files · ~9,244 words
- Verdict: corpus is large enough that graph structure adds value.

## Summary
- 121 nodes · 148 edges · 14 communities (12 shown, 2 thin omitted)
- Extraction: 100% EXTRACTED · 0% INFERRED · 0% AMBIGUOUS
- Token cost: 0 input · 0 output

## Community Hubs (Navigation)
- [[_COMMUNITY_Community 0|Community 0]]
- [[_COMMUNITY_Community 1|Community 1]]
- [[_COMMUNITY_Community 2|Community 2]]
- [[_COMMUNITY_Community 3|Community 3]]
- [[_COMMUNITY_Community 4|Community 4]]
- [[_COMMUNITY_Community 5|Community 5]]
- [[_COMMUNITY_Community 6|Community 6]]
- [[_COMMUNITY_Community 7|Community 7]]
- [[_COMMUNITY_Community 8|Community 8]]
- [[_COMMUNITY_Community 9|Community 9]]

## God Nodes (most connected - your core abstractions)
1. `Worklog.xlsx — Hướng dẫn sử dụng` - 13 edges
2. `run()` - 9 edges
3. `load_state()` - 7 edges
4. `cmd_start()` - 7 edges
5. `cmd_stop()` - 7 edges
6. `Worklog Auto-formulas Implementation Plan` - 7 edges
7. `Sheets` - 7 edges
8. `get_elapsed()` - 6 edges
9. `run()` - 6 edges
10. `append_entry()` - 5 edges

## Surprising Connections (you probably didn't know these)
- None detected - all connections are within the same source files.

## Communities (14 total, 2 thin omitted)

### Community 0 - "Community 0"
Cohesion: 0.14
Nodes (26): append_entry(), clear_state(), cmd_cancel(), cmd_continue(), cmd_pause(), cmd_pomo(), cmd_start(), cmd_status() (+18 more)

### Community 1 - "Community 1"
Cohesion: 0.11
Nodes (17): 5 Bảng Tính, Bước 1: Thiết lập KPIs (làm một lần), Bước 2: Ghi nhật ký trong Time Entries, Bước 3: Xem Weekly Summary, Bước 4: Xem Daily Detail, Bước 5: Monthly Performance, code:bash (# Kiểm tra dữ liệu hợp lệ), code:bash (# Xây lại workbook + seed dữ liệu mẫu) (+9 more)

### Community 2 - "Community 2"
Cohesion: 0.13
Nodes (14): code:python (import openpyxl), code:block2 (=SORT(UNIQUE(FILTER('Time Entries'!J:J&": "&'Time Entries'!K), code:block3 (=SORT(UNIQUE(FILTER('Time Entries'!J:J,), code:block4 (=FILTER('Time Entries'!F:G&"-"&'Time Entries'!G:G, 'Time Ent), code:block5 (Total Hours: =SUM(E4#)), code:block6 (=FILTER(KPIs!A:A,), code:block7 (Total hours this month: =SUM(C3#)), Task 1: Add helper columns J-K to Time Entries (+6 more)

### Community 3 - "Community 3"
Cohesion: 0.13
Nodes (14): 0. Data Normalization (helper columns in Time Entries), 1. Time Entries (minimal change), 2. KPIs (new sheet), 3. Weekly Summary (revamped — fully auto-populated), 4. Daily Detail (revamped — fully auto-populated), 5. Monthly (new — performance sheet), Auto-population Logic, code:block1 (Row 1: Week label (e.g. "Week 27 (Jun 29 - Jul 05, 2026)")) (+6 more)

### Community 4 - "Community 4"
Cohesion: 0.38
Nodes (9): compute_consistency(), compute_daily_summary(), compute_scores(), compute_task_performance(), get_time_entries(), load_gamify_data(), run(), save_gamify_data() (+1 more)

### Community 5 - "Community 5"
Cohesion: 0.25
Nodes (8): Auto-detect, code:bash (# Bắt đầu làm việc (dùng thư mục git hiện tại)), code:bash (# Hôm nay đã làm gì?), code:bash (# 25 phút làm + 5 phút nghỉ (mặc định)), Lệnh cơ bản, Pomodoro, Sử dụng CLI Timer (khuyên dùng thay vì gõ tay), Xem báo cáo nhanh

### Community 6 - "Community 6"
Cohesion: 0.25
Nodes (7): code:bash (# Build workbook + seed data mẫu), code:block2 (WorkLog/), Cấu trúc thư mục, Quick Start, Tính năng, WorkLog, Yêu cầu

### Community 7 - "Community 7"
Cohesion: 0.52
Nodes (6): empty_rows(), load_ws(), missing_categories(), missing_descriptions(), run(), utilization_gaps()

## Knowledge Gaps
- **44 isolated node(s):** `Get total elapsed seconds from state. Works with paused or running state.`, `Find first empty row in Time Entries (col A check).`, `Append one row to Time Entries sheet and save.`, `Get numeric duration from row. Computes from start/end if formula not evaluated.`, `Countdown timer with progress display.` (+39 more)
  These have ≤1 connection - possible missing edges or undocumented components.
- **2 thin communities (<3 nodes) omitted from report** — run `graphify query` to explore isolated nodes.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **Why does `Worklog.xlsx — Hướng dẫn sử dụng` connect `Community 1` to `Community 5`?**
  _High betweenness centrality (0.037) - this node is a cross-community bridge._
- **Why does `Sử dụng CLI Timer (khuyên dùng thay vì gõ tay)` connect `Community 5` to `Community 1`?**
  _High betweenness centrality (0.020) - this node is a cross-community bridge._
- **What connects `Get total elapsed seconds from state. Works with paused or running state.`, `Find first empty row in Time Entries (col A check).`, `Append one row to Time Entries sheet and save.` to the rest of the system?**
  _44 weakly-connected nodes found - possible documentation gaps or missing edges._
- **Should `Community 0` be split into smaller, more focused modules?**
  _Cohesion score 0.14 - nodes in this community are weakly interconnected._
- **Should `Community 1` be split into smaller, more focused modules?**
  _Cohesion score 0.11 - nodes in this community are weakly interconnected._
- **Should `Community 2` be split into smaller, more focused modules?**
  _Cohesion score 0.13 - nodes in this community are weakly interconnected._
- **Should `Community 3` be split into smaller, more focused modules?**
  _Cohesion score 0.13 - nodes in this community are weakly interconnected._