# Changelog

All notable changes to the WorkLog RPG Quest System will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.3.0] - 2026-08-12

### Added
- **Estimate vs Actual Auto-Calculation**: Added optional `estimate` parameter (default 7.5h) to `/start`, `/subtask start`, `/subtask add`, and `/kpi add`. Automatically sums total logged actual hours upon timer stop (`/stop`, `/sub done`, `/kpi done`) and syncs to Excel & Google Sheets.
- **Dynamic EXP & 5 Level Milestones**: Dynamic EXP formula based on $\frac{\text{Actual}}{\text{Estimate}}$ ratio, +30% on-time Debug bonus & heavy quadratic overtime penalty, mapped across 101 thresholds to 5 Level Milestones (`Novice`, `Adventurer`, `Expert`, `Master`, `Grandmaster Legend`).
- **Automated Discord Reminders Scheduler**: Off-hours and daily quest reminders (08:30 morning quest start, 12:00 lunch campfire rest, 13:30 afternoon resume, 17:30 workday summary) with 1-fire per day deduplication lock.
- **5th Google Sheets Worksheet (`Gamify Summary`)**: Automatically syncs Hero Rank, EXP score, level (0..100), streak, and consistency stats to Google Sheets.
- **Expanded Unit Test Suite**: `14/14 tests PASSED` (`0.060s`) verifying reminder deduplication, estimate parsing, and actual hour summation.

### Changed
- Standardized category presets with autocomplete suggestions (`Development`, `Debug / Bug Fix`, `Refactoring`, `Code Review`, `Testing / QA`, `DevOps / CI-CD`, `Documentation`).

## [1.2.0] - 2026-08-12

### Added
- **Gamified RPG Quest System**: Overhauled 13 slash commands with unified RPG Quest UI cards (`🏰 Realm`, `📜 Main Quest`, `🧩 Side Quest`, `⚔️ Quest Battles`, `🛡️ Hero Stats`).
- **Strict Multi-User Isolation**: Every timer, summary, and status request is bound strictly to unique `Discord User ID` (`interaction.user.id`).
- **Ephemeral Private Mode**: All slash command responses return private ephemeral messages (`ephemeral=True`) visible only to the invoking user.
- **Off-Hours Safety Catch-Up**: Automatic pause guard during lunch breaks (12:00-13:30) and night off (17:30-08:30).
- **Automated Unit Test Suite**: `tests/test_timer_user_isolation.py` protecting timer user isolation and preventing cross-user action leaks.
- **Silent OS Auto-Start**: Background VBScript launcher (`tools/run_bot.vbs`) integrated with Windows Startup folder.

### Changed
- Capped slash command autocompletes (`_ac_proj`, `_ac_task`, `_ac_sub`) strictly to 5 choices sorted newest first (`reversed(...)`).
- Standardized Google Sheets dates to ISO `YYYY-MM-DD` format.
- Upgraded `/stop` to mark subtask as `Hung` and `/sub done` / `/kpi done` to stop running timers before marking `Done`.

### Fixed
- Fixed bug where non-existent timer queries fell back to returning other users' active timers.
- Fixed `get_duration()` time string parsing and overnight timer midnight crossing (`+86400s`).
- Fixed module import trigger that previously terminated running bot instances.
