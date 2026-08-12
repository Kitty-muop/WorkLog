# 📋 Implementation Plan: Mandatory Fields Enforcement

**Spec:** [`docs/specs/spec_mandatory_fields.md`](file:///D:/WorkLog/docs/specs/spec_mandatory_fields.md)  
**Date:** 2026-08-12  

---

## 🎯 Architecture & Plan Overview

1. **Phase 1: TDD Unit Test Suite (`tests/test_mandatory_fields.py`)**:
   - Write tests enforcing non-null/non-empty checks for `project`, `task`, `subtask`, `category`, `estimate` in `cmd_start`, `cmd_sub_add`, `cmd_sub_start`, `cmd_kpi_add`.

2. **Phase 2: Core Command Layer (`tools/timer.py`)**:
   - Add explicit mandatory parameter validation checks returning exit code `1` if any required parameter is missing or empty.

3. **Phase 3: Discord Slash Command Layer (`tools/discord_bot.py`)**:
   - Update Discord Slash Command signature signatures (`project: str`, `task: str`, `subtask: str`, `category: str`, `estimate: str`) without default `= None` values.
   - Update parameter `@app_commands.describe` tags to explicitly mark them `[REQUIRED]`.

4. **Phase 4: Verification & Deployment**:
   - Run unit test suite, syntax compile, validate Excel schema, full sync to Google Sheets, and restart OS background bot process.
