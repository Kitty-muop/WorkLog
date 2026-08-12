# 📝 Task List: Enforce Mandatory Slash Command Fields

- [ ] **Task 1**: Create TDD Unit Test Suite `tests/test_mandatory_fields.py`
  - Acceptance: Unit tests verify missing mandatory fields return error code 1.
  - Verify: `.\.venv\Scripts\python.exe -m unittest discover -s tests -p "test_*.py"`
  - Files: `tests/test_mandatory_fields.py`

- [ ] **Task 2**: Implement Mandatory Field Checks in `tools/timer.py`
  - Acceptance: `cmd_start`, `cmd_sub_add`, `cmd_sub_start`, `cmd_kpi_add` validate required fields before execution.
  - Verify: `.\.venv\Scripts\python.exe -m unittest discover -s tests -p "test_*.py"`
  - Files: `tools/timer.py`

- [ ] **Task 3**: Update Discord Slash Command Signatures in `tools/discord_bot.py`
  - Acceptance: Slash command signatures require positional arguments without default `= None`.
  - Verify: `.\.venv\Scripts\python.exe -m py_compile tools/discord_bot.py`
  - Files: `tools/discord_bot.py`

- [ ] **Task 4**: Full System Verification & OS Background Bot Restart
  - Acceptance: All 14+ tests pass, Google Sheets synced, OS bot restarted with PID active.
  - Verify: `powershell -Command "Get-Process -Id (Get-Content 'D:\WorkLog\tools\bot.pid')"`
  - Files: All
