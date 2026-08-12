# 📜 Specification: Enforce Mandatory Slash Command Fields

**Version:** 1.0.0  
**Date:** 2026-08-12  
**Status:** DRAFT (Awaiting Approval)  

---

## 🎯 1. Objective

Transform optional parameters into **Strict Mandatory (Required) Fields** across WorkLog creation and timer slash commands (`/start`, `/subtask add`, `/subtask start`, `/kpi add`) so that no incomplete, untracked, or default-fallback entries can be created without explicit project, task, category, and estimate parameters.

---

## 🎯 2. Assumptions & Surface Requirements

1. **`/start` Command Requirements**:
   - `project`: **Required** (e.g. `"AI-AutoTest"`)
   - `task`: **Required** (e.g. `"Create MCP"`)
   - `subtask`: **Required** (e.g. `"Add tool"`)
   - `category`: **Required** (Autocomplete: `Development`, `Debug / Bug Fix`, `Refactoring`, `Code Review`, `Testing / QA`, `DevOps / CI-CD`, `Documentation`)
   - `estimate`: **Required** (Positive float in hours, e.g. `3.5`)

2. **`/subtask add` Command Requirements**:
   - `name`: **Required**
   - `father_task`: **Required**
   - `project`: **Required**
   - `estimate`: **Required**

3. **`/subtask start` Command Requirements**:
   - `name`: **Required**
   - `project`: **Required**
   - `task`: **Required**
   - `category`: **Required**
   - `estimate`: **Required**

4. **`/kpi add` Command Requirements**:
   - `code`: **Required**
   - `name`: **Required**
   - `project`: **Required**
   - `deadline_days`: **Required**

---

## 💻 3. Technical Commands

- **Run Unit Tests**: `.\.venv\Scripts\python.exe -m unittest discover -s tests -p "test_*.py"`
- **Validate Excel Schema**: `.\.venv\Scripts\python.exe -m tools.validate`
- **Full Sync Google Sheets**: `.\.venv\Scripts\python.exe -m tools.sync_gsheets`
- **Restart OS Background Bot**: `wscript.exe "C:\Users\Admin\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\WorkLogBot.vbs"`

---

## 🏗️ 4. Project Structure

```text
D:\WorkLog\
├── tools/
│   ├── discord_bot.py   # Slash Command Decorators & Descriptions
│   ├── timer.py         # Command Implementations & Required Field Checks
│   └── gamify.py        # EXP & Performance Engine
├── tests/
│   ├── test_mandatory_fields.py # TDD Unit Tests for Mandatory Parameters
│   └── ...
├── docs/specs/
│   └── spec_mandatory_fields.md # This Specification File
└── tasks/
    ├── plan.md          # Implementation Plan
    └── todo.md          # Discrete Task List
```

---

## 📐 5. Code Style & Example Snippet

```python
# GOOD: Enforce non-null mandatory parameter checks
@bot.tree.command(name="start", description="Start timer (All parameters are mandatory)")
@app_commands.describe(
    project="[REQUIRED] Project name",
    task="[REQUIRED] Father task name",
    subtask="[REQUIRED] Subtask name",
    category="[REQUIRED] Work category",
    estimate="[REQUIRED] Estimated hours (e.g. 3.5)"
)
@app_commands.autocomplete(project=_ac_all_proj, task=_ac_all_task, subtask=_ac_all_sub, category=_ac_cat)
async def start(
    interaction: discord.Interaction,
    project: str,
    task: str,
    subtask: str,
    category: str,
    estimate: str
):
    ...
```

---

## 🧪 6. Testing Strategy

- Unit test cases in `tests/test_mandatory_fields.py` verifying:
  - Missing mandatory fields return clean error code `1` and descriptive error message.
  - Slash command parameters declare non-default positional arguments in Discord tree.

---

## 🚧 7. Boundaries

- **Always Do**: Run full test suite before committing, validate inputs at command boundary, return private ephemeral responses (`ephemeral=True`).
- **Ask First**: Changing Excel sheet columns or database schema.
- **Never Do**: Silently fall back to placeholder values when mandatory fields are missing.

---

## ✅ 8. Success Criteria

1. 100% of start and creation commands require explicit user parameters.
2. Missing mandatory parameters rejected at interaction boundary.
3. 100% unit tests pass cleanly in `< 100ms`.
