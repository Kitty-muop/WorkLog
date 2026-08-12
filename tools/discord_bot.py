#!/usr/bin/env python3
"""WorkLog Discord Bot — slash commands with autocomplete."""

import argparse
import datetime
import io
import os
import sys
from pathlib import Path
from typing import Optional

import discord
from discord import app_commands
from discord.ext import commands, tasks as ext_tasks

sys.path.insert(0, os.path.dirname(__file__))
import timer as tm
import gamify as gm

# Load .env from project root
env_path = Path(__file__).resolve().parent.parent / '.env'
if env_path.exists():
    for line in env_path.read_text(encoding='utf-8').splitlines():
        line = line.strip()
        if line and not line.startswith('#'):
            k, _, v = line.partition('=')
            os.environ[k.strip()] = v.strip()

import subprocess

PID_FILE = Path(__file__).resolve().parent / 'bot.pid'

def enforce_single_instance():
    current_pid = os.getpid()
    if PID_FILE.exists():
        try:
            old_pid = int(PID_FILE.read_text(encoding='utf-8').strip())
            if old_pid != current_pid:
                cmd = f'taskkill /F /PID {old_pid}'
                subprocess.run(cmd, shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                print(f"[SINGLE INSTANCE] Terminated old bot process PID {old_pid}", flush=True)
        except Exception as e:
            print(f"[SINGLE INSTANCE] Note: {e}", flush=True)
    PID_FILE.write_text(str(current_pid), encoding='utf-8')

TOKEN = os.environ.get('WORKLOG_DISCORD_TOKEN', '')
intents = discord.Intents.default()
intents.message_content = True  # Requires enabling in Discord Developer Portal
intents.members = True  # Requires enabling in Discord Developer Portal
# Note: If intents fail, disable the above lines and ensure they're enabled in:
# https://discord.com/developers/applications -> Bot -> Privileged Gateway Intents
bot = commands.Bot(command_prefix='/', intents=intents, help_command=None)


import re

def _clean_param(val):
    if not val:
        return val
    s = str(val).strip()
    s = re.sub(r'^[^\w\d\s\-_\.]+\s*', '', s)
    s = re.sub(r'\s*\([^)]*?(?:running|paused|break|min|h|m)[^)]*?\)$', '', s, flags=re.IGNORECASE)
    s = re.sub(r'\s*\[[^\]]*?\]$', '', s)
    return s.strip()


def _run(func, **ns):
    cleaned = {}
    for k, v in ns.items():
        if isinstance(v, str):
            cleaned[k] = _clean_param(v)
        else:
            cleaned[k] = v

    cleaned.setdefault('subtask_code', None)
    cleaned.setdefault('deadline', None)
    cleaned.setdefault('description', None)
    cleaned.setdefault('notes', None)
    cleaned.setdefault('category', None)
    args = argparse.Namespace(**cleaned)
    buf = io.StringIO()
    old = sys.stdout
    sys.stdout = buf
    try:
        ret = func(args)
    except SystemExit:
        ret = 1
    except Exception as e:
        print(f"Error: {e}", file=buf)
        ret = 1
    finally:
        sys.stdout = old
    return buf.getvalue().strip() or "(done)", ret


import time

# ── Cache & Robust Response Helpers ──
_cache = {}
def _cached(key, ttl_sec, func):
    now = time.time()
    if key in _cache:
        val, ts = _cache[key]
        if now - ts < ttl_sec:
            return val
    val = func()
    _cache[key] = (val, now)
    return val


def clear_bot_cache():
    _cache.clear()


async def _defer(interaction: discord.Interaction, ephemeral: bool = True) -> bool:
    try:
        if not interaction.response.is_done():
            await interaction.response.defer(ephemeral=ephemeral)
        return True
    except Exception as e:
        print(f"Defer note: {e}")
        return False


async def _send_embed(
    interaction: discord.Interaction,
    embed: discord.Embed,
    raw_text: str = None,
    deferred: bool = True,
    filename: str = "worklog_report.md",
    ephemeral: bool = True
):
    """Send a rich Discord Embed card with attached .md file backup (STRICTLY EPHEMERAL/PRIVATE)."""
    if interaction and interaction.user and not embed.author.name:
        avatar_url = interaction.user.display_avatar.url if hasattr(interaction.user, 'display_avatar') else None
        embed.set_author(name=f"🛡️ Hero: {interaction.user.display_name}", icon_url=avatar_url)
    if not embed.timestamp:
        embed.timestamp = datetime.datetime.now(datetime.timezone.utc)

    file_bytes = None
    if raw_text:
        md_content = f"# WorkLog Report\n\n{raw_text}\n"
        file_bytes = io.BytesIO(md_content.encode('utf-8'))

    def _get_file():
        if file_bytes:
            file_bytes.seek(0)
            return discord.File(fp=file_bytes, filename=filename)
        return None

    if deferred or interaction.response.is_done():
        try:
            f = _get_file()
            if f:
                await interaction.followup.send(embed=embed, file=f, ephemeral=ephemeral)
            else:
                await interaction.followup.send(embed=embed, ephemeral=ephemeral)
            return
        except Exception as e:
            print(f"Followup send_embed error ({e}), retrying response.send_message...")

    try:
        f = _get_file()
        if f:
            await interaction.response.send_message(embed=embed, file=f, ephemeral=ephemeral)
        else:
            await interaction.response.send_message(embed=embed, ephemeral=ephemeral)
        return
    except Exception as e:
        print(f"Response send_message error: {e}")


async def _send_output(interaction: discord.Interaction, out_text: str, deferred: bool = True, filename: str = "worklog_report.md"):
    text = str(out_text).strip() or "(done)"
    embed = build_action_card("📊 WorkLog Report", 0x3498db, user=interaction.user if interaction else None, raw_output=text)
    await _send_embed(interaction, embed, text, deferred=deferred, filename=filename)


async def _reply(interaction, text):
    await _send_output(interaction, text, deferred=True)


# ── Rich Discord Embed Builders matching Excel Sheet Designs ──


def make_progress_bar(pct: float, length: int = 10) -> str:
    """Create a high-resolution ASCII progress bar."""
    pct = max(0.0, min(1.0, float(pct)))
    filled = int(round(pct * length))
    empty = length - filled
    return "█" * filled + "░" * empty


def make_game_exp_bar(score: int) -> tuple[str, str, int, int]:
    """Calculate level, level name, progress bar, current exp, and next level exp."""
    thresholds = gm.SCORE_THRESHOLDS if (gm and hasattr(gm, 'SCORE_THRESHOLDS')) else [0, 10, 25, 50, 100, 200, 350, 500]

    lvl = 0
    for i, t in reversed(list(enumerate(thresholds))):
        if score >= t:
            lvl = i
            break

    current_threshold = thresholds[lvl]
    next_threshold = thresholds[lvl + 1] if lvl + 1 < len(thresholds) else current_threshold * 2

    exp_in_level = score - current_threshold
    level_range = max(1, next_threshold - current_threshold)
    pct = exp_in_level / level_range

    bar = make_progress_bar(pct, length=10)
    tier_name = gm.get_tier_name(lvl) if (gm and hasattr(gm, 'get_tier_name')) else "Novice"
    return f"{tier_name} (Lv.{lvl})", bar, score, next_threshold


def build_gamified_hud_embed(user_id=None):
    """Build a RPG Quest Live HUD embed."""
    g_res = gm.run() if gm else {}
    score = g_res.get('total_score', 0)
    level_name, exp_bar, curr_xp, next_xp = make_game_exp_bar(score)
    streak = g_res.get('streak', 0)
    max_streak = g_res.get('max_streak', 0)

    today_hours = 0.0
    daily_summary = g_res.get('daily_summary', {})
    today_str = str(datetime.date.today())
    if today_str in daily_summary:
        today_hours = daily_summary[today_str].get('hours', 0.0)

    state = tm.load_state()
    timers = state.get('timers', [])
    if user_id:
        timers = [t for t in timers if str(t.get('user_id', '')) == str(user_id)]

    active_timers = [t for t in timers if not t.get('paused')]
    paused_timers = [t for t in timers if t.get('paused')]

    embed = discord.Embed(
        title="🎮 WORKLOG QUEST HUD — LIVE DASHBOARD",
        color=0x9b59b6 if active_timers else (0xf1c40f if paused_timers else 0x34495e),
        timestamp=datetime.datetime.now(datetime.timezone.utc)
    )

    daily_target = gm.DAILY_TARGET_HOURS if gm else 7.5
    daily_pct = min(1.0, today_hours / daily_target)
    daily_bar = make_progress_bar(daily_pct, length=10)

    stats_val = (
        f"🏆 **Rank:** `Lv.{g_res.get('level', 0)} {level_name}`\n"
        f"⭐ **EXP:** `{score} / {next_xp}` `[{exp_bar}]`\n"
        f"🔥 **Streak:** `{streak} Days` *(Best: {max_streak}d)*\n"
        f"🎯 **Daily Goal:** `{today_hours:.2f}h / {daily_target}h` `[{daily_bar}]` ({int(daily_pct*100)}%)"
    )
    embed.add_field(name="🛡️ Hero Stats & Progress", value=stats_val, inline=False)

    if active_timers:
        q_lines = []
        for t in active_timers:
            acc = tm.get_elapsed(t)
            h, m, s = int(acc) // 3600, (int(acc) % 3600) // 60, int(acc) % 60
            dur_str = f"{h}h {m:02d}m {s:02d}s" if h else f"{m:02d}m {s:02d}s"

            p_name = t.get('project', 'General')
            ft_name = t.get('task', 'Work')
            st_name = t.get('subtask', '')
            cat = t.get('category', 'Development')

            q_str = f"⚔️ **[{p_name}] {ft_name}**"
            if st_name:
                q_str += f" ── *{st_name}*"
            q_str += f"\n⏱️ **Duration:** `{dur_str}` │ 🏷️ `{cat}`"
            q_lines.append(q_str)

        embed.add_field(name=f"⚡ Active Quests ({len(active_timers)})", value="\n\n".join(q_lines)[:1024], inline=False)

    elif paused_timers:
        p_lines = []
        for t in paused_timers:
            acc = tm.get_elapsed(t)
            h, m = int(acc) // 3600, (int(acc) % 3600) // 60
            dur_str = f"{h}h{m:02d}m" if h else f"{m}m"
            r = t.get('pause_reason', 'resting')
            p_lines.append(f"⏸️ **[{t.get('project')}] {t.get('task')}** (`{dur_str}`) — *{r}*")
        embed.add_field(name=f"☕ Checkpoint / Paused ({len(paused_timers)})", value="\n".join(p_lines)[:1024], inline=False)
    else:
        embed.add_field(
            name="💤 Campfire Rest",
            value="No active quests running. Use `/start` in Discord to launch a new quest!",
            inline=False
        )

    embed.set_footer(text="🎮 WorkLog RPG System • Real-Time Progress HUD")
    return embed


def build_action_card(
    title: str,
    color: int,
    user: discord.User | discord.Member = None,
    fields: list[tuple[str, str, bool]] = None,
    description: str = None,
    raw_output: str = None
) -> discord.Embed:
    g_res = gm.run() if gm else {}
    score = g_res.get('total_score', 0)
    level_name, exp_bar, curr_xp, next_xp = make_game_exp_bar(score)
    streak = g_res.get('streak', 0)

    title_map = {
        "▶️ Timer Started": "⚔️ QUEST ACCEPTED & BATTLE STARTED",
        "⏹️ Timer Stopped & Saved": "🎉 QUEST COMPLETED & REWARDS SAVED!",
        "⏹️ Timer Stopped": "⏹️ QUEST ENDED & LOGGED",
        "⏸️ Timer Paused": "⏸️ QUEST PAUSED (Campfire Rest)",
        "▶️ Timer Resumed": "▶️ QUEST RESUMED (Back to Battle)",
        "❌ Timer Cancelled": "❌ QUEST ABANDONED (Time Discarded)",
        "📁 Project Added": "🏰 REALM UNLOCKED (Project Created)",
        "📁 Project Renamed": "🏰 REALM RENAMED",
        "📦 Project Archived": "📦 REALM ARCHIVED",
        "🎯 Father Task (KPI) Added": "📜 MAIN QUEST CREATED (KPI Added)",
        "✅ Father Task Completed": "🏆 MAIN QUEST ACCOMPLISHED!",
        "🎯 Father Task Updated": "📜 MAIN QUEST UPDATED",
        "🎯 Father Task Renamed": "📜 MAIN QUEST RENAMED",
        "🗑️ Father Task Deleted": "🗑️ MAIN QUEST DELETED",
        "🧩 Subtask Added": "🧩 SIDE QUEST CREATED (Subtask Added)",
        "✅ Subtask Marked as Done": "✨ SIDE QUEST COMPLETED!",
        "🧩 Subtask Updated": "🧩 SIDE QUEST UPDATED",
        "🧩 Subtask Renamed": "🧩 SIDE QUEST RENAMED",
        "🗑️ Subtask Deleted": "🗑️ SIDE QUEST DELETED",
        "▶️ Subtask Timer Started": "⚔️ SIDE QUEST BATTLE STARTED",
        "⏹️ Subtask Timer Stopped": "🎉 SIDE QUEST COMPLETED & SAVED!",
        "⏹️ All Timers Stopped & Saved": "⏹️ ALL QUEST BATTLES ENDED & REWARDS SAVED!"
    }
    game_title = title_map.get(title, title)

    embed = discord.Embed(
        title=game_title,
        color=color,
        timestamp=datetime.datetime.now(datetime.timezone.utc)
    )
    if user:
        avatar_url = user.display_avatar.url if hasattr(user, 'display_avatar') else None
        embed.set_author(
            name=f"🛡️ Hero: {user.display_name} │ Lv.{g_res.get('level', 0)} {level_name} │ 🔥 {streak}d Streak",
            icon_url=avatar_url
        )

    if description:
        embed.description = str(description)[:4096]
    elif raw_output:
        raw_str = str(raw_output)
        if len(raw_str) > 1800:
            raw_str = raw_str[:1800] + "\n... [output truncated, see attached .md file]"
        embed.description = f"```markdown\n{raw_str}\n```"

    if fields:
        for name, val, inline in fields:
            if val:
                embed.add_field(name=str(name)[:256], value=str(val)[:1024], inline=inline)

    embed.set_footer(text=f"🎮 WorkLog RPG Engine v2.0 │ Multi-User System │ EXP: {score}/{next_xp} [{exp_bar}]")
    return embed


def build_projects_embed():
    projects = tm.get_projects()
    if not projects:
        return discord.Embed(title="🏰 REALMS & KINGDOMS (No Realms Found)", description="No projects found. Use `/project add` to unlock a new realm!", color=0x3498db)

    embed = discord.Embed(
        title="🏰 REALMS & KINGDOMS DIRECTORY (Projects)",
        description="**Columns:** `Code` │ `Realm Name` │ `Status` │ `Description` │ `Created Date`",
        color=0x3498db,
        timestamp=datetime.datetime.now(datetime.timezone.utc)
    )

    lines = []
    for p in projects:
        st = p.get('status', 'Active')
        emoji = "🏰" if st.lower() == 'active' else "📦"
        code = p.get('code', 'PRJ')
        name = p.get('name', '')
        desc = f" — *{p['description']}*" if p.get('description') else ""
        c_date = f" ({p['created_date'].strftime('%d-%m-%Y')})" if hasattr(p.get('created_date'), 'strftime') else (f" ({p['created_date']})" if p.get('created_date') else "")
        lines.append(f"{emoji} `{code}` **{name}** [{st}]{c_date}{desc}")

    embed.add_field(name=f"Registered Realms ({len(projects)})", value="\n".join(lines)[:1024], inline=False)
    embed.set_footer(text=f"Total: {len(projects)} Realms Registered │ WorkLog RPG System │ Excel & GSheets Synced")
    return embed


def build_tasks_embed(task_filter=None):
    kpis = tm.get_kpis()
    subtasks = tm.get_subtasks()
    if not kpis:
        return discord.Embed(title="🎯 QUEST MAP & ADVENTURE DIRECTORY", description="No Main Quests (KPIs) found.", color=0x3498db)

    embed = discord.Embed(
        title="🎯 QUEST MAP & ADVENTURE DIRECTORY (Tasks Grouped)",
        color=0x3498db,
        timestamp=datetime.datetime.now(datetime.timezone.utc)
    )

    sub_by_kpi = {}
    for s in subtasks:
        ft = s.get('father_task', 'General')
        sub_by_kpi.setdefault(ft.lower(), []).append(s)

    for k in kpis:
        if task_filter and task_filter.lower() not in k['name'].lower():
            continue
        hours = tm.get_total_hours_for_task(k['name'])
        st_list = sub_by_kpi.get(k['name'].lower(), [])

        lines = [f"🏰 **Realm:** `{k['project']}` │ **Deadline:** `{k['deadline_days']}d` │ **EXP Watched:** `{hours:.2f}h`"]
        if st_list:
            lines.append("**Side Quests:**")
            for s in st_list:
                s_emoji = "✨" if s['status'].lower() in ('done', 'completed') else ("⚔️" if s['status'].lower() in ('active', 'in progress') else "⏳")
                lines.append(f"  • {s_emoji} `{s['code']}` **{s['name']}** [{s['status']}]")
        else:
            lines.append("*(No side quests registered under this Main Quest)*")

        k_emoji = "🏆" if k['status'].lower() == 'done' else ("📜" if k['status'].lower() in ('active', 'in progress') else "⏳")
        embed.add_field(name=f"{k_emoji} `{k['code']}` {k['name']} [{k['status']}]", value="\n".join(lines)[:1024], inline=False)

    embed.set_footer(text="WorkLog Quest Hierarchy: Side Quest < Main Quest < Realm │ RPG Engine v2.0")
    return embed


def build_subtask_status_embed():
    subtasks = tm.get_subtasks()
    if not subtasks:
        return discord.Embed(title="🧩 SIDE QUEST LOGBOOK (Subtasks)", description="No side quests found.", color=0x3498db)

    by_father = {}
    for s in subtasks:
        key = f"{s['project']} / {s['father_task']}"
        by_father.setdefault(key, []).append(s)

    embed = discord.Embed(
        title="🧩 SIDE QUEST LOGBOOK (Subtasks)",
        description="**Columns:** `Code` │ `Side Quest` │ `Status` │ `Completed Date` │ `Notes`",
        color=0x3498db
    )

    total_all = 0
    done_all = 0
    for key in sorted(by_father.keys()):
        items = by_father[key]
        lines = []
        done_sub = 0
        for s in items:
            st = s['status']
            if st.lower() in ('done', 'completed'):
                emoji = "✨"
                done_sub += 1
            elif st.lower() in ('active', 'in progress'):
                emoji = "⚔️"
            else:
                emoji = "⏳"

            comp = f" (Cleared: {s['completed_date'].strftime('%d-%m-%Y')})" if hasattr(s.get('completed_date'), 'strftime') else (f" (Cleared: {s['completed_date']})" if s.get('completed_date') else "")
            notes = f" — *{s['notes']}*" if s.get('notes') else ""
            lines.append(f"{emoji} `{s['code']}` **{s['name']}** — `{st}`{comp}{notes}")
        total_all += len(items)
        done_all += done_sub
        val_text = "\n".join(lines)
        val_text += f"\n*Cleared: {done_sub}/{len(items)}*"
        embed.add_field(name=f"🏰 {key}", value=val_text[:1024], inline=False)

    pct_done = int((done_all / total_all) * 100) if total_all > 0 else 0
    bar = make_progress_bar(done_all / total_all if total_all > 0 else 0, length=10)
    embed.set_footer(text=f"Quest Cleared: {done_all}/{total_all} [{bar}] ({pct_done}%) │ WorkLog RPG System")
    return embed


def build_timer_status_embed(user_id=None):
    state = tm.load_state()
    timers = state.get('timers', [])
    if user_id:
        timers = [t for t in timers if str(t.get('user_id', '')) == str(user_id)]
    if not timers:
        return discord.Embed(title="⚔️ ACTIVE QUEST BATTLES & CHECKPOINTS", description="No active quest battles currently running for your account. Use `/start` to begin!", color=0x34495e)

    today = datetime.date.today()
    timers_sorted = sorted(timers, key=lambda t: t.get('segment_start', ''))

    active_count = 0
    paused_count = 0
    hung_count = 0

    embed = discord.Embed(
        title="⚔️ ACTIVE QUEST BATTLES & CHECKPOINTS",
        color=0x2ecc71 if any(not t.get('paused') for t in timers) else 0xf1c40f
    )

    for t in timers_sorted:
        sd = datetime.datetime.fromisoformat(t['segment_start'])
        is_hung = sd.date() < today
        acc = tm.get_elapsed(t)
        h, m = int(acc) // 3600, (int(acc) % 3600) // 60
        dur = f"{h}h{m:02d}m" if h else f"{m}m"

        is_paused = t.get('paused', False)
        if is_hung and not is_paused:
            emoji = "⚠️"
            status_str = f"HUNG BATTLE since {sd.strftime('%d-%m %H:%M')}"
            hung_count += 1
        elif is_paused:
            emoji = "⏸️"
            reason = t.get('pause_reason', 'resting')
            status_str = f"CAMPFIRE REST ({reason}) since {sd.strftime('%H:%M')}"
            paused_count += 1
        else:
            emoji = "⚔️"
            status_str = f"IN BATTLE since {sd.strftime('%H:%M')}"
            active_count += 1

        p_name = t.get('project', '') or 'General'
        ft_name = t.get('task', '') or 'Task'
        st_name = t.get('subtask', '')
        user_name = t.get('user_name', '')

        hier = tm.get_hierarchy_watched_time(p_name, ft_name, st_name, acc)

        user_str = f" 🛡️ {user_name}" if user_name else ""
        field_name = f"{emoji} #{t['id']}{user_str} │ {dur} │ {p_name} / {ft_name}"
        details = []
        if user_name:
            details.append(f"**Hero:** {user_name}")
        details.append(f"**Quest Status:** {status_str}")

        if st_name:
            sub_d = f"{st_name} (`{t['subtask_code']}`)" if t.get('subtask_code') else st_name
            details.append(f"**Side Quest:** {sub_d}")

        if t.get('category'):
            details.append(f"**Category:** {t['category']}")
        if t.get('description'):
            details.append(f"**Objective:** {t['description']}")

        hier_str = f"📊 **3-Level Watched Time:**\n• 🏰 Realm (`{p_name}`): `{hier['project'][1]:.2f}h`\n• 📜 Main Quest (`{ft_name}`): `{hier['kpi'][1]:.2f}h`"
        if st_name:
            hier_str += f"\n• 🧩 Side Quest (`{st_name}`): `{hier['subtask'][1]:.2f}h`"
        details.append(hier_str)

        embed.add_field(name=str(field_name)[:256], value="\n".join(details)[:1024], inline=False)

    embed.set_footer(text=f"⚔️ Battles: {active_count} │ ⏸️ Campfire: {paused_count} │ ⚠️ Hung: {hung_count} │ Quest Hierarchy: Side < Main < Realm")
    return embed


def build_performance_embed():
    if not os.path.exists(tm.WORKLOG_FILE):
        return discord.Embed(title="👑 LEGENDARY PERFORMANCE AUDIT", description="No worklog.xlsx found.", color=0xe74c3c)

    wb = tm.load_workbook(tm.WORKLOG_FILE)
    ws = wb['Time Entries']
    today = datetime.date.today()
    monday = today - datetime.timedelta(days=today.weekday())
    sunday = monday + datetime.timedelta(days=6)

    proj_hours = {}
    kpi_hours = {}
    total_week_hours = 0.0
    active_days = set()

    for r in range(2, ws.max_row + 1):
        raw_d = ws.cell(r, 1).value
        d = tm.parse_date_val(raw_d)
        if not d or not (monday <= d <= sunday):
            continue
        active_days.add(d)
        dur = tm.get_duration(ws, r)
        total_week_hours += dur

        p = str(ws.cell(r, 3).value or 'General')
        ft = str(ws.cell(r, 4).value or 'General')

        proj_hours[p] = proj_hours.get(p, 0.0) + dur
        kpi_hours[ft] = kpi_hours.get(ft, 0.0) + dur

    subtasks = tm.get_subtasks()
    completed_this_week = []
    in_progress_subtasks = []
    for s in subtasks:
        c_date = s.get('completed_date')
        if c_date:
            if hasattr(c_date, 'date'):
                c_date = c_date.date()
            if isinstance(c_date, datetime.date) and monday <= c_date <= sunday:
                completed_this_week.append(s)
        if str(s.get('status', '')).lower() in ('in progress', 'active'):
            in_progress_subtasks.append(s)

    target_hours = 37.5
    perf_pct = min(int((total_week_hours / target_hours) * 100), 100) if total_week_hours > 0 else 0

    is_friday = today.weekday() == 4
    title_prefix = "👑 FRIDAY LEGENDARY PERFORMANCE AUDIT" if is_friday else "📊 WEEKLY QUEST PERFORMANCE REPORT"

    embed = discord.Embed(
        title=f"{title_prefix} ({monday.strftime('%d-%m')} → {sunday.strftime('%d-%m-%Y')})",
        color=0x2ecc71 if perf_pct >= 80 else (0xf1c40f if perf_pct >= 50 else 0xe74c3c)
    )

    bar = make_progress_bar(total_week_hours / target_hours if target_hours > 0 else 0, length=10)
    embed.description = (
        f"🏆 **Performance Score:** `{perf_pct}%` `[{bar}]`\n"
        f"⏱️ **Weekly Total:** `{total_week_hours:.2f}h / {target_hours:.1f}h` │ 📅 **Days Worked:** `{len(active_days)}/5 days`\n"
        f"✨ **Side Quests Cleared:** `{len(completed_this_week)}` │ ⚔️ **In Battle:** `{len(in_progress_subtasks)}`"
    )

    p_lines = []
    for p, h in sorted(proj_hours.items(), key=lambda x: -x[1]):
        p_lines.append(f"• 🏰 **{p}**: `{h:.2f}h` watched")
    if p_lines:
        embed.add_field(name="🏰 Level 1 — Realms Watched Time", value="\n".join(p_lines)[:1024], inline=False)

    k_lines = []
    for ft, h in sorted(kpi_hours.items(), key=lambda x: -x[1]):
        k_lines.append(f"• 📜 **{ft}**: `{h:.2f}h` watched")
    if k_lines:
        embed.add_field(name="📜 Level 2 — Main Quests Watched Time", value="\n".join(k_lines)[:1024], inline=False)

    s_lines = []
    if completed_this_week:
        for s in completed_this_week:
            s_lines.append(f"✨ `{s['code']}` **{s['name']}** [{s['project']}]")
    for s in in_progress_subtasks[:5]:
        s_lines.append(f"⚔️ `{s['code']}` **{s['name']}** [{s['project']}]")
    if s_lines:
        embed.add_field(name="🧩 Level 3 — Side Quests Progress", value="\n".join(s_lines)[:1024], inline=False)

    embed.set_footer(text="3-Level Quest Hierarchy: Side Quest < Main Quest < Realm │ RPG System")
    return embed


def build_kpi_embed():
    kpis = tm.get_kpis()
    if not kpis:
        return discord.Embed(title="📜 MAIN QUESTS BOARD (KPIs)", description="No Main Quests found.", color=0x3498db)

    by_project = {}
    for k in kpis:
        by_project.setdefault(k['project'], []).append(k)

    embed = discord.Embed(
        title="📜 MAIN QUESTS BOARD (KPIs)",
        description="**Columns:** `Code` │ `Main Quest` │ `Deadline` │ `Hours Logged` │ `Status`",
        color=0x3498db
    )

    total_all = 0.0
    for proj_name in sorted(by_project.keys()):
        items = by_project[proj_name]
        lines = []
        proj_total = 0.0
        for k in items:
            hours = tm.get_total_hours_for_task(k['name'])
            proj_total += hours
            dl = k['deadline_days']
            dl_str = f" ({dl}d deadline)" if dl else ""
            st = k['status']
            emoji = "🏆" if st.lower() in ('done', 'completed') else ("📜" if st.lower() in ('active', 'in progress') else "⏳")
            lines.append(f"{emoji} `{k['code']}` **{k['name']}** — `{hours:.2f}h` [{st}]{dl_str}")

        total_all += proj_total
        embed.add_field(name=f"🏰 Realm: {proj_name} (Subtotal: {proj_total:.2f}h)", value="\n".join(lines)[:1024], inline=False)

    embed.set_footer(text=f"Grand Total: {total_all:.2f}h logged across {len(kpis)} Main Quests │ WorkLog RPG System")
    return embed


def build_today_embed():
    if not os.path.exists(tm.WORKLOG_FILE):
        return discord.Embed(title="🛡️ DAILY BATTLE REPORT", description="No worklog file found.", color=0xe74c3c)

    wb = tm.load_workbook(tm.WORKLOG_FILE)
    ws = wb['Time Entries']
    today = datetime.date.today()
    total = 0.0

    entries = []
    for r in range(2, ws.max_row + 1):
        raw_d = ws.cell(r, 1).value
        d = tm.parse_date_val(raw_d)
        if not d or d != today:
            continue
        dur = tm.get_duration(ws, r)
        total += dur
        entries.append({
            'proj': str(ws.cell(r, 3).value or ''),
            'task': str(ws.cell(r, 4).value or ''),
            'sub': str(ws.cell(r, 5).value or ''),
            'desc': str(ws.cell(r, 11).value or ''),
            'dur': dur,
        })

    embed = discord.Embed(
        title=f"🛡️ DAILY BATTLE REPORT ({today.strftime('%d-%m-%Y')})",
        color=0x2ecc71 if total >= 7.5 else 0x3498db
    )

    if not entries:
        embed.description = "No quest battles logged today yet. Use `/start` to launch a new quest!"
        return embed

    by_project = {}
    for e in entries:
        by_project.setdefault(e['proj'], []).append(e)

    for proj, proj_entries in sorted(by_project.items()):
        proj_total = sum(e['dur'] for e in proj_entries)
        lines = []
        for e in proj_entries:
            sub_str = f" (`{e['sub']}`)" if e['sub'] else ""
            desc_str = f" — *{e['desc']}*" if e['desc'] else ""
            lines.append(f"• **{e['task']}**{sub_str}: `{e['dur']:.2f}h`{desc_str}")
        embed.add_field(name=f"🏰 {proj} ({proj_total:.2f}h)", value="\n".join(lines)[:1024], inline=False)

    pct = int(min((total / 7.5) * 100, 100)) if total > 0 else 0
    bar = make_progress_bar(total / 7.5 if total > 0 else 0, length=10)
    embed.description = f"🎯 **DAILY GOAL:** `{total:.2f}h / 7.5h` `[{bar}] {pct}%`"
    return embed


def build_week_embed():
    if not os.path.exists(tm.WORKLOG_FILE):
        return discord.Embed(title="🏆 WEEKLY CONQUEST SUMMARY", description="No worklog file found.", color=0xe74c3c)

    wb = tm.load_workbook(tm.WORKLOG_FILE)
    ws = wb['Time Entries']

    today = datetime.date.today()
    monday = today - datetime.timedelta(days=today.weekday())
    sunday = monday + datetime.timedelta(days=6)

    by_project = {}
    total = 0.0
    days_set = set()
    for r in range(2, ws.max_row + 1):
        raw_d = ws.cell(r, 1).value
        d = tm.parse_date_val(raw_d)
        if not d or not (monday <= d <= sunday):
            continue
        days_set.add(d)
        dur = tm.get_duration(ws, r)
        total += dur
        proj = str(ws.cell(r, 3).value or '?')
        task = str(ws.cell(r, 4).value or '?')
        by_project.setdefault(proj, {})
        by_project[proj].setdefault(task, 0)
        by_project[proj][task] += dur

    embed = discord.Embed(
        title=f"🏆 WEEKLY CONQUEST SUMMARY ({monday.strftime('%d-%m')} → {sunday.strftime('%d-%m-%Y')})",
        color=0x9b59b6
    )

    for proj in sorted(by_project.keys()):
        tasks = by_project[proj]
        proj_total = sum(tasks.values())
        lines = []
        for t, h in sorted(tasks.items(), key=lambda x: -x[1]):
            lines.append(f"• **{t}**: `{h:.2f}h`")
        embed.add_field(name=f"🏰 {proj} ({proj_total:.2f}h)", value="\n".join(lines)[:1024], inline=False)

    days_worked = len(days_set)
    bar = make_progress_bar(total / 37.5 if total > 0 else 0, length=10)
    pct = int(min((total / 37.5) * 100, 100)) if total > 0 else 0
    embed.description = (
        f"🏆 **WEEKLY TOTAL:** `{total:.2f}h / 37.5h` `[{bar}] {pct}%`\n"
        f"📅 **Days Worked:** `{days_worked}/5 days` │ 🔥 **Consistency Score:** `{days_worked * 20}%`"
    )
    return embed


def _fmt_timer_status(timer):
    """Format timer status for autocomplete label."""
    elapsed = tm.get_elapsed(timer)
    h, m = int(elapsed) // 3600, (int(elapsed) % 3600) // 60
    time_str = f"{h}h{m:02d}m" if h else f"{m}m"
    u_prefix = f"[{timer['user_name']}] " if timer.get('user_name') else ""

    seg = timer.get('segment_start', '')
    if seg:
        seg_date = datetime.datetime.fromisoformat(seg).date()
        if seg_date < datetime.date.today():
            return f"⚠️ {u_prefix}{{}} ({time_str} HUNG)"

    if timer.get('paused', False):
        reason = timer.get('pause_reason', 'break')
        return f"⏸️ {u_prefix}{{}} ({time_str} paused: {reason})"

    return f"🟢 {u_prefix}{{}} ({time_str} running)"


async def _ac_proj(interaction: discord.Interaction, current: str):
    projects = list(reversed(_cached('projects', 2, tm.get_projects)))
    state = tm.load_state()
    active_projects = {t.get('project') for t in state.get('timers', []) if t.get('project')}

    choices = []
    c_lower = current.strip().lower()
    for p in projects:
        p_name = p.get('name', '')
        if not p_name:
            continue
        if not c_lower or c_lower in p_name.lower():
            st = p.get('status', 'Active')
            if st.lower() in ('archived', 'done', 'completed'):
                continue
            emoji = "🟢" if p_name in active_projects else _status_emoji(st)
            label = f"{emoji} {p_name}"[:100]
            choices.append(app_commands.Choice(name=label, value=p_name[:100]))

    choices.sort(key=lambda c: 0 if c.name.startswith("🟢") else 1)
    return choices[:5]


async def _ac_task(interaction: discord.Interaction, current: str):
    kpis = list(reversed(_cached('kpis', 2, tm.get_kpis)))
    state = tm.load_state()
    active_tasks = {t.get('task') for t in state.get('timers', []) if t.get('task')}

    choices = []
    c_lower = current.strip().lower()
    for k in kpis:
        k_name = k.get('name', '')
        if not k_name:
            continue
        if not c_lower or c_lower in k_name.lower():
            st = k.get('status', 'Active')
            if st.lower() in ('archived', 'done', 'completed'):
                continue
            emoji = "🟢" if k_name in active_tasks else _status_emoji(st)
            proj_str = f" — {k.get('project')}" if k.get('project') else ""
            label = f"{emoji} {k_name}{proj_str}"[:100]
            choices.append(app_commands.Choice(name=label, value=k_name[:100]))

    choices.sort(key=lambda c: 0 if c.name.startswith("🟢") else 1)
    return choices[:5]


async def _ac_sub(interaction: discord.Interaction, current: str):
    subtasks = list(reversed(_cached('subtasks', 2, tm.get_subtasks)))
    state = tm.load_state()
    active_subtasks = {t.get('subtask') for t in state.get('timers', []) if t.get('subtask')}

    choices = []
    c_lower = current.strip().lower()
    for s in subtasks:
        s_name = s.get('name', '')
        if not s_name:
            continue
        if not c_lower or c_lower in s_name.lower():
            st = s.get('status', 'Pending')
            if st.lower() in ('archived', 'done', 'completed'):
                continue
            emoji = "🟢" if s_name in active_subtasks else _status_emoji(st)
            father_str = f" — {s.get('father_task')}" if s.get('father_task') else ""
            label = f"{emoji} {s_name}{father_str}"[:100]
            choices.append(app_commands.Choice(name=label, value=s_name[:100]))

    choices.sort(key=lambda c: 0 if c.name.startswith("🟢") else 1)
    return choices[:5]


def _status_emoji(status):
    """Map status string to emoji."""
    s = (status or '').strip().lower()
    if s in ('done', 'completed'):
        return '✅'
    if s in ('active', 'in progress'):
        return '🔵'
    if s in ('archived',):
        return '📦'
    if s in ('pending', 'not started'):
        return '⏳'
def check_reminder_dedup(sent_dict, user_id, date_str, window_key):
    """Return True if reminder has not been sent yet for this user, date, and window."""
    u_sent = sent_dict.get(str(user_id), {})
    key = f"{date_str}_{window_key}"
    return not u_sent.get(key, False)


def mark_reminder_sent(sent_dict, user_id, date_str, window_key):
    """Mark reminder as sent for this user, date, and window."""
    u_key = str(user_id)
    if u_key not in sent_dict:
        sent_dict[u_key] = {}
    key = f"{date_str}_{window_key}"
    sent_dict[u_key][key] = True


async def _ac_cat(interaction: discord.Interaction, current: str):
    cats = gm.PRODUCTION_CATEGORIES if (gm and hasattr(gm, 'PRODUCTION_CATEGORIES')) else [
        "Development", "Debug / Bug Fix", "Refactoring", "Code Review", "Testing / QA", "DevOps / CI-CD", "Documentation"
    ]
    c_lower = current.strip().lower()
    choices = []
    for c in cats:
        if not c_lower or c_lower in c.lower():
            emoji = "🐛" if "debug" in c.lower() or "bug" in c.lower() else "🏷️"
            choices.append(app_commands.Choice(name=f"{emoji} {c}", value=c))
    return choices[:5]


_ac_all_proj = _ac_proj
_ac_all_task = _ac_task
_ac_all_sub = _ac_sub


# ════════════════════════  Timer  ════════════════════════

@bot.tree.command(name="start", description="Start timer (All parameters are mandatory)")
@app_commands.describe(project="[REQUIRED] Project name", task="[REQUIRED] Father task name", subtask="[REQUIRED] Subtask name", category="[REQUIRED] Work category", estimate="[REQUIRED] Estimated hours (e.g. 3.5)")
@app_commands.autocomplete(project=_ac_all_proj, task=_ac_all_task, subtask=_ac_all_sub, category=_ac_cat)
async def start(interaction: discord.Interaction, project: str, task: str, subtask: str, category: str, estimate: str):
    deferred = await _defer(interaction, ephemeral=True)
    u_id = str(interaction.user.id)
    u_name = interaction.user.display_name
    out, ret = _run(tm.cmd_start, project=project, task=task, subtask=subtask, category=category, estimate=estimate, user_id=u_id, user_name=u_name)
    clear_bot_cache()

    state = tm.load_state()
    timer = tm.find_timer(state, user_id=u_id, paused=False)

    if timer and ret == 0:
        if subtask_name := timer.get('subtask'):
            tm.gsheets.sync_timer_start(
                timer.get('project') or '', timer.get('task') or '',
                subtask_name, timer.get('subtask_code') or '',
                timer.get('category') or 'Development', timer.get('description') or ''
            )
        fields = [
            ("👤 User", u_name, True),
            ("🆔 Timer ID", f"`#{timer.get('id')}`", True),
            ("📁 Project", f"`{timer.get('project') or 'General'}`", True),
            ("🎯 Father Task", f"`{timer.get('task') or 'Task'}`", True),
        ]
        if timer.get('subtask'):
            sub_d = f"{timer['subtask']} (`{timer['subtask_code']}`)" if timer.get('subtask_code') else timer['subtask']
            fields.append(("🧩 Subtask", sub_d, True))
        fields.append(("🏷️ Category", timer.get('category', 'Development'), True))
        if timer.get('description'):
            fields.append(("📝 Description", timer['description'], False))
        active_cnt = len(tm.find_timers(state, paused=False))
        embed = build_action_card("▶️ Timer Started", 0x2ecc71, user=interaction.user, fields=fields)
        embed.set_footer(text=f"⚔️ Active Quests: {active_cnt} │ 🎮 WorkLog RPG Engine v2.0")
    else:
        embed = build_action_card("▶️ Timer Started", 0x2ecc71 if ret == 0 else 0xe74c3c, user=interaction.user, raw_output=out)

    await _send_embed(interaction, embed, out, deferred, "start.md")


@bot.tree.command(name="stop", description="Stop timer and save entry")
@app_commands.describe(project="Project name", task="Father task name", subtask="Subtask name")
@app_commands.autocomplete(project=_ac_proj, task=_ac_task, subtask=_ac_sub)
async def stop(interaction: discord.Interaction, project: str = None, task: str = None, subtask: str = None):
    deferred = await _defer(interaction)
    u_id = str(interaction.user.id)
    u_name = interaction.user.display_name
    state_before = tm.load_state()
    timer = tm.find_timer(state_before, project=project, task=task, subtask=subtask, user_id=u_id, paused=None)

    out, ret = _run(tm.cmd_stop, project=project, task=task, subtask=subtask, user_id=u_id, user_name=u_name)
    clear_bot_cache()

    if timer and ret == 0:
        if subtask_name := timer.get('subtask'):
            tm.gsheets.sync_timer_start(
                timer.get('project') or '', timer.get('task') or '',
                subtask_name, timer.get('subtask_code') or '',
                timer.get('category') or 'Development',
                (timer.get('description') or '') + " [STOPPED]"
            )
        acc = tm.get_elapsed(timer)
        h, m = int(acc) // 3600, (int(acc) % 3600) // 60
        dur_str = f"{h}h{m:02d}m (`{acc/3600:.2f}h`)" if h else f"{m}m (`{acc/3600:.2f}h`)"

        fields = [
            ("👤 User", u_name, True),
            ("🆔 Timer ID", f"`#{timer.get('id')}`", True),
            ("⏱️ Duration", dur_str, True),
            ("📁 Project", f"`{timer.get('project') or 'General'}`", True),
            ("🎯 Father Task", f"`{timer.get('task') or 'Task'}`", True),
        ]
        if timer.get('subtask'):
            sub_d = f"{timer['subtask']} (`{timer['subtask_code']}`)" if timer.get('subtask_code') else timer['subtask']
            fields.append(("🧩 Subtask", sub_d, True))
        fields.append(("🏷️ Category", timer.get('category', 'Development'), True))
        if timer.get('description'):
            fields.append(("📝 Description", timer['description'], False))

        embed = build_action_card("⏹️ Timer Stopped & Saved", 0x3498db, user=interaction.user, fields=fields)
    else:
        embed = build_action_card("⏹️ Timer Stopped", 0x3498db if ret == 0 else 0xe74c3c, user=interaction.user, raw_output=out)

    await _send_embed(interaction, embed, out, deferred, "stop.md")


@bot.tree.command(name="stop_all", description="Stop ALL active and hung timers for your user and save to WorkLog")
async def stop_all_cmd(interaction: discord.Interaction):
    deferred = await _defer(interaction)
    u_id = str(interaction.user.id)
    u_name = interaction.user.display_name
    out, ret = _run(tm.cmd_stop, stop_all=True, user_id=u_id, user_name=u_name)
    clear_bot_cache()

    embed = build_action_card(
        "⏹️ All Timers Stopped & Saved",
        0x3498db if ret == 0 else 0xe74c3c,
        user=interaction.user,
        raw_output=out
    )
    await _send_embed(interaction, embed, out, deferred, "stop_all.md")


@bot.tree.command(name="pause", description="Pause running timer")
@app_commands.describe(project="Project name", task="Father task name", subtask="Subtask name", reason="Reason: meeting, lunch, break, review, context-switch...")
@app_commands.autocomplete(project=_ac_proj, task=_ac_task, subtask=_ac_sub)
async def pause(interaction: discord.Interaction, project: str = None, task: str = None, subtask: str = None, reason: str = None):
    deferred = await _defer(interaction)
    u_id = str(interaction.user.id)
    u_name = interaction.user.display_name

    out, ret = _run(tm.cmd_pause, reason=reason, project=project, task=task, subtask=subtask, user_id=u_id, user_name=u_name)
    clear_bot_cache()

    state = tm.load_state()
    timer = tm.find_timer(state, project=project, task=task, subtask=subtask, user_id=u_id, paused=True)
    if timer and ret == 0:
        if subtask_name := timer.get('subtask'):
            r = timer.get('pause_reason', reason or 'break')
            tm.gsheets.sync_timer_start(
                timer.get('project') or '', timer.get('task') or '',
                subtask_name, timer.get('subtask_code') or '',
                timer.get('category') or 'Development',
                (timer.get('description') or '') + f" [PAUSED: {r}]"
            )
        acc = timer.get('accumulated_seconds', 0)
        h, m = int(acc) // 3600, (int(acc) % 3600) // 60
        acc_str = f"{h}h{m:02d}m" if h else f"{m}m"

        fields = [
            ("👤 User", u_name, True),
            ("🆔 Timer ID", f"`#{timer.get('id')}`", True),
            ("⏳ Accumulated", acc_str, True),
            ("💬 Reason", timer.get('pause_reason', reason or 'break'), True),
            ("📁 Project", f"`{timer.get('project') or 'General'}`", True),
            ("🎯 Father Task", f"`{timer.get('task') or 'Task'}`", True),
        ]
        embed = build_action_card("⏸️ Timer Paused", 0xf1c40f, user=interaction.user, fields=fields)
    else:
        embed = build_action_card("⏸️ Timer Paused", 0xf1c40f if ret == 0 else 0xe74c3c, user=interaction.user, raw_output=out)

    await _send_embed(interaction, embed, out, deferred, "pause.md")


@bot.tree.command(name="continue", description="Resume paused timer")
@app_commands.describe(project="Project name", task="Father task name", subtask="Subtask name")
@app_commands.autocomplete(project=_ac_proj, task=_ac_task, subtask=_ac_sub)
async def resume(interaction: discord.Interaction, project: str = None, task: str = None, subtask: str = None):
    deferred = await _defer(interaction)
    u_id = str(interaction.user.id)
    u_name = interaction.user.display_name

    out, ret = _run(tm.cmd_continue, project=project, task=task, subtask=subtask, user_id=u_id, user_name=u_name)
    clear_bot_cache()

    state = tm.load_state()
    timer = tm.find_timer(state, project=project, task=task, subtask=subtask, user_id=u_id, paused=False)
    if timer and ret == 0:
        if subtask_name := timer.get('subtask'):
            tm.gsheets.sync_timer_start(
                timer.get('project') or '', timer.get('task') or '',
                subtask_name, timer.get('subtask_code') or '',
                timer.get('category') or 'Development',
                (timer.get('description') or '') + " [RESUMED]"
            )
        fields = [
            ("👤 User", u_name, True),
            ("🆔 Timer ID", f"`#{timer.get('id')}`", True),
            ("📁 Project", f"`{timer.get('project') or 'General'}`", True),
            ("🎯 Father Task", f"`{timer.get('task') or 'Task'}`", True),
        ]
        embed = build_action_card("▶️ Timer Resumed", 0x2ecc71, user=interaction.user, fields=fields)
    else:
        embed = build_action_card("▶️ Timer Resumed", 0x2ecc71 if ret == 0 else 0xe74c3c, user=interaction.user, raw_output=out)

    await _send_embed(interaction, embed, out, deferred, "resume.md")


@bot.tree.command(name="cancel", description="Cancel timer and discard time")
async def cancel(interaction: discord.Interaction):
    deferred = await _defer(interaction)
    u_id = str(interaction.user.id)
    u_name = interaction.user.display_name
    out, ret = _run(tm.cmd_cancel, user_id=u_id, user_name=u_name)
    clear_bot_cache()
    fields = [("👤 User", u_name, True)]
    embed = build_action_card("❌ Timer Cancelled", 0xe74c3c, user=interaction.user, fields=fields, raw_output=out)
    await _send_embed(interaction, embed, out, deferred, "cancel.md")


@bot.tree.command(name="status", description="Show all running/paused timers")
async def status(interaction: discord.Interaction):
    deferred = await _defer(interaction, ephemeral=True)
    u_id = str(interaction.user.id)
    u_name = interaction.user.display_name
    out, _ = _run(tm.cmd_status, user_id=u_id, user_name=u_name)
    embed = build_timer_status_embed(user_id=u_id)
    await _send_embed(interaction, embed, out, deferred, "status.md", ephemeral=True)


@bot.tree.command(name="today", description="Show today's time summary")
async def today(interaction: discord.Interaction):
    deferred = await _defer(interaction)
    out, _ = _run(tm.cmd_today)
    embed = build_today_embed()
    await _send_embed(interaction, embed, out, deferred, "today.md")


@bot.tree.command(name="week", description="Show this week's time summary")
async def week(interaction: discord.Interaction):
    deferred = await _defer(interaction)
    out, _ = _run(tm.cmd_week)
    embed = build_week_embed()
    await _send_embed(interaction, embed, out, deferred, "week.md")


@bot.tree.command(name="performance", description="Show Friday Weekly Performance Report across 3 levels (Subtask < KPI < Project)")
async def performance_cmd(interaction: discord.Interaction):
    deferred = await _defer(interaction)
    out, _ = _run(tm.cmd_performance)
    embed = build_performance_embed()
    await _send_embed(interaction, embed, out, deferred, "weekly_performance_report.md")


@bot.tree.command(name="tasks", description="Show tasks grouped by father task")
@app_commands.describe(task="Filter by father task name")
async def tasks(interaction: discord.Interaction, task: str = None):
    deferred = await _defer(interaction)
    out, _ = _run(tm.cmd_tasks, task=task)
    embed = build_tasks_embed(task_filter=task)
    await _send_embed(interaction, embed, out, deferred, "tasks.md")


# ════════════════════════  Project  ════════════════════════

project = app_commands.Group(name="project", description="Manage projects")


@project.command(name="list", description="List all projects")
async def proj_list(interaction: discord.Interaction):
    deferred = await _defer(interaction)
    out, _ = _run(tm.cmd_project_list)
    embed = build_projects_embed()
    await _send_embed(interaction, embed, out, deferred, "projects.md")


@project.command(name="add", description="Add a new project")
@app_commands.describe(name="Project name", description="Project description")
async def proj_add(interaction: discord.Interaction, name: str, description: str = None):
    deferred = await _defer(interaction)
    out, ret = _run(tm.cmd_project_add, name=name, description=description)
    clear_bot_cache()
    fields = [
        ("👤 User", interaction.user.display_name, True),
        ("📁 Project Name", f"`{name}`", True),
        ("📊 Status", "`Active`", True),
    ]
    if description:
        fields.append(("📝 Description", description, False))
    embed = build_action_card("📁 Project Added", 0x2ecc71 if ret == 0 else 0xe74c3c, user=interaction.user, fields=fields, raw_output=out if ret != 0 else None)
    await _send_embed(interaction, embed, out, deferred)


@project.command(name="rename", description="Rename a project")
@app_commands.describe(name="Current project name", new_name="New project name")
@app_commands.autocomplete(name=_ac_all_proj)
async def proj_rename(interaction: discord.Interaction, name: str, new_name: str):
    deferred = await _defer(interaction)
    out, ret = _run(tm.cmd_project_rename, name=name, new_name=new_name)
    clear_bot_cache()
    fields = [
        ("👤 User", interaction.user.display_name, True),
        ("📁 Old Name", f"`{name}`", True),
        ("📁 New Name", f"`{new_name}`", True),
    ]
    embed = build_action_card("📁 Project Renamed", 0x3498db if ret == 0 else 0xe74c3c, user=interaction.user, fields=fields, raw_output=out if ret != 0 else None)
    await _send_embed(interaction, embed, out, deferred)


@project.command(name="archive", description="Archive a project")
@app_commands.describe(name="Project name to archive")
@app_commands.autocomplete(name=_ac_all_proj)
async def proj_archive(interaction: discord.Interaction, name: str):
    deferred = await _defer(interaction)
    out, ret = _run(tm.cmd_project_archive, name=name)
    clear_bot_cache()
    fields = [
        ("👤 User", interaction.user.display_name, True),
        ("📁 Project Name", f"`{name}`", True),
        ("📊 New Status", "`Archived`", True),
    ]
    embed = build_action_card("📦 Project Archived", 0xe67e22 if ret == 0 else 0xe74c3c, user=interaction.user, fields=fields, raw_output=out if ret != 0 else None)
    await _send_embed(interaction, embed, out, deferred)


# ════════════════════════  KPI (Father Task)  ════════════════════════

kpi = app_commands.Group(name="kpi", description="Manage father tasks (KPIs)")


@kpi.command(name="list", description="List all father tasks with status")
async def kpi_list(interaction: discord.Interaction):
    deferred = await _defer(interaction)
    out, _ = _run(tm.cmd_kpi_list)
    embed = build_kpi_embed()
    await _send_embed(interaction, embed, out, deferred, "kpi_status.md")


@kpi.command(name="add", description="Add a new father task")
@app_commands.describe(name="Father task name", project="Project name", deadline="Deadline in days (default: 7)")
@app_commands.autocomplete(project=_ac_all_proj)
async def kpi_add(interaction: discord.Interaction, name: str, project: str, deadline: int = 7):
    deferred = await _defer(interaction)
    out, ret = _run(tm.cmd_kpi_add, name=name, project=project, deadline=deadline)
    clear_bot_cache()
    fields = [
        ("👤 User", interaction.user.display_name, True),
        ("🎯 Father Task", f"`{name}`", True),
        ("📁 Project", f"`{project}`", True),
        ("📅 Deadline", f"`{deadline} days` (`{deadline * 7.5:.1f}h`)", True),
    ]
    embed = build_action_card("🎯 Father Task (KPI) Added", 0x2ecc71 if ret == 0 else 0xe74c3c, user=interaction.user, fields=fields, raw_output=out if ret != 0 else None)
    await _send_embed(interaction, embed, out, deferred)


@kpi.command(name="done", description="Mark a father task as completed")
@app_commands.describe(name="Father task name")
@app_commands.autocomplete(name=_ac_all_task)
async def kpi_done(interaction: discord.Interaction, name: str):
    deferred = await _defer(interaction)
    out, ret = _run(tm.cmd_kpi_done, name=name)
    clear_bot_cache()
    fields = [
        ("👤 User", interaction.user.display_name, True),
        ("🎯 Father Task", f"`{name}`", True),
        ("📊 Status", "`Done ✅`", True),
    ]
    embed = build_action_card("✅ Father Task Completed", 0x2ecc71 if ret == 0 else 0xe74c3c, user=interaction.user, fields=fields, raw_output=out if ret != 0 else None)
    await _send_embed(interaction, embed, out, deferred)


@kpi.command(name="edit", description="Edit a father task's project or deadline")
@app_commands.describe(name="Father task name", project="New project name", deadline="New deadline in days")
@app_commands.autocomplete(name=_ac_all_task, project=_ac_all_proj)
async def kpi_edit(interaction: discord.Interaction, name: str, project: str = None, deadline: int = None):
    deferred = await _defer(interaction)
    out, ret = _run(tm.cmd_kpi_edit, name=name, project=project, deadline=deadline)
    clear_bot_cache()
    fields = [
        ("👤 User", interaction.user.display_name, True),
        ("🎯 Father Task", f"`{name}`", True),
    ]
    if project:
        fields.append(("📁 Project", f"`{project}`", True))
    if deadline:
        fields.append(("📅 Deadline", f"`{deadline} days`", True))
    embed = build_action_card("🎯 Father Task Updated", 0x3498db if ret == 0 else 0xe74c3c, user=interaction.user, fields=fields, raw_output=out if ret != 0 else None)
    await _send_embed(interaction, embed, out, deferred)


@kpi.command(name="rename", description="Rename a father task")
@app_commands.describe(name="Current father task name", new_name="New father task name")
@app_commands.autocomplete(name=_ac_all_task)
async def kpi_rename(interaction: discord.Interaction, name: str, new_name: str):
    deferred = await _defer(interaction)
    out, ret = _run(tm.cmd_kpi_rename, name=name, new_name=new_name)
    clear_bot_cache()
    fields = [
        ("👤 User", interaction.user.display_name, True),
        ("🎯 Old Name", f"`{name}`", True),
        ("🎯 New Name", f"`{new_name}`", True),
    ]
    embed = build_action_card("🎯 Father Task Renamed", 0x3498db if ret == 0 else 0xe74c3c, user=interaction.user, fields=fields, raw_output=out if ret != 0 else None)
    await _send_embed(interaction, embed, out, deferred)


@kpi.command(name="delete", description="Delete a father task")
@app_commands.describe(name="Father task name to delete")
@app_commands.autocomplete(name=_ac_all_task)
async def kpi_delete(interaction: discord.Interaction, name: str):
    deferred = await _defer(interaction)
    out, ret = _run(tm.cmd_kpi_delete, name=name)
    clear_bot_cache()
    fields = [
        ("👤 User", interaction.user.display_name, True),
        ("🎯 Father Task", f"`{name}`", True),
    ]
    embed = build_action_card("🗑️ Father Task Deleted", 0xe74c3c if ret == 0 else 0xe74c3c, user=interaction.user, fields=fields, raw_output=out if ret != 0 else None)
    await _send_embed(interaction, embed, out, deferred)


@kpi.command(name="status", description="Show overall task progress")
async def kpi_status(interaction: discord.Interaction):
    deferred = await _defer(interaction)
    out, _ = _run(tm.cmd_kpi_status)
    embed = build_kpi_embed()
    await _send_embed(interaction, embed, out, deferred, "kpi_status.md")


# ════════════════════════  Subtask  ════════════════════════

subtask = app_commands.Group(name="subtask", description="Manage subtasks")


@subtask.command(name="list", description="List all subtasks")
@app_commands.describe(task="Filter by father task name")
@app_commands.autocomplete(task=_ac_all_task)
async def sub_list(interaction: discord.Interaction, task: str = None):
    deferred = await _defer(interaction)
    out, _ = _run(tm.cmd_sub_list, task=task)
    embed = build_subtask_status_embed()
    await _send_embed(interaction, embed, out, deferred, "subtask_list.md")


@subtask.command(name="add", description="Add a new subtask (All parameters are mandatory)")
@app_commands.describe(name="[REQUIRED] Subtask name", father_task="[REQUIRED] Father task name", project="[REQUIRED] Project name", estimate="[REQUIRED] Estimated hours (e.g. 3.5)")
@app_commands.autocomplete(father_task=_ac_all_task, project=_ac_all_proj)
async def sub_add(interaction: discord.Interaction, name: str, father_task: str, project: str, estimate: str):
    deferred = await _defer(interaction)
    out, ret = _run(tm.cmd_sub_add, subtask=name, task=father_task, project=project, estimate=estimate)
    clear_bot_cache()
    fields = [
        ("👤 User", interaction.user.display_name, True),
        ("🧩 Subtask Name", f"`{name}`", True),
        ("🎯 Father Task", f"`{father_task}`", True),
        ("🎯 Estimate", f"`{tm.parse_estimate(estimate):.1f}h`", True),
    ]
    if project:
        fields.append(("📁 Project", f"`{project}`", True))
    embed = build_action_card("🧩 Subtask Added", 0x2ecc71 if ret == 0 else 0xe74c3c, user=interaction.user, fields=fields, raw_output=out if ret != 0 else None)
    await _send_embed(interaction, embed, out, deferred)


@subtask.command(name="done", description="Mark a subtask as completed")
@app_commands.describe(name="Subtask name")
@app_commands.autocomplete(name=_ac_all_sub)
async def sub_done(interaction: discord.Interaction, name: str):
    deferred = await _defer(interaction)
    out, ret = _run(tm.cmd_sub_done, subtask=name)
    clear_bot_cache()
    fields = [
        ("👤 User", interaction.user.display_name, True),
        ("🧩 Subtask Name", f"`{name}`", True),
        ("📊 Status", "`Done ✅`", True),
    ]
    embed = build_action_card("✅ Subtask Marked as Done", 0x2ecc71 if ret == 0 else 0xe74c3c, user=interaction.user, fields=fields, raw_output=out if ret != 0 else None)
    await _send_embed(interaction, embed, out, deferred)


@subtask.command(name="edit", description="Edit a subtask")
@app_commands.describe(name="Subtask name", new_name="New name", father_task="New father task", project="New project", estimate="New estimated hours (e.g. 4.5)", notes="New notes")
@app_commands.autocomplete(name=_ac_all_sub, father_task=_ac_all_task, project=_ac_all_proj)
async def sub_edit(interaction: discord.Interaction, name: str, new_name: str = None, father_task: str = None, project: str = None, estimate: str = None, notes: str = None):
    deferred = await _defer(interaction)
    out, ret = _run(tm.cmd_sub_edit, subtask=name, new_name=new_name, task=father_task, project=project, estimate=estimate, notes=notes)
    clear_bot_cache()
    fields = [
        ("👤 User", interaction.user.display_name, True),
        ("🧩 Subtask Name", f"`{name}`", True),
    ]
    if new_name:
        fields.append(("🧩 New Name", f"`{new_name}`", True))
    if father_task:
        fields.append(("🎯 Father Task", f"`{father_task}`", True))
    if project:
        fields.append(("📁 Project", f"`{project}`", True))
    if estimate:
        fields.append(("🎯 Estimate", f"`{tm.parse_estimate(estimate):.1f}h`", True))
    embed = build_action_card("🧩 Subtask Updated", 0x3498db if ret == 0 else 0xe74c3c, user=interaction.user, fields=fields, raw_output=out if ret != 0 else None)
    await _send_embed(interaction, embed, out, deferred)


@subtask.command(name="rename", description="Rename a subtask in all entries")
@app_commands.describe(name="Current subtask name", new_name="New subtask name")
@app_commands.autocomplete(name=_ac_all_sub)
async def sub_rename(interaction: discord.Interaction, name: str, new_name: str):
    deferred = await _defer(interaction)
    out, ret = _run(tm.cmd_sub_rename, subtask=name, new_name=new_name)
    clear_bot_cache()
    fields = [
        ("👤 User", interaction.user.display_name, True),
        ("🧩 Old Name", f"`{name}`", True),
        ("🧩 New Name", f"`{new_name}`", True),
    ]
    embed = build_action_card("🧩 Subtask Renamed", 0x3498db if ret == 0 else 0xe74c3c, user=interaction.user, fields=fields, raw_output=out if ret != 0 else None)
    await _send_embed(interaction, embed, out, deferred)


@subtask.command(name="delete", description="Delete subtask from all entries")
@app_commands.describe(name="Subtask name to delete")
@app_commands.autocomplete(name=_ac_all_sub)
async def sub_delete(interaction: discord.Interaction, name: str):
    deferred = await _defer(interaction)
    out, ret = _run(tm.cmd_sub_delete, subtask=name)
    clear_bot_cache()
    fields = [
        ("👤 User", interaction.user.display_name, True),
        ("🧩 Subtask Name", f"`{name}`", True),
    ]
    embed = build_action_card("🗑️ Subtask Deleted", 0xe74c3c if ret == 0 else 0xe74c3c, user=interaction.user, fields=fields, raw_output=out if ret != 0 else None)
    await _send_embed(interaction, embed, out, deferred)


@subtask.command(name="status", description="Show subtask progress")
async def sub_status(interaction: discord.Interaction):
    deferred = await _defer(interaction)
    out, _ = _run(tm.cmd_sub_status)
    embed = build_subtask_status_embed()
    await _send_embed(interaction, embed, out, deferred, "subtask_status.md")


@subtask.command(name="start", description="Start timer for a subtask (All parameters are mandatory)")
@app_commands.describe(name="[REQUIRED] Subtask name or code", project="[REQUIRED] Project name", task="[REQUIRED] Father task name", category="[REQUIRED] Category", description="Description", estimate="[REQUIRED] Estimated hours (e.g. 3.5)")
@app_commands.autocomplete(name=_ac_all_sub, project=_ac_all_proj, task=_ac_all_task, category=_ac_cat)
async def sub_start(interaction: discord.Interaction, name: str, project: str, task: str, category: str, estimate: str, description: str = None):
    deferred = await _defer(interaction)
    u_id = str(interaction.user.id)
    u_name = interaction.user.display_name
    out, ret = _run(tm.cmd_sub_start, subtask=name, project=project, task=task, category=category, description=description, user_id=u_id, user_name=u_name)
    clear_bot_cache()

    subtasks = tm.get_subtasks()
    match = None
    for s in subtasks:
        if s['name'].strip().lower() == name.strip().lower() or s['code'].strip().lower() == name.strip().lower():
            match = s
            break
    if match:
        tm.gsheets.sync_timer_start(
            project or match.get('project') or '',
            task or match.get('father_task') or '',
            match['name'],
            match['code'],
            category or 'Development',
            description or ''
        )

    fields = [
        ("👤 User", u_name, True),
        ("🧩 Subtask", f"`{name}`", True),
    ]
    if project:
        fields.append(("📁 Project", f"`{project}`", True))
    if task:
        fields.append(("🎯 Father Task", f"`{task}`", True))

    embed = build_action_card("▶️ Subtask Timer Started", 0x2ecc71 if ret == 0 else 0xe74c3c, user=interaction.user, fields=fields, raw_output=out if ret != 0 else None)
    await _send_embed(interaction, embed, out, deferred, "subtask_start.md")


@subtask.command(name="stop", description="Stop timer for a subtask")
@app_commands.describe(name="Subtask name to stop", project="Project override", task="Father task override")
@app_commands.autocomplete(name=_ac_sub, project=_ac_proj, task=_ac_task)
async def sub_stop(interaction: discord.Interaction, name: str, project: str = None, task: str = None):
    deferred = await _defer(interaction)
    u_id = str(interaction.user.id)
    u_name = interaction.user.display_name
    out, ret = _run(tm.cmd_stop, subtask=name, project=project, task=task, user_id=u_id, user_name=u_name)
    clear_bot_cache()

    fields = [
        ("👤 User", u_name, True),
        ("🧩 Subtask", f"`{name}`", True),
    ]
    embed = build_action_card("⏹️ Subtask Timer Stopped", 0x3498db if ret == 0 else 0xe74c3c, user=interaction.user, fields=fields, raw_output=out if ret != 0 else None)
    await _send_embed(interaction, embed, out, deferred, "subtask_stop.md")


# ════════════════════════  Run Airtest  ════════════════════════

@bot.tree.command(name="run-airtest", description="Run Airtest with parameters")
@app_commands.describe(game="Game name", project="Project name", version="Version", test_scripts="Test scripts path")
async def run_airtest(interaction: discord.Interaction, game: str, project: str, version: str, test_scripts: str):
    deferred = await _defer(interaction)
    cmd = f"airtest run {test_scripts} --game {game} --project {project} --version {version}"
    import subprocess
    try:
        result = subprocess.run(cmd, shell=True, capture_output=True, text=True, timeout=300)
        out = result.stdout or result.stderr or "(done)"
        ret = result.returncode
    except Exception as e:
        out = f"Error: {e}"
        ret = 1

    fields = [
        ("👤 User", interaction.user.display_name, True),
        ("🎮 Game", f"`{game}`", True),
        ("📁 Project", f"`{project}`", True),
        ("🏷️ Version", f"`{version}`", True),
        ("📜 Test Script", f"`{test_scripts}`", False),
    ]
    embed = build_action_card("🧪 Airtest Execution", 0x9b59b6 if ret == 0 else 0xe74c3c, user=interaction.user, fields=fields, raw_output=out)
    await _send_embed(interaction, embed, out, deferred, "airtest_result.md")


# ════════════════════════  Help  ════════════════════════

HELP_TEXT = """**WORKLOG — Hướng dẫn sử dụng**

**Flow làm việc:**
  Project → KPI (Father Task) → Subtask → Start/Stop

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

**Bước 1 — Tạo Project**
  `/project add name:"Tên project"`
  `/project list`
  `/project rename name:"Cũ" new_name:"Mới"`
  `/project archive name:"Tên"`

**Bước 2 — Tạo KPI (Father Task)**
  `/kpi add name:"Tên task" project:"Tên project" deadline:7`
  `/kpi list`
  `/kpi done name:"Tên task"`
  `/kpi status`
  `/kpi edit name:"Tên" project:"Mới" deadline:14`
  `/kpi rename name:"Cũ" new_name:"Mới"`
  `/kpi delete name:"Tên"`

**Bước 3 — Tạo Subtask**
  `/subtask add name:"Tên sub" father_task:"Tên Father Task"`
  `/subtask list`
  `/subtask done name:"Tên sub"`
  `/subtask status`
  `/subtask rename name:"Cũ" new_name:"Mới"`
  `/subtask delete name:"Tên"`
  `/subtask edit name:"Tên" new_name:"Mới" father_task:"FT" project:"P"`

**Bước 4 — Timer**
  `/start project:"P" task:"FT" subtask:"ST"`
  `/stop` — Lưu vào worklog + gửi webhook
  `/pause reason:"lunch"` — Tạm dừng
  `/continue` — Tiếp tục
  `/cancel` — Hủy bỏ (không lưu)
  `/status` — Xem timer đang chạy

**Báo cáo**
  `/today` — Hôm nay
  `/week` — Tuần này
  `/tasks task:"FT"` — Chi tiết theo Father Task
"""


@bot.tree.command(name="help-log", description="Show full usage guide")
async def help_cmd(interaction: discord.Interaction):
    deferred = await _defer(interaction)
    embed = build_action_card("📖 WorkLog Usage Guide & Command Directory", 0x3498db, user=interaction.user, description=HELP_TEXT)
    await _send_embed(interaction, embed, HELP_TEXT, deferred, "help_guide.md")


# ════════════════════════  Run  ════════════════════════

bot.tree.add_command(project)
bot.tree.add_command(kpi)
bot.tree.add_command(subtask)


@bot.tree.error
async def on_tree_error(interaction: discord.Interaction, error: app_commands.AppCommandError):
    print(f"[ERROR] Slash command error: {error}", flush=True)
    try:
        if not interaction.response.is_done():
            await interaction.response.send_message(f"⚠️ **Error executing command:** `{error}`", ephemeral=True)
        else:
            await interaction.followup.send(f"⚠️ **Error executing command:** `{error}`", ephemeral=True)
    except Exception as ex:
        print(f"[ERROR] Failed to send error feedback: {ex}", flush=True)


def get_ict_now():
    """Get current datetime in UTC+7 (Asia/Ho_Chi_Minh)."""
    tz_ict = datetime.timezone(datetime.timedelta(hours=7))
    return datetime.datetime.now(tz_ict)


class WorklogOvertimeView(discord.ui.View):
    def __init__(self, session_name: str):
        super().__init__(timeout=1800)
        self.session_name = session_name

    @discord.ui.button(label="▶️ Continue Working", style=discord.ButtonStyle.green, custom_id="overtime_continue")
    async def continue_working(self, interaction: discord.Interaction, button: discord.ui.Button):
        out, ret = tm.cmd_continue()
        clear_bot_cache()

        for item in self.children:
            item.disabled = True

        embed = interaction.message.embeds[0] if interaction.message.embeds else discord.Embed()
        embed.title = f"▶️ Overtime Session Continued ({self.session_name})"
        embed.description = f"**{interaction.user.display_name}** selected to **continue working**. Timer resumed!\n```\n{out}\n```"
        embed.color = 0x2ecc71

        await interaction.response.edit_message(embed=embed, view=self)

    @discord.ui.button(label="⏹️ Stop & Save Session", style=discord.ButtonStyle.red, custom_id="overtime_stop")
    async def stop_working(self, interaction: discord.Interaction, button: discord.ui.Button):
        out, ret = tm.cmd_stop(description=f"Work session ended ({self.session_name})")
        clear_bot_cache()

        for item in self.children:
            item.disabled = True

        embed = interaction.message.embeds[0] if interaction.message.embeds else discord.Embed()
        embed.title = f"⏹️ Work Session Stopped ({self.session_name})"
        embed.description = f"**{interaction.user.display_name}** selected to **stop & save session**. Saved to WorkLog!\n```\n{out}\n```"
        embed.color = 0xe74c3c

        await interaction.response.edit_message(embed=embed, view=self)


_last_triggered_slot = None


async def trigger_schedule_event(time_str: str, now_ict: datetime.datetime):
    channel = None
    channel_id_env = os.environ.get('WORKLOG_DISCORD_CHANNEL_ID', '').strip()
    if channel_id_env and channel_id_env.isdigit():
        channel = bot.get_channel(int(channel_id_env))

    if not channel:
        for g in bot.guilds:
            for ch in g.text_channels:
                if ch.permissions_for(g.me).send_messages:
                    channel = ch
                    break
            if channel:
                break

    if not channel:
        print(f"[SCHEDULE] No accessible text channel found to send notification for {time_str}", flush=True)
        return

    if time_str == "08:30":
        embed = discord.Embed(
            title="⚔️ MORNING QUEST BEGINS (8:30 AM UTC+7)",
            description="Good morning, Hero! 🌅 A new day of adventure awaits!\nUse `/start` to launch your first quest of the day.",
            color=0x3498db,
            timestamp=now_ict
        )
        embed.set_footer(text="🎮 WorkLog RPG Scheduler • Mon-Fri")
        await channel.send(embed=embed)

    elif time_str == "13:30":
        embed = discord.Embed(
            title="⚔️ AFTERNOON QUEST RESUMES (1:30 PM UTC+7)",
            description="Good afternoon, Hero! ☀️ Time to resume the adventure!\nUse `/start` or `/continue` to rejoin your quest.",
            color=0x3498db,
            timestamp=now_ict
        )
        embed.set_footer(text="🎮 WorkLog RPG Scheduler • Mon-Fri")
        await channel.send(embed=embed)

    elif time_str in ("12:00", "17:30"):
        session_name = "Morning Session (12:00 PM)" if time_str == "12:00" else "End of Workday (5:30 PM)"
        state = tm.load_state()
        timers = state.get('timers', [])
        running_timers = [t for t in timers if not t.get('paused')]

        if running_timers:
            out, ret = tm.cmd_pause(reason=f"Auto-pause: {session_name}")
            clear_bot_cache()

            embed = discord.Embed(
                title=f"⏸️ CAMPFIRE REST — Quest Auto-Paused ({session_name})",
                description=(
                    f"The scheduled quest session has ended at **{time_str} (UTC+7)**.\n"
                    f"Your active quest battles have been **automatically paused** at the nearest campfire.\n\n"
                    f"**Do you want to continue the adventure / enter overtime quest?**"
                ),
                color=0xf1c40f,
                timestamp=now_ict
            )
            out_str = str(out)[:950]
            embed.add_field(name="⏸️ Campfire Checkpoint Output", value=f"```\n{out_str}\n```", inline=False)
            embed.set_footer(text="Select an option below. If no action is taken, quest remains safely paused at campfire.")

            view = WorklogOvertimeView(session_name=session_name)
            await channel.send(embed=embed, view=view)
        else:
            embed = discord.Embed(
                title=f"🔔 QUEST SESSION ENDED ({session_name})",
                description=f"The scheduled quest session ended at **{time_str} (UTC+7)**. No active quest battles were running.",
                color=0x95a5a6,
                timestamp=now_ict
            )
            embed.set_footer(text="🎮 WorkLog RPG Scheduler • Mon-Fri")
            await channel.send(embed=embed)


@ext_tasks.loop(seconds=30)
async def check_work_schedule():
    global _last_triggered_slot
    try:
        now_ict = get_ict_now()
        is_weekend = now_ict.weekday() >= 5
        time_num = now_ict.hour * 60 + now_ict.minute

        # Off-hours defined as:
        # Weekend OR Lunch break (12:00 <= time < 13:30) OR Night off (time >= 17:30 or time < 08:30)
        is_lunch_break = (12 * 60 <= time_num < 13 * 60 + 30)
        is_night_off = (time_num >= 17 * 60 + 30 or time_num < 8 * 60 + 30)
        is_off_hours = is_weekend or is_lunch_break or is_night_off

        # Safety Catch-Up: If we are in off-hours AND there are active running timers, auto-pause them immediately!
        if is_off_hours:
            state = tm.load_state()
            timers = state.get('timers', [])
            running_timers = [t for t in timers if not t.get('paused')]
            if running_timers:
                session_label = "Weekend" if is_weekend else ("Lunch Break (12:00-13:30)" if is_lunch_break else "Off-Hours")
                print(f"[SCHEDULE AUTO-GUARD] Auto-pausing {len(running_timers)} timer(s) running during {session_label}", flush=True)
                slot_time = "12:00" if is_lunch_break else "17:30"
                await trigger_schedule_event(slot_time, now_ict)
                return

        if is_weekend:
            return

        time_str = now_ict.strftime("%H:%M")
        date_slot = f"{now_ict.strftime('%Y-%m-%d')}_{time_str}"

        if _last_triggered_slot == date_slot:
            return

        if time_str in ("08:30", "12:00", "13:30", "17:30"):
            _last_triggered_slot = date_slot
            await trigger_schedule_event(time_str, now_ict)
    except Exception as e:
        print(f"[SCHEDULE ERROR] Exception in work schedule loop: {e}", flush=True)


@bot.tree.command(name="test_schedule", description="Test schedule event trigger (08:30, 12:00, 13:30, 17:30)")
@app_commands.describe(time_slot="Time slot: 08:30, 12:00, 13:30, 17:30")
async def test_schedule(interaction: discord.Interaction, time_slot: str = "12:00"):
    deferred = await _defer(interaction)
    now_ict = get_ict_now()
    await trigger_schedule_event(time_slot, now_ict)
    await interaction.followup.send(f"✅ Triggered test schedule event for `{time_slot}` (UTC+7)", ephemeral=True)


_hud_msg_channel_id = None
_hud_msg_id = None


@bot.tree.command(name="hud", description="Spawn or view the Live Gamified RPG Quest HUD")
async def hud_cmd(interaction: discord.Interaction):
    global _hud_msg_channel_id, _hud_msg_id
    u_id = str(interaction.user.id)
    deferred = await _defer(interaction, ephemeral=True)
    embed = build_gamified_hud_embed(user_id=u_id)
    msg = await interaction.followup.send(embed=embed, ephemeral=True)
    if msg:
        _hud_msg_channel_id = msg.channel.id
        _hud_msg_id = msg.id


@ext_tasks.loop(seconds=15)
async def update_live_hud_loop():
    global _hud_msg_channel_id, _hud_msg_id
    if not _hud_msg_channel_id or not _hud_msg_id:
        return
    try:
        channel = bot.get_channel(_hud_msg_channel_id)
        if not channel:
            return
        msg = await channel.fetch_message(_hud_msg_id)
        if not msg:
            return

        embed = build_gamified_hud_embed()
        await msg.edit(embed=embed)
    except Exception:
        pass


@bot.event
async def on_ready():
    print(f"WorkLog Bot logged in as {bot.user}", flush=True)
    if not check_work_schedule.is_running():
        check_work_schedule.start()
        print("[SCHEDULE] Work schedule background loop started (UTC+7 Mon-Fri)", flush=True)
    if not update_live_hud_loop.is_running():
        update_live_hud_loop.start()
        print("[HUD] Live progress HUD loop started (every 15s)", flush=True)
    try:
        for guild in bot.guilds:
            bot.tree.copy_global_to(guild=guild)
            synced = await bot.tree.sync(guild=guild)
            print(f"Synced {len(synced)} command(s) to guild {guild.name} ({guild.id})", flush=True)
    except Exception as e:
        print(f"Failed to sync slash commands: {e}", flush=True)


if __name__ == '__main__':
    enforce_single_instance()
    if not TOKEN:
        print("Error: WORKLOG_DISCORD_TOKEN not set.")
        sys.exit(1)
    bot.run(TOKEN)
