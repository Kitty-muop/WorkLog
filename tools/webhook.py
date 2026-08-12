"""Webhook notifications for WorkLog.

Sends a POST to WORKLOG_WEBHOOK_URL on timer stop.
Supports Slack/Discord compatible JSON format.
"""
import os
import json
import datetime
import urllib.request
import urllib.error
import threading

WEBHOOK_URL = os.environ.get('WORKLOG_WEBHOOK_URL', '')


def _build_payload(project=None, task=None, duration_h=0, description='',
                   category='', subtask='', break_info='', start_time='', end_time=''):
    """Build payload in Discord embed format or plain JSON."""
    is_discord = 'discord.com/api/webhooks' in WEBHOOK_URL.lower()
    duration_m = int(duration_h * 60)
    duration_str = f'{int(duration_h)}h{int(duration_m % 60)}m' if duration_h >= 1 else f'{duration_m}m'

    fields = []
    if project:
        fields.append({'name': 'Project', 'value': project, 'inline': True})
    if task:
        fields.append({'name': 'Father Task', 'value': task, 'inline': True})
    if subtask:
        fields.append({'name': 'Sub Task', 'value': subtask, 'inline': True})
    if category:
        fields.append({'name': 'Category', 'value': category, 'inline': True})
    if start_time and end_time:
        fields.append({'name': 'Time', 'value': f'{start_time} - {end_time}', 'inline': True})
    if break_info:
        fields.append({'name': 'Breaks', 'value': break_info, 'inline': False})

    color = 0x57F287 if duration_h >= 1 else 0xFEE75C  # green / yellow

    if is_discord:
        return {
            'embeds': [{
                'title': 'Timer Stopped',
                'description': description or 'No description',
                'color': color,
                'fields': fields,
                'footer': {'text': f'Duration: {duration_str}'},
                'timestamp': datetime.datetime.now(datetime.timezone.utc).isoformat(),
            }]
        }
    else:
        return {
            'project': project or '',
            'task': task or '',
            'duration_h': round(duration_h, 2),
            'duration_m': duration_m,
            'description': description or '',
            'category': category or '',
            'subtask': subtask or '',
            'break_info': break_info or '',
            'start_time': start_time,
            'end_time': end_time,
            'source': 'worklog',
        }


def send(project=None, task=None, duration_h=0, description='',
         category='', subtask='', break_info='', start_time='', end_time=''):
    if not WEBHOOK_URL:
        return False
    payload = _build_payload(project, task, duration_h, description,
                             category, subtask, break_info, start_time, end_time)
    threading.Thread(target=_post, args=(payload,), daemon=True).start()
    return True


def send_test(args=None):
    if not WEBHOOK_URL:
        print("No WORKLOG_WEBHOOK_URL set.")
        print("  Set environment variable: $env:WORKLOG_WEBHOOK_URL = 'https://hooks.example.com/webhook'")
        return 1
    print(f"Sending test to {WEBHOOK_URL}...")
    ok = send(
        project='TestProject',
        task='TestTask',
        duration_h=1.5,
        description='Webhook test notification',
        category='Testing',
        break_info='break(5m)',
        start_time='09:00',
        end_time='10:30',
    )
    if ok:
        print("Test webhook sent (async). Check your webhook endpoint.")
    return 0


def _post(payload):
    try:
        data = json.dumps(payload, ensure_ascii=False).encode('utf-8')
        req = urllib.request.Request(
            WEBHOOK_URL,
            data=data,
            headers={'Content-Type': 'application/json'},
            method='POST'
        )
        urllib.request.urlopen(req, timeout=5)
    except urllib.error.HTTPError as e:
        print(f"  [webhook] HTTP {e.code}: {e.reason}")
    except urllib.error.URLError as e:
        print(f"  [webhook] Connection failed: {e.reason}")
    except Exception as e:
        print(f"  [webhook] {e}")
