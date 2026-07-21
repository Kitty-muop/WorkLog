"""Webhook notifications for WorkLog.

Sends a POST to WORKLOG_WEBHOOK_URL on timer stop.
Supports Slack/Discord compatible JSON format.
"""
import os
import json
import urllib.request
import urllib.error
import threading

WEBHOOK_URL = os.environ.get('WORKLOG_WEBHOOK_URL', '')


def send(project=None, task=None, duration_h=0, description='',
         category='', subtask='', break_info='', start_time='', end_time=''):
    if not WEBHOOK_URL:
        return False
    payload = {
        'project': project or '',
        'task': task or '',
        'duration_h': round(duration_h, 2),
        'duration_m': int(duration_h * 60),
        'description': description or '',
        'category': category or '',
        'subtask': subtask or '',
        'break_info': break_info or '',
        'start_time': start_time,
        'end_time': end_time,
        'source': 'worklog',
    }
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
