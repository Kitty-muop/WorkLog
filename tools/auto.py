#!/usr/bin/env python3
"""
WorkLog Automation Suite - Daily Scorecard
Runs validate + gamify + report in a single process.
"""
import sys, json
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

import validate
import gamify
import db
from report import run as run_report

def main():
    db.init()
    print("\n\x1b[1m=== WORKLOG AUTO-RUN ===\x1b[0m\n")

    print("--- Running validate.py ---")
    issues = validate.run()
    for i in issues:
        print(f"  [{i['severity'].upper()}] {i['type']}: {i['detail']}")
    if not issues:
        print("  No issues found.")
    print()

    print("--- Running gamify.py ---")
    g = gamify.run()
    print(json.dumps(g, indent=2, default=str))
    print()

    print("--- Running report.py ---")
    run_report(gamify_result=g)

if __name__ == '__main__':
    main()
