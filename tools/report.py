import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
import validate
import gamify

BOLD = "\033[1m"
END = "\033[0m"
RED = "\033[91m"
GREEN = "\033[92m"
YELLOW = "\033[93m"

def title(text):
    print(f"\n{BOLD}=== {text} ==={END}")

def run(gamify_result=None):
    if gamify_result is None:
        gamify_result = gamify.run()

    print(f"{BOLD}{'='*50}{END}")
    print(f"{BOLD}{' ' * 15}W O R K L O G  R E P O R T{END}")
    print(f"{BOLD}{'='*50}{END}")

    issues = validate.run()
    title("VALIDATION")
    if issues:
        for i in issues:
            c = RED if i["severity"] == "high" else (YELLOW if i["severity"] == "medium" else GREEN)
            print(f"  [{i['severity'].upper()}] {i['type']}: {i['detail']}")
    else:
        print("  No issues found.")

    g = gamify_result
    title("GAMIFICATION SCORES")
    print(f"  Level:       {g['level']} - {g['level_name']}")
    print(f"  Total Score: {g['total_score']}")
    print(f"  Today Score: +{g['today_score']}")
    print(f"  Streak:       {g['streak']}d (max: {g['max_streak']}d)")
    print(f"  Consistency:  {g['consistency_pct']}% ({g['logged_weekdays']}/{g['total_weekdays']})")

    rating = "Excellent" if g['consistency_pct'] >= 90 else ("Good" if g['consistency_pct'] >= 70 else ("Fair" if g['consistency_pct'] >= 40 else "Needs improvement"))
    print(f"  Rating:      {rating}")

    title("TASK PERFORMANCE")
    by_project = {}
    for task, t in g['tasks'].items():
        by_project.setdefault(t['project'], []).append((task, t))
    for proj in sorted(by_project.keys()):
        print(f"  [{proj}]")
        for task, t in sorted(by_project[proj]):
            pct = t['performance_pct']
            c = GREEN if pct >= 80 else (YELLOW if pct >= 50 else RED)
            s = "DONE" if t['status'] == "Done" else f"{c}{t['status']}{END}"
            code = t.get('code', '')
            print(f"    {code:<18} {task:<25} {t['logged_hours']:.1f}/{t['deadline_hours']}h = {c}{pct}%{END} - {s}")

    total = sum(t['logged_hours'] for t in g['tasks'].values())
    print(f"\n{BOLD}{'='*50}{END}")
    print(f"  {BOLD}Total Hours: {total}h{END}")
    print(f"{'='*50}{END}")

if __name__ == '__main__':
    run()
