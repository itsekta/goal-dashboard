from flask import Flask, jsonify, render_template
import openpyxl
import hashlib
import re
import os
from datetime import datetime

app = Flask(__name__)

EXCEL_FILE = "/Users/ektachaudhary/Downloads/Goal_progress_June.xlsx"
MIN_TICKETS = 6

DATE_LABELS = {
    1: "Mon Jun 1",  2: "Tue Jun 2",  3: "Wed Jun 3",
    4: "Thu Jun 4",  5: "Fri Jun 5",
    8: "Mon Jun 9",  9: "Tue Jun 10", 10: "Wed Jun 11",
    11: "Thu Jun 12", 12: "Fri Jun 13",
    13: "Mon Jun 16", 14: "Tue Jun 17", 15: "Wed Jun 18",
    16: "Thu Jun 19", 17: "Fri Jun 20",
    18: "Mon Jun 23", 19: "Tue Jun 24",
}
WORK_COLS = sorted(DATE_LABELS.keys())

# In-memory state
_state = {"hash": None, "rows": None, "changed_cells": set(), "change_log": []}


def file_hash(path):
    with open(path, "rb") as f:
        return hashlib.md5(f.read()).hexdigest()


def load_rows(path):
    wb = openpyxl.load_workbook(path)
    return [list(r) for r in wb["Sheet1"].iter_rows(values_only=True)]


def detect_changes(old, new):
    changed = set()
    if not old:
        return changed
    for ri in range(min(len(old), len(new))):
        for ci in range(min(len(old[ri]), len(new[ri]))):
            if old[ri][ci] != new[ri][ci]:
                changed.add((ri, ci))
    return changed


def count_tickets(val):
    if val is None:
        return 0, []
    s = str(val).strip()
    if not s:
        return 0, []
    if "on leave" in s.lower():
        return -1, ["On Leave"]
    items = [t.strip() for t in re.split(r"\n+", s) if t.strip()]
    return len(items), items


def is_done(val):
    return bool(val and str(val).strip().lower().startswith("done"))


def get_data():
    current_hash = file_hash(EXCEL_FILE)
    file_changed = current_hash != _state["hash"]

    if file_changed:
        new_rows = load_rows(EXCEL_FILE)
        new_changed = detect_changes(_state["rows"], new_rows)

        for (ri, ci) in new_changed:
            old_v = _state["rows"][ri][ci] if _state["rows"] and ri < len(_state["rows"]) and ci < len(_state["rows"][ri]) else None
            new_v = new_rows[ri][ci] if ri < len(new_rows) and ci < len(new_rows[ri]) else None
            date = DATE_LABELS.get(ci, f"col {ci}")
            _state["change_log"].insert(0, {
                "date": date, "row": ri, "col": ci,
                "from": str(old_v) if old_v is not None else "—",
                "to": str(new_v)[:60] if new_v is not None else "—",
            })

        _state["rows"] = new_rows
        _state["hash"] = current_hash
        _state["changed_cells"] = _state["changed_cells"] | new_changed

    rows = _state["rows"]
    changed = _state["changed_cells"]

    def safe(idx):
        return rows[idx] if rows and idx < len(rows) else []

    r1 = safe(1)   # Goal 1 tickets
    r4 = safe(4)   # Goal 2 engagement
    r6 = safe(6)   # Goal 2 responsiveness
    r9 = safe(9)   # Goal 3 RTO
    r12 = safe(12) # Goal 4 growth
    r14 = safe(14) # Others

    def cell(row, ci):
        return row[ci] if ci < len(row) else None

    # Goal 1
    goal1 = []
    for ci in WORK_COLS:
        cnt, tix = count_tickets(cell(r1, ci))
        goal1.append({
            "date": DATE_LABELS[ci], "count": cnt,
            "tickets": tix, "is_leave": cnt == -1,
            "changed": (1, ci) in changed,
        })

    # Goal 2
    engagement, responsiveness = [], []
    for ci in WORK_COLS:
        engagement.append({"date": DATE_LABELS[ci], "done": is_done(cell(r4, ci)), "changed": (4, ci) in changed})
        responsiveness.append({"date": DATE_LABELS[ci], "done": is_done(cell(r6, ci)), "changed": (6, ci) in changed})

    # Goal 3
    rto = []
    for ci in WORK_COLS:
        rto.append({"date": DATE_LABELS[ci], "done": is_done(cell(r9, ci)), "changed": (9, ci) in changed})

    # Goal 4
    growth = []
    for ci in WORK_COLS:
        val = cell(r12, ci)
        if val and str(val).strip():
            tools = [t.strip() for t in str(val).split("\n") if t.strip()]
            growth.append({"date": DATE_LABELS[ci], "tools": tools, "changed": (12, ci) in changed})

    # Others
    others = []
    for ci in WORK_COLS:
        val = cell(r14, ci)
        if val and str(val).strip():
            acts = [a.strip() for a in str(val).split("\n") if a.strip()]
            others.append({"date": DATE_LABELS[ci], "activities": acts, "changed": (14, ci) in changed})

    # Stats
    tracked = [d for d in goal1 if d["count"] != 0]
    leave = [d for d in goal1 if d["is_leave"]]
    met = [d for d in goal1 if d["count"] >= MIN_TICKETS]

    mod_time = datetime.fromtimestamp(os.path.getmtime(EXCEL_FILE)).strftime("%b %d, %Y %H:%M")

    return {
        "meta": {
            "last_modified": mod_time,
            "file_changed": file_changed,
            "change_count": len(_state["change_log"]),
        },
        "stats": {
            "goal1_met": len(met),
            "goal1_tracked": len(tracked) - len(leave),
            "goal2_engage": sum(1 for d in engagement if d["done"]),
            "goal2_respond": sum(1 for d in responsiveness if d["done"]),
            "goal3_rto": sum(1 for d in rto if d["done"]),
            "goal4_training": len(growth),
            "total_days": len(WORK_COLS),
        },
        "goal1": goal1,
        "goal2": {"engagement": engagement, "responsiveness": responsiveness},
        "goal3": rto,
        "goal4": growth,
        "others": others,
        "change_log": _state["change_log"][:20],
    }


@app.route("/")
def index():
    return render_template("dashboard.html")


@app.route("/api/data")
def api_data():
    return jsonify(get_data())


@app.route("/api/clear-changes", methods=["POST"])
def clear_changes():
    _state["changed_cells"] = set()
    _state["change_log"] = []
    return jsonify({"ok": True})


if __name__ == "__main__":
    print("\n  Goal Progress Dashboard")
    print("  Running at: http://localhost:8080\n")
    app.run(debug=True, port=8080, use_reloader=False)
