from flask import Flask, jsonify, render_template
import openpyxl
import hashlib
import re
import os
from datetime import datetime, date, timedelta

app = Flask(__name__)

EXCEL_FILE = "/Users/ektachaudhary/Downloads/Goal_progress_June.xlsx"
MIN_TICKETS = 6

# Anchor: col 4 in the header row = "Thurs - 4th" = June 4, 2026
ANCHOR_COL  = 4
ANCHOR_DATE = date(2026, 6, 4)


def build_date_labels(header_row):
    """Dynamically derive date labels from the header row.
    Uses the header's day-of-week text to snap each column to the correct date,
    so it handles weeks where Sat/Sun columns are absent. Any new columns added
    to the spreadsheet are automatically picked up on the next reload."""
    # Each entry: (list_of_substrings_to_match, day_of_week_int)
    # Thursday checks both 'thu' and 'thru' to handle "Thrusday" typos
    DOW = [
        (['mon'],        0),
        (['tue'],        1),
        (['wed'],        2),
        (['thu', 'thru'], 3),
        (['fri'],        4),
        (['sat'],        5),
        (['sun'],        6),
    ]

    # Build (col_index, day_of_week) list for all non-label columns
    col_dows = []
    for ci, val in enumerate(header_row):
        if ci == 0 or not val:
            continue
        s = str(val).lower()
        for keys, d in DOW:
            if any(k in s for k in keys):
                col_dows.append((ci, d))
                break

    anchor_idx = next((i for i, (ci, _) in enumerate(col_dows) if ci == ANCHOR_COL), None)
    if anchor_idx is None:
        return {}, []

    date_labels = {}
    work_cols   = []

    # Walk forward from anchor — advance current date until day-of-week matches header
    cur = ANCHOR_DATE
    for ci, dow in col_dows[anchor_idx:]:
        while cur.weekday() != dow:
            cur += timedelta(days=1)
        if dow < 5:  # Mon–Fri only
            date_labels[ci] = f"{cur.strftime('%a')} {cur.strftime('%b')} {cur.day}"
            work_cols.append(ci)
        cur += timedelta(days=1)

    # Walk backward from anchor
    cur = ANCHOR_DATE - timedelta(days=1)
    for ci, dow in reversed(col_dows[:anchor_idx]):
        while cur.weekday() != dow:
            cur -= timedelta(days=1)
        if dow < 5:
            date_labels[ci] = f"{cur.strftime('%a')} {cur.strftime('%b')} {cur.day}"
            work_cols.append(ci)
        cur -= timedelta(days=1)

    return date_labels, sorted(work_cols)

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

        # Build dynamic labels from the new header row
        new_labels, _ = build_date_labels(new_rows[0])
        for (ri, ci) in new_changed:
            old_v = _state["rows"][ri][ci] if _state["rows"] and ri < len(_state["rows"]) and ci < len(_state["rows"][ri]) else None
            new_v = new_rows[ri][ci] if ri < len(new_rows) and ci < len(new_rows[ri]) else None
            _state["change_log"].insert(0, {
                "date": new_labels.get(ci, f"col {ci}"), "row": ri, "col": ci,
                "from": str(old_v) if old_v is not None else "—",
                "to": str(new_v)[:60] if new_v is not None else "—",
            })

        _state["rows"] = new_rows
        _state["hash"] = current_hash
        _state["changed_cells"] = _state["changed_cells"] | new_changed

    rows = _state["rows"]
    changed = _state["changed_cells"]

    # Always rebuild from the current header so new columns are picked up live
    date_labels, work_cols = build_date_labels(rows[0])

    def safe(idx):
        return rows[idx] if rows and idx < len(rows) else []

    r1  = safe(1)   # Goal 1 tickets
    r4  = safe(4)   # Goal 2 engagement
    r6  = safe(6)   # Goal 2 responsiveness
    r9  = safe(9)   # Goal 3 RTO
    r12 = safe(12)  # Goal 4 growth
    r14 = safe(14)  # Others

    def cell(row, ci):
        return row[ci] if ci < len(row) else None

    # Goal 1
    goal1 = []
    for ci in work_cols:
        cnt, tix = count_tickets(cell(r1, ci))
        goal1.append({
            "date": date_labels[ci], "count": cnt,
            "tickets": tix, "is_leave": cnt == -1,
            "changed": (1, ci) in changed,
        })

    # Goal 2
    engagement, responsiveness = [], []
    for ci in work_cols:
        engagement.append({"date": date_labels[ci], "done": is_done(cell(r4, ci)), "changed": (4, ci) in changed})
        responsiveness.append({"date": date_labels[ci], "done": is_done(cell(r6, ci)), "changed": (6, ci) in changed})

    # Goal 3
    rto = []
    for ci in work_cols:
        rto.append({"date": date_labels[ci], "done": is_done(cell(r9, ci)), "changed": (9, ci) in changed})

    # Goal 4
    growth = []
    for ci in work_cols:
        val = cell(r12, ci)
        if val and str(val).strip():
            tools = [t.strip() for t in str(val).split("\n") if t.strip()]
            growth.append({"date": date_labels[ci], "tools": tools, "changed": (12, ci) in changed})

    # Others
    others = []
    for ci in work_cols:
        val = cell(r14, ci)
        if val and str(val).strip():
            acts = [a.strip() for a in str(val).split("\n") if a.strip()]
            others.append({"date": date_labels[ci], "activities": acts, "changed": (14, ci) in changed})

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
            "total_days": len(work_cols),
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
