#!/usr/bin/env python3
"""
generate.py
Fetches open Action Items and Risks from the Marketing Studio / Targeted Offer RAID Log
(Smartsheet sheet ID 980305451110276) and regenerates index.html for GitHub Pages.

Required env var:  SMARTSHEET_TOKEN  (a Smartsheet personal access token)
"""

import os, json, time, requests
from datetime import date

# ─── Config ───────────────────────────────────────────────────────────────────
SHEET_ID    = "980305451110276"
TOKEN       = os.environ["SMARTSHEET_TOKEN"]
TODAY       = date.today()
LAST_SYNCED = TODAY.strftime("%B %d, %Y")
SS_ALPHA_ID = "3XFh8vH6VwcrWhH2Jw54J44hMX5G7JXfHmGQX8x1"
SS_BASE     = f"https://app.smartsheet.com/sheets/{SS_ALPHA_ID}"
SHEET_URL   = f"{SS_BASE}?view=grid"
BINDER_URL  = "https://docs.google.com/spreadsheets/d/1ThmYPfgTH_zhlmuJr63H47UzonJquO6QWHxkvXrP9tw/edit?gid=565365486#gid=565365486"

MS_CONV_SHEET_ID = "905349279207300"
CO_WBS_SHEET_ID  = "3526224222572420"
MS_CONV_SS_URL   = "https://app.smartsheet.com/sheets/WcJRFcqhF6qJR8fHWMjJ5Jx99q8qrCvxFFg3Pc41"
CO_WBS_SS_URL    = "https://app.smartsheet.com/sheets/vp5r93RjCf9QvVPw38mJM627Xg66gVPpFwR97Rh1"

INCLUDE_TYPES = {"Action Item", "Risk"}

# ─── Fetch sheet (with retries for transient 5xx errors) ─────────────────────
def fetch_sheet(retries=4, backoff=5):
    headers = {"Authorization": f"Bearer {TOKEN}", "Accept": "application/json"}
    for attempt in range(1, retries + 1):
        try:
            r = requests.get(
                f"https://api.smartsheet.com/2.0/sheets/{SHEET_ID}",
                headers=headers,
                params={"include": "objectValue"},
                timeout=30
            )
            r.raise_for_status()
            return r.json()
        except requests.exceptions.HTTPError as e:
            status = e.response.status_code if e.response is not None else None
            if status and status >= 500 and attempt < retries:
                wait = backoff * attempt
                print(f"  Smartsheet returned {status} on attempt {attempt}/{retries} — retrying in {wait}s...")
                time.sleep(wait)
            else:
                raise

# ─── Surrogate-safe string sanitizer ─────────────────────────────────────────
def safe_str(text):
    """Remove lone surrogate characters that can't be encoded as UTF-8."""
    return text.encode("utf-8", errors="replace").decode("utf-8")

# ─── Safe cell value extraction ───────────────────────────────────────────────
def cell_text(cell):
    if not cell:
        return ""
    obj = cell.get("objectValue") or {}
    if isinstance(obj, dict):
        name = obj.get("name") or obj.get("email") or ""
        if name:
            return safe_str(name.strip())
    raw = str(cell.get("displayValue") or cell.get("value") or "").strip()
    return safe_str(raw)

# ─── Parse owners ─────────────────────────────────────────────────────────────
def parse_owners(sheet):
    col = {}
    for c in sheet.get("columns", []):
        col[c["title"].strip()] = c["id"]

    type_col   = col.get("Type")
    action_col = col.get("Action Item")
    owner_col  = col.get("Owner")
    status_col = col.get("Status")
    due_col    = col.get("Est. Completion")

    print(f"  Column IDs found: Type={type_col}, Action Item={action_col}, "
          f"Owner={owner_col}, Status={status_col}, Due={due_col}")

    owners_dict = {}
    skipped = 0

    for row in sheet.get("rows", []):
        cells = {}
        for cell in row.get("cells", []):
            cells[cell.get("columnId")] = cell

        type_val   = cell_text(cells.get(type_col))
        action_txt = cell_text(cells.get(action_col))
        owner_name = cell_text(cells.get(owner_col)) or "Unassigned"
        status_val = cell_text(cells.get(status_col))

        if type_val not in INCLUDE_TYPES:
            skipped += 1
            continue
        if status_val == "Completed":
            skipped += 1
            continue
        if not action_txt:
            skipped += 1
            continue

        is_overdue    = False
        overdue_label = ""
        due_cell      = cells.get(due_col)
        if due_cell:
            due_str = str(due_cell.get("value") or "")[:10]
            if due_str:
                try:
                    due_date = date.fromisoformat(due_str)
                    if due_date < TODAY:
                        is_overdue    = True
                        overdue_label = f"Due {due_date.strftime('%b %-d')} passed"
                except Exception:
                    pass

        row_id = str(row["id"])
        if owner_name not in owners_dict:
            owners_dict[owner_name] = {"items": [], "first_overdue": None}

        item = {"t": action_txt, "rowId": row_id}
        if is_overdue:
            item["od"] = True
            if owners_dict[owner_name]["first_overdue"] is None:
                owners_dict[owner_name]["first_overdue"] = overdue_label

        owners_dict[owner_name]["items"].append(item)

    print(f"  Parsed {len(owners_dict)} owners, skipped {skipped} rows")

    result = []
    for name, data in sorted(owners_dict.items(), key=lambda x: len(x[1]["items"]), reverse=True):
        items         = data["items"]
        overdue_count = sum(1 for i in items if i.get("od"))
        result.append({
            "name":         name,
            "total":        len(items),
            "overdue":      overdue_count,
            "overdueDetail": data["first_overdue"] or "",
            "items":        items,
        })

    for o in result:
        print(f"    {o['name']}: {o['total']} items, {o['overdue']} overdue")

    return result

# ─── Fetch WBS sheet (same retry logic, sheet_id as parameter) ───────────────
def fetch_wbs(sheet_id, retries=4, backoff=5):
    headers = {"Authorization": f"Bearer {TOKEN}", "Accept": "application/json"}
    for attempt in range(1, retries + 1):
        try:
            r = requests.get(
                f"https://api.smartsheet.com/2.0/sheets/{sheet_id}",
                headers=headers,
                params={"include": "objectValue"},
                timeout=30
            )
            r.raise_for_status()
            return r.json()
        except requests.exceptions.HTTPError as e:
            status = e.response.status_code if e.response is not None else None
            if status and status >= 500 and attempt < retries:
                wait = backoff * attempt
                print(f"  Smartsheet returned {status} on attempt {attempt}/{retries} — retrying in {wait}s...")
                time.sleep(wait)
            else:
                raise

# ─── Parse WBS sheet ──────────────────────────────────────────────────────────
def parse_wbs(sheet):
    """Parse WBS sheet into milestone groups.
    Order-independent: groups tasks by bracket-extracted milestone name so the
    result is correct whether the API returns rows in natural or alphabetical order.
    """
    STATUS_NORM = {
        "Completed":     "completed",
        "In Progress":   "inProgress",
        "Not Started":   "notStarted",
        "Blocked":       "blocked",
        "Not Applicable":"notStarted",
        "":              "notStarted",
    }

    col = {}
    for c in sheet.get("columns", []):
        col[c["title"].strip()] = c["id"]

    primary_col = col.get("Milestone / Task")
    owner_col   = col.get("Owner")
    status_col  = col.get("Status")
    date_col    = col.get("Target Date")

    # Pass 1: collect section header names in sheet order, and bucket tasks by milestone
    from collections import defaultdict
    header_order = []   # milestone names in appearance order
    task_map = defaultdict(list)  # milestone_name -> [task, ...]

    for row in sheet.get("rows", []):
        cells = {}
        for cell in row.get("cells", []):
            cells[cell.get("columnId")] = cell

        primary_val = cell_text(cells.get(primary_col)) if primary_col else ""
        if not primary_val:
            continue
        if primary_val.startswith("▶"):
            continue

        if not primary_val.startswith("["):
            # Section header row
            if primary_val not in header_order:
                header_order.append(primary_val)
        else:
            # Task row — extract milestone name from [MilestoneName] prefix
            bracket_end = primary_val.find("]")
            if bracket_end == -1:
                ms_name   = primary_val
                task_desc = primary_val
            else:
                ms_name   = primary_val[1:bracket_end].strip()
                task_desc = primary_val[bracket_end + 1:].strip()

            owner      = cell_text(cells.get(owner_col))  if owner_col  else ""
            status_raw = cell_text(cells.get(status_col)) if status_col else ""
            status     = STATUS_NORM.get(status_raw, "notStarted")
            date_val   = cell_text(cells.get(date_col))   if date_col   else ""

            task_map[ms_name].append({
                "task": safe_str(task_desc),
                "owner": safe_str(owner),
                "status": status,
                "date": safe_str(date_val),
            })

    # Pass 2: assemble groups in header order; append any orphan groups last
    seen = set()
    groups = []
    totals = {"completed": 0, "inProgress": 0, "notStarted": 0, "blocked": 0}

    for name in header_order:
        if name in seen:
            continue
        seen.add(name)
        tasks = task_map.get(name, [])
        if not tasks:
            continue
        counts = {"completed": 0, "inProgress": 0, "notStarted": 0, "blocked": 0}
        for t in tasks:
            counts[t["status"]] = counts.get(t["status"], 0) + 1
            totals[t["status"]] = totals.get(t["status"], 0) + 1
        groups.append({"name": name, "tasks": tasks, "counts": counts})

    # Orphan tasks whose header wasn't in the sheet
    for ms_name, tasks in task_map.items():
        if ms_name in seen:
            continue
        counts = {"completed": 0, "inProgress": 0, "notStarted": 0, "blocked": 0}
        for t in tasks:
            counts[t["status"]] = counts.get(t["status"], 0) + 1
            totals[t["status"]] = totals.get(t["status"], 0) + 1
        groups.append({"name": ms_name, "tasks": tasks, "counts": counts})

    return {"groups": groups, "totals": totals}

# ─── HTML template ────────────────────────────────────────────────────────────
HTML_TEMPLATE = """\
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Marketing Studio &middot; Program Dashboard</title>
<style>
:root {
  --brand:#4f46e5; --brand-light:#ede9fe; --brand-mid:#c7d2fe;
  --red:#ef4444; --red-light:#fef2f2; --red-mid:#fecaca;
  --green:#10b981; --green-light:#ecfdf5; --green-mid:#6ee7b7;
  --amber:#f59e0b; --amber-light:#fffbeb; --amber-mid:#fde68a;
  --blue:#3b82f6; --blue-light:#eff6ff;
  --purple:#7c3aed; --purple-light:#f5f3ff;
  --gray:#6b7280; --gray-light:#f3f4f6; --border:#e5e7eb;
  --bg:#f9fafb; --surface:#ffffff; --text:#111827; --text-2:#374151;
}
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;background:var(--bg);color:var(--text);padding:20px 20px 60px;min-height:100vh}
.tab-bar{display:flex;gap:4px;margin-bottom:24px;border-bottom:2px solid var(--border);flex-wrap:wrap}
.tab{padding:8px 16px;font-size:12.5px;font-weight:600;cursor:pointer;border-radius:8px 8px 0 0;color:var(--gray);background:none;border:none;border-bottom:3px solid transparent;margin-bottom:-2px;transition:color .15s,border-color .15s;white-space:nowrap}
.tab:hover{color:var(--brand)}
.tab.active{color:var(--brand);border-bottom-color:var(--brand);background:var(--brand-light)}
.page{display:none}.page.active{display:block}
.page-header{display:flex;align-items:flex-start;justify-content:space-between;flex-wrap:wrap;gap:12px;margin-bottom:20px}
.header-left h1{font-size:19px;font-weight:800;letter-spacing:-.3px;color:var(--text)}
.header-left .sub{font-size:11.5px;color:var(--gray);margin-top:3px}
.ext-link{display:inline-flex;align-items:center;gap:5px;font-size:11.5px;font-weight:600;color:var(--brand);text-decoration:none;border:1.5px solid var(--brand-mid);border-radius:8px;padding:6px 13px;background:var(--brand-light);white-space:nowrap;transition:background .15s}
.ext-link:hover{background:#ddd6fe}
.section-label{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.09em;color:var(--gray);margin-bottom:8px;display:flex;align-items:center;gap:5px}
.section-title{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:var(--gray);margin-bottom:12px}
.section-divider{border:none;border-top:1px solid var(--border);margin:18px 0}
.summary-row{display:flex;gap:12px;margin-bottom:24px;flex-wrap:wrap}
.summary-card{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:12px 18px;text-align:center;box-shadow:0 1px 3px rgba(0,0,0,.05);min-width:110px;flex:1}
.summary-card .val{font-size:28px;font-weight:900;line-height:1}
.summary-card .lbl{font-size:9.5px;color:var(--gray);text-transform:uppercase;letter-spacing:.07em;margin-top:4px}
.val-brand{color:var(--brand)}.val-green{color:var(--green)}.val-red{color:var(--red)}
.rel-tiles{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:10px}
.rel-tile{display:flex;align-items:center;gap:10px;background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:10px 16px;box-shadow:0 1px 3px rgba(0,0,0,.04);flex:1;min-width:0}
.rel-tile-icon{font-size:18px;line-height:1}
.rel-tile-val{font-size:20px;font-weight:900;line-height:1}
.rel-tile-lbl{font-size:9px;color:var(--gray);text-transform:uppercase;letter-spacing:.06em;margin-top:2px}
.metric-strip{display:flex;flex-wrap:wrap;gap:0;background:var(--surface);border:1px solid var(--border);border-radius:14px;overflow:hidden;margin-bottom:22px;box-shadow:0 1px 4px rgba(0,0,0,.04)}
.metric-strip-item{flex:1;min-width:80px;padding:12px 14px;text-align:center;border-right:1px solid var(--border)}
.metric-strip-item:last-child{border-right:none}
.metric-strip-val{font-size:22px;font-weight:900;line-height:1}
.metric-strip-lbl{font-size:9px;color:var(--gray);text-transform:uppercase;letter-spacing:.06em;margin-top:3px}
.rel-two-col{display:grid;grid-template-columns:1fr 1.3fr;gap:16px;margin-bottom:22px;align-items:start}
.chart-card{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:18px 20px;box-shadow:0 1px 4px rgba(0,0,0,.05)}
.chart-card h3{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.07em;color:var(--gray);margin-bottom:14px}
.chart-wrap canvas{max-height:220px}
.status-legend{display:flex;flex-wrap:wrap;gap:7px;margin-top:14px}
.s-chip{display:inline-flex;align-items:center;gap:5px;font-size:10.5px;font-weight:600;padding:3px 9px;border-radius:20px}
.s-chip-dot{width:8px;height:8px;border-radius:50%;flex-shrink:0}
.s-done{background:var(--green-light);color:#065f46}
.s-prog{background:var(--amber-light);color:#92400e}
.s-block{background:var(--red-light);color:#991b1b}
.s-ns{background:var(--purple-light);color:#5b21b6}
.proj-card{background:var(--surface);border:1.5px solid var(--border);border-radius:14px;padding:14px 16px;cursor:pointer;transition:box-shadow .18s,border-color .18s,transform .18s;margin-bottom:10px;position:relative}
.proj-card:last-child{margin-bottom:0}
.proj-card:hover{box-shadow:0 6px 18px rgba(79,70,229,.1);border-color:var(--brand-mid);transform:translateY(-1px)}
.proj-card.active{border-color:var(--brand);box-shadow:0 0 0 3px var(--brand-mid);transform:none}
.proj-card-top{display:flex;align-items:center;justify-content:space-between;margin-bottom:10px}
.proj-card-name{font-size:13.5px;font-weight:800;color:var(--text)}
.proj-card-dates{font-size:10px;color:var(--gray);margin-top:1px}
.proj-card-right{display:flex;align-items:center;gap:8px}
.proj-total{font-size:20px;font-weight:900}
.proj-total-lbl{font-size:9px;color:var(--gray);text-transform:uppercase;letter-spacing:.06em;line-height:1.2}
.chevron{font-size:10px;color:var(--gray);transition:transform .2s;margin-left:4px}
.proj-card.active .chevron{transform:rotate(180deg);color:var(--brand)}
.proj-progress-bar{height:6px;border-radius:4px;background:var(--gray-light);overflow:hidden;margin-bottom:10px;display:flex;gap:1px}
.pb-seg{height:100%;border-radius:4px}
.proj-status-row{display:flex;gap:6px;flex-wrap:wrap}
.proj-detail{display:none;background:var(--surface);border:2px solid var(--brand);border-radius:14px;padding:18px 20px;margin-bottom:22px;box-shadow:0 4px 20px rgba(79,70,229,.1);animation:slideIn .2s ease}
.proj-detail.visible{display:block}
@keyframes slideIn{from{opacity:0;transform:translateY(-6px)}to{opacity:1;transform:none}}
.proj-detail-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:14px}
.proj-detail-title{font-size:15px;font-weight:800;color:var(--text)}
.close-btn{font-size:20px;cursor:pointer;color:var(--gray);border:none;background:none;padding:2px 7px;border-radius:6px;line-height:1}
.close-btn:hover{background:#f3f4f6}
.milestone-list{list-style:none}
.milestone-li{display:flex;align-items:center;gap:10px;padding:8px 0;border-bottom:1px solid #f3f4f6;font-size:12.5px}
.milestone-li:last-child{border-bottom:none}
.ms-num{flex-shrink:0;width:22px;height:22px;background:var(--brand-light);color:var(--brand);border-radius:6px;font-size:10px;font-weight:800;display:flex;align-items:center;justify-content:center}
.ms-label{flex:1;font-weight:500;color:var(--text-2)}
.ms-owner{font-size:11px;color:var(--gray);flex-shrink:0;min-width:80px;text-align:right}
.ms-end{font-size:11px;font-weight:600;color:var(--text);margin-left:8px;flex-shrink:0}
.chip{display:inline-block;font-size:9.5px;font-weight:700;padding:2px 7px;border-radius:20px;flex-shrink:0;margin-left:6px}
.chip-done{background:var(--green-light);color:#065f46}
.chip-prog{background:var(--amber-light);color:#92400e}
.chip-ns{background:var(--purple-light);color:#5b21b6}
.chip-block{background:var(--red-light);color:#991b1b}
.owner-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:12px;margin-bottom:24px}
.owner-card{background:var(--surface);border:1.5px solid var(--border);border-radius:12px;padding:14px 16px;cursor:pointer;position:relative;transition:box-shadow .18s,border-color .18s,transform .18s;user-select:none}
.owner-card:hover{box-shadow:0 6px 18px rgba(79,70,229,.12);border-color:var(--brand-mid);transform:translateY(-2px)}
.owner-card.active{border-color:var(--brand);box-shadow:0 0 0 3px var(--brand-mid);transform:translateY(-2px)}
.status-bar{position:absolute;top:0;left:0;right:0;height:4px;border-radius:12px 12px 0 0}
.status-bar.overdue{background:linear-gradient(90deg,var(--red),var(--amber))}
.status-bar.ok{background:linear-gradient(90deg,var(--green),#34d399)}
.owner-top{display:flex;align-items:center;justify-content:space-between;margin:6px 0 10px}
.owner-avatar{width:32px;height:32px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:800;color:#fff;flex-shrink:0}
.owner-name-wrap{flex:1;padding:0 9px}
.owner-name{font-size:13px;font-weight:700;line-height:1.2}
.owner-stats{display:flex;gap:8px}
.stat-box{flex:1;text-align:center;background:var(--bg);border-radius:8px;padding:7px 4px}
.stat-box .num{font-size:20px;font-weight:900;line-height:1}
.stat-box .lbl{font-size:9px;color:var(--gray);text-transform:uppercase;letter-spacing:.05em;margin-top:2px}
.num-brand{color:var(--brand)}.num-red{color:var(--red)}.num-green{color:var(--green)}
.risk-pill{margin-top:8px;font-size:10px;font-weight:600;color:var(--red);background:var(--red-light);border:1px solid var(--red-mid);border-radius:6px;padding:3px 8px;display:flex;align-items:center;gap:4px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.detail-panel{display:none;background:var(--surface);border:2px solid var(--brand);border-radius:14px;padding:20px 22px;margin-bottom:24px;box-shadow:0 4px 24px rgba(79,70,229,.1);animation:slideIn .2s ease}
.detail-panel.visible{display:block}
.detail-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:14px}
.detail-title{display:flex;align-items:center;gap:10px}
.detail-avatar{width:36px;height:36px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:900;color:#fff}
.detail-name{font-size:15px;font-weight:800}
.detail-count{font-size:11.5px;color:var(--gray);margin-top:2px}
.item-list{list-style:none}
.item-list li{display:flex;align-items:flex-start;gap:10px;padding:9px 0;border-bottom:1px solid #f3f4f6;font-size:12.5px;color:var(--text-2);line-height:1.5}
.item-list li:last-child{border-bottom:none}
.item-num{flex-shrink:0;min-width:22px;height:22px;background:var(--brand-light);color:var(--brand);border-radius:6px;font-size:10px;font-weight:800;display:flex;align-items:center;justify-content:center;margin-top:2px}
.item-num.od{background:var(--red-light);color:var(--red)}
.overdue-badge{display:inline-flex;align-items:center;gap:3px;font-size:9.5px;font-weight:700;color:var(--red);background:var(--red-light);border:1px solid var(--red-mid);border-radius:4px;padding:1px 5px;margin-left:6px;vertical-align:middle;flex-shrink:0}
.item-link{color:inherit;text-decoration:none}
.item-link:hover{color:var(--brand);text-decoration:underline;text-decoration-style:dotted;text-underline-offset:2px}
.item-link .ss-icon{font-size:10px;color:var(--brand);opacity:.5;margin-left:4px}
.item-link:hover .ss-icon{opacity:1}
.link-hint{font-size:10px;color:var(--gray);margin-bottom:10px;padding:6px 10px;background:var(--bg);border-radius:6px;border:1px solid var(--border);display:flex;align-items:center;gap:5px}
.gantt-section{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:18px 20px;margin-bottom:20px;box-shadow:0 1px 4px rgba(0,0,0,.04)}
.gantt-section h3{font-size:13px;font-weight:800;margin-bottom:14px;color:var(--text)}
.gantt-header{display:flex;margin-bottom:6px;margin-left:160px}
.gantt-month{flex:1;font-size:9.5px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:var(--gray);text-align:center;border-left:1px dashed var(--border);padding-left:4px}
.gantt-row{display:flex;align-items:center;margin-bottom:7px}
.gantt-label{width:155px;flex-shrink:0;font-size:11px;color:var(--text-2);font-weight:500;padding-right:10px;line-height:1.3}
.gantt-track{flex:1;height:22px;background:var(--gray-light);border-radius:4px;position:relative;overflow:visible}
.gantt-bar{position:absolute;height:100%;border-radius:4px;min-width:6px}
.gantt-bar.completed{background:linear-gradient(90deg,#059669,#34d399)}
.gantt-bar.in-progress{background:linear-gradient(90deg,#d97706,#fbbf24)}
.gantt-bar.not-started{background:linear-gradient(90deg,#6366f1,#818cf8)}
.gantt-bar.blocked{background:linear-gradient(90deg,#dc2626,#f87171)}
.milestone-marker{position:absolute;width:14px;height:14px;background:var(--text);border-radius:3px;transform:rotate(45deg);top:4px;z-index:2}
.gantt-divider{border:none;border-top:2px dashed #e5e7eb;margin:10px 0}
.legend{display:flex;gap:14px;flex-wrap:wrap;margin-bottom:16px}
.leg{display:flex;align-items:center;gap:5px;font-size:10.5px;color:var(--gray)}
.leg-dot{width:10px;height:10px;border-radius:3px;flex-shrink:0}
.ms-table{width:100%;border-collapse:collapse;font-size:12px}
.ms-table th{background:var(--bg);font-size:9.5px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:var(--gray);padding:8px 10px;text-align:left;border-bottom:2px solid var(--border)}
.ms-table td{padding:8px 10px;border-bottom:1px solid #f3f4f6;vertical-align:top;color:var(--text-2)}
.ms-table tr:last-child td{border-bottom:none}
.ms-table tr:hover td{background:#fafbff}
.track-badge{display:inline-block;font-size:10px;font-weight:700;padding:2px 7px;border-radius:5px}
.track-conv{background:var(--brand-light);color:var(--brand)}
.track-co1{background:#ecfdf5;color:#065f46}
.track-co2{background:#fff7ed;color:#9a3412}
.footer{font-size:10.5px;color:#d1d5db;text-align:right;margin-top:14px}
.blank-page{display:flex;flex-direction:column;align-items:center;justify-content:center;min-height:340px;text-align:center;gap:14px;color:var(--gray)}
.blank-page-icon{font-size:48px;opacity:.35}
.blank-page-title{font-size:16px;font-weight:700;color:var(--text-2)}
.blank-page-sub{font-size:12.5px;color:var(--gray);max-width:320px;line-height:1.6}
.health-strip{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:12px}
.health-card{flex:1;min-width:120px;background:var(--surface);border:1.5px solid var(--border);border-radius:12px;padding:14px 18px;text-align:center;box-shadow:0 1px 3px rgba(0,0,0,.05);cursor:pointer;transition:box-shadow .18s,border-color .18s,transform .15s}
.health-card:hover{box-shadow:0 6px 18px rgba(79,70,229,.1);border-color:var(--brand-mid);transform:translateY(-2px)}
.health-card.active{border-color:var(--brand);box-shadow:0 0 0 3px var(--brand-mid);transform:none}
.health-card .h-val{font-size:32px;font-weight:900;line-height:1}
.health-card .h-lbl{font-size:9.5px;color:var(--gray);text-transform:uppercase;letter-spacing:.07em;margin-top:4px}
.health-card .h-sub{font-size:10px;color:var(--gray);margin-top:3px}
.health-detail{display:none;background:var(--surface);border:2px solid var(--brand);border-radius:14px;padding:18px 20px;margin-bottom:20px;box-shadow:0 4px 20px rgba(79,70,229,.1);animation:slideIn .2s ease}
.health-detail.visible{display:block}
.health-item-row{display:flex;align-items:flex-start;gap:10px;padding:8px 0;border-bottom:1px solid #f3f4f6;font-size:12px}
.health-item-row:last-child{border-bottom:none}
.health-item-track{font-size:10px;font-weight:700;padding:2px 7px;border-radius:5px;flex-shrink:0;margin-top:1px}
.ht-conv{background:var(--brand-light);color:var(--brand)}
.ht-co{background:#ecfdf5;color:#065f46}
.ht-raid{background:#fff7ed;color:#9a3412}
.health-item-ms{font-size:10px;color:var(--gray);flex-shrink:0;min-width:90px}
.health-item-task{flex:1;color:var(--text-2);font-weight:500;line-height:1.4}
.health-item-owner{font-size:11px;color:var(--gray);flex-shrink:0;text-align:right;min-width:80px}
.health-item-date{font-size:11px;font-weight:600;color:var(--text);flex-shrink:0;margin-left:6px}
.ms-group-card{background:var(--surface);border:1.5px solid var(--border);border-radius:14px;padding:14px 16px;margin-bottom:10px;cursor:pointer;transition:box-shadow .18s,border-color .18s}
.ms-group-card:hover{box-shadow:0 4px 14px rgba(79,70,229,.1);border-color:var(--brand-mid)}
.ms-group-card.open{border-color:var(--brand);box-shadow:0 0 0 3px var(--brand-mid)}
.ms-group-header{display:flex;align-items:center;justify-content:space-between}
.ms-group-name{font-size:13.5px;font-weight:800;color:var(--text)}
.ms-group-chips{display:flex;gap:6px;align-items:center;flex-wrap:wrap}
.ms-group-body{display:none;margin-top:12px;border-top:1px solid var(--border);padding-top:12px}
.ms-group-body.open{display:block}
.wbs-task-row{display:flex;align-items:center;gap:10px;padding:7px 0;border-bottom:1px solid #f3f4f6;font-size:12px}
.wbs-task-row:last-child{border-bottom:none}
.wbs-task-name{flex:1;color:var(--text-2);font-weight:500}
.wbs-task-owner{font-size:11px;color:var(--gray);min-width:100px;text-align:right;flex-shrink:0}
.wbs-task-date{font-size:11px;font-weight:600;color:var(--text);margin-left:6px;flex-shrink:0}
.wbs-summary{background:var(--surface);border:1px solid var(--border);border-radius:14px;overflow:hidden;margin-bottom:20px;box-shadow:0 1px 4px rgba(0,0,0,.04)}
</style>
</head>
<body>

<div class="tab-bar">
  <button class="tab active" onclick="showTab('main',this)">&#128202; MS Dashboard</button>
  <button class="tab" onclick="showTab('release',this)">&#128197; Release Calendar</button>
  <button class="tab" onclick="showTab('msconv',this)">&#128260; MS Convergence</button>
  <button class="tab" onclick="showTab('co1',this)">&#128994; Change Order Ph 1</button>
  <button class="tab" onclick="showTab('co2',this)">&#128993; Change Order Ph 2</button>
  <button class="tab" onclick="showTab('to2',this)">&#127919; Targeted Offer 2.0</button>
</div>

<!-- PAGE 1: MS DASHBOARD -->
<div id="page-main" class="page active">
  <div class="page-header">
    <div class="header-left">
      <h1>Marketing Studio &middot; MS Dashboard</h1>
      <div class="sub">Targeted Offer Program &middot; RAID Log &middot; Auto-refreshed from Smartsheet &middot; LAST_SYNCED_PLACEHOLDER</div>
    </div>
    <a class="ext-link" href="SHEET_URL_PLACEHOLDER" target="_blank">&#8599; Open RAID Log</a>
  </div>

  <div class="section-label">&#128200; Overall Program Health &mdash; click a card to see details</div>
  <div class="health-strip" id="healthStrip"></div>
  <div class="health-detail" id="healthDetail">
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:14px">
      <div id="healthDetailTitle" style="font-size:15px;font-weight:800;color:var(--text)"></div>
      <button class="close-btn" onclick="closeHealthDetail()">&#215;</button>
    </div>
    <div id="healthDetailContent"></div>
  </div>

  <div class="section-title" style="margin-bottom:10px">Track Breakdown &mdash; click to drill down</div>
  <div id="projCards" style="margin-bottom:20px"></div>

  <div class="proj-detail" id="projDetail">
    <div class="proj-detail-header">
      <div class="proj-detail-title" id="projDetailTitle"></div>
      <button class="close-btn" onclick="closeProjDetail()">&#215;</button>
    </div>
    <ul class="milestone-list" id="projMilestoneList"></ul>
  </div>

  <hr class="section-divider">

  <div class="section-label">&#128197; Release Calendar</div>
  <div class="rel-tiles">
    <div class="rel-tile">
      <div class="rel-tile-icon">&#128203;</div>
      <div><div class="rel-tile-val" style="color:#4f46e5">3</div><div class="rel-tile-lbl">Active Tracks</div></div>
    </div>
    <div class="rel-tile">
      <div class="rel-tile-icon">&#127919;</div>
      <div><div class="rel-tile-val" style="color:#10b981">May 11</div><div class="rel-tile-lbl">MS Conv Go-Live</div></div>
    </div>
    <div class="rel-tile">
      <div class="rel-tile-icon">&#127937;</div>
      <div><div class="rel-tile-val" style="color:#059669">May 30</div><div class="rel-tile-lbl">CO Ph1 Launch</div></div>
    </div>
    <div class="rel-tile">
      <div class="rel-tile-icon">&#128640;</div>
      <div><div class="rel-tile-val" style="color:#f59e0b">Jun 12</div><div class="rel-tile-lbl">CO Ph2 Launch</div></div>
    </div>
  </div>

  <hr class="section-divider">

  <div class="section-label">&#128203; RAID Log</div>
  <div class="summary-row">
    <div class="summary-card"><div class="val val-brand">TOTAL_ITEMS_PLACEHOLDER</div><div class="lbl">Open Items</div></div>
    <div class="summary-card"><div class="val val-green">NUM_OWNERS_PLACEHOLDER</div><div class="lbl">Owners</div></div>
    <div class="summary-card"><div class="val val-red">TOTAL_OVERDUE_PLACEHOLDER</div><div class="lbl">Overdue</div></div>
  </div>

  <div class="section-title">Owner Breakdown &mdash; click a card to drill down</div>
  <div class="owner-grid" id="ownerGrid"></div>

  <div class="detail-panel" id="detailPanel">
    <div class="detail-header">
      <div class="detail-title">
        <div class="detail-avatar" id="detailAvatar"></div>
        <div><div class="detail-name" id="detailName"></div><div class="detail-count" id="detailCount"></div></div>
      </div>
      <button class="close-btn" onclick="closeDetail()">&#215;</button>
    </div>
    <div class="link-hint">&#128279; Click any item to open the RAID Log in Smartsheet</div>
    <ul class="item-list" id="detailList"></ul>
  </div>

  <div class="footer">Auto-refreshed from Smartsheet &middot; Last sync: LAST_SYNCED_PLACEHOLDER</div>
</div>

<!-- PAGE 2: RELEASE CALENDAR -->
<div id="page-release" class="page">
  <div class="page-header">
    <div class="header-left">
      <h1>Marketing Studio &middot; Release Calendar</h1>
      <div class="sub">MS Convergence + Change Orders Ph1 &amp; Ph2 &middot; Q2 2026</div>
    </div>
    <a class="ext-link" href="BINDER_URL_PLACEHOLDER" target="_blank">&#8599; Program Binder</a>
  </div>

  <div class="metric-strip">
    <div class="metric-strip-item"><div class="metric-strip-val" style="color:#4f46e5">3</div><div class="metric-strip-lbl">Active Tracks</div></div>
    <div class="metric-strip-item"><div class="metric-strip-val" style="color:#374151">20</div><div class="metric-strip-lbl">Milestones</div></div>
    <div class="metric-strip-item"><div class="metric-strip-val" style="color:#10b981">1</div><div class="metric-strip-lbl">Completed</div></div>
    <div class="metric-strip-item"><div class="metric-strip-val" style="color:#f59e0b">4</div><div class="metric-strip-lbl">In Progress</div></div>
    <div class="metric-strip-item"><div class="metric-strip-val" style="color:#6366f1">14</div><div class="metric-strip-lbl">Not Started</div></div>
    <div class="metric-strip-item"><div class="metric-strip-val" style="color:#ef4444">1</div><div class="metric-strip-lbl">Blocked</div></div>
  </div>

  <div class="legend">
    <div class="leg"><div class="leg-dot" style="background:linear-gradient(90deg,#059669,#34d399)"></div>Completed</div>
    <div class="leg"><div class="leg-dot" style="background:linear-gradient(90deg,#d97706,#fbbf24)"></div>In Progress</div>
    <div class="leg"><div class="leg-dot" style="background:linear-gradient(90deg,#6366f1,#818cf8)"></div>Not Started</div>
    <div class="leg"><div class="leg-dot" style="background:linear-gradient(90deg,#dc2626,#f87171)"></div>Blocked</div>
    <div class="leg"><div class="leg-dot" style="background:#111827;transform:rotate(45deg);border-radius:2px"></div>Key Milestone</div>
  </div>
  <div class="gantt-section">
    <h3>Timeline &middot; April &ndash; June 2026</h3>
    <div id="ganttContainer"></div>
  </div>

  <div class="section-title">Key Milestones</div>
  <div style="background:var(--surface);border:1px solid var(--border);border-radius:14px;overflow:hidden;margin-bottom:20px;box-shadow:0 1px 4px rgba(0,0,0,.04)">
    <table class="ms-table">
      <thead><tr><th>Track</th><th>Milestone</th><th>Owner</th><th>Start</th><th>Target</th><th>Status</th></tr></thead>
      <tbody id="milestoneTable"></tbody>
    </table>
  </div>
  <div class="footer">Source: Marketing Studio Program Binder &middot; Last updated LAST_SYNCED_PLACEHOLDER</div>
</div>

<!-- PAGE 3: MS CONVERGENCE -->
<div id="page-msconv" class="page">
  <div class="page-header">
    <div class="header-left">
      <h1>Marketing Studio &middot; MS Convergence</h1>
      <div class="sub">Q2 2026 &middot; Live from Smartsheet &middot; LAST_SYNCED_PLACEHOLDER</div>
    </div>
    <a class="ext-link" href="MS_CONV_SS_URL_PLACEHOLDER" target="_blank">&#8599; Open WBS</a>
  </div>
  <div class="wbs-summary" id="msconvSummary"></div>
  <div id="msconvGroups"></div>
</div>

<!-- PAGE 4: CHANGE ORDER -->
<div id="page-co1" class="page">
  <div class="page-header">
    <div class="header-left">
      <h1>Marketing Studio &middot; Change Order WBS</h1>
      <div class="sub">Ph1 &amp; Ph2 Combined &middot; Q2 2026 &middot; Live from Smartsheet &middot; LAST_SYNCED_PLACEHOLDER</div>
    </div>
    <a class="ext-link" href="CO_WBS_SS_URL_PLACEHOLDER" target="_blank">&#8599; Open WBS</a>
  </div>
  <div class="wbs-summary" id="cowbsSummary"></div>
  <div id="cowbsGroups"></div>
</div>

<!-- PAGE 5: CHANGE ORDER PH 2 -->
<div id="page-co2" class="page">
  <div class="page-header">
    <div class="header-left">
      <h1>Marketing Studio &middot; Change Order Ph 2</h1>
      <div class="sub">Q2 2026</div>
    </div>
  </div>
  <div class="blank-page">
    <div class="blank-page-icon">&#128993;</div>
    <div class="blank-page-title">View Combined WBS</div>
    <div class="blank-page-sub">Change Order Ph1 &amp; Ph2 combined WBS is available on the Change Order tab.</div>
  </div>
</div>

<!-- PAGE 6: TARGETED OFFER 2.0 -->
<div id="page-to2" class="page">
  <div class="page-header">
    <div class="header-left">
      <h1>Marketing Studio &middot; Targeted Offer 2.0</h1>
      <div class="sub">RAID Log &middot; Live from Smartsheet &middot; LAST_SYNCED_PLACEHOLDER</div>
    </div>
    <a class="ext-link" href="SHEET_URL_PLACEHOLDER" target="_blank">&#8599; Open RAID Log</a>
  </div>
  <div class="wbs-summary" id="toSummary"></div>
  <div class="section-title" style="margin-top:16px">Owner Breakdown &mdash; click a card to drill down</div>
  <div class="owner-grid" id="toOwnerGrid"></div>
  <div class="detail-panel" id="toDetailPanel">
    <div class="detail-header">
      <div class="detail-title">
        <div class="detail-avatar" id="toDetailAvatar"></div>
        <div><div class="detail-name" id="toDetailName"></div><div class="detail-count" id="toDetailCount"></div></div>
      </div>
      <button class="close-btn" onclick="closeToDetail()">&#215;</button>
    </div>
    <div class="link-hint">&#128279; Click any item to open the RAID Log in Smartsheet</div>
    <ul class="item-list" id="toDetailList"></ul>
  </div>
  <div class="footer">Auto-refreshed from Smartsheet &middot; Last sync: LAST_SYNCED_PLACEHOLDER</div>
</div>

<script>
const SHEET_URL = 'SHEET_URL_PLACEHOLDER';
const MS_CONV_SS_URL = 'MS_CONV_SS_URL_PLACEHOLDER';
const CO_WBS_SS_URL  = 'CO_WBS_SS_URL_PLACEHOLDER';
const msConvData = MS_CONV_WBS_PLACEHOLDER;
const coWbsData  = CO_WBS_PLACEHOLDER;

function showTab(id, el) {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.getElementById('page-' + id).classList.add('active');
  if (el) el.classList.add('active');
}

/* ── ACTION ITEMS ── */
const PALETTE = ['#4f46e5','#7c3aed','#2563eb','#0891b2','#059669','#d97706','#dc2626','#be185d','#6366f1','#0e7490','#7e22ce','#b45309'];
const owners = OWNERS_JSON_PLACEHOLDER;

function initials(n) { return n.split(' ').slice(0,2).map(w=>w[0]).join('').toUpperCase(); }

let activeOwnerIdx = null;
(function() {
  const g = document.getElementById('ownerGrid');
  owners.forEach((o, i) => {
    const c = document.createElement('div');
    c.className = 'owner-card'; c.id = 'ocard-' + i; c.onclick = () => toggleOwner(i);
    c.innerHTML = '<div class="status-bar ' + (o.overdue>0?'overdue':'ok') + '"></div>'
      + '<div class="owner-top">'
      + '<div class="owner-avatar" style="background:' + PALETTE[i%PALETTE.length] + '">' + initials(o.name) + '</div>'
      + '<div class="owner-name-wrap"><div class="owner-name">' + o.name + '</div></div>'
      + '<span class="chevron" id="ochev-' + i + '">&#9660;</span>'
      + '</div>'
      + '<div class="owner-stats">'
      + '<div class="stat-box"><div class="num num-brand">' + o.total + '</div><div class="lbl">Total</div></div>'
      + '<div class="stat-box"><div class="num ' + (o.overdue>0?'num-red':'num-green') + '">' + o.overdue + '</div><div class="lbl">Overdue</div></div>'
      + '</div>'
      + (o.overdue>0 ? '<div class="risk-pill">&#9888; ' + o.overdueDetail + '</div>' : '');
    g.appendChild(c);
  });
})();

function toggleOwner(idx) {
  const p = document.getElementById('detailPanel');
  if (activeOwnerIdx===idx) { closeDetail(); return; }
  if (activeOwnerIdx!==null) document.getElementById('ocard-'+activeOwnerIdx).classList.remove('active');
  activeOwnerIdx = idx; document.getElementById('ocard-'+idx).classList.add('active');
  const o = owners[idx];
  const av = document.getElementById('detailAvatar'); av.textContent = initials(o.name); av.style.background = PALETTE[idx%PALETTE.length];
  document.getElementById('detailName').textContent = o.name;
  document.getElementById('detailCount').textContent = o.total + ' action item' + (o.total!==1?'s':'') + (o.overdue>0?' \u00b7 '+o.overdue+' overdue':'');
  const list = document.getElementById('detailList'); list.innerHTML = '';
  o.items.forEach((item, n) => {
    const li = document.createElement('li');
    const badge = item.od ? '<span class="overdue-badge">&#9888; OVERDUE</span>' : '';
    li.innerHTML = '<div class="item-num ' + (item.od?'od':'') + '">' + (n+1) + '</div>'
      + '<span><a class="item-link" href="' + SHEET_URL + '" target="_blank">' + item.t + '<span class="ss-icon">&#8599;</span></a>' + badge + '</span>';
    list.appendChild(li);
  });
  p.classList.add('visible');
  setTimeout(() => p.scrollIntoView({behavior:'smooth', block:'nearest'}), 60);
}
function closeDetail() {
  if (activeOwnerIdx!==null) document.getElementById('ocard-'+activeOwnerIdx).classList.remove('active');
  activeOwnerIdx = null; document.getElementById('detailPanel').classList.remove('visible');
}

/* ── RELEASE CALENDAR ── */
const projects = [
  {
    id:'conv', name:'MS Convergence', dates:'Apr 3 – May 29, 2026',
    color:'#4f46e5',
    counts:{completed:1, inProgress:1, blocked:1, notStarted:5}, total:8,
    milestones:[
      {label:'Code Complete (P0/CM)', owner:'Gaurav Misra', end:'Apr 17', status:'completed'},
      {label:'Code Complete (PPM/IXP)', owner:'Gaurav Misra', end:'Apr 23', status:'in-progress'},
      {label:'XD Figma Design Spec', owner:'Atul Joshi', end:'Apr 23', status:'blocked'},
      {label:'E2E Testing', owner:'Gaurav Misra', end:'Apr 30', status:'not-started'},
      {label:'Production Deployment', owner:'Gaurav Misra', end:'May 8', status:'not-started'},
      {label:'UAT (PM + MOps)', owner:'Bharatkumar', end:'May 8', status:'not-started'},
      {label:'\U0001F3AF Go Live', owner:'Gaurav Misra', end:'May 11', status:'not-started'},
      {label:'Success Metrics', owner:'Bharatkumar', end:'May 29', status:'not-started'},
    ]
  },
  {
    id:'co1', name:'Change Orders Ph1', dates:'Apr 20 – Jun 5, 2026',
    color:'#059669',
    counts:{completed:0, inProgress:2, blocked:0, notStarted:4}, total:6,
    milestones:[
      {label:'Solution Design Sign-off', owner:'Srinivas Kommu', end:'Apr 27', status:'in-progress'},
      {label:'Code Complete', owner:'Kumaran / Abhay / Minal', end:'May 15', status:'in-progress'},
      {label:'Integration + E2E Testing', owner:'Abhay Kumar Singh', end:'May 22', status:'not-started'},
      {label:'UAT', owner:'Bharatkumar', end:'May 27', status:'not-started'},
      {label:'\U0001F3AF Production Launch', owner:'Srinivas Kommu', end:'May 28', status:'not-started'},
      {label:'Post Launch Monitoring', owner:'Srinivas Kommu', end:'Jun 5', status:'not-started'},
    ]
  },
  {
    id:'co2', name:'Change Orders Ph2', dates:'Apr 22 – Jun 19, 2026',
    color:'#d97706',
    counts:{completed:0, inProgress:1, blocked:0, notStarted:5}, total:6,
    milestones:[
      {label:'Solution Design Complete', owner:'NEERAJ GANG', end:'Apr 30', status:'in-progress'},
      {label:'Code Complete', owner:'NEERAJ GANG / Firdosh / Sunil', end:'May 15', status:'not-started'},
      {label:'E2E Testing', owner:'NEERAJ GANG', end:'May 29', status:'not-started'},
      {label:'UAT (PM + MOps)', owner:'Bharatkumar', end:'Jun 5', status:'not-started'},
      {label:'\U0001F3AF Production Launch', owner:'NEERAJ GANG', end:'Jun 12', status:'not-started'},
      {label:'Post Launch + Metrics', owner:'Bharatkumar', end:'Jun 19', status:'not-started'},
    ]
  },
  {
    id:'to2', name:'Targeted Offer 2.0', dates:'Q2 2026',
    color:'#be185d',
    get counts() {
      const open = owners.length;
      const overdue = owners.filter(o => o.overdue > 0).length;
      return {completed: 0, inProgress: open - overdue, blocked: overdue, notStarted: 0};
    },
    get total() { return owners.length; },
    get milestones() {
      return owners.slice(0,8).map(o => ({
        label: o.name + ' — ' + o.total + ' item' + (o.total!==1?'s':''),
        owner: o.name,
        end: o.overdue > 0 ? o.overdueDetail : 'On track',
        status: o.overdue > 0 ? 'blocked' : 'in-progress'
      }));
    }
  }
];

const statusLabels = {completed:'Done','in-progress':'In Progress','not-started':'Not Started',blocked:'Blocked'};
const chipClass = {completed:'chip-done','in-progress':'chip-prog','not-started':'chip-ns',blocked:'chip-block'};

/* ── HEALTH STRIP ── */
// Build health items lookup from all tracks
const healthItems = {completed:[], inProgress:[], notStarted:[], blocked:[]};
[['MS Convergence', msConvData], ['Change Orders', coWbsData]].forEach(function(pair) {
  var track = pair[0], wbs = pair[1];
  wbs.groups.forEach(function(grp) {
    grp.tasks.forEach(function(t) {
      if (healthItems[t.status]) {
        healthItems[t.status].push({track: track, milestone: grp.name, task: t.task, owner: t.owner, date: t.date});
      }
    });
  });
});
owners.forEach(function(o) {
  o.items.forEach(function(item) {
    var key = item.od ? 'blocked' : 'inProgress';
    healthItems[key].push({track: 'Targeted Offer', milestone: 'RAID Log', task: item.t, owner: o.name, date: ''});
  });
});

let activeHealthKey = null;
(function buildHealthStrip() {
  const totals = {completed:0, inProgress:0, notStarted:0, blocked:0};
  [msConvData, coWbsData].forEach(d => {
    Object.keys(totals).forEach(k => { totals[k] += (d.totals[k]||0); });
  });
  const raidOpen = owners.reduce((s,o) => s + o.total, 0);
  const raidOD   = owners.reduce((s,o) => s + o.overdue, 0);
  totals.inProgress += raidOpen - raidOD;
  totals.blocked    += raidOD;

  const cfg = [
    {k:'completed',  label:'Completed',        color:'var(--green)',  sub:'tasks done'},
    {k:'inProgress', label:'In Progress',       color:'var(--amber)',  sub:'actively tracked'},
    {k:'blocked',    label:'Blocked / At Risk', color:'var(--red)',    sub:'need attention'},
    {k:'notStarted', label:'Not Started',       color:'var(--brand)',  sub:'upcoming'},
  ];
  const strip = document.getElementById('healthStrip');
  cfg.forEach(c => {
    const div = document.createElement('div');
    div.className = 'health-card';
    div.dataset.key = c.k;
    div.innerHTML = '<div class="h-val" style="color:'+c.color+'">'+totals[c.k]+'</div>'
      + '<div class="h-lbl">'+c.label+'</div>'
      + '<div class="h-sub">'+c.sub+'</div>';
    div.addEventListener('click', (function(key){ return function(){ showHealthDetail(key); }; })(c.k));
    strip.appendChild(div);
  });
})();

function showHealthDetail(key) {
  const detail  = document.getElementById('healthDetail');
  const title   = document.getElementById('healthDetailTitle');
  const content = document.getElementById('healthDetailContent');
  document.querySelectorAll('.health-card').forEach(c => c.classList.remove('active'));
  const activeCard = document.querySelector('.health-card[data-key="'+key+'"]');
  if (activeCard) activeCard.classList.add('active');
  if (activeHealthKey === key && detail.classList.contains('visible')) { closeHealthDetail(); return; }
  activeHealthKey = key;
  const labels = {completed:'Completed', inProgress:'In Progress', notStarted:'Not Started', blocked:'Blocked / At Risk'};
  const colors = {completed:'var(--green)', inProgress:'var(--amber)', notStarted:'var(--brand)', blocked:'var(--red)'};
  title.innerHTML = '<span style="color:'+colors[key]+'">'+labels[key]+'</span> &mdash; All Items ('+(healthItems[key]||[]).length+')';
  const items = healthItems[key] || [];
  if (items.length === 0) {
    content.innerHTML = '<div style="color:var(--gray);font-size:13px;padding:12px 0">No items in this category.</div>';
  } else {
    const trackCss = {'MS Convergence':'ht-conv','Change Orders':'ht-co','Targeted Offer':'ht-raid'};
    content.innerHTML = items.map(function(item) {
      var tc = trackCss[item.track] || 'ht-raid';
      return '<div class="health-item-row">'
        + '<span class="health-item-track '+tc+'">'+item.track+'</span>'
        + '<span class="health-item-ms">'+item.milestone+'</span>'
        + '<span class="health-item-task">'+item.task+'</span>'
        + '<span class="health-item-owner">'+item.owner+'</span>'
        + (item.date ? '<span class="health-item-date">&#128197; '+item.date+'</span>' : '')
        + '</div>';
    }).join('');
  }
  detail.classList.add('visible');
  setTimeout(function(){ detail.scrollIntoView({behavior:'smooth', block:'nearest'}); }, 60);
}

function closeHealthDetail() {
  document.querySelectorAll('.health-card').forEach(c => c.classList.remove('active'));
  activeHealthKey = null;
  document.getElementById('healthDetail').classList.remove('visible');
}

(function() {
  const g = document.getElementById('projCards');
  projects.forEach((proj, i) => {
    const card = document.createElement('div');
    card.className = 'proj-card'; card.id = 'pcard-'+i; card.onclick = () => toggleProj(i);
    const t = proj.counts; const tot = proj.total;
    const pDone = (t.completed/tot*100).toFixed(1);
    const pProg = (t.inProgress/tot*100).toFixed(1);
    const pBlk  = (t.blocked/tot*100).toFixed(1);
    const pNS   = (t.notStarted/tot*100).toFixed(1);
    const progBar = '<div class="proj-progress-bar">'
      + (t.completed>0  ? '<div class="pb-seg" style="width:'+pDone+'%;background:#10b981"></div>' : '')
      + (t.inProgress>0 ? '<div class="pb-seg" style="width:'+pProg+'%;background:#f59e0b"></div>' : '')
      + (t.blocked>0    ? '<div class="pb-seg" style="width:'+pBlk+'%;background:#ef4444"></div>'  : '')
      + (t.notStarted>0 ? '<div class="pb-seg" style="width:'+pNS+'%;background:#6366f1"></div>'   : '')
      + '</div>';
    const statRow = [
      t.completed>0  ? '<span class="chip chip-done">&#10003; '+t.completed+' Done</span>'               : '',
      t.inProgress>0 ? '<span class="chip chip-prog">&#8635; '+t.inProgress+' In Progress</span>'        : '',
      t.blocked>0    ? '<span class="chip chip-block">&#215; '+t.blocked+' Blocked</span>'               : '',
      t.notStarted>0 ? '<span class="chip chip-ns">&#9675; '+t.notStarted+' Not Started</span>'          : '',
    ].filter(Boolean).join('');
    card.innerHTML = '<div class="proj-card-top">'
      + '<div><div class="proj-card-name">'+proj.name+'</div><div class="proj-card-dates">'+proj.dates+'</div></div>'
      + '<div class="proj-card-right">'
      + '<div style="text-align:center"><div class="proj-total" style="color:'+proj.color+'">'+proj.total+'</div><div class="proj-total-lbl">Milestones</div></div>'
      + '<span class="chevron" id="pchev-'+i+'">&#9660;</span>'
      + '</div></div>'
      + progBar
      + '<div class="proj-status-row">'+statRow+'</div>';
    g.appendChild(card);
  });
})();

let activeProjIdx = null;
function toggleProj(idx) {
  const panel = document.getElementById('projDetail');
  if (activeProjIdx===idx) { closeProjDetail(); return; }
  if (activeProjIdx!==null) document.getElementById('pcard-'+activeProjIdx).classList.remove('active');
  activeProjIdx = idx; document.getElementById('pcard-'+idx).classList.add('active');
  const proj = projects[idx];
  document.getElementById('projDetailTitle').textContent = proj.name + ' \u2014 Milestone Detail';
  const list = document.getElementById('projMilestoneList'); list.innerHTML = '';
  proj.milestones.forEach((ms, n) => {
    const li = document.createElement('li');
    li.className = 'milestone-li';
    li.innerHTML = '<div class="ms-num">'+(n+1)+'</div>'
      + '<div class="ms-label">'+ms.label+'</div>'
      + '<div class="ms-owner">'+ms.owner+'</div>'
      + '<div class="ms-end">&#128197; '+ms.end+'</div>'
      + '<span class="chip '+chipClass[ms.status]+'">'+statusLabels[ms.status]+'</span>';
    list.appendChild(li);
  });
  panel.classList.add('visible');
  setTimeout(() => panel.scrollIntoView({behavior:'smooth', block:'nearest'}), 60);
}
function closeProjDetail() {
  if (activeProjIdx!==null) document.getElementById('pcard-'+activeProjIdx).classList.remove('active');
  activeProjIdx = null; document.getElementById('projDetail').classList.remove('visible');
}

/* ── GANTT ── */
const RS = new Date('2026-04-01'), RD = 90;
function pct(d) { if(!d) return null; const diff=(new Date(d)-RS)/864e5; return Math.max(0,Math.min(100,diff/RD*100)); }
function bw(s,e) { const sp=pct(s),ep=pct(e); return sp==null||ep==null?0:Math.max(.8,ep-sp); }
const ganttData = [
  {label:'MS Conv \u00b7 Code (P0/CM)',track:'conv',status:'completed',start:'2026-04-03',end:'2026-04-17'},
  {label:'MS Conv \u00b7 Code (PPM/IXP)',track:'conv',status:'in-progress',start:'2026-04-03',end:'2026-04-23'},
  {label:'MS Conv \u00b7 XD Figma',track:'conv',status:'blocked',start:'2026-04-10',end:'2026-04-23'},
  {label:'MS Conv \u00b7 E2E Testing',track:'conv',status:'not-started',start:'2026-04-17',end:'2026-04-30'},
  {label:'MS Conv \u00b7 Prod Deploy',track:'conv',status:'not-started',start:'2026-05-01',end:'2026-05-08'},
  {label:'MS Conv \u00b7 UAT',track:'conv',status:'not-started',start:'2026-05-04',end:'2026-05-08'},
  {label:'MS Conv \u00b7 Go Live \U0001F3AF',track:'conv',status:'not-started',start:'2026-05-08',end:'2026-05-11',key:true},
  {label:'MS Conv \u00b7 Post Launch',track:'conv',status:'not-started',start:'2026-05-11',end:'2026-05-29'},
  {label:'CO Ph1 \u00b7 Solution Design',track:'co1',status:'in-progress',start:'2026-04-20',end:'2026-04-27'},
  {label:'CO Ph1 \u00b7 Code Complete',track:'co1',status:'in-progress',start:'2026-04-22',end:'2026-05-15'},
  {label:'CO Ph1 \u00b7 Integration Test',track:'co1',status:'not-started',start:'2026-05-13',end:'2026-05-15'},
  {label:'CO Ph1 \u00b7 E2E Testing',track:'co1',status:'not-started',start:'2026-05-18',end:'2026-05-22'},
  {label:'CO Ph1 \u00b7 UAT',track:'co1',status:'not-started',start:'2026-05-25',end:'2026-05-27'},
  {label:'CO Ph1 \u00b7 Launch \U0001F3AF',track:'co1',status:'not-started',start:'2026-05-28',end:'2026-05-29',key:true},
  {label:'CO Ph1 \u00b7 Post Launch',track:'co1',status:'not-started',start:'2026-05-29',end:'2026-06-05'},
  {label:'CO Ph2 \u00b7 Solution Design',track:'co2',status:'in-progress',start:'2026-04-22',end:'2026-04-30'},
  {label:'CO Ph2 \u00b7 Code Complete',track:'co2',status:'not-started',start:'2026-05-05',end:'2026-05-15'},
  {label:'CO Ph2 \u00b7 E2E Testing',track:'co2',status:'not-started',start:'2026-05-18',end:'2026-05-29'},
  {label:'CO Ph2 \u00b7 UAT',track:'co2',status:'not-started',start:'2026-05-25',end:'2026-06-05'},
  {label:'CO Ph2 \u00b7 Launch \U0001F3AF',track:'co2',status:'not-started',start:'2026-06-12',end:'2026-06-13',key:true},
  {label:'CO Ph2 \u00b7 Post Launch',track:'co2',status:'not-started',start:'2026-06-15',end:'2026-06-19'},
];
(function() {
  const cont = document.getElementById('ganttContainer');
  const hdr = document.createElement('div'); hdr.className = 'gantt-header';
  ['April','May','June'].forEach(m => { const d=document.createElement('div'); d.className='gantt-month'; d.textContent=m; hdr.appendChild(d); });
  cont.appendChild(hdr);
  const tp = pct(new Date().toISOString().slice(0,10));
  let lastTrack = null;
  ganttData.forEach(row => {
    if (row.track!==lastTrack && lastTrack!==null) { const hr=document.createElement('hr'); hr.className='gantt-divider'; cont.appendChild(hr); }
    lastTrack = row.track;
    const rd = document.createElement('div'); rd.className = 'gantt-row';
    const lbl = document.createElement('div'); lbl.className='gantt-label'; lbl.textContent=row.label;
    const trk = document.createElement('div'); trk.className='gantt-track';
    const tl = document.createElement('div');
    tl.style.cssText = 'position:absolute;left:'+tp+'%;top:-3px;bottom:-3px;width:2px;background:rgba(239,68,68,.45);z-index:1;';
    trk.appendChild(tl);
    const bar = document.createElement('div'); bar.className='gantt-bar '+row.status;
    bar.style.left = pct(row.start)+'%'; bar.style.width = bw(row.start,row.end)+'%';
    trk.appendChild(bar);
    if (row.key) { const m=document.createElement('div'); m.className='milestone-marker'; m.style.left=(pct(row.end)-.8)+'%'; trk.appendChild(m); }
    rd.appendChild(lbl); rd.appendChild(trk); cont.appendChild(rd);
  });
})();

const allMilestones = [
  {track:'MS Convergence',label:'Code Complete (P0/CM)',owner:'Gaurav Misra',start:'Apr 3',end:'Apr 17',status:'completed'},
  {track:'MS Convergence',label:'Code Complete (PPM/IXP)',owner:'Gaurav Misra',start:'Apr 3',end:'Apr 23',status:'in-progress'},
  {track:'MS Convergence',label:'XD Figma Design Spec',owner:'Atul Joshi',start:'\u2014',end:'Apr 23',status:'blocked'},
  {track:'MS Convergence',label:'E2E Testing',owner:'Gaurav Misra',start:'Apr 17',end:'Apr 30',status:'not-started'},
  {track:'MS Convergence',label:'Production Deployment',owner:'Gaurav Misra',start:'May 1',end:'May 8',status:'not-started'},
  {track:'MS Convergence',label:'UAT (PM + MOps)',owner:'Bharatkumar',start:'May 4',end:'May 8',status:'not-started'},
  {track:'MS Convergence',label:'\U0001F3AF Go Live',owner:'Gaurav Misra',start:'May 8',end:'May 11',status:'not-started'},
  {track:'MS Convergence',label:'Success Metrics',owner:'Bharatkumar',start:'May 11',end:'May 29',status:'not-started'},
  {track:'Change Orders Ph1',label:'Solution Design Sign-off',owner:'Srinivas Kommu',start:'Apr 20',end:'Apr 27',status:'in-progress'},
  {track:'Change Orders Ph1',label:'Code Complete',owner:'Kumaran / Abhay / Minal',start:'Apr 22',end:'May 15',status:'in-progress'},
  {track:'Change Orders Ph1',label:'Integration + E2E Testing',owner:'Abhay Kumar Singh',start:'May 13',end:'May 22',status:'not-started'},
  {track:'Change Orders Ph1',label:'UAT',owner:'Bharatkumar',start:'May 20',end:'May 27',status:'not-started'},
  {track:'Change Orders Ph1',label:'\U0001F3AF Production Launch',owner:'Srinivas Kommu',start:'May 28',end:'May 28',status:'not-started'},
  {track:'Change Orders Ph1',label:'Post Launch Monitoring',owner:'Srinivas Kommu',start:'May 29',end:'Jun 5',status:'not-started'},
  {track:'Change Orders Ph2',label:'Solution Design Complete',owner:'NEERAJ GANG',start:'Apr 22',end:'Apr 30',status:'in-progress'},
  {track:'Change Orders Ph2',label:'Code Complete',owner:'NEERAJ GANG / Firdosh / Sunil',start:'May 5',end:'May 15',status:'not-started'},
  {track:'Change Orders Ph2',label:'E2E Testing',owner:'NEERAJ GANG',start:'May 18',end:'May 29',status:'not-started'},
  {track:'Change Orders Ph2',label:'UAT (PM + MOps)',owner:'Bharatkumar',start:'May 25',end:'Jun 5',status:'not-started'},
  {track:'Change Orders Ph2',label:'\U0001F3AF Production Launch',owner:'NEERAJ GANG',start:'Jun 12',end:'Jun 12',status:'not-started'},
  {track:'Change Orders Ph2',label:'Post Launch + Metrics',owner:'Bharatkumar',start:'Jun 15',end:'Jun 19',status:'not-started'},
];
const statusMap = {completed:'<span class="chip chip-done">\u2713 Done</span>','in-progress':'<span class="chip chip-prog">\u21bb In Progress</span>','not-started':'<span class="chip chip-ns">\u25cb Not Started</span>',blocked:'<span class="chip chip-block">\u2715 Blocked</span>'};
const trackMap = {'MS Convergence':'<span class="track-badge track-conv">MS Conv</span>','Change Orders Ph1':'<span class="track-badge track-co1">CO Ph1</span>','Change Orders Ph2':'<span class="track-badge track-co2">CO Ph2</span>'};
const tbody = document.getElementById('milestoneTable');
allMilestones.forEach(m => {
  const tr = document.createElement('tr');
  tr.innerHTML = '<td>'+(trackMap[m.track]||m.track)+'</td><td style="font-weight:500">'+m.label+'</td><td style="color:#6b7280;font-size:11px">'+m.owner+'</td><td style="font-size:11px">'+m.start+'</td><td style="font-weight:600">'+m.end+'</td><td>'+(statusMap[m.status]||m.status)+'</td>';
  tbody.appendChild(tr);
});

/* ── WBS PAGE RENDERER ── */
const STATUS_COLORS_WBS = {completed:'#10b981',inProgress:'#f59e0b',notStarted:'#6366f1',blocked:'#ef4444'};
const STATUS_LABELS_WBS = {completed:'Completed',inProgress:'In Progress',notStarted:'Not Started',blocked:'Blocked'};
const STATUS_CHIP_WBS   = {completed:'chip-done',inProgress:'chip-prog',notStarted:'chip-ns',blocked:'chip-block'};

function renderWbsSummary(data, containerId) {
  const t = data.totals;
  const total = Object.values(t).reduce((a,b)=>a+b,0);
  const el = document.getElementById(containerId);
  if (!el) return;
  el.innerHTML = '<div class="metric-strip">'
    + ['completed','inProgress','notStarted','blocked'].map(k =>
        '<div class="metric-strip-item">'
        +'<div class="metric-strip-val" style="color:'+STATUS_COLORS_WBS[k]+'">'+t[k]+'</div>'
        +'<div class="metric-strip-lbl">'+STATUS_LABELS_WBS[k]+'</div>'
        +'</div>'
      ).join('')
    + '</div>';
}

function renderWbsGroups(data, containerId) {
  const container = document.getElementById(containerId);
  if (!container) return;
  container.innerHTML = '';
  data.groups.forEach((group, gi) => {
    const total = group.tasks.length;
    const card = document.createElement('div');
    card.className = 'ms-group-card';
    card.id = 'wbscard-'+containerId+'-'+gi;
    // progress bar segments
    const keys = ['completed','inProgress','blocked','notStarted'];
    const pbSegs = keys.map(k => {
      if (!group.counts[k]) return '';
      const pct = (group.counts[k]/total*100).toFixed(1);
      return '<div class="pb-seg" style="width:'+pct+'%;background:'+STATUS_COLORS_WBS[k]+'"></div>';
    }).join('');
    // chips
    const chips = keys.filter(k => group.counts[k]>0).map(k =>
      '<span class="chip '+STATUS_CHIP_WBS[k]+'">'+group.counts[k]+' '+STATUS_LABELS_WBS[k]+'</span>'
    ).join('');
    // tasks
    const taskRows = group.tasks.map(t =>
      '<div class="wbs-task-row">'
      +'<span class="chip '+STATUS_CHIP_WBS[t.status]+'" style="flex-shrink:0;margin-left:0">'+STATUS_LABELS_WBS[t.status]+'</span>'
      +'<span class="wbs-task-name">'+t.task+'</span>'
      +'<span class="wbs-task-owner">'+t.owner+'</span>'
      +(t.date ? '<span class="wbs-task-date">&#128197; '+t.date+'</span>' : '')
      +'</div>'
    ).join('');
    card.innerHTML =
      '<div class="ms-group-header" id="whdr-'+containerId+'-'+gi+'">'
      +'<div><div class="ms-group-name">'+group.name+'</div>'
      +'<div class="proj-card-dates">'+total+' task'+(total!==1?'s':'')+'</div></div>'
      +'<div class="ms-group-chips">'+chips+'<span class="chevron" id="wchev-'+containerId+'-'+gi+'">&#9660;</span></div>'
      +'</div>'
      +'<div class="proj-progress-bar" style="margin:10px 0 6px">'+pbSegs+'</div>'
      +'<div class="ms-group-body" id="wbsbody-'+containerId+'-'+gi+'">'+taskRows+'</div>';
    card.querySelector('.ms-group-header').addEventListener('click', (function(cid,g){return function(){toggleWbsCard(cid,g);};})(containerId,gi));
    container.appendChild(card);
  });
}

function toggleWbsCard(cid, gi) {
  const body = document.getElementById('wbsbody-'+cid+'-'+gi);
  const chev = document.getElementById('wchev-'+cid+'-'+gi);
  const card = document.getElementById('wbscard-'+cid+'-'+gi);
  const isOpen = body.classList.contains('open');
  body.classList.toggle('open', !isOpen);
  card.classList.toggle('open', !isOpen);
  if (chev) chev.style.transform = isOpen ? '' : 'rotate(180deg)';
}

/* ── TARGETED OFFER tab owner grid (duplicate of main, different IDs) ── */
let activeToOwnerIdx = null;
(function() {
  const toSummary = document.getElementById('toSummary');
  if (toSummary) {
    const total = owners.reduce((s,o)=>s+o.total,0);
    const overdue = owners.reduce((s,o)=>s+o.overdue,0);
    const numOwners = owners.length;
    toSummary.innerHTML = '<div class="metric-strip">'
      +'<div class="metric-strip-item"><div class="metric-strip-val" style="color:var(--brand)">'+total+'</div><div class="metric-strip-lbl">Open Items</div></div>'
      +'<div class="metric-strip-item"><div class="metric-strip-val" style="color:var(--green)">'+numOwners+'</div><div class="metric-strip-lbl">Owners</div></div>'
      +'<div class="metric-strip-item"><div class="metric-strip-val" style="color:var(--red)">'+overdue+'</div><div class="metric-strip-lbl">Overdue</div></div>'
      +'</div>';
  }
  const g = document.getElementById('toOwnerGrid');
  if (!g) return;
  owners.forEach((o, i) => {
    const c = document.createElement('div');
    c.className = 'owner-card'; c.id = 'tocard-' + i; c.onclick = () => toggleToOwner(i);
    c.innerHTML = '<div class="status-bar ' + (o.overdue>0?'overdue':'ok') + '"></div>'
      + '<div class="owner-top">'
      + '<div class="owner-avatar" style="background:' + PALETTE[i%PALETTE.length] + '">' + initials(o.name) + '</div>'
      + '<div class="owner-name-wrap"><div class="owner-name">' + o.name + '</div></div>'
      + '<span class="chevron" id="tochev-' + i + '">&#9660;</span>'
      + '</div>'
      + '<div class="owner-stats">'
      + '<div class="stat-box"><div class="num num-brand">' + o.total + '</div><div class="lbl">Total</div></div>'
      + '<div class="stat-box"><div class="num ' + (o.overdue>0?'num-red':'num-green') + '">' + o.overdue + '</div><div class="lbl">Overdue</div></div>'
      + '</div>'
      + (o.overdue>0 ? '<div class="risk-pill">&#9888; ' + o.overdueDetail + '</div>' : '');
    g.appendChild(c);
  });
})();

function toggleToOwner(idx) {
  const p = document.getElementById('toDetailPanel');
  if (activeToOwnerIdx===idx) { closeToDetail(); return; }
  if (activeToOwnerIdx!==null) document.getElementById('tocard-'+activeToOwnerIdx).classList.remove('active');
  activeToOwnerIdx = idx; document.getElementById('tocard-'+idx).classList.add('active');
  const o = owners[idx];
  const av = document.getElementById('toDetailAvatar'); av.textContent = initials(o.name); av.style.background = PALETTE[idx%PALETTE.length];
  document.getElementById('toDetailName').textContent = o.name;
  document.getElementById('toDetailCount').textContent = o.total + ' action item' + (o.total!==1?'s':'') + (o.overdue>0?' · '+o.overdue+' overdue':'');
  const list = document.getElementById('toDetailList'); list.innerHTML = '';
  o.items.forEach((item, n) => {
    const li = document.createElement('li');
    const badge = item.od ? '<span class="overdue-badge">&#9888; OVERDUE</span>' : '';
    li.innerHTML = '<div class="item-num ' + (item.od?'od':'') + '">' + (n+1) + '</div>'
      + '<span><a class="item-link" href="' + SHEET_URL + '" target="_blank">' + item.t + '<span class="ss-icon">&#8599;</span></a>' + badge + '</span>';
    list.appendChild(li);
  });
  p.classList.add('visible');
  setTimeout(() => p.scrollIntoView({behavior:'smooth', block:'nearest'}), 60);
}
function closeToDetail() {
  if (activeToOwnerIdx!==null) document.getElementById('tocard-'+activeToOwnerIdx).classList.remove('active');
  activeToOwnerIdx = null; document.getElementById('toDetailPanel').classList.remove('visible');
}

/* ── INIT WBS PAGES ── */
renderWbsSummary(msConvData, 'msconvSummary');
renderWbsGroups(msConvData, 'msconvGroups');
renderWbsSummary(coWbsData, 'cowbsSummary');
renderWbsGroups(coWbsData, 'cowbsGroups');
</script>
</body>
</html>
"""

# ─── HTML generation ──────────────────────────────────────────────────────────
def generate_html(owners, ms_conv_wbs, co_wbs):
    total_items   = sum(o["total"]   for o in owners)
    total_overdue = sum(o["overdue"] for o in owners)
    num_owners    = len(owners)
    owners_json       = json.dumps(owners,       ensure_ascii=True)
    ms_conv_wbs_json  = json.dumps(ms_conv_wbs,  ensure_ascii=True)
    co_wbs_json       = json.dumps(co_wbs,       ensure_ascii=True)

    html = HTML_TEMPLATE
    html = html.replace("OWNERS_JSON_PLACEHOLDER",  owners_json)
    html = html.replace("MS_CONV_WBS_PLACEHOLDER",  ms_conv_wbs_json)
    html = html.replace("CO_WBS_PLACEHOLDER",       co_wbs_json)
    html = html.replace("LAST_SYNCED_PLACEHOLDER",  LAST_SYNCED)
    html = html.replace("TOTAL_ITEMS_PLACEHOLDER",  str(total_items))
    html = html.replace("NUM_OWNERS_PLACEHOLDER",   str(num_owners))
    html = html.replace("TOTAL_OVERDUE_PLACEHOLDER",str(total_overdue))
    html = html.replace("SHEET_URL_PLACEHOLDER",    SHEET_URL)
    html = html.replace("BINDER_URL_PLACEHOLDER",   BINDER_URL)
    html = html.replace("MS_CONV_SS_URL_PLACEHOLDER", MS_CONV_SS_URL)
    html = html.replace("CO_WBS_SS_URL_PLACEHOLDER",  CO_WBS_SS_URL)
    return html

# ─── Main ─────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("Fetching RAID log from Smartsheet...")
    sheet = fetch_sheet()
    print(f"  Sheet: {sheet.get('name')}")
    print(f"  Total rows: {len(sheet.get('rows', []))}")

    owners = parse_owners(sheet)
    print(f"  {len(owners)} owners, {sum(o['total'] for o in owners)} open items.")

    print("Fetching MS Convergence WBS...")
    ms_conv_sheet = fetch_wbs(MS_CONV_SHEET_ID)
    print(f"  Sheet: {ms_conv_sheet.get('name')}, rows: {len(ms_conv_sheet.get('rows',[]))}")
    ms_conv_wbs = parse_wbs(ms_conv_sheet)
    print(f"  Groups: {len(ms_conv_wbs['groups'])}, totals: {ms_conv_wbs['totals']}")

    print("Fetching Change Order WBS...")
    co_wbs_sheet = fetch_wbs(CO_WBS_SHEET_ID)
    print(f"  Sheet: {co_wbs_sheet.get('name')}, rows: {len(co_wbs_sheet.get('rows',[]))}")
    co_wbs = parse_wbs(co_wbs_sheet)
    print(f"  Groups: {len(co_wbs['groups'])}, totals: {co_wbs['totals']}")

    html = generate_html(owners, ms_conv_wbs, co_wbs)
    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html)
    print("Done. Written -> index.html")
