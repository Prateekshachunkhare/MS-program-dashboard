#!/usr/bin/env python3
"""
generate.py
Fetches open Action Items from the Marketing Studio / Targeted Offer RAID Log
(Smartsheet sheet ID 980305451110276) and regenerates index.html for GitHub Pages.

Required env var:  SMARTSHEET_TOKEN  (a Smartsheet personal access token)
"""

import os, json, requests
from datetime import date

# ─── Config ───────────────────────────────────────────────────────────────────
SHEET_ID   = "980305451110276"
TOKEN      = os.environ["SMARTSHEET_TOKEN"]
TODAY      = date.today()
LAST_SYNCED = TODAY.strftime("%B %d, %Y")
SS_BASE    = f"https://app.smartsheet.com/sheets/{SHEET_ID}"
SHEET_URL  = "https://app.smartsheet.com/sheets/3XFh8vH6VwcrWhH2Jw54J44hMX5G7JXfHmGQX8x1?view=grid"
BINDER_URL = "https://docs.google.com/spreadsheets/d/1ThmYPfgTH_zhlmuJr63H47UzonJquO6QWHxkvXrP9tw/edit?gid=565365486#gid=565365486"

# ─── Fetch sheet ──────────────────────────────────────────────────────────────
def fetch_sheet():
    headers = {"Authorization": f"Bearer {TOKEN}", "Accept": "application/json"}
    r = requests.get(
        f"https://api.smartsheet.com/2.0/sheets/{SHEET_ID}",
        headers=headers, timeout=30
    )
    r.raise_for_status()
    return r.json()

# ─── Parse owners ─────────────────────────────────────────────────────────────
def parse_owners(sheet):
    col = {c["title"].strip(): c["id"] for c in sheet.get("columns", [])}
    type_col   = col.get("Type")
    action_col = col.get("Action Item")
    owner_col  = col.get("Owner")
    status_col = col.get("Status")
    due_col    = col.get("Est. Completion")

    owners_dict = {}

    for row in sheet.get("rows", []):
        cells = {cell["columnId"]: cell for cell in row.get("cells", [])}

        # Only Action Item rows
        if type_col:
            type_val = cells.get(type_col, {}).get("displayValue", "")
            if type_val != "Action Item":
                continue

        # Skip completed
        status_val = ""
        if status_col:
            status_val = cells.get(status_col, {}).get("displayValue", "")
        if status_val == "Completed":
            continue

        # Action text
        action_text = ""
        if action_col:
            ac = cells.get(action_col, {})
            action_text = ac.get("displayValue") or ac.get("value") or ""
        if not action_text:
            continue

        # Owner
        owner_name = "Unassigned"
        if owner_col:
            oc = cells.get(owner_col, {})
            owner_name = oc.get("displayValue") or oc.get("value") or "Unassigned"

        # Overdue check
        is_overdue = False
        overdue_label = ""
        if due_col:
            dc = cells.get(due_col, {})
            due_str = dc.get("value") or ""
            if due_str:
                try:
                    due_date = date.fromisoformat(due_str[:10])
                    if due_date < TODAY:
                        is_overdue = True
                        overdue_label = f"Due {due_date.strftime('%b %-d')} passed"
                except Exception:
                    pass

        row_id = str(row["id"])

        if owner_name not in owners_dict:
            owners_dict[owner_name] = {"items": [], "first_overdue": None}

        item = {"t": action_text, "rowId": row_id}
        if is_overdue:
            item["od"] = True
            if owners_dict[owner_name]["first_overdue"] is None:
                owners_dict[owner_name]["first_overdue"] = overdue_label

        owners_dict[owner_name]["items"].append(item)

    # Build sorted list
    result = []
    for name, data in sorted(owners_dict.items(), key=lambda x: len(x[1]["items"]), reverse=True):
        items = data["items"]
        overdue_count = sum(1 for i in items if i.get("od"))
        result.append({
            "name": name,
            "total": len(items),
            "overdue": overdue_count,
            "overdueDetail": data["first_overdue"] or "",
            "items": items
        })
    return result

# ─── HTML generation ──────────────────────────────────────────────────────────
def generate_html(owners):
    total_items   = sum(o["total"]   for o in owners)
    total_overdue = sum(o["overdue"] for o in owners)
    num_owners    = len(owners)
    owners_json   = json.dumps(owners, ensure_ascii=False, indent=2)

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Marketing Studio · Program Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.5.0/dist/chart.umd.js" integrity="sha384-iU8HYtnGQ8Cy4zl7gbNMOhsDTTKX02BTXptVP/vqAWIaTfM7isw76iyZCsjL2eVi" crossorigin="anonymous"></script>
<style>
  :root {{
    color-scheme: light;
    --brand:#4f46e5;--brand-light:#ede9fe;--brand-mid:#c7d2fe;
    --red:#ef4444;--red-light:#fef2f2;--red-mid:#fecaca;
    --green:#10b981;--green-light:#ecfdf5;
    --amber:#f59e0b;--amber-light:#fffbeb;
    --blue:#3b82f6;--blue-light:#eff6ff;
    --gray:#6b7280;--border:#e5e7eb;
    --bg:#ffffff;--surface:#f8f9ff;--text:#111827;
  }}
  *,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
  body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;background:var(--bg);color:var(--text);padding:24px 20px 48px}}
  .tab-bar{{display:flex;gap:4px;margin-bottom:24px;border-bottom:2px solid var(--border)}}
  .tab{{padding:9px 18px;font-size:13px;font-weight:600;cursor:pointer;border-radius:8px 8px 0 0;color:var(--gray);background:none;border:none;border-bottom:3px solid transparent;margin-bottom:-2px;transition:color .15s,border-color .15s}}
  .tab:hover{{color:var(--brand)}}
  .tab.active{{color:var(--brand);border-bottom-color:var(--brand);background:var(--brand-light)}}
  .page{{display:none}}.page.active{{display:block}}
  .page-header{{display:flex;align-items:flex-start;justify-content:space-between;flex-wrap:wrap;gap:12px;margin-bottom:24px}}
  .header-left h1{{font-size:20px;font-weight:800;letter-spacing:-.3px}}
  .header-left .sub{{font-size:12px;color:var(--gray);margin-top:4px}}
  .ext-link{{display:inline-flex;align-items:center;gap:6px;font-size:12px;font-weight:600;color:var(--brand);text-decoration:none;border:1.5px solid var(--brand-mid);border-radius:8px;padding:7px 14px;background:var(--brand-light);white-space:nowrap}}
  .ext-link:hover{{background:#ddd6fe}}
  .summary-row{{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:26px}}
  .summary-card{{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:16px 20px;text-align:center;box-shadow:0 1px 4px rgba(0,0,0,.04)}}
  .summary-card .val{{font-size:34px;font-weight:900;line-height:1}}
  .summary-card .lbl{{font-size:10px;color:var(--gray);text-transform:uppercase;letter-spacing:.07em;margin-top:5px}}
  .val-brand{{color:var(--brand)}}.val-green{{color:var(--green)}}.val-red{{color:var(--red)}}
  .charts-row{{display:grid;grid-template-columns:1.5fr 1fr;gap:14px;margin-bottom:26px}}
  .chart-card{{background:#fff;border:1px solid var(--border);border-radius:14px;padding:18px 20px;box-shadow:0 1px 4px rgba(0,0,0,.04)}}
  .chart-card h3{{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.07em;color:var(--gray);margin-bottom:14px}}
  .chart-wrap canvas{{max-height:210px}}
  .section-title{{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:var(--gray);margin-bottom:12px}}
  .owner-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(228px,1fr));gap:12px;margin-bottom:26px}}
  .owner-card{{background:#fff;border:1.5px solid var(--border);border-radius:12px;padding:14px 16px;cursor:pointer;position:relative;transition:box-shadow .18s,border-color .18s,transform .18s;user-select:none}}
  .owner-card:hover{{box-shadow:0 6px 18px rgba(79,70,229,.12);border-color:var(--brand-mid);transform:translateY(-2px)}}
  .owner-card.active{{border-color:var(--brand);box-shadow:0 0 0 3px var(--brand-mid);transform:translateY(-2px)}}
  .status-bar{{position:absolute;top:0;left:0;right:0;height:4px;border-radius:12px 12px 0 0}}
  .status-bar.overdue{{background:linear-gradient(90deg,var(--red),var(--amber))}}
  .status-bar.ok{{background:linear-gradient(90deg,var(--green),#34d399)}}
  .owner-top{{display:flex;align-items:center;justify-content:space-between;margin:6px 0 10px}}
  .owner-avatar{{width:32px;height:32px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:800;color:#fff;flex-shrink:0}}
  .owner-name-wrap{{flex:1;padding:0 9px}}
  .owner-name{{font-size:13px;font-weight:700;line-height:1.2}}
  .chevron{{font-size:10px;color:var(--gray);transition:transform .2s}}
  .owner-card.active .chevron{{transform:rotate(180deg);color:var(--brand)}}
  .owner-stats{{display:flex;gap:8px}}
  .stat-box{{flex:1;text-align:center;background:var(--surface);border-radius:8px;padding:7px 4px}}
  .stat-box .num{{font-size:20px;font-weight:900;line-height:1}}
  .stat-box .lbl{{font-size:9px;color:var(--gray);text-transform:uppercase;letter-spacing:.05em;margin-top:2px}}
  .num-brand{{color:var(--brand)}}.num-red{{color:var(--red)}}.num-green{{color:var(--green)}}
  .risk-pill{{margin-top:9px;font-size:10px;font-weight:600;color:var(--red);background:var(--red-light);border:1px solid var(--red-mid);border-radius:6px;padding:3px 8px;display:flex;align-items:center;gap:4px;width:100%;overflow:hidden;white-space:nowrap;text-overflow:ellipsis}}
  .detail-panel{{display:none;background:#fff;border:2px solid var(--brand);border-radius:14px;padding:20px 22px;margin-bottom:26px;box-shadow:0 4px 24px rgba(79,70,229,.1);animation:slideIn .2s ease}}
  .detail-panel.visible{{display:block}}
  @keyframes slideIn{{from{{opacity:0;transform:translateY(-8px)}}to{{opacity:1;transform:none}}}}
  .detail-header{{display:flex;align-items:center;justify-content:space-between;margin-bottom:16px}}
  .detail-title{{display:flex;align-items:center;gap:10px}}
  .detail-avatar{{width:36px;height:36px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:900;color:#fff}}
  .detail-name{{font-size:16px;font-weight:800}}
  .detail-count{{font-size:12px;color:var(--gray);margin-top:1px}}
  .close-btn{{font-size:20px;cursor:pointer;color:var(--gray);border:none;background:none;padding:3px 7px;border-radius:6px}}
  .close-btn:hover{{background:#f3f4f6}}
  .item-list{{list-style:none}}
  .item-list li{{display:flex;align-items:flex-start;gap:10px;padding:10px 0;border-bottom:1px solid #f3f4f6;font-size:13px;color:#374151;line-height:1.5}}
  .item-list li:last-child{{border-bottom:none}}
  .item-num{{flex-shrink:0;min-width:24px;height:24px;background:var(--brand-light);color:var(--brand);border-radius:6px;font-size:11px;font-weight:800;display:flex;align-items:center;justify-content:center;margin-top:1px}}
  .item-num.od{{background:var(--red-light);color:var(--red)}}
  .overdue-badge{{display:inline-flex;align-items:center;gap:3px;font-size:10px;font-weight:700;color:var(--red);background:var(--red-light);border:1px solid var(--red-mid);border-radius:4px;padding:1px 6px;margin-left:7px;vertical-align:middle;flex-shrink:0}}
  .item-link{{color:inherit;text-decoration:none}}
  .item-link:hover{{color:var(--brand);text-decoration:underline;text-decoration-style:dotted;text-underline-offset:2px}}
  .item-link .ss-icon{{font-size:10px;color:var(--brand);opacity:.6;margin-left:5px;vertical-align:middle;transition:opacity .15s}}
  .item-link:hover .ss-icon{{opacity:1}}
  .link-hint{{font-size:10px;color:var(--gray);margin-bottom:10px;padding:6px 10px;background:var(--surface);border-radius:6px;border:1px solid var(--border);display:flex;align-items:center;gap:5px}}
  .footer{{font-size:11px;color:#d1d5db;text-align:right;margin-top:16px}}
  /* Release calendar */
  .rel-summary{{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:24px}}
  .rel-card{{border-radius:14px;border:1px solid var(--border);padding:14px 18px;text-align:center;background:var(--surface)}}
  .rel-card .big{{font-size:28px;font-weight:900;line-height:1}}
  .rel-card .sm{{font-size:10px;color:var(--gray);text-transform:uppercase;letter-spacing:.06em;margin-top:4px}}
  .legend{{display:flex;gap:16px;flex-wrap:wrap;margin-bottom:18px}}
  .leg{{display:flex;align-items:center;gap:6px;font-size:11px;color:var(--gray)}}
  .leg-dot{{width:12px;height:12px;border-radius:3px;flex-shrink:0}}
  .gantt-section{{background:#fff;border:1px solid var(--border);border-radius:14px;padding:18px 20px;margin-bottom:20px;box-shadow:0 1px 4px rgba(0,0,0,.04)}}
  .gantt-section h3{{font-size:13px;font-weight:800;margin-bottom:14px}}
  .gantt-header{{display:flex;margin-bottom:6px;margin-left:160px}}
  .gantt-month{{flex:1;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:var(--gray);text-align:center;border-left:1px dashed var(--border);padding-left:4px}}
  .gantt-row{{display:flex;align-items:center;margin-bottom:7px}}
  .gantt-label{{width:155px;flex-shrink:0;font-size:11px;color:#374151;font-weight:500;padding-right:10px;line-height:1.3}}
  .gantt-track{{flex:1;height:22px;background:#f3f4f6;border-radius:4px;position:relative;overflow:visible}}
  .gantt-bar{{position:absolute;height:100%;border-radius:4px;display:flex;align-items:center;padding:0 6px;font-size:9px;font-weight:700;color:#fff;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;min-width:6px}}
  .gantt-bar.completed{{background:linear-gradient(90deg,#059669,#34d399)}}
  .gantt-bar.in-progress{{background:linear-gradient(90deg,#d97706,#fbbf24)}}
  .gantt-bar.not-started{{background:linear-gradient(90deg,#6366f1,#818cf8)}}
  .gantt-bar.blocked{{background:linear-gradient(90deg,#dc2626,#f87171)}}
  .milestone-marker{{position:absolute;width:14px;height:14px;background:var(--text);border-radius:3px;transform:rotate(45deg);top:4px;z-index:2}}
  .gantt-divider{{border:none;border-top:2px dashed #e5e7eb;margin:12px 0}}
  .chip{{display:inline-block;font-size:10px;font-weight:700;padding:2px 8px;border-radius:20px}}
  .chip-done{{background:var(--green-light);color:#065f46}}
  .chip-prog{{background:var(--amber-light);color:#92400e}}
  .chip-none{{background:var(--blue-light);color:#1e40af}}
  .chip-block{{background:var(--red-light);color:#991b1b}}
  .ms-table{{width:100%;border-collapse:collapse;font-size:12px;margin-top:4px}}
  .ms-table th{{background:var(--surface);font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:var(--gray);padding:8px 10px;text-align:left;border-bottom:2px solid var(--border)}}
  .ms-table td{{padding:8px 10px;border-bottom:1px solid #f3f4f6;vertical-align:top;color:#374151}}
  .ms-table tr:last-child td{{border-bottom:none}}
  .ms-table tr:hover td{{background:#fafbff}}
  .track-badge{{display:inline-block;font-size:10px;font-weight:700;padding:2px 7px;border-radius:5px}}
  .track-conv{{background:var(--brand-light);color:var(--brand)}}
  .track-co1{{background:#ecfdf5;color:#065f46}}
  .track-co2{{background:#fff7ed;color:#9a3412}}
</style>
</head>
<body>

<div class="tab-bar">
  <button class="tab active" onclick="showTab('actions')">📋 Action Items</button>
  <button class="tab" onclick="showTab('release')">🗓 Release Calendar</button>
</div>

<!-- PAGE 1: ACTION ITEMS -->
<div id="page-actions" class="page active">
  <div class="page-header">
    <div class="header-left">
      <h1>Targeted Offer Program · RAID Log</h1>
      <div class="sub">Action items by owner · Auto-refreshed from Smartsheet · {LAST_SYNCED}</div>
    </div>
    <a class="ext-link" href="{SHEET_URL}" target="_blank">↗ Open RAID Log</a>
  </div>
  <div class="summary-row">
    <div class="summary-card"><div class="val val-brand">{total_items}</div><div class="lbl">Total Log Items</div></div>
    <div class="summary-card"><div class="val val-green">{num_owners}</div><div class="lbl">Owners Tracked</div></div>
    <div class="summary-card"><div class="val val-red">{total_overdue}</div><div class="lbl">Overdue Items</div></div>
  </div>
  <div class="charts-row">
    <div class="chart-card"><h3>Items by Owner</h3><div class="chart-wrap"><canvas id="donutChart"></canvas></div></div>
    <div class="chart-card"><h3>On-Track vs Overdue</h3><div class="chart-wrap"><canvas id="pieChart"></canvas></div></div>
  </div>
  <div class="section-title">Owner Breakdown — click a card to drill down</div>
  <div class="owner-grid" id="ownerGrid"></div>
  <div class="detail-panel" id="detailPanel">
    <div class="detail-header">
      <div class="detail-title">
        <div class="detail-avatar" id="detailAvatar"></div>
        <div><div class="detail-name" id="detailName"></div><div class="detail-count" id="detailCount"></div></div>
      </div>
      <button class="close-btn" onclick="closeDetail()">×</button>
    </div>
    <div class="link-hint" id="linkHint" style="display:none">🔗 Underlined items open directly in Smartsheet — click to view or update</div>
    <ul class="item-list" id="detailList"></ul>
  </div>
  <div class="footer">Auto-refreshed from Smartsheet · Last sync: {LAST_SYNCED}</div>
</div>

<!-- PAGE 2: RELEASE CALENDAR -->
<div id="page-release" class="page">
  <div class="page-header">
    <div class="header-left">
      <h1>Marketing Studio · Release Calendar</h1>
      <div class="sub">MS Convergence + Change Orders Ph1 & Ph2 · Source: Marketing Studio Program Binder</div>
    </div>
    <a class="ext-link" href="{BINDER_URL}" target="_blank">↗ Open Program Binder</a>
  </div>
  <div class="rel-summary">
    <div class="rel-card"><div class="big" style="color:#4f46e5">3</div><div class="sm">Active Tracks</div></div>
    <div class="rel-card"><div class="big" style="color:#10b981">May 11</div><div class="sm">MS Convergence Go-Live</div></div>
    <div class="rel-card"><div class="big" style="color:#f59e0b">Jun 12</div><div class="sm">Change Orders Ph2 Launch</div></div>
  </div>
  <div class="legend">
    <div class="leg"><div class="leg-dot" style="background:linear-gradient(90deg,#059669,#34d399)"></div>Completed</div>
    <div class="leg"><div class="leg-dot" style="background:linear-gradient(90deg,#d97706,#fbbf24)"></div>In Progress</div>
    <div class="leg"><div class="leg-dot" style="background:linear-gradient(90deg,#6366f1,#818cf8)"></div>Not Started</div>
    <div class="leg"><div class="leg-dot" style="background:linear-gradient(90deg,#dc2626,#f87171)"></div>Blocked</div>
    <div class="leg"><div class="leg-dot" style="background:#111827;transform:rotate(45deg);border-radius:2px"></div>Key Milestone</div>
  </div>
  <div class="gantt-section">
    <h3>Timeline · April – June 2026</h3>
    <div id="ganttContainer"></div>
  </div>
  <div class="section-title">Key Milestones</div>
  <div style="background:#fff;border:1px solid var(--border);border-radius:14px;overflow:hidden;margin-bottom:20px;box-shadow:0 1px 4px rgba(0,0,0,.04)">
    <table class="ms-table">
      <thead><tr><th>Track</th><th>Milestone</th><th>Owner</th><th>Start</th><th>End / Target</th><th>Status</th></tr></thead>
      <tbody id="milestoneTable"></tbody>
    </table>
  </div>
  <div class="footer">Source: Marketing Studio Program Binder · Last updated {LAST_SYNCED}</div>
</div>

<script>
/* ── Tab logic ── */
function showTab(id){{
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
  document.getElementById('page-'+id).classList.add('active');
  event.currentTarget.classList.add('active');
}}

/* ── Data ── */
const SHEET_ID = '{SHEET_ID}';
const SS_BASE  = 'https://app.smartsheet.com/sheets/' + SHEET_ID;
const PALETTE  = ['#4f46e5','#7c3aed','#2563eb','#0891b2','#059669','#d97706','#dc2626','#be185d','#6366f1','#0e7490','#7e22ce','#b45309'];
const owners   = {owners_json};

function initials(n){{ return n.split(' ').slice(0,2).map(w=>w[0]).join('').toUpperCase(); }}
function ssLink(rowId){{ return SS_BASE + '?rowId=' + rowId; }}

/* ── Charts ── */
(function(){{
  const dc = document.getElementById('donutChart').getContext('2d');
  new Chart(dc,{{type:'doughnut',data:{{labels:owners.map(o=>o.name.split(' ')[0]),datasets:[{{data:owners.map(o=>o.total),backgroundColor:PALETTE,borderColor:'#fff',borderWidth:3,hoverOffset:8}}]}},options:{{cutout:'62%',plugins:{{legend:{{position:'right',labels:{{font:{{size:11}},color:'#374151',boxWidth:12,padding:10}}}},tooltip:{{callbacks:{{label:c=>` ${{c.label}}: ${{c.parsed}} items`}}}}}}}}}});
  const pc = document.getElementById('pieChart').getContext('2d');
  const od=owners.reduce((s,o)=>s+o.overdue,0), tot=owners.reduce((s,o)=>s+o.total,0);
  new Chart(pc,{{type:'pie',data:{{labels:['On Track','Overdue'],datasets:[{{data:[tot-od,od],backgroundColor:['#10b981','#ef4444'],borderColor:'#fff',borderWidth:4,hoverOffset:8}}]}},options:{{plugins:{{legend:{{position:'bottom',labels:{{font:{{size:12}},color:'#374151',boxWidth:14,padding:14}}}},tooltip:{{callbacks:{{label:c=>` ${{c.label}}: ${{c.parsed}} items`}}}}}}}}}});
}})();

/* ── Owner cards ── */
let activeIdx=null;
(function(){{
  const g=document.getElementById('ownerGrid');
  owners.forEach((o,i)=>{{
    const c=document.createElement('div');
    c.className='owner-card'; c.id='card-'+i; c.onclick=()=>toggleOwner(i);
    c.innerHTML=`<div class="status-bar ${{o.overdue>0?'overdue':'ok'}}"></div>
      <div class="owner-top">
        <div class="owner-avatar" style="background:${{PALETTE[i%PALETTE.length]}}">${{initials(o.name)}}</div>
        <div class="owner-name-wrap"><div class="owner-name">${{o.name}}</div></div>
        <span class="chevron" id="chev-${{i}}">▼</span>
      </div>
      <div class="owner-stats">
        <div class="stat-box"><div class="num num-brand">${{o.total}}</div><div class="lbl">Total</div></div>
        <div class="stat-box"><div class="num ${{o.overdue>0?'num-red':'num-green'}}">${{o.overdue}}</div><div class="lbl">Overdue</div></div>
      </div>
      ${{o.overdue>0?`<div class="risk-pill">⚠ ${{o.overdueDetail}}</div>`:''}}\`;
    g.appendChild(c);
  }});
}})();

function toggleOwner(idx){{
  const p=document.getElementById('detailPanel');
  if(activeIdx===idx){{closeDetail();return;}}
  if(activeIdx!==null) document.getElementById('card-'+activeIdx).classList.remove('active');
  activeIdx=idx; document.getElementById('card-'+idx).classList.add('active');
  const o=owners[idx];
  const av=document.getElementById('detailAvatar'); av.textContent=initials(o.name); av.style.background=PALETTE[idx%PALETTE.length];
  document.getElementById('detailName').textContent=o.name;
  document.getElementById('detailCount').textContent=`${{o.total}} action item${{o.total!==1?'s':''}}${{o.overdue>0?'  ·  '+o.overdue+' overdue':''}}`;
  const hasLinks=o.items.some(i=>i.rowId);
  document.getElementById('linkHint').style.display=hasLinks?'flex':'none';
  const list=document.getElementById('detailList'); list.innerHTML='';
  o.items.forEach((item,n)=>{{
    const li=document.createElement('li');
    const numDiv=`<div class="item-num ${{item.od?'od':''}}">${{n+1}}</div>`;
    const badge=item.od?'<span class="overdue-badge">⚠ OVERDUE</span>':'';
    const content=item.rowId
      ? `<span><a class="item-link" href="${{ssLink(item.rowId)}}" target="_blank">${{item.t}}<span class="ss-icon">↗</span></a>${{badge}}</span>`
      : `<span>${{item.t}}${{badge}}</span>`;
    li.innerHTML=numDiv+content;
    list.appendChild(li);
  }});
  p.classList.add('visible');
  setTimeout(()=>p.scrollIntoView({{behavior:'smooth',block:'nearest'}}),60);
}}
function closeDetail(){{
  if(activeIdx!==null) document.getElementById('card-'+activeIdx).classList.remove('active');
  activeIdx=null; document.getElementById('detailPanel').classList.remove('visible');
}}

/* ── Gantt ── */
const RANGE_START=new Date('2026-04-01'); const RANGE_DAYS=90;
function pct(d){{if(!d)return null;const diff=(new Date(d)-RANGE_START)/(864e5);return Math.max(0,Math.min(100,diff/RANGE_DAYS*100));}}
function bw(s,e){{const sp=pct(s),ep=pct(e);return sp==null||ep==null?0:Math.max(.8,ep-sp);}}
const ganttData=[
  {{label:'MS Conv · Code (P0/CM)',track:'conv',status:'completed',start:'2026-04-03',end:'2026-04-17'}},
  {{label:'MS Conv · Code (PPM/IXP)',track:'conv',status:'in-progress',start:'2026-04-03',end:'2026-04-23'}},
  {{label:'MS Conv · XD Figma',track:'conv',status:'blocked',start:'2026-04-10',end:'2026-04-23'}},
  {{label:'MS Conv · E2E Testing',track:'conv',status:'not-started',start:'2026-04-17',end:'2026-04-30'}},
  {{label:'MS Conv · Prod Deploy',track:'conv',status:'not-started',start:'2026-05-01',end:'2026-05-08'}},
  {{label:'MS Conv · UAT',track:'conv',status:'not-started',start:'2026-05-04',end:'2026-05-08'}},
  {{label:'MS Conv · Go Live 🎯',track:'conv',status:'not-started',start:'2026-05-08',end:'2026-05-11',key:true}},
  {{label:'MS Conv · Post Launch',track:'conv',status:'not-started',start:'2026-05-11',end:'2026-05-29'}},
  {{label:'CO Ph1 · Solution Design',track:'co1',status:'in-progress',start:'2026-04-20',end:'2026-04-27'}},
  {{label:'CO Ph1 · Code Complete',track:'co1',status:'in-progress',start:'2026-04-22',end:'2026-05-15'}},
  {{label:'CO Ph1 · Integration Test',track:'co1',status:'not-started',start:'2026-05-13',end:'2026-05-15'}},
  {{label:'CO Ph1 · E2E Testing',track:'co1',status:'not-started',start:'2026-05-18',end:'2026-05-22'}},
  {{label:'CO Ph1 · UAT',track:'co1',status:'not-started',start:'2026-05-25',end:'2026-05-27'}},
  {{label:'CO Ph1 · Launch 🎯',track:'co1',status:'not-started',start:'2026-05-28',end:'2026-05-29',key:true}},
  {{label:'CO Ph1 · Post Launch',track:'co1',status:'not-started',start:'2026-05-29',end:'2026-06-05'}},
  {{label:'CO Ph2 · Solution Design',track:'co2',status:'in-progress',start:'2026-04-22',end:'2026-04-30'}},
  {{label:'CO Ph2 · Code Complete',track:'co2',status:'not-started',start:'2026-05-05',end:'2026-05-15'}},
  {{label:'CO Ph2 · E2E Testing',track:'co2',status:'not-started',start:'2026-05-18',end:'2026-05-29'}},
  {{label:'CO Ph2 · UAT',track:'co2',status:'not-started',start:'2026-05-25',end:'2026-06-05'}},
  {{label:'CO Ph2 · Launch 🎯',track:'co2',status:'not-started',start:'2026-06-12',end:'2026-06-13',key:true}},
  {{label:'CO Ph2 · Post Launch',track:'co2',status:'not-started',start:'2026-06-15',end:'2026-06-19'}},
];
(function(){{
  const cont=document.getElementById('ganttContainer');
  const hdr=document.createElement('div'); hdr.className='gantt-header';
  ['April','May','June'].forEach(m=>{{const d=document.createElement('div');d.className='gantt-month';d.textContent=m;hdr.appendChild(d);}});
  cont.appendChild(hdr);
  const todayPct=pct(new Date().toISOString().slice(0,10));
  let lastTrack=null;
  ganttData.forEach(row=>{{
    if(row.track!==lastTrack&&lastTrack!==null){{const hr=document.createElement('hr');hr.className='gantt-divider';cont.appendChild(hr);}}
    lastTrack=row.track;
    const rowDiv=document.createElement('div'); rowDiv.className='gantt-row';
    const lbl=document.createElement('div'); lbl.className='gantt-label'; lbl.textContent=row.label;
    const track=document.createElement('div'); track.className='gantt-track';
    const tline=document.createElement('div');
    tline.style.cssText=`position:absolute;left:${{todayPct}}%;top:-3px;bottom:-3px;width:2px;background:rgba(239,68,68,.5);z-index:1;`;
    track.appendChild(tline);
    const bar=document.createElement('div'); bar.className='gantt-bar '+row.status;
    bar.style.left=pct(row.start)+'%'; bar.style.width=bw(row.start,row.end)+'%';
    track.appendChild(bar);
    if(row.key){{const m=document.createElement('div');m.className='milestone-marker';m.style.left=(pct(row.end)-.8)+'%';track.appendChild(m);}}
    rowDiv.appendChild(lbl); rowDiv.appendChild(track); cont.appendChild(rowDiv);
  }});
}})();

/* ── Milestone table ── */
const milestones=[
  {{track:'MS Convergence',label:'Code Complete (P0/CM)',owner:'Gaurav Misra',start:'Apr 3',end:'Apr 17',status:'completed'}},
  {{track:'MS Convergence',label:'Code Complete (PPM/IXP)',owner:'Gaurav Misra',start:'Apr 3',end:'Apr 23',status:'in-progress'}},
  {{track:'MS Convergence',label:'XD Figma Design Spec',owner:'Atul Joshi',start:'—',end:'Apr 23',status:'blocked'}},
  {{track:'MS Convergence',label:'E2E Testing',owner:'Gaurav Misra',start:'Apr 17',end:'Apr 30',status:'not-started'}},
  {{track:'MS Convergence',label:'Production Deployment',owner:'Gaurav Misra',start:'May 1',end:'May 8',status:'not-started'}},
  {{track:'MS Convergence',label:'UAT (PM + MOps)',owner:'Bharatkumar',start:'May 4',end:'May 8',status:'not-started'}},
  {{track:'MS Convergence',label:'🎯 Go Live',owner:'Gaurav Misra',start:'May 8',end:'May 11',status:'not-started'}},
  {{track:'MS Convergence',label:'Success Metrics',owner:'Bharatkumar',start:'May 11',end:'May 29',status:'not-started'}},
  {{track:'Change Orders Ph1',label:'Solution Design Sign-off',owner:'Srinivas Kommu',start:'Apr 20',end:'Apr 27',status:'in-progress'}},
  {{track:'Change Orders Ph1',label:'Code Complete',owner:'Kumaran / Abhay / Minal',start:'Apr 22',end:'May 15',status:'in-progress'}},
  {{track:'Change Orders Ph1',label:'Integration + E2E Testing',owner:'Abhay Kumar Singh',start:'May 13',end:'May 22',status:'not-started'}},
  {{track:'Change Orders Ph1',label:'UAT',owner:'Bharatkumar',start:'May 20',end:'May 27',status:'not-started'}},
  {{track:'Change Orders Ph1',label:'🎯 Production Launch',owner:'Srinivas Kommu',start:'May 28',end:'May 28',status:'not-started'}},
  {{track:'Change Orders Ph1',label:'Post Launch Monitoring',owner:'Srinivas Kommu',start:'May 29',end:'Jun 5',status:'not-started'}},
  {{track:'Change Orders Ph2',label:'Solution Design Complete',owner:'NEERAJ GANG',start:'Apr 22',end:'Apr 30',status:'in-progress'}},
  {{track:'Change Orders Ph2',label:'Code Complete',owner:'NEERAJ GANG / Firdosh / Sunil',start:'May 5',end:'May 15',status:'not-started'}},
  {{track:'Change Orders Ph2',label:'E2E Testing',owner:'NEERAJ GANG',start:'May 18',end:'May 29',status:'not-started'}},
  {{track:'Change Orders Ph2',label:'UAT (PM + MOps)',owner:'Bharatkumar',start:'May 25',end:'Jun 5',status:'not-started'}},
  {{track:'Change Orders Ph2',label:'🎯 Production Launch',owner:'NEERAJ GANG',start:'Jun 12',end:'Jun 12',status:'not-started'}},
  {{track:'Change Orders Ph2',label:'Post Launch + Metrics',owner:'Bharatkumar',start:'Jun 15',end:'Jun 19',status:'not-started'}},
];
const statusMap={{completed:'<span class="chip chip-done">✓ Done</span>','in-progress':'<span class="chip chip-prog">⟳ In Progress</span>','not-started':'<span class="chip chip-none">○ Not Started</span>',blocked:'<span class="chip chip-block">✕ Blocked</span>'}};
const trackMap={{'MS Convergence':'<span class="track-badge track-conv">MS Conv</span>','Change Orders Ph1':'<span class="track-badge track-co1">CO Ph1</span>','Change Orders Ph2':'<span class="track-badge track-co2">CO Ph2</span>'}};
const tbody=document.getElementById('milestoneTable');
milestones.forEach(m=>{{
  const tr=document.createElement('tr');
  tr.innerHTML=`<td>${{trackMap[m.track]||m.track}}</td><td style="font-weight:500">${{m.label}}</td><td style="color:#6b7280">${{m.owner}}</td><td>${{m.start}}</td><td style="font-weight:600">${{m.end}}</td><td>${{statusMap[m.status]||m.status}}</td>`;
  tbody.appendChild(tr);
}});
</script>
</body>
</html>"""

# ─── Main ─────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("Fetching sheet from Smartsheet...")
    sheet = fetch_sheet()
    print(f"  Sheet: {sheet.get('name')}, rows: {len(sheet.get('rows', []))}")

    owners = parse_owners(sheet)
    print(f"  Found {len(owners)} owners, {sum(o['total'] for o in owners)} open action items")

    html = generate_html(owners)

    out_path = "index.html"
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"  Written → {out_path}")
