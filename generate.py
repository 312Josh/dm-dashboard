#!/usr/bin/env python3
"""
Generate the dm.mellender.io sales dashboard HTML from daily CSV data.

Usage:
    python3 generate.py /path/to/MM.DD.Data/

Reads the 5 CSVs from the data folder and outputs index.html in this script's directory.
"""

import csv
import json
import os
import sys
from datetime import datetime, date
from pathlib import Path


def read_csv(path):
    with open(path, newline="", encoding="utf-8-sig", errors="replace") as f:
        return list(csv.DictReader(f))


def parse_pct(s):
    """Parse '28.4%' -> 28.4, or decimal fraction '0.284' -> 28.4, or return 0.0"""
    if not s:
        return 0.0
    has_pct_sign = "%" in s
    val = float(s.replace("%", "").replace(",", ""))
    if not has_pct_sign and -1 <= val <= 1:
        val *= 100
    return val


def parse_money(s):
    """Parse '$17,625', '17625.00', or 'USD 17,625.00' -> 17625.0"""
    if not s:
        return 0.0
    cleaned = s.replace("$", "").replace(",", "").replace("USD", "").strip()
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def fmt_money(v):
    """Format as $XX,XXX"""
    if v >= 1000:
        return f"${v:,.0f}"
    return f"${v:,.2f}"


def fmt_pct(v):
    return f"{v:.1f}%"


def pacing_badge(pct, day_of_month, total_days):
    expected = (day_of_month / total_days) * 100 if total_days else 0
    if pct >= expected:
        return '<span class="badge badge-green">On Track</span>'
    elif pct >= expected * 0.8:
        return '<span class="badge badge-yellow">At Risk</span>'
    else:
        return '<span class="badge badge-red">Behind</span>'


def pace_actual_vs_expected(actual, monthly_goal, biz_days_passed, biz_days_total):
    """Return (expected_by_now, pace_pct, daily_goal, weekly_goal)."""
    if not monthly_goal or not biz_days_total:
        return (0, 0, 0, 0)
    daily = monthly_goal / biz_days_total
    weekly = daily * 5
    expected = daily * biz_days_passed
    pace_pct = (actual / expected * 100) if expected else 0
    return (round(expected), pace_pct, round(daily), round(weekly))


def pace_color_and_label(actual, monthly_goal, biz_days_passed, biz_days_total):
    """Return HTML showing actual/expected with pace color."""
    expected, pace_pct, daily, weekly = pace_actual_vs_expected(actual, monthly_goal, biz_days_passed, biz_days_total)
    if not expected:
        return f'{actual}'
    if pace_pct >= 100:
        color = "var(--green)"
    elif pace_pct >= 80:
        color = "var(--yellow)"
    else:
        color = "var(--red)"
    return f'<span style="color:{color};font-weight:600">{actual}/{expected}</span> <span style="font-size:11px;color:var(--gray-400)">({pace_pct:.0f}%)</span>'


# Normalize name mismatches across CSVs (Salesforce uses full names, others use short)
NAME_MAP = {
    "Christopher Byrnes": "Chris Byrnes",
    "Chris Donovan": "Chris Donovan",  # separate person, kept as-is
}


def normalize_name(name):
    return NAME_MAP.get(name, name)


def pct_color(pct, good=70, warn=50):
    if pct >= good:
        return "var(--green)"
    elif pct >= warn:
        return "var(--yellow)"
    return "var(--red)"


def bar_color(pct, expected_pct):
    if pct >= expected_pct:
        return "var(--green)"
    elif pct >= expected_pct * 0.8:
        return "var(--yellow)"
    return "var(--red)"


def rep_link(name):
    slug = name.lower().replace(" ", "-")
    return f'<a href="reps/{slug}.html" style="color:inherit;text-decoration:none;border-bottom:1px dashed var(--gray-300)">{name}</a>'


BADGE_ON_TRACK = '<span class="badge badge-green">On Track</span>'
BADGE_AT_RISK = '<span class="badge badge-red">At Risk</span>'


def icon(name, size=18):
    """Return an inline img tag for a Toast icon asset."""
    return f'<img src="assets/{name}.svg" alt="" style="width:{size}px;height:{size}px;vertical-align:middle;opacity:0.7;">'


def sparkline(values, width=80, height=20, color="var(--toast-orange)"):
    """Generate an inline SVG sparkline from a list of numeric values."""
    if not values or len(values) < 2:
        return ""
    vmin, vmax = min(values), max(values)
    rng = vmax - vmin if vmax > vmin else 1
    n = len(values)
    points = []
    for i, v in enumerate(values):
        x = (i / (n - 1)) * width
        y = height - ((v - vmin) / rng) * height
        points.append(f"{x:.1f},{y:.1f}")
    points_str = " ".join(points)
    last_x = width
    last_y = height - ((values[-1] - vmin) / rng) * height
    return (
        f'<svg width="{width}" height="{height}" style="vertical-align:middle;overflow:visible">'
        f'<polyline points="{points_str}" fill="none" stroke="{color}" stroke-width="1.5" stroke-linejoin="round"/>'
        f'<circle cx="{last_x:.1f}" cy="{last_y:.1f}" r="2" fill="{color}"/>'
        f'</svg>'
    )


def pipeline_health(stale, total):
    if total == 0:
        return '<span class="badge badge-green">Healthy</span>'
    ratio = stale / total
    if ratio <= 0.1:
        return '<span class="badge badge-green">Healthy</span>'
    elif ratio <= 0.25:
        return '<span class="badge badge-yellow">Monitor</span>'
    return '<span class="badge badge-red">Needs attention</span>'


def generate(data_dir):
    data_dir = Path(data_dir)
    folder_name = data_dir.name  # e.g. "04.08.Data"
    date_prefix = folder_name.replace(".Data", "")  # "04.08"
    month_num, day_num = date_prefix.split(".")

    # Determine current month/year for display
    now = date.today()
    month_name = datetime(now.year, int(month_num), 1).strftime("%B %Y")
    data_date = f"{month_name[:3]} {int(day_num)}, {now.year}"

    # Calculate actual business days (Mon-Fri) in the month and how many have passed
    import calendar
    year = now.year
    m = int(month_num)
    day_of_month = int(day_num)
    total_days_in_month = calendar.monthrange(year, m)[1]
    total_biz_days = sum(1 for d in range(1, total_days_in_month + 1)
                         if calendar.weekday(year, m, d) < 5)
    biz_days_passed = sum(1 for d in range(1, day_of_month + 1)
                          if calendar.weekday(year, m, d) < 5)
    expected_pct = (biz_days_passed / total_biz_days) * 100 if total_biz_days else 0

    # Read CSVs
    km_file = list(data_dir.glob(f"*Key.Metrics*"))[0]
    wins_file = list(data_dir.glob(f"*Opp.Wins*"))[0]
    opps_file = list(data_dir.glob(f"*Open.Opps*"))[0]
    calls_file = list(data_dir.glob(f"*SL.Calls*"))[0]
    emails_file = list(data_dir.glob(f"*SL.Emails*"))[0]

    metrics = read_csv(km_file)
    wins = read_csv(wins_file)
    opps = read_csv(opps_file)
    calls = read_csv(calls_file)
    emails = read_csv(emails_file)

    # Sigma exports % columns as decimal fractions (1.16 = 116%). parse_pct only
    # auto-scales when val <= 1, so reps over goal are read as ~1% and miscounted.
    # Normalize known percent columns to "NNN.N%" form on load.
    pct_cols = (
        "Demos Held % to Goal", "EC ARR % to Goal", "Opp:Win % to Goal",
        "Opp % to Goal", "Demo:Win", "ARR + Ads % to Goal",
        "ARR + Ads Pacing in Month", "ARR % to Goal (Xactly)",
        "SS Opps Created % to Goal", "ARR Pacing in Month (Xactly)",
        "Wins % to Goal", "% to Goal", "Opp to Win", "Opp:Demo % to Goal",
        "Opp:Demo", "Opp:Demo Goal", "Avg Opp:Win Goal",
    )
    for r in metrics:
        for col in pct_cols:
            v = r.get(col, "")
            if not v or "%" in v:
                continue
            try:
                f = float(v.replace(",", ""))
            except ValueError:
                continue
            r[col] = f"{f * 100}%"

    # Pre-build lookup dicts
    calls_by_name = {r["user_name"]: r for r in calls}
    emails_by_name = {r["user_name"]: r for r in emails}

    # Rep order from metrics
    reps = [r["Rep Name"] for r in metrics]

    # --- UE Promo deals (not captured in Sigma/SF ARR — manually tracked) ---
    # Pull the most recent "UE Promo Deals {Month} - Sheet1*.csv" from Downloads
    # and inject $ into rep ARR so team sales reflect reality.
    ue_by_rep = {}
    ue_deals = []
    ue_csv_path = None
    downloads_dir = Path.home() / "Downloads"
    ue_candidates = sorted(downloads_dir.glob("UE Promo Deals*.csv"), key=os.path.getmtime, reverse=True)
    if ue_candidates:
        ue_csv_path = ue_candidates[0]
        for r in read_csv(ue_csv_path):
            rep = normalize_name((r.get("Rep") or "").strip())
            if not rep or rep not in reps:
                continue
            arr = parse_money(r.get("Software (Annual)", "0"))
            if arr <= 0:
                continue
            ue_by_rep[rep] = ue_by_rep.get(rep, 0.0) + arr
            ue_deals.append({
                "rep": rep,
                "account": r.get("Account Name", ""),
                "opp": r.get("Opportunity Name", ""),
                "arr": arr,
                "close": (r.get("Close Date") or next((v for k, v in r.items() if k and k.startswith("Close Date")), "")),
                "stage": r.get("Stage", ""),
            })

    # Inject UE Promo $ into per-rep ARR and recompute ARR % to Goal
    for r in metrics:
        name = r.get("Rep Name", "")
        ue = ue_by_rep.get(name, 0.0)
        if ue > 0:
            cur = parse_money(r.get("Total Booked Saas ARR", "0"))
            r["Total Booked Saas ARR"] = str(cur + ue)
            quota = parse_money(r.get("Booked SaaS Quota (Xactly)", "0"))
            if quota:
                r["ARR % to Goal (Xactly)"] = f"{(cur + ue) / quota * 100}%"

    # Overwrite $0 Opp Wins rows matching UE Promo accounts so Recent Wins tables show real $
    ue_by_account = {d["account"]: d["arr"] for d in ue_deals}
    for w in wins:
        acct = w.get("Account Name", "")
        if acct in ue_by_account and parse_money(w.get("Software (Annual)", "0")) == 0:
            w["Software (Annual)"] = str(ue_by_account[acct])

    # --- ROE clawbacks (deals taken away post-close per Rules of Engagement) ---
    # Reads ~/dm-dashboard/data/roe-clawbacks.csv with columns:
    #   Rep, Account Name, Opportunity Name, Software (Annual), Close Date, Reason
    # Subtracts ARR from rep totals, decrements Wins, removes the matching Opp Wins row.
    roe_by_rep = {}
    roe_deals = []
    roe_csv_path = Path(__file__).parent / "data" / "roe-clawbacks.csv"
    if roe_csv_path.exists():
        for r in read_csv(roe_csv_path):
            rep = normalize_name((r.get("Rep") or "").strip())
            if not rep or rep not in reps:
                continue
            arr = parse_money(r.get("Software (Annual)", "0"))
            if arr <= 0:
                continue
            roe_by_rep[rep] = roe_by_rep.get(rep, 0.0) + arr
            roe_deals.append({
                "rep": rep,
                "account": r.get("Account Name", ""),
                "opp": r.get("Opportunity Name", ""),
                "arr": arr,
                "close": r.get("Close Date", ""),
                "reason": r.get("Reason", ""),
            })

    for r in metrics:
        name = r.get("Rep Name", "")
        clawback = roe_by_rep.get(name, 0.0)
        if clawback > 0:
            cur = parse_money(r.get("Total Booked Saas ARR", "0"))
            new_arr = max(0.0, cur - clawback)
            r["Total Booked Saas ARR"] = str(new_arr)
            quota = parse_money(r.get("Booked SaaS Quota (Xactly)", "0"))
            if quota:
                r["ARR % to Goal (Xactly)"] = f"{new_arr / quota * 100}%"
            wins_now = int(r.get("Wins", "0") or 0)
            num_clawbacks = sum(1 for d in roe_deals if d["rep"] == name)
            r["Wins"] = str(max(0, wins_now - num_clawbacks))

    # Drop matching wins rows so Recent Wins tables don't show clawed-back deals
    roe_keys = {(d["rep"], d["account"], d["opp"]) for d in roe_deals}
    wins = [
        w for w in wins
        if (
            normalize_name(w.get("Opportunity Owner", "")),
            w.get("Account Name", ""),
            w.get("Opportunity Name", ""),
        ) not in roe_keys
    ]

    # --- Compute KPI totals ---
    team_arr = sum(parse_money(r.get("Total Booked Saas ARR", "0")) for r in metrics)
    # Team quota by month (includes waterfall adjustments)
    m = int(month_num)
    if m <= 3:  # Jan-Mar
        team_quota = 400440
    elif m <= 6:  # Apr-Jun
        team_quota = 447320
    else:  # Jul-Dec
        team_quota = 467000
    team_attainment = (team_arr / team_quota * 100) if team_quota else 0
    team_wins_count = sum(int(r.get("Wins", "0") or 0) for r in metrics)

    # Pipeline from open opps (only current team members)
    team_opps = [r for r in opps if normalize_name(r.get("Opportunity Owner", "")) in reps]
    total_pipeline = sum(parse_money(r.get("Software (Annual)", "0")) for r in team_opps)
    total_open_deals = len(team_opps)

    # Stale deals (>30 days old based on stage duration)
    stale_count = sum(1 for r in team_opps if float(r.get("Stage Duration", "0") or 0) > 30)

    # Per-rep pipeline aggregation (used by forecast + pipeline health sections)
    rep_pipeline = {}
    for r_opp in opps:
        owner = normalize_name(r_opp.get("Opportunity Owner", "Unknown"))
        if owner not in reps:
            continue
        if owner not in rep_pipeline:
            rep_pipeline[owner] = {"deals": 0, "value": 0.0, "stale": 0}
        rep_pipeline[owner]["deals"] += 1
        rep_pipeline[owner]["value"] += parse_money(r_opp.get("Software (Annual)", "0"))
        if float(r_opp.get("Stage Duration", "0") or 0) > 30:
            rep_pipeline[owner]["stale"] += 1

    # Load historical snapshots early (used by sparklines + forecast + trend tables)
    history_dir_early = Path(__file__).parent / "data" / "history"
    all_history = []
    if history_dir_early.exists():
        for hf in sorted(history_dir_early.glob("*.json")):
            with open(hf) as f:
                all_history.append(json.load(f))

    # Pacing
    team_pacing_val = (team_arr / team_quota * 100) if team_quota else 0

    # ========== BUILD HTML ==========
    html = []
    html.append("""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Growth Sales Team Dashboard — Josh Mellender</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Source+Sans+3:wght@400;600;700;800&display=swap" rel="stylesheet">
<style>
  :root {
    --toast-orange: #FF4C00; --toast-orange-light: #FF6A2B; --toast-orange-bg: #FFF3ED;
    --toast-navy: #2B4FB9;
    --toast-dark: #252525;
    --green: #22C55E; --green-bg: #F0FDF4;
    --yellow: #EAB308; --yellow-bg: #FEFCE8;
    --red: #EF4444; --red-bg: #FEF2F2;
    --blue: #3B82F6; --blue-bg: #EFF6FF;
    --warm-50: #F9F5F3; --warm-100: #F6F1EE;
    --gray-50: #F7FAFC; --gray-100: #F3F4F6; --gray-200: #E5E7EB;
    --gray-300: #D1D5DB; --gray-400: #9CA3AF; --gray-500: #6B7280;
    --gray-600: #4B5563; --gray-700: #374151; --gray-800: #252525;
    --shadow: 0 1px 3px rgba(0,0,0,0.06);
    --radius: 8px;
  }
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body { font-family: 'Source Sans 3', 'Source Sans Pro', system-ui, -apple-system, sans-serif; background: var(--warm-50); color: var(--gray-800); line-height: 1.5; font-size: 16px; }
  .header { background: white; color: var(--toast-dark); padding: 18px 32px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 12px; border-bottom: 1px solid var(--gray-200); }
  .header-left { display: flex; align-items: center; gap: 16px; }
  .header-left .logo { height: 28px; width: auto; }
  .header-left h1 { font-size: 22px; font-weight: 500; letter-spacing: -0.3px; color: var(--toast-dark); }
  .header-left p { font-size: 14px; color: var(--gray-500); margin-top: 2px; }
  .header-right { text-align: right; font-size: 14px; color: var(--gray-500); }
  .header-right .date { font-size: 16px; color: var(--toast-dark); font-weight: 600; }
  .kpi-banner { display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 12px; padding: 20px 32px; background: white; border-bottom: 1px solid var(--gray-200); }
  .kpi-card { text-align: center; padding: 14px 12px; border-radius: var(--radius); border: 1px solid var(--gray-200); }
  .kpi-card .label { font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px; color: var(--gray-500); font-weight: 600; }
  .kpi-card .value { font-size: 24px; font-weight: 600; margin: 4px 0; color: var(--toast-dark); }
  .kpi-card .sub { font-size: 12px; color: var(--gray-500); }
  .section { padding: 24px 32px; }
  .section-title { font-size: 18px; font-weight: 500; margin-bottom: 14px; display: flex; align-items: center; gap: 8px; color: var(--toast-dark); }
  table { width: 100%; border-collapse: collapse; background: white; border-radius: var(--radius); overflow: hidden; box-shadow: var(--shadow); margin-bottom: 8px; }
  th { background: var(--gray-50); font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px; color: var(--gray-600); font-weight: 600; padding: 10px 14px; text-align: left; border-bottom: 2px solid var(--gray-200); }
  td { padding: 10px 14px; border-top: 1px solid var(--gray-100); font-size: 14px; }
  tr:hover { background: var(--warm-100); }
  .progress-bar { background: var(--gray-200); border-radius: 999px; height: 8px; width: 100px; display: inline-block; vertical-align: middle; }
  .progress-fill { height: 100%; border-radius: 999px; }
  .badge { display: inline-block; padding: 3px 10px; border-radius: 999px; font-size: 11px; font-weight: 600; }
  .badge-green { background: var(--green-bg); color: var(--green); }
  .badge-yellow { background: var(--yellow-bg); color: var(--yellow); }
  .badge-red { background: var(--red-bg); color: var(--red); }
  .badge-blue { background: var(--blue-bg); color: var(--blue); }
  .badge-orange { background: var(--toast-orange-bg); color: var(--toast-orange); }
  .forecast-bar { display: flex; align-items: center; gap: 8px; margin: 8px 0; }
  .forecast-bar-track { flex: 1; background: var(--gray-200); border-radius: 6px; height: 24px; position: relative; }
  .forecast-bar-fill { height: 100%; border-radius: 6px; }
  .forecast-bar-label { font-size: 14px; font-weight: 600; min-width: 50px; text-align: right; }
  .forecast-bar-name { font-size: 15px; font-weight: 500; min-width: 130px; }
  .pace-line { position: absolute; top: -4px; bottom: -4px; width: 2px; background: var(--toast-orange); z-index: 2; }
  .pace-line::after { content: "\\25BC"; position: absolute; top: -14px; left: -4px; font-size: 9px; color: var(--toast-orange); }
  .two-col { display: grid; grid-template-columns: 1fr 1fr; gap: 24px; }
  @media (max-width: 900px) { .two-col { grid-template-columns: 1fr; } }
  .footer { text-align: center; padding: 20px; font-size: 13px; color: var(--gray-400); border-top: 1px solid var(--gray-200); margin-top: 8px; }
</style>
</head>
<body>
""")

    # Header
    html.append(f"""
<div class="header">
  <div class="header-left">
    <img class="logo" src="assets/toast-logo_color.svg" alt="Toast" style="height:28px;width:auto;">
    <div>
      <h1>Growth Sales Dashboard</h1>
      <p>Josh Mellender — Growth Sales District Manager | Reports to Katie Patterson (RVP)</p>
    </div>
  </div>
  <div class="header-right">
    <div class="date">{month_name} — MTD Performance</div>
    <div>Data as of {data_date}</div>
  </div>
</div>
""")

    # KPI Banner
    pacing_color = "var(--red)" if team_pacing_val < expected_pct else "var(--green)"
    pace_target = team_quota * expected_pct / 100
    pace_delta = team_arr - pace_target
    if pace_delta < 0:
        pacing_label = f"Behind by {fmt_money(-pace_delta)}"
    else:
        pacing_label = f"Ahead by {fmt_money(pace_delta)}"
    html.append(f"""
<div class="kpi-banner">
  <div class="kpi-card"><div class="label">Team MTD ACV</div><div class="value">{fmt_money(team_arr)}</div><div class="sub">of {fmt_money(team_quota)} quota</div></div>
  <div class="kpi-card"><div class="label">Attainment</div><div class="value">{fmt_pct(team_attainment)}</div><div class="sub">Working Day {biz_days_passed}/{total_biz_days} &middot; pace {fmt_pct(expected_pct)}</div></div>
  <div class="kpi-card"><div class="label">Pace Target</div><div class="value">{fmt_money(pace_target)}</div><div class="sub" style="color:{pacing_color};font-weight:600">{pacing_label}</div></div>
  <div class="kpi-card"><div class="label">Closed Deals</div><div class="value">{team_wins_count}</div><div class="sub">MTD</div></div>
  <div class="kpi-card"><div class="label">Open Pipeline</div><div class="value">{fmt_money(total_pipeline)}</div><div class="sub">{total_open_deals} opps</div></div>
  <div class="kpi-card"><div class="label">Stale (&gt;30d)</div><div class="value" style="color: var(--red)">{stale_count}</div><div class="sub">stuck deals</div></div>
</div>
""")

    # ---- Team Overview Table ----
    html.append("""
<div class="section">
  <div class="section-title"><img src="assets/bar-chart.svg" alt="" style="width:18px;height:18px;vertical-align:middle;opacity:0.7;"> Team Overview — MTD ACV vs Quota</div>
  <table>
    <thead><tr><th>Rep</th><th>Role</th><th>Quota</th><th>MTD ACV</th><th>%</th><th>Progress</th><th>Trend (8d)</th><th>Wins</th><th>Pacing</th></tr></thead>
    <tbody>
""")
    for r in metrics:
        name = r["Rep Name"]
        role = r.get("Role in Month", "AE")
        cohort = r.get("Ramping Cohort", "")
        role_display = f"{role} ({cohort})" if cohort and cohort != "-" else role
        quota = parse_money(r.get("Booked SaaS Quota (Xactly)", "0"))
        arr = parse_money(r.get("Total Booked Saas ARR", "0"))
        pct = parse_pct(r.get("ARR % to Goal (Xactly)", "0"))
        wins_count = int(r.get("Wins", "0") or 0)
        bar_w = min(pct, 100)
        color = bar_color(pct, expected_pct)
        badge = pacing_badge(pct, biz_days_passed, total_biz_days)
        # Sparkline: last up-to-8 historical ARR values + today's actual
        spark_values = [h.get("reps", {}).get(name, {}).get("arr", 0) for h in all_history[-7:]]
        spark_values.append(arr)
        spark = sparkline(spark_values, width=90, height=22, color=color)
        html.append(f'<tr><td><strong>{rep_link(name)}</strong></td><td style="font-size:12px;color:var(--gray-500)">{role_display}</td><td>{fmt_money(quota)}</td><td><strong>{fmt_money(arr)}</strong></td><td>{fmt_pct(pct)}</td><td><div class="progress-bar"><div class="progress-fill" style="width:{bar_w:.1f}%;background:{color}"></div></div></td><td>{spark}</td><td>{wins_count}</td><td>{badge}</td></tr>\n')

    team_spark_values = [h.get("team_arr", 0) for h in all_history[-7:]]
    team_spark_values.append(team_arr)
    team_spark = sparkline(team_spark_values, width=90, height=22, color="var(--toast-orange)")
    html.append(f'<tr style="background:var(--gray-100);font-weight:700"><td>TEAM</td><td>{len(metrics)} AEs</td><td>{fmt_money(team_quota)}</td><td>{fmt_money(team_arr)}</td><td>{fmt_pct(team_attainment)}</td><td></td><td>{team_spark}</td><td>{team_wins_count}</td><td></td></tr>\n')
    html.append("    </tbody></table>\n</div>\n")

    # ---- Close the Gap Calculator ----
    remaining_biz_days = total_biz_days - biz_days_passed
    remaining_weeks = remaining_biz_days / 5 if remaining_biz_days > 0 else 0.01
    html.append("""
<div class="section">
  <div class="section-title"><img src="assets/trending-up.svg" alt="" style="width:18px;height:18px;vertical-align:middle;opacity:0.7;"> Close the Gap — What It Takes to Hit Quota</div>
  <p style="font-size:13px;color:var(--gray-500);margin-bottom:14px">Based on each rep's current conversion rates and remaining working days ({remaining} biz days, {weeks:.1f} weeks left).</p>
  <table>
    <thead><tr><th>Rep</th><th>Gap to Quota</th><th>Deals Needed</th><th>Demos Needed</th><th>Demos/Week</th><th>Calls/Day</th><th>Feasibility</th></tr></thead>
    <tbody>
""".replace("{remaining}", str(remaining_biz_days)).replace("{weeks}", f"{remaining_weeks:.1f}"))
    for r in metrics:
        name = r["Rep Name"]
        quota = parse_money(r.get("Booked SaaS Quota (Xactly)", "0"))
        arr = parse_money(r.get("Total Booked Saas ARR", "0"))
        gap = quota - arr
        if gap <= 0 or quota == 0:
            html.append(f'<tr><td><strong>{rep_link(name)}</strong></td><td style="color:var(--green);font-weight:600">HIT QUOTA</td><td colspan="5" style="color:var(--green)">At {fmt_pct(arr/quota*100) if quota else 0} — exceeding target</td></tr>\n')
            continue
        avg_arr = parse_money(r.get("ARR Won per Opp", "0"))
        if avg_arr <= 0:
            avg_arr = (team_arr / team_wins_count) if team_wins_count else 1800
        demo_win_rate = parse_pct(r.get("Demo:Win", "0")) / 100
        if demo_win_rate <= 0:
            demo_win_rate = 0.5
        deals_needed = gap / avg_arr
        demos_needed = deals_needed / demo_win_rate
        demos_per_week = demos_needed / remaining_weeks if remaining_weeks > 0 else 999
        call_goal_mo = int(r.get("Call Goal", "0") or 0)
        calls_per_day = round(call_goal_mo / total_biz_days) if total_biz_days else 0
        demo_goal_mo = int(r.get("Total Demo Goal", "0") or 0)
        normal_demos_wk = round(demo_goal_mo / 4.3) if demo_goal_mo else 0
        if demos_per_week <= normal_demos_wk * 1.0:
            feasibility = '<span class="badge badge-green">Achievable</span>'
        elif demos_per_week <= normal_demos_wk * 1.5:
            feasibility = '<span class="badge badge-yellow">Stretch</span>'
        else:
            feasibility = '<span class="badge badge-red">Needs Wins Now</span>'
        html.append(f'<tr><td><strong>{rep_link(name)}</strong></td><td style="font-weight:600">{fmt_money(gap)}</td><td>{deals_needed:.1f} <span style="font-size:11px;color:var(--gray-400)">@ {fmt_money(avg_arr)} avg</span></td><td>{demos_needed:.1f} <span style="font-size:11px;color:var(--gray-400)">@ {demo_win_rate*100:.0f}% D:W</span></td><td style="font-weight:600;color:{pct_color(normal_demos_wk / max(demos_per_week, 0.01) * 100, 100, 67)}">{demos_per_week:.1f}/wk</td><td>{calls_per_day}/day</td><td>{feasibility}</td></tr>\n')
    html.append("    </tbody></table>\n</div>\n")

    # ---- Month-End Forecast Simulator ----
    # Compute 3 forecast methodologies per rep:
    # 1. Linear pace: project current pace to end of month
    # 2. Recent run-rate: last 3 snapshots' daily delta, extrapolated
    # 3. Pipeline-based: current ARR + (open pipeline * rep's demo:win rate)
    history_for_forecast = all_history  # reuse early-loaded history
    team_forecast_linear = 0
    team_forecast_runrate = 0
    team_forecast_pipeline = 0
    forecast_rows = []
    for r in metrics:
        name = r["Rep Name"]
        quota = parse_money(r.get("Booked SaaS Quota (Xactly)", "0"))
        arr_cur = parse_money(r.get("Total Booked Saas ARR", "0"))

        # Method 1: Linear pace
        if biz_days_passed > 0:
            daily_pace = arr_cur / biz_days_passed
            forecast_linear = daily_pace * total_biz_days
        else:
            forecast_linear = arr_cur

        # Method 2: Recent run-rate (last 3 snapshots)
        forecast_runrate = arr_cur
        if len(history_for_forecast) >= 3:
            recent = history_for_forecast[-3:]
            first_arr = recent[0].get("reps", {}).get(name, {}).get("arr", arr_cur)
            last_arr = recent[-1].get("reps", {}).get(name, {}).get("arr", arr_cur)
            recent_daily = (last_arr - first_arr) / max(len(recent) - 1, 1)
            remaining_days = total_biz_days - biz_days_passed
            forecast_runrate = arr_cur + (recent_daily * remaining_days)

        # Method 3: Pipeline-based (arr + pipeline * demo_win)
        rep_pipe_val = rep_pipeline.get(name, {"value": 0})["value"] if name in rep_pipeline else 0
        rep_dw = parse_pct(r.get("Demo:Win", "0")) / 100
        if rep_dw <= 0:
            rep_dw = 0.3  # conservative default
        # Assume only a fraction of pipeline can close this month (scaled by time left)
        time_left_factor = (total_biz_days - biz_days_passed) / max(total_biz_days, 1) * 2  # aggressive: 2x remaining fraction
        forecast_pipeline = arr_cur + (rep_pipe_val * rep_dw * min(time_left_factor, 0.8))

        team_forecast_linear += forecast_linear
        team_forecast_runrate += forecast_runrate
        team_forecast_pipeline += forecast_pipeline
        forecast_rows.append({
            "name": name, "quota": quota, "cur": arr_cur,
            "linear": forecast_linear, "runrate": forecast_runrate, "pipeline": forecast_pipeline
        })

    team_linear_pct = (team_forecast_linear / team_quota * 100) if team_quota else 0
    team_runrate_pct = (team_forecast_runrate / team_quota * 100) if team_quota else 0
    team_pipeline_pct = (team_forecast_pipeline / team_quota * 100) if team_quota else 0

    html.append(f"""
<div class="section">
  <div class="section-title"><img src="assets/trending-up.svg" alt="" style="width:18px;height:18px;vertical-align:middle;opacity:0.7;"> Month-End Forecast — Where Will We Land?</div>
  <p style="font-size:13px;color:var(--gray-500);margin-bottom:14px">Three projection methods to bracket the likely outcome at month-end (Working Day {biz_days_passed}/{total_biz_days}).</p>
  <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:16px;margin-bottom:16px">
    <div style="padding:16px;background:var(--blue-bg);border-radius:var(--radius);border-left:4px solid var(--blue)">
      <div style="font-size:11px;text-transform:uppercase;letter-spacing:0.5px;color:var(--gray-600);font-weight:600">Linear Pace</div>
      <div style="font-size:24px;font-weight:600;color:var(--toast-dark);margin:4px 0">{fmt_money(team_forecast_linear)}</div>
      <div style="font-size:13px;color:{pct_color(team_linear_pct, 100, 80)};font-weight:600">{fmt_pct(team_linear_pct)} of quota</div>
      <div style="font-size:11px;color:var(--gray-500);margin-top:4px">If today's daily rate continues</div>
    </div>
    <div style="padding:16px;background:var(--warm-100);border-radius:var(--radius);border-left:4px solid var(--toast-orange)">
      <div style="font-size:11px;text-transform:uppercase;letter-spacing:0.5px;color:var(--gray-600);font-weight:600">Recent Run-Rate</div>
      <div style="font-size:24px;font-weight:600;color:var(--toast-dark);margin:4px 0">{fmt_money(team_forecast_runrate)}</div>
      <div style="font-size:13px;color:{pct_color(team_runrate_pct, 100, 80)};font-weight:600">{fmt_pct(team_runrate_pct)} of quota</div>
      <div style="font-size:11px;color:var(--gray-500);margin-top:4px">Last 3 snapshots daily delta</div>
    </div>
    <div style="padding:16px;background:var(--green-bg);border-radius:var(--radius);border-left:4px solid var(--green)">
      <div style="font-size:11px;text-transform:uppercase;letter-spacing:0.5px;color:var(--gray-600);font-weight:600">Pipeline-Based</div>
      <div style="font-size:24px;font-weight:600;color:var(--toast-dark);margin:4px 0">{fmt_money(team_forecast_pipeline)}</div>
      <div style="font-size:13px;color:{pct_color(team_pipeline_pct, 100, 80)};font-weight:600">{fmt_pct(team_pipeline_pct)} of quota</div>
      <div style="font-size:11px;color:var(--gray-500);margin-top:4px">Current + pipeline × D:W</div>
    </div>
  </div>
  <table>
    <thead><tr><th>Rep</th><th>Quota</th><th>Current</th><th>Linear</th><th>Run-Rate</th><th>Pipeline</th><th>Range</th></tr></thead>
    <tbody>
""")
    for fr in forecast_rows:
        low = min(fr["linear"], fr["runrate"], fr["pipeline"])
        high = max(fr["linear"], fr["runrate"], fr["pipeline"])
        if fr["quota"] == 0:
            pct_of_quota_lo = 0
            pct_of_quota_hi = 0
            range_display = f'{fmt_money(low)} – {fmt_money(high)}'
            range_color = "var(--gray-500)"
        else:
            pct_of_quota_lo = low / fr["quota"] * 100
            pct_of_quota_hi = high / fr["quota"] * 100
            range_display = f'{fmt_money(low)} – {fmt_money(high)}<br><span style="font-size:11px;color:var(--gray-500)">{fmt_pct(pct_of_quota_lo)} – {fmt_pct(pct_of_quota_hi)}</span>'
            range_color = pct_color(pct_of_quota_hi, 100, 80)
        html.append(f'<tr><td><strong>{rep_link(fr["name"])}</strong></td><td>{fmt_money(fr["quota"])}</td><td>{fmt_money(fr["cur"])}</td><td>{fmt_money(fr["linear"])}</td><td>{fmt_money(fr["runrate"])}</td><td>{fmt_money(fr["pipeline"])}</td><td style="color:{range_color};font-weight:600">{range_display}</td></tr>\n')
    team_low = min(team_forecast_linear, team_forecast_runrate, team_forecast_pipeline)
    team_high = max(team_forecast_linear, team_forecast_runrate, team_forecast_pipeline)
    team_range_pct_lo = team_low / team_quota * 100 if team_quota else 0
    team_range_pct_hi = team_high / team_quota * 100 if team_quota else 0
    html.append(f'<tr style="background:var(--gray-100);font-weight:700"><td>TEAM</td><td>{fmt_money(team_quota)}</td><td>{fmt_money(team_arr)}</td><td>{fmt_money(team_forecast_linear)}</td><td>{fmt_money(team_forecast_runrate)}</td><td>{fmt_money(team_forecast_pipeline)}</td><td>{fmt_money(team_low)} – {fmt_money(team_high)}<br><span style="font-size:11px">{fmt_pct(team_range_pct_lo)} – {fmt_pct(team_range_pct_hi)}</span></td></tr>\n')
    html.append("    </tbody></table>\n</div>\n")

    # ---- MBO Tracker ----
    # Compute MBO metrics from data
    total_opps = sum(int(r.get("Opps", "0") or 0) for r in metrics)
    total_demos = sum(int(r.get("Demos", "0") or 0) for r in metrics)
    total_ss = sum(int(r.get("Self-Sourced Created Opps", "0") or 0) for r in metrics)
    team_ss_goal = sum(int(r.get("Total SS Opp Goal", "0") or 0) for r in metrics)
    total_nbr = sum(int(r.get("NB Referral Created Opps", "0") or 0) for r in metrics)
    avg_arr_per_opp = (team_arr / team_wins_count) if team_wins_count else 0
    reps_at_quota = sum(1 for r in metrics if parse_pct(r.get("ARR % to Goal (Xactly)", "0")) >= 100)
    accountable_reps = sum(1 for r in metrics if parse_money(r.get("Booked SaaS Quota (Xactly)", "0")) > 0)
    opps_per_rep = total_opps / len(metrics) if metrics else 0

    # Opp:Win and Opp:Demo conversion (team averages)
    opp_win_vals = [parse_pct(r.get("Opp to Win", "0")) for r in metrics if r.get("Opp to Win")]
    opp_demo_vals = [parse_pct(r.get("Opp:Demo", "0")) for r in metrics if r.get("Opp:Demo")]
    team_opp_win = (sum(opp_win_vals) / len(opp_win_vals)) if opp_win_vals else 0
    team_opp_demo = (sum(opp_demo_vals) / len(opp_demo_vals)) if opp_demo_vals else 0

    # Payroll ARR as % of total ARR
    total_ec_arr = sum(parse_money(r.get("EC ARR", "0")) for r in metrics)
    payroll_pct_of_arr = (total_ec_arr / team_arr * 100) if team_arr else 0

    # NB Referral Demo:Win — use Demo:Win column for reps with NBR activity
    demo_win_vals = [parse_pct(r.get("Demo:Win", "0")) for r in metrics if int(r.get("NB Referral Created Opps", "0") or 0) > 0]
    nbr_demo_win = (sum(demo_win_vals) / len(demo_win_vals)) if demo_win_vals else 0

    # Headcount / People (configurable)
    total_headcount = len(metrics)
    total_seats = 7  # current team size = filled seats
    aes_in_pm = 1  # manually tracked
    attrition = 0

    html.append("""
<div class="section">
  <div class="section-title"><img src="assets/Other_Award.svg" alt="" style="width:18px;height:18px;vertical-align:middle;opacity:0.7;"> MBO Tracker — District Scorecard</div>
  <div class="two-col">
    <div>
      <table>
        <thead><tr><th>Attainment to Quota</th><th>Goal</th><th>Actual</th><th>District % to Goal</th></tr></thead>
        <tbody>
""")
    html.append(f'<tr><td>Total ARR (Controlled)</td><td>{fmt_money(team_quota)}</td><td>{fmt_money(team_arr)}</td><td style="color:{pct_color(team_attainment, 100, 80)};font-weight:600">{fmt_pct(team_attainment)}</td></tr>\n')
    html.append(f'<tr><td>% of AEs at ARR Quota</td><td>100%</td><td>{reps_at_quota}/{accountable_reps}</td><td style="font-weight:600">{fmt_pct(reps_at_quota/accountable_reps*100) if accountable_reps else "N/A"}</td></tr>\n')
    html.append("""        </tbody></table>
      <table style="margin-top:12px">
        <thead><tr><th>MBOs</th><th>Goal</th><th>Actual</th><th>Status</th></tr></thead>
        <tbody>
""")
    payroll_goal_pct = 40  # 36%/40% — using 40% as target
    payroll_color = pct_color(payroll_pct_of_arr, 40, 36)
    payroll_badge = BADGE_ON_TRACK if payroll_pct_of_arr >= 36 else BADGE_AT_RISK
    html.append(f'<tr><td>Payroll ARR as % of Total ARR</td><td>36% / 40%</td><td style="font-weight:600">{fmt_pct(payroll_pct_of_arr)}</td><td>{payroll_badge}</td></tr>\n')
    nbr_goal = 83  # 80%/83%
    nbr_color = pct_color(nbr_demo_win, 83, 80)
    nbr_badge = BADGE_ON_TRACK if nbr_demo_win >= 80 else BADGE_AT_RISK
    html.append(f'<tr><td>NB Referral Demo:Win Conv</td><td>80% / 83%</td><td style="font-weight:600">{fmt_pct(nbr_demo_win)}</td><td>{nbr_badge}</td></tr>\n')
    html.append("        </tbody></table>\n    </div>\n    <div>\n")

    total_opp_goal = sum(int(r.get("Opp Goal", "0") or 0) for r in metrics)
    total_demo_goal = sum(int(r.get("Total Demo Goal", "0") or 0) for r in metrics)
    total_wins_goal = sum(int(r.get("Total Wins Goal", "0") or 0) for r in metrics)
    total_call_goal = sum(int(r.get("Call Goal", "0") or 0) for r in metrics)

    html.append("""      <table>
        <thead><tr><th>Key Metrics</th><th>Goal (Mo)</th><th>Actual</th><th>% to Goal</th><th>Pace</th></tr></thead>
        <tbody>
""")
    opp_pct = (total_opps / total_opp_goal * 100) if total_opp_goal else 0
    html.append(f'<tr><td>Total Opps Created</td><td>{total_opp_goal}</td><td>{total_opps}</td><td>{fmt_pct(opp_pct)}</td><td>{pacing_badge(opp_pct, biz_days_passed, total_biz_days)}</td></tr>\n')
    html.append(f'<tr><td>Opps Created Per Rep</td><td>{total_opp_goal/len(metrics):.0f}</td><td>{opps_per_rep:.1f}</td><td></td><td></td></tr>\n')
    avg_arr_color = pct_color(avg_arr_per_opp / 1800 * 100, 100, 80)
    html.append(f'<tr><td>Avg ARR</td><td>$1,800</td><td style="font-weight:600">{fmt_money(avg_arr_per_opp)}</td><td style="color:{avg_arr_color};font-weight:600">{fmt_pct(avg_arr_per_opp/1800*100)}</td><td></td></tr>\n')
    html.append(f'<tr><td>Opp:Win Conversion</td><td></td><td>{fmt_pct(team_opp_win)}</td><td></td><td></td></tr>\n')
    html.append(f'<tr><td>Opp:Demo Conversion</td><td></td><td>{fmt_pct(team_opp_demo)}</td><td></td><td></td></tr>\n')
    html.append(f'<tr><td>NBR Referral Opps (total)</td><td></td><td>{total_nbr}</td><td></td><td></td></tr>\n')
    ss_pct = (total_ss / team_ss_goal * 100) if team_ss_goal else 0
    html.append(f'<tr><td>SS Opps (total)</td><td>{team_ss_goal}</td><td>{total_ss}</td><td>{fmt_pct(ss_pct)}</td><td>{pacing_badge(ss_pct, biz_days_passed, total_biz_days)}</td></tr>\n')
    wins_pct = (team_wins_count / total_wins_goal * 100) if total_wins_goal else 0
    html.append(f'<tr><td>Total Wins</td><td>{total_wins_goal}</td><td>{team_wins_count}</td><td>{fmt_pct(wins_pct)}</td><td>{pacing_badge(wins_pct, biz_days_passed, total_biz_days)}</td></tr>\n')
    total_demos_actual = sum(int(r.get("Demos", "0") or 0) for r in metrics)
    demos_pct = (total_demos_actual / total_demo_goal * 100) if total_demo_goal else 0
    html.append(f'<tr><td>Total Demos</td><td>{total_demo_goal}</td><td>{total_demos_actual}</td><td>{fmt_pct(demos_pct)}</td><td>{pacing_badge(demos_pct, biz_days_passed, total_biz_days)}</td></tr>\n')
    total_calls_actual = sum(int(calls_by_name.get(r["Rep Name"], {}).get("calls", "0") or 0) for r in metrics)
    calls_pct = (total_calls_actual / total_call_goal * 100) if total_call_goal else 0
    html.append(f'<tr><td>Total Calls</td><td>{total_call_goal}</td><td>{total_calls_actual}</td><td>{fmt_pct(calls_pct)}</td><td>{pacing_badge(calls_pct, biz_days_passed, total_biz_days)}</td></tr>\n')
    html.append("""        </tbody></table>
      <table style="margin-top:12px">
        <thead><tr><th>People</th><th>Goal</th><th>Actual</th><th>Status</th></tr></thead>
        <tbody>
""")
    html.append(f'<tr><td>Headcount</td><td>{total_seats}</td><td>{total_headcount}/{total_seats}</td><td>{BADGE_ON_TRACK}</td></tr>\n')
    html.append(f'<tr><td>OOT Attrition</td><td>0</td><td>{attrition}</td><td>{BADGE_ON_TRACK}</td></tr>\n')
    pm_badge = BADGE_AT_RISK if aes_in_pm > 0 else BADGE_ON_TRACK
    html.append(f'<tr><td>AEs in PM</td><td>0</td><td>{aes_in_pm}</td><td>{pm_badge}</td></tr>\n')
    html.append("        </tbody></table>\n    </div>\n  </div>\n</div>\n")

    # ---- Self-Sourced + NBR Referrals (two-col) ----
    html.append("""
<div class="section">
  <div class="two-col">
    <div>
      <div class="section-title"><img src="assets/star.svg" alt="" style="width:18px;height:18px;vertical-align:middle;opacity:0.7;"> Self-Sourced Opps Created</div>
      <table>
        <thead><tr><th>Rep</th><th>Created</th><th>Goal</th><th>% to Goal</th><th>Pace</th></tr></thead>
        <tbody>
""")
    team_ss = 0
    team_ss_goal = 0
    for r in metrics:
        name = r["Rep Name"]
        ss = int(r.get("Self-Sourced Created Opps", "0") or 0)
        goal = int(r.get("Total SS Opp Goal", "0") or 0)
        pct = parse_pct(r.get("SS Opps Created % to Goal", "0"))
        team_ss += ss
        team_ss_goal += goal
        _, ss_pace, _, _ = pace_actual_vs_expected(ss, goal, biz_days_passed, total_biz_days)
        pace_badge = pacing_badge(pct, biz_days_passed, total_biz_days)
        html.append(f'<tr><td><strong>{name}</strong></td><td>{ss}</td><td>{goal}</td><td>{fmt_pct(pct)}</td><td>{pace_badge}</td></tr>\n')
    html.append(f'<tr style="background:var(--gray-100);font-weight:700"><td>TEAM</td><td>{team_ss}</td><td>{team_ss_goal}</td><td></td><td></td></tr>\n')
    html.append("        </tbody></table>\n    </div>\n")

    # NBR Referrals
    html.append("""    <div>
      <div class="section-title"><img src="assets/restaurant-group.svg" alt="" style="width:18px;height:18px;vertical-align:middle;opacity:0.7;"> NBR Referrals</div>
      <table>
        <thead><tr><th>Rep</th><th>Created</th><th>Closed Won</th><th>Demo:Win %</th></tr></thead>
        <tbody>
""")
    team_nbr_created = 0
    team_nbr_won = 0
    team_demos_total = 0
    team_wins_from_demos = 0
    for r in metrics:
        name = r["Rep Name"]
        created = int(r.get("NB Referral Created Opps", "0") or 0)
        won = int(r.get("NB Referral Closed Won", "0") or 0)
        demo_win = parse_pct(r.get("Demo:Win", "0"))
        demos = int(r.get("Demos", "0") or 0)
        wins_count = int(r.get("Wins", "0") or 0)
        team_nbr_created += created
        team_nbr_won += won
        team_demos_total += demos
        team_wins_from_demos += wins_count
        dw_color = pct_color(demo_win, 80, 60)
        html.append(f'<tr><td><strong>{name}</strong></td><td>{created}</td><td>{won}</td><td style="color:{dw_color};font-weight:600">{fmt_pct(demo_win)}</td></tr>\n')
    team_demo_win = (team_wins_from_demos / team_demos_total * 100) if team_demos_total else 0
    team_dw_color = pct_color(team_demo_win, 80, 60)
    nbr_mbo_badge = BADGE_ON_TRACK if team_demo_win >= 80 else BADGE_AT_RISK
    html.append(f'<tr style="background:var(--gray-100);font-weight:700"><td>TEAM</td><td>{team_nbr_created}</td><td>{team_nbr_won}</td><td style="color:{team_dw_color}">{fmt_pct(team_demo_win)}</td></tr>\n')
    html.append(f'    </tbody></table>\n')
    html.append(f'    <p style="margin-top:6px;font-size:12px;color:var(--gray-500)">MBO Target: Demo:Win 80% / 83% — Currently {fmt_pct(team_demo_win)} {nbr_mbo_badge}</p>\n')
    html.append("    </div>\n  </div>\n</div>\n")

    # ---- EC / Payroll Leaderboard ----
    ec_data = []
    total_ec_units = 0
    total_ec_arr_sum = 0
    total_ec_goal_sum = 0
    for r in metrics:
        name = r["Rep Name"]
        units = int(r.get("EC Units", "0") or 0)
        ec_arr = parse_money(r.get("EC ARR", "0"))
        ec_goal = parse_money(r.get("EC ARR Goal", "0"))
        ec_pct = parse_pct(r.get("EC ARR % to Goal", "0"))
        rep_arr = parse_money(r.get("Total Booked Saas ARR", "0"))
        rep_ec_attach = (ec_arr / rep_arr * 100) if rep_arr else 0
        total_ec_units += units
        total_ec_arr_sum += ec_arr
        total_ec_goal_sum += ec_goal
        ec_data.append({"name": name, "units": units, "ec_arr": ec_arr, "ec_goal": ec_goal, "ec_pct": ec_pct, "attach": rep_ec_attach, "rep_arr": rep_arr})
    ec_data.sort(key=lambda x: x["ec_arr"], reverse=True)
    total_ec_pct = (total_ec_arr_sum / total_ec_goal_sum * 100) if total_ec_goal_sum else 0
    ec_pct_of_total = (total_ec_arr_sum / team_arr * 100) if team_arr else 0
    ec_mbo_color = pct_color(ec_pct_of_total, 40, 36)
    ec_mbo_badge = BADGE_ON_TRACK if ec_pct_of_total >= 36 else BADGE_AT_RISK
    ec_gap_to_36 = max(0, team_arr * 0.36 - total_ec_arr_sum)

    html.append(f"""
<div class="section">
  <div class="section-title"><img src="assets/pay-check.svg" alt="" style="width:18px;height:18px;vertical-align:middle;opacity:0.7;"> EC / Payroll Leaderboard <span style="font-size:13px;font-weight:400;color:var(--gray-500)">— MBO Target: 36% / 40% of Total ARR</span></div>
  <div style="display:flex;gap:16px;margin-bottom:14px;flex-wrap:wrap">
    <div style="padding:10px 16px;background:{'var(--green-bg)' if ec_pct_of_total >= 36 else 'var(--red-bg)'};border-radius:var(--radius);font-size:14px"><strong style="color:{'var(--green)' if ec_pct_of_total >= 36 else 'var(--red)'}">{fmt_pct(ec_pct_of_total)}</strong> EC Attach Rate {ec_mbo_badge}</div>
    <div style="padding:10px 16px;background:var(--gray-100);border-radius:var(--radius);font-size:14px"><strong>{fmt_money(total_ec_arr_sum)}</strong> EC ARR of <strong>{fmt_money(team_arr)}</strong> total</div>
""")
    if ec_gap_to_36 > 0:
        html.append(f'    <div style="padding:10px 16px;background:var(--yellow-bg);border-radius:var(--radius);font-size:14px">Need <strong style="color:var(--yellow)">{fmt_money(ec_gap_to_36)}</strong> more EC ARR to hit 36%</div>\n')
    html.append("""  </div>
  <table>
    <thead><tr><th>#</th><th>Rep</th><th>EC Units</th><th>EC ARR</th><th>Attach Rate</th><th>% of Team EC</th><th>EC Goal</th><th>% to Goal</th><th>Pace</th></tr></thead>
    <tbody>
""")
    for rank, ed in enumerate(ec_data, 1):
        team_contribution = (ed["ec_arr"] / total_ec_arr_sum * 100) if total_ec_arr_sum else 0
        attach_color = pct_color(ed["attach"], 40, 25)
        ec_badge = pacing_badge(ed["ec_pct"], biz_days_passed, total_biz_days)
        medal = {1: " &#x1F947;", 2: " &#x1F948;", 3: " &#x1F949;"}.get(rank, "")
        html.append(f'<tr><td style="font-weight:700;font-size:16px">{rank}{medal}</td><td><strong>{ed["name"]}</strong></td><td>{ed["units"]}</td><td style="font-weight:600">{fmt_money(ed["ec_arr"])}</td><td style="color:{attach_color};font-weight:600">{fmt_pct(ed["attach"])}</td><td>{fmt_pct(team_contribution)}</td><td>{fmt_money(ed["ec_goal"])}</td><td>{fmt_pct(ed["ec_pct"])}</td><td>{ec_badge}</td></tr>\n')
    html.append(f'<tr style="background:var(--gray-100);font-weight:700"><td></td><td>TEAM</td><td>{total_ec_units}</td><td>{fmt_money(total_ec_arr_sum)}</td><td style="color:{ec_mbo_color}">{fmt_pct(ec_pct_of_total)}</td><td>100%</td><td>{fmt_money(total_ec_goal_sum)}</td><td>{fmt_pct(total_ec_pct)}</td><td></td></tr>\n')
    html.append("    </tbody></table>\n</div>\n")

    # ---- Enablement MBO Tracker (from WorkRamp data) ----
    enablement_path = Path(__file__).parent / "data" / "enablement.json"
    if enablement_path.exists():
        with open(enablement_path) as f:
            enablement = json.load(f)

        e_date = enablement.get("scraped_date", "unknown")
        team_comp = enablement.get("team_completion", 0)
        team_grade = enablement.get("team_avg_grade", 0)
        mbo_threshold = 85
        mbo_met = team_comp >= mbo_threshold

        html.append(f"""
<div class="section">
  <div class="section-title"><img src="assets/analytics-check.svg" alt="" style="width:18px;height:18px;vertical-align:middle;opacity:0.7;"> Enablement MBO Tracker <span style="font-size:12px;color:var(--gray-500);font-weight:400">(WorkRamp — as of {e_date})</span></div>
  <table>
    <thead><tr><th>Rep</th><th>Completion</th><th>Avg Grade</th><th>Failed</th><th>Overdue</th><th>Needs Grading</th></tr></thead>
    <tbody>
""")
        for rep in enablement.get("reps", []):
            comp = rep["completion"]
            grade = rep["avg_grade"]
            failed = rep.get("failed", 0)
            overdue = rep.get("overdue", 0)
            needs = rep.get("needs_grading", 0)
            comp_color = pct_color(comp, 85, 75)
            grade_color = pct_color(grade, 85, 75)
            fail_color = "var(--red)" if failed > 0 else "var(--gray-600)"
            over_color = "var(--red)" if overdue > 0 else "var(--gray-600)"
            html.append(f'<tr><td><strong>{rep["name"]}</strong></td><td style="color:{comp_color};font-weight:600">{comp}%</td><td>{grade}%</td><td style="color:{fail_color}">{failed}</td><td style="color:{over_color}">{overdue}</td><td>{needs}</td></tr>\n')

        team_comp_color = "var(--green)" if mbo_met else "var(--red)"
        mbo_icon = BADGE_ON_TRACK if mbo_met else BADGE_AT_RISK
        html.append(f'<tr style="background:var(--gray-100);font-weight:700"><td>TEAM</td><td style="color:{team_comp_color}">{team_comp}%</td><td>{team_grade}%</td><td></td><td></td><td></td></tr>\n')
        html.append("    </tbody></table>\n")
        html.append(f'  <p style="margin-top:8px;font-size:12px;color:var(--gray-500)">{mbo_icon} MBO threshold: {mbo_threshold}% team completion. Currently at {team_comp}% — {"meeting target" if mbo_met else "below target"}.</p>\n')
        html.append("</div>\n")

        # ---- Enablement Impact: training completion vs ARR attainment ----
        impact_rows = []
        for rep in enablement.get("reps", []):
            rep_name = rep["name"]
            metric_row = next((m for m in metrics if m["Rep Name"] == rep_name), None)
            if not metric_row:
                continue
            comp = rep["completion"]
            grade = rep["avg_grade"]
            arr_pct_r = parse_pct(metric_row.get("ARR % to Goal (Xactly)", "0"))
            rep_arr = parse_money(metric_row.get("Total Booked Saas ARR", "0"))
            dw = parse_pct(metric_row.get("Demo:Win", "0"))
            impact_rows.append({"name": rep_name, "comp": comp, "grade": grade, "arr_pct": arr_pct_r, "arr": rep_arr, "dw": dw})
        if impact_rows:
            html.append("""
<div class="section">
  <div class="section-title"><img src="assets/analytics-check.svg" alt="" style="width:18px;height:18px;vertical-align:middle;opacity:0.7;"> Enablement Impact — Training vs Performance</div>
  <p style="font-size:13px;color:var(--gray-500);margin-bottom:14px">Side-by-side view of training completion and performance. Look for reps who are strong on one but weak on the other.</p>
  <table>
    <thead><tr><th>Rep</th><th>Completion</th><th>Grade</th><th>ARR % Goal</th><th>Demo:Win</th><th>Pattern</th></tr></thead>
    <tbody>
""")
            for ir in impact_rows:
                comp_color = pct_color(ir["comp"], 85, 75)
                arr_color = pct_color(ir["arr_pct"], expected_pct, expected_pct * 0.8)
                training_strong = ir["comp"] >= 85
                performing = ir["arr_pct"] >= expected_pct * 0.8
                if training_strong and performing:
                    pattern = '<span class="badge badge-green">Aligned — performing</span>'
                elif training_strong and not performing:
                    pattern = '<span class="badge badge-orange">Trained, not hitting</span>'
                elif not training_strong and performing:
                    pattern = '<span class="badge badge-blue">Hitting despite gap</span>'
                else:
                    pattern = '<span class="badge badge-red">Both need attention</span>'
                html.append(f'<tr><td><strong>{rep_link(ir["name"])}</strong></td><td style="color:{comp_color};font-weight:600">{ir["comp"]}%</td><td>{ir["grade"]}%</td><td style="color:{arr_color};font-weight:600">{fmt_pct(ir["arr_pct"])}</td><td>{fmt_pct(ir["dw"])}</td><td>{pattern}</td></tr>\n')
            html.append("    </tbody></table>\n")
            html.append('  <p style="margin-top:8px;font-size:11px;color:var(--gray-400)">Small sample (n=7): use this as a sanity check, not a statistical analysis. If someone is "Trained, not hitting" — the coaching conversation is about application, not skills gap.</p>\n')
            html.append("</div>\n")

    # ---- Pipeline Health ----
    html.append("""
<div class="section">
  <div class="section-title"><img src="assets/percentage.svg" alt="" style="width:18px;height:18px;vertical-align:middle;opacity:0.7;"> Pipeline Health</div>
  <table>
    <thead><tr><th>Rep</th><th>Open Deals</th><th>Pipeline Value</th><th>Stale (&gt;30d)</th><th>Health</th></tr></thead>
    <tbody>
""")
    # rep_pipeline is computed earlier (near KPI totals) and reused here
    team_deals = 0
    team_pipe_val = 0
    team_stale = 0
    for name in reps:
        p = rep_pipeline.get(name, {"deals": 0, "value": 0.0, "stale": 0})
        stale_color = "var(--red)" if p["stale"] > 0 else "var(--gray-600)"
        health = pipeline_health(p["stale"], p["deals"])
        team_deals += p["deals"]
        team_pipe_val += p["value"]
        team_stale += p["stale"]
        html.append(f'<tr><td><strong>{name}</strong></td><td>{p["deals"]}</td><td>{fmt_money(p["value"])}</td><td style="color:{stale_color}">{p["stale"]}</td><td>{health}</td></tr>\n')
    html.append(f'<tr style="background:var(--gray-100);font-weight:700"><td>TEAM</td><td>{team_deals}</td><td>{fmt_money(team_pipe_val)}</td><td>{team_stale}</td><td></td></tr>\n')
    html.append("    </tbody></table>\n</div>\n")

    # ---- Pipeline Risk Scoring ----
    scored_deals = []
    for o in team_opps:
        owner = normalize_name(o.get("Opportunity Owner", ""))
        opp_name = o.get("Opportunity Name", "")
        acct = o.get("Account Name", "")
        stage = o.get("Stage", "")
        days = float(o.get("Stage Duration", "0") or 0)
        opp_arr = parse_money(o.get("Software (Annual)", "0"))
        next_step = (o.get("Next Step", "") or "").strip()
        risk_score = 0
        risk_flags = []
        if days > 60:
            risk_score += 40
            risk_flags.append("60+ days stale")
        elif days > 30:
            risk_score += 25
            risk_flags.append("30+ days stale")
        elif days > 14:
            risk_score += 10
            risk_flags.append("aging")
        if not next_step or next_step.lower() in ("", "n/a", "none", "-"):
            risk_score += 20
            risk_flags.append("no next step")
        if opp_arr < 500:
            risk_score += 10
            risk_flags.append("low value")
        late_stages = ["Negotiation", "Closed", "Contract"]
        early_stages = ["Prospecting", "Qualification", "Discovery"]
        if any(s.lower() in stage.lower() for s in early_stages) and days > 21:
            risk_score += 15
            risk_flags.append("early stage + aging")
        scored_deals.append({
            "owner": owner, "name": opp_name[:45], "acct": acct, "stage": stage,
            "days": int(days), "arr": opp_arr, "next_step": next_step[:60],
            "risk_score": risk_score, "flags": risk_flags
        })
    scored_deals.sort(key=lambda x: x["risk_score"], reverse=True)

    html.append("""
<div class="section">
  <div class="section-title"><img src="assets/percentage.svg" alt="" style="width:18px;height:18px;vertical-align:middle;opacity:0.7;"> Pipeline Risk — Prioritized Action List</div>
  <p style="font-size:13px;color:var(--gray-500);margin-bottom:14px">Deals scored by risk (stage duration, missing next steps, value). Focus pipeline reviews on the top of this list.</p>
  <table>
    <thead><tr><th>Risk</th><th>Rep</th><th>Account</th><th>Stage</th><th>Days</th><th>ARR</th><th>Flags</th><th>Next Step</th></tr></thead>
    <tbody>
""")
    for d in scored_deals[:15]:
        if d["risk_score"] >= 30:
            risk_badge = f'<span class="badge badge-red">{d["risk_score"]}</span>'
        elif d["risk_score"] >= 15:
            risk_badge = f'<span class="badge badge-yellow">{d["risk_score"]}</span>'
        else:
            risk_badge = f'<span class="badge badge-green">{d["risk_score"]}</span>'
        days_color = "var(--red)" if d["days"] > 30 else ("var(--yellow)" if d["days"] > 14 else "var(--gray-600)")
        flags_str = ", ".join(d["flags"]) if d["flags"] else "—"
        ns_style = 'style="font-size:11px;color:var(--gray-500);max-width:180px;white-space:normal"'
        ns_text = d["next_step"] if d["next_step"] else '<span style="color:var(--red);font-weight:600">MISSING</span>'
        html.append(f'<tr><td>{risk_badge}</td><td><strong>{d["owner"]}</strong></td><td>{d["acct"]}</td><td style="font-size:12px">{d["stage"]}</td><td style="color:{days_color};font-weight:600">{d["days"]}d</td><td>{fmt_money(d["arr"])}</td><td style="font-size:11px;color:var(--gray-500)">{flags_str}</td><td {ns_style}>{ns_text}</td></tr>\n')
    html.append("    </tbody></table>\n")
    if len(scored_deals) > 15:
        html.append(f'  <p style="margin-top:8px;font-size:12px;color:var(--gray-400)">Showing top 15 of {len(scored_deals)} open deals by risk score.</p>\n')
    html.append("</div>\n")

    # ---- Stale Deal Aging Heatmap ----
    # Bucket each rep's open deals by age: 0-14, 15-30, 31-60, 60+
    heatmap = {}
    for r_opp in team_opps:
        owner = normalize_name(r_opp.get("Opportunity Owner", ""))
        if owner not in reps:
            continue
        days = float(r_opp.get("Stage Duration", "0") or 0)
        if owner not in heatmap:
            heatmap[owner] = {"0-14": 0, "15-30": 0, "31-60": 0, "60+": 0, "total": 0}
        if days <= 14:
            heatmap[owner]["0-14"] += 1
        elif days <= 30:
            heatmap[owner]["15-30"] += 1
        elif days <= 60:
            heatmap[owner]["31-60"] += 1
        else:
            heatmap[owner]["60+"] += 1
        heatmap[owner]["total"] += 1

    def heat_cell(count, bucket):
        if count == 0:
            return '<td style="text-align:center;color:var(--gray-300)">0</td>'
        if bucket == "0-14":
            bg, fg = "var(--green-bg)", "var(--green)"
        elif bucket == "15-30":
            bg, fg = "var(--yellow-bg)", "var(--yellow)"
        elif bucket == "31-60":
            bg, fg = "var(--toast-orange-bg)", "var(--toast-orange)"
        else:
            bg, fg = "var(--red-bg)", "var(--red)"
        intensity = min(count / 5, 1.0)  # saturate at 5+
        opacity = 0.3 + intensity * 0.7
        return f'<td style="text-align:center;background:{bg};color:{fg};font-weight:700;opacity:{opacity:.2f}">{count}</td>'

    html.append("""
<div class="section">
  <div class="section-title"><img src="assets/percentage.svg" alt="" style="width:18px;height:18px;vertical-align:middle;opacity:0.7;"> Stale Deal Aging Heatmap</div>
  <p style="font-size:13px;color:var(--gray-500);margin-bottom:14px">Distribution of each rep's open deals by age. Red columns = deals that have been sitting too long.</p>
  <table>
    <thead><tr><th>Rep</th><th style="text-align:center">0-14 days<br><span style="font-weight:400;font-size:10px;color:var(--green)">Fresh</span></th><th style="text-align:center">15-30 days<br><span style="font-weight:400;font-size:10px;color:var(--yellow)">Aging</span></th><th style="text-align:center">31-60 days<br><span style="font-weight:400;font-size:10px;color:var(--toast-orange)">Stale</span></th><th style="text-align:center">60+ days<br><span style="font-weight:400;font-size:10px;color:var(--red)">Critical</span></th><th style="text-align:center">Total</th></tr></thead>
    <tbody>
""")
    team_buckets = {"0-14": 0, "15-30": 0, "31-60": 0, "60+": 0, "total": 0}
    for name in reps:
        h = heatmap.get(name, {"0-14": 0, "15-30": 0, "31-60": 0, "60+": 0, "total": 0})
        html.append(f'<tr><td><strong>{rep_link(name)}</strong></td>')
        for b in ["0-14", "15-30", "31-60", "60+"]:
            html.append(heat_cell(h[b], b))
            team_buckets[b] += h[b]
        team_buckets["total"] += h["total"]
        html.append(f'<td style="text-align:center;font-weight:600">{h["total"]}</td></tr>\n')
    html.append('<tr style="background:var(--gray-100);font-weight:700"><td>TEAM</td>')
    for b in ["0-14", "15-30", "31-60", "60+"]:
        html.append(f'<td style="text-align:center">{team_buckets[b]}</td>')
    html.append(f'<td style="text-align:center">{team_buckets["total"]}</td></tr>\n')
    html.append("    </tbody></table>\n</div>\n")

    # ---- Calls & Emails Activity (two-col) ----
    html.append("""
<div class="section">
  <div class="two-col">
    <div>
      <div class="section-title"><img src="assets/phone.svg" alt="" style="width:18px;height:18px;vertical-align:middle;opacity:0.7;"> Calls Activity</div>
      <table>
        <thead><tr><th>Rep</th><th>Calls</th><th>Avg/Day</th><th>Conversations</th><th>Conv Rate</th><th>Avg Duration</th></tr></thead>
        <tbody>
""")
    for name in reps:
        c = calls_by_name.get(name, {})
        total = c.get("calls", "0")
        avg_day = c.get("average_calls_logged_per_day", "0")
        convos = c.get("conversations", "0")
        conv_rate = c.get("conversations_rate", "0")
        avg_dur = c.get("average_calls_duration", "N/A")
        html.append(f'<tr><td><strong>{name}</strong></td><td>{total}</td><td>{avg_day}</td><td>{convos}</td><td>{conv_rate}%</td><td>{avg_dur}</td></tr>\n')
    html.append("        </tbody></table>\n    </div>\n")

    # Emails
    html.append("""    <div>
      <div class="section-title"><img src="assets/email.svg" alt="" style="width:18px;height:18px;vertical-align:middle;opacity:0.7;"> Email Activity</div>
      <table>
        <thead><tr><th>Rep</th><th>Sent</th><th>Opened</th><th>Open Rate</th><th>Replied</th><th>Reply Rate</th></tr></thead>
        <tbody>
""")
    for name in reps:
        e = emails_by_name.get(name, {})
        sent = e.get("sent", "0")
        opened = e.get("opened", "0")
        open_rate = e.get("opened_rate", "0")
        replied = e.get("replied", "0")
        reply_rate = e.get("replied_rate", "0")
        html.append(f'<tr><td><strong>{name}</strong></td><td>{sent}</td><td>{opened}</td><td>{open_rate}%</td><td>{replied}</td><td>{reply_rate}%</td></tr>\n')
    html.append("        </tbody></table>\n    </div>\n  </div>\n</div>\n")

    # ---- Activity-to-Outcome Correlations ----
    html.append("""
<div class="section">
  <div class="section-title"><img src="assets/analytics-check.svg" alt="" style="width:18px;height:18px;vertical-align:middle;opacity:0.7;"> Activity → Outcome Efficiency</div>
  <p style="font-size:13px;color:var(--gray-500);margin-bottom:14px">High activity + low conversion = skill gap. Low activity + high conversion = capacity gap. This table helps diagnose where to coach.</p>
  <table>
    <thead><tr><th>Rep</th><th>Calls</th><th>Conv %</th><th>Emails</th><th>Reply %</th><th>Opps</th><th>Demo:Win</th><th>ARR/Win</th><th>ARR</th><th>Diagnosis</th></tr></thead>
    <tbody>
""")
    for r in metrics:
        name = r["Rep Name"]
        c = calls_by_name.get(name, {})
        e = emails_by_name.get(name, {})
        call_total = int(c.get("calls", "0") or 0)
        conv_rate = float(c.get("conversations_rate", "0") or 0)
        email_total = int(e.get("sent", "0") or 0)
        reply_rate = float(e.get("replied_rate", "0") or 0)
        rep_opps = int(r.get("Opps", "0") or 0)
        demo_win = parse_pct(r.get("Demo:Win", "0"))
        arr = parse_money(r.get("Total Booked Saas ARR", "0"))
        avg_per_win = parse_money(r.get("ARR Won per Opp", "0"))
        pct = parse_pct(r.get("ARR % to Goal (Xactly)", "0"))
        call_expected = round(int(r.get("Call Goal", "0") or 0) * biz_days_passed / total_biz_days) if total_biz_days else 0
        demo_expected = round(int(r.get("Total Demo Goal", "0") or 0) * biz_days_passed / total_biz_days) if total_biz_days else 0
        activity_ok = call_total >= call_expected * 0.8 if call_expected else True
        conversion_ok = demo_win >= 50
        if activity_ok and conversion_ok:
            diagnosis = '<span class="badge badge-green">Performing</span>'
        elif activity_ok and not conversion_ok:
            diagnosis = '<span class="badge badge-orange">Skill Gap</span>'
        elif not activity_ok and conversion_ok:
            diagnosis = '<span class="badge badge-blue">Capacity Gap</span>'
        else:
            diagnosis = '<span class="badge badge-red">Both</span>'
        conv_color = pct_color(conv_rate, 15, 8)
        reply_color = pct_color(reply_rate, 15, 8)
        dw_color = pct_color(demo_win, 60, 40)
        html.append(f'<tr><td><strong>{rep_link(name)}</strong></td><td>{call_total}</td><td style="color:{conv_color};font-weight:600">{conv_rate}%</td><td>{email_total}</td><td style="color:{reply_color};font-weight:600">{reply_rate}%</td><td>{rep_opps}</td><td style="color:{dw_color};font-weight:600">{fmt_pct(demo_win)}</td><td>{fmt_money(avg_per_win)}</td><td style="font-weight:600">{fmt_money(arr)}</td><td>{diagnosis}</td></tr>\n')
    html.append("    </tbody></table>\n</div>\n")

    # ---- Recent Wins ----
    html.append("""
<div class="section">
  <div class="section-title"><img src="assets/dollar.svg" alt="" style="width:18px;height:18px;vertical-align:middle;opacity:0.7;"> Recent Wins</div>
  <table>
    <thead><tr><th>Account</th><th>Type</th><th>ARR</th><th>Close Date</th><th>Rep</th></tr></thead>
    <tbody>
""")
    # Sort by close date descending, show top 20
    sorted_wins = sorted(wins, key=lambda r: r.get("Close Date", ""), reverse=True)[:20]
    for w in sorted_wins:
        acct = w.get("Account Name", "")
        wtype = w.get("Type", "")
        arr = parse_money(w.get("Software (Annual)", "0"))
        close = w.get("Close Date", "")
        owner = normalize_name(w.get("Opportunity Owner", ""))
        html.append(f'<tr><td><strong>{acct}</strong></td><td style="font-size:12px">{wtype}</td><td>{fmt_money(arr)}</td><td>{close}</td><td>{owner}</td></tr>\n')
    html.append("    </tbody></table>\n</div>\n")

    # ---- Forecast vs Quota bars ----
    pace_pct = expected_pct
    html.append("""
<div class="section">
  <div class="section-title"><img src="assets/trending-up.svg" alt="" style="width:18px;height:18px;vertical-align:middle;opacity:0.7;"> Forecast vs Quota — Pacing</div>
""")
    for r in metrics:
        name = r["Rep Name"]
        quota = parse_money(r.get("Booked SaaS Quota (Xactly)", "0"))
        pct = parse_pct(r.get("ARR % to Goal (Xactly)", "0"))
        bar_w = min(pct, 100)
        color = bar_color(pct, expected_pct)
        html.append(f'<div class="forecast-bar"><div class="forecast-bar-name">{name}</div><div class="forecast-bar-track"><div class="forecast-bar-fill" style="width:{bar_w:.1f}%;background:{color}"></div><div class="pace-line" style="left:{pace_pct:.1f}%"></div></div><div class="forecast-bar-label">{fmt_money(quota)}</div><div class="forecast-bar-label">{fmt_pct(pct)}</div></div>\n')

    html.append(f'<div class="forecast-bar" style="margin-top:12px;padding-top:12px;border-top:2px solid var(--gray-300)"><div class="forecast-bar-name" style="font-weight:800">TEAM TOTAL</div><div class="forecast-bar-track"><div class="forecast-bar-fill" style="width:{min(team_attainment, 100):.1f}%;background:var(--toast-orange)"></div><div class="pace-line" style="left:{pace_pct:.1f}%"></div></div><div class="forecast-bar-label">{fmt_money(team_quota)}</div><div class="forecast-bar-label">{fmt_pct(team_attainment)}</div></div>\n')
    html.append(f'  <p style="margin-top:12px;font-size:13px;color:var(--gray-500)">&#x25BC; = Expected pace (Working Day {biz_days_passed}/{total_biz_days})</p>\n</div>\n')

    # Footer
    generated = datetime.now().strftime("%b %d, %Y")
    html.append(f'\n<div class="footer">Growth Sales Team Dashboard — Built for Josh Mellender | Data: Sigma + Xactly + SalesLoft | Generated {generated}</div>\n</body></html>\n')

    # Write main dashboard
    out_path = Path(__file__).parent / "index.html"
    out_path.write_text("".join(html))
    print(f"Dashboard generated: {out_path}")
    print(f"  Month: {month_name} | Data date: {data_date}")
    print(f"  Team ARR: {fmt_money(team_arr)} / {fmt_money(team_quota)} ({fmt_pct(team_attainment)})")
    print(f"  Wins: {team_wins_count} | Pipeline: {fmt_money(total_pipeline)} ({total_open_deals} opps)")
    if ue_deals:
        ue_total = sum(d["arr"] for d in ue_deals)
        print(f"  UE Promos injected: {fmt_money(ue_total)} across {len(ue_deals)} deals ({ue_csv_path.name})")
        for rep_name, rep_total in sorted(ue_by_rep.items(), key=lambda x: -x[1]):
            print(f"    - {rep_name}: {fmt_money(rep_total)}")
    if roe_deals:
        roe_total = sum(d["arr"] for d in roe_deals)
        print(f"  ROE clawbacks: -{fmt_money(roe_total)} across {len(roe_deals)} deals")
        for d in roe_deals:
            print(f"    - {d['rep']}: -{fmt_money(d['arr'])} ({d['account']}) — {d['reason']}")

    # ========== SAVE HISTORICAL SNAPSHOT ==========
    history_dir = Path(__file__).parent / "data" / "history"
    history_dir.mkdir(parents=True, exist_ok=True)
    snapshot = {
        "date": f"{now.year}-{month_num}-{day_num.zfill(2)}",
        "team_arr": team_arr,
        "team_quota": team_quota,
        "team_wins": team_wins_count,
        "team_pipeline": total_pipeline,
        "team_open_deals": total_open_deals,
        "reps": {}
    }
    for r in metrics:
        name = r["Rep Name"]
        snapshot["reps"][name] = {
            "arr": parse_money(r.get("Total Booked Saas ARR", "0")),
            "quota": parse_money(r.get("Booked SaaS Quota (Xactly)", "0")),
            "arr_pct": parse_pct(r.get("ARR % to Goal (Xactly)", "0")),
            "wins": int(r.get("Wins", "0") or 0),
            "demos": int(r.get("Demos", "0") or 0),
            "opps": int(r.get("Opps", "0") or 0),
            "opp_demo": parse_pct(r.get("Opp:Demo", "0")),
            "opp_win": parse_pct(r.get("Opp to Win", "0")),
            "demo_win": parse_pct(r.get("Demo:Win", "0")),
            "avg_arr_per_opp": parse_money(r.get("ARR Won per Opp", "0")),
            "ss_opps": int(r.get("Self-Sourced Created Opps", "0") or 0),
            "nbr_created": int(r.get("NB Referral Created Opps", "0") or 0),
            "nbr_won": int(r.get("NB Referral Closed Won", "0") or 0),
            "ec_units": int(r.get("EC Units", "0") or 0),
            "ec_arr": parse_money(r.get("EC ARR", "0")),
            "calls": int(calls_by_name.get(name, {}).get("calls", "0") or 0),
            "conversations": int(calls_by_name.get(name, {}).get("conversations", "0") or 0),
            "emails_sent": int(emails_by_name.get(name, {}).get("sent", "0") or 0),
            "emails_replied": int(emails_by_name.get(name, {}).get("replied", "0") or 0),
        }
    snap_path = history_dir / f"{now.year}-{month_num}-{day_num.zfill(2)}.json"
    snap_path.write_text(json.dumps(snapshot, indent=2))
    print(f"  Snapshot saved: {snap_path.name}")

    # ========== LOAD HISTORICAL DATA FOR TRENDS ==========
    history_files = sorted(history_dir.glob("*.json"))
    all_history = []
    for hf in history_files:
        with open(hf) as f:
            all_history.append(json.load(f))

    # Get previous week's data (7+ days ago) for comparison
    prev_snap = None
    if len(all_history) >= 2:
        prev_snap = all_history[-2]  # most recent before today

    # ========== GENERATE REP DETAIL PAGES ==========
    reps_dir = Path(__file__).parent / "reps"
    reps_dir.mkdir(parents=True, exist_ok=True)

    for r in metrics:
        name = r["Rep Name"]
        slug = name.lower().replace(" ", "-")
        rep_data = snapshot["reps"][name]
        prev_rep = prev_snap["reps"].get(name, {}) if prev_snap else {}

        # Get rep's open opps (stale = >14 days)
        rep_opps_all = [o for o in opps if normalize_name(o.get("Opportunity Owner", "")) == name]
        rep_stale = [o for o in rep_opps_all if float(o.get("Stage Duration", "0") or 0) > 14]
        rep_stale.sort(key=lambda o: float(o.get("Stage Duration", "0") or 0), reverse=True)

        # Get rep's wins
        rep_wins = [w for w in wins if normalize_name(w.get("Opportunity Owner", "")) == name]
        rep_wins.sort(key=lambda w: w.get("Close Date", ""), reverse=True)

        # Calls/emails data
        c = calls_by_name.get(name, {})
        e = emails_by_name.get(name, {})

        # Compute deltas vs previous snapshot
        def delta(key, fmt="num"):
            cur = rep_data.get(key, 0)
            prev = prev_rep.get(key, 0)
            diff = cur - prev
            if diff == 0 or not prev_rep:
                return ""
            sign = "+" if diff > 0 else ""
            if fmt == "money":
                return f' <span style="font-size:11px;color:{"var(--green)" if diff > 0 else "var(--red)"}">{sign}{fmt_money(diff)}</span>'
            elif fmt == "pct":
                return f' <span style="font-size:11px;color:{"var(--green)" if diff > 0 else "var(--red)"}">{sign}{diff:.1f}pp</span>'
            else:
                return f' <span style="font-size:11px;color:{"var(--green)" if diff > 0 else "var(--red)"}">{sign}{diff}</span>'

        # ---- Sandler Coach to Success Analysis ----
        role = r.get("Role in Month", "AE")
        cohort = r.get("Ramping Cohort", "")
        quota = rep_data["quota"]
        arr_pct = rep_data["arr_pct"]

        # Behavior analysis (activity) — pace by working day
        # ~22 working days/month, figure out how many working days have passed
        pace_factor = biz_days_passed / total_biz_days if total_biz_days else 1

        behavior_items = []
        call_goal_mo = int(r.get("Call Goal", "0") or 0)
        call_goal_daily = round(call_goal_mo / total_biz_days) if total_biz_days else 0
        call_goal_weekly = call_goal_daily * 5
        call_expected = round(call_goal_mo * pace_factor)
        call_actual = rep_data["calls"]
        call_pace_pct = (call_actual / call_expected * 100) if call_expected else 0
        if call_expected > 0 and call_pace_pct < 70:
            behavior_items.append(f"Calls at {call_actual}/{call_expected} expected by day {biz_days_passed} ({call_pace_pct:.0f}% to pace) — significantly behind. Goal: {call_goal_daily}/day, {call_goal_weekly}/week. Are there time management barriers?")
        elif call_expected > 0 and call_pace_pct < 90:
            behavior_items.append(f"Calls at {call_actual}/{call_expected} expected by day {biz_days_passed} ({call_pace_pct:.0f}% to pace) — slightly behind. Goal: {call_goal_daily}/day, {call_goal_weekly}/week.")
        elif call_expected > 0:
            behavior_items.append(f"Calls at {call_actual}/{call_expected} expected ({call_pace_pct:.0f}% to pace) — on track. Goal: {call_goal_daily}/day, {call_goal_weekly}/week.")

        demo_goal_mo = int(r.get("Total Demo Goal", "0") or 0)
        demo_goal_weekly = round(demo_goal_mo / 4.3) if demo_goal_mo else 0
        demo_expected = round(demo_goal_mo * pace_factor)
        demo_actual = rep_data["demos"]
        demo_pace_pct = (demo_actual / demo_expected * 100) if demo_expected else 0
        if demo_expected > 0 and demo_pace_pct < 70:
            behavior_items.append(f"Demos at {demo_actual}/{demo_expected} expected by day {biz_days_passed} ({demo_pace_pct:.0f}% to pace) — not getting enough at-bats. Goal: ~{demo_goal_weekly}/week. Focus on booking more demos.")
        elif demo_expected > 0 and demo_pace_pct < 90:
            behavior_items.append(f"Demos at {demo_actual}/{demo_expected} expected by day {biz_days_passed} ({demo_pace_pct:.0f}% to pace) — close. Goal: ~{demo_goal_weekly}/week.")
        elif demo_expected > 0:
            behavior_items.append(f"Demos at {demo_actual}/{demo_expected} expected ({demo_pace_pct:.0f}% to pace) — on track. Goal: ~{demo_goal_weekly}/week.")

        emails_sent = rep_data["emails_sent"]
        if emails_sent < 20:
            behavior_items.append(f"Only {emails_sent} emails sent this week — low outbound volume.")

        if not behavior_items:
            behavior_items.append("Activity levels are strong. Maintain current cadence.")

        # Technique analysis (conversion rates)
        technique_items = []
        opp_demo = rep_data["opp_demo"]
        if opp_demo > 0 and opp_demo < 70:
            technique_items.append(f"Opp:Demo conversion at {opp_demo:.0f}% — opps not converting to demos. Qualify harder upfront or improve scheduling follow-up.")
        demo_win = rep_data["demo_win"]
        if demo_win > 0 and demo_win < 50:
            technique_items.append(f"Demo:Win at {demo_win:.0f}% — low close rate from demos. Work on post-demo follow-up, urgency, and handling stalls.")
        elif demo_win > 0 and demo_win < 70:
            technique_items.append(f"Demo:Win at {demo_win:.0f}% — room to improve close rate. Focus on commitment at end of demo.")

        avg_arr = rep_data["avg_arr_per_opp"]
        if avg_arr > 0 and avg_arr < 1500:
            technique_items.append(f"Avg ARR per opp ${avg_arr:,.0f} (goal: $1,800) — selling smaller deals. Coach on attaching more products/bundling.")
        elif avg_arr > 0 and avg_arr < 1800:
            technique_items.append(f"Avg ARR per opp ${avg_arr:,.0f} — close to $1,800 target. Look for upsell opportunities.")

        if not technique_items:
            technique_items.append("Conversion rates and deal size are solid. Focus on maintaining consistency.")

        # Results analysis — compare against working day pace, not raw monthly %
        results_items = []
        is_ramping = quota == 0 or (cohort and cohort.startswith("1-3"))
        pace_ratio = arr_pct / expected_pct if expected_pct and not is_ramping else 0

        if is_ramping:
            # Manual goals for ramping reps (not in Sigma/Xactly)
            ramp_goals = {
                "Thomas Jensen": {"demos": 30, "ec_units": 3},
            }
            ramp = ramp_goals.get(name, {})
            results_items.append(f"Ramping rep ({cohort if cohort and cohort != '-' else 'new'}). {fmt_money(rep_data['arr'])} booked with {rep_data['wins']} wins so far. No ARR quota yet — focus on hitting activity and EC goals.")

            # Demo pacing against manual goal
            ramp_demo_goal = ramp.get("demos", int(r.get("Total Demo Goal", "0") or 0))
            rep_demos = rep_data["demos"]
            if ramp_demo_goal:
                _, demo_p, demo_d, demo_wk = pace_actual_vs_expected(rep_demos, ramp_demo_goal, biz_days_passed, total_biz_days)
                exp_demos = round(ramp_demo_goal * biz_days_passed / total_biz_days) if total_biz_days else 0
                if demo_p >= 100:
                    results_items.append(f"Demos: {rep_demos}/{exp_demos} expected by day {biz_days_passed} (goal: {ramp_demo_goal}/mo, ~{demo_wk}/wk) — on track.")
                else:
                    results_items.append(f"Demos: {rep_demos}/{exp_demos} expected by day {biz_days_passed} (goal: {ramp_demo_goal}/mo, ~{demo_wk}/wk) — behind pace.")

            # EC units pacing against manual goal
            ramp_ec_goal = ramp.get("ec_units", 0)
            rep_ec_units = rep_data["ec_units"]
            if ramp_ec_goal:
                exp_ec = round(ramp_ec_goal * biz_days_passed / total_biz_days, 1) if total_biz_days else 0
                if rep_ec_units >= ramp_ec_goal:
                    results_items.append(f"EC Units: {rep_ec_units}/{ramp_ec_goal} — goal met.")
                elif rep_ec_units >= exp_ec:
                    results_items.append(f"EC Units: {rep_ec_units}/{ramp_ec_goal} goal — on pace.")
                else:
                    results_items.append(f"EC Units: {rep_ec_units}/{ramp_ec_goal} goal — needs {ramp_ec_goal - rep_ec_units} more this month.")
        elif arr_pct >= 100:
            results_items.append(f"At {arr_pct:.0f}% to quota — already hit goal. What's working that we can share with the team?")
        elif pace_ratio >= 1.2:
            results_items.append(f"At {arr_pct:.0f}% to quota ({pace_ratio:.1f}x pace on working day {biz_days_passed}/{total_biz_days}) — well ahead of pace. Keep momentum.")
        elif pace_ratio >= 1.0:
            results_items.append(f"At {arr_pct:.0f}% to quota (on pace for working day {biz_days_passed}/{total_biz_days}) — tracking well. What deals can we push to close this month?")
        elif pace_ratio >= 0.8:
            results_items.append(f"At {arr_pct:.0f}% to quota ({pace_ratio:.1f}x pace on working day {biz_days_passed}/{total_biz_days}) — slightly behind pace. Review pipeline and prioritize highest-probability deals.")
        else:
            results_items.append(f"At {arr_pct:.0f}% to quota ({pace_ratio:.1f}x pace on working day {biz_days_passed}/{total_biz_days}) — behind pace. Need to diagnose: is it activity, pipeline, or conversion?")

        stale_count_rep = len(rep_stale)
        if stale_count_rep > 5:
            results_items.append(f"{stale_count_rep} deals stale >14 days — pipeline needs cleaning. Which should be closed-lost vs. re-engaged?")
        elif stale_count_rep > 0:
            results_items.append(f"{stale_count_rep} deals stale >14 days — review next steps on each.")

        if not is_ramping and arr_pct < 100 and quota > 0 and remaining_biz_days > 0:
            rep_gap = quota - rep_data["arr"]
            rep_avg = rep_data["avg_arr_per_opp"] if rep_data["avg_arr_per_opp"] > 0 else (team_arr / team_wins_count if team_wins_count else 1800)
            rep_dw = rep_data["demo_win"] / 100 if rep_data["demo_win"] > 0 else 0.5
            rep_deals_need = rep_gap / rep_avg
            rep_demos_need = rep_deals_need / rep_dw
            rep_demos_wk = rep_demos_need / remaining_weeks if remaining_weeks > 0 else 999
            results_items.append(f"<strong>Close the Gap:</strong> Need {fmt_money(rep_gap)} more → {rep_deals_need:.1f} deals (@ {fmt_money(rep_avg)} avg) → {rep_demos_need:.0f} demos (@ {rep_dw*100:.0f}% D:W) → <strong>{rep_demos_wk:.1f} demos/week</strong> for remaining {remaining_biz_days} working days.")

        # Build rep page HTML
        rp = []
        rp.append(f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{name} — Growth Sales Dashboard</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Source+Sans+3:wght@400;600;700;800&display=swap" rel="stylesheet">
<style>
  :root {{
    --toast-orange: #FF4C00; --toast-orange-bg: #FFF3ED;
    --toast-navy: #2B4FB9;
    --toast-dark: #252525;
    --green: #22C55E; --green-bg: #F0FDF4;
    --yellow: #EAB308; --yellow-bg: #FEFCE8;
    --red: #EF4444; --red-bg: #FEF2F2;
    --blue: #3B82F6; --blue-bg: #EFF6FF;
    --warm-50: #F9F5F3; --warm-100: #F6F1EE;
    --gray-50: #F7FAFC; --gray-100: #F3F4F6; --gray-200: #E5E7EB;
    --gray-300: #D1D5DB; --gray-400: #9CA3AF; --gray-500: #6B7280;
    --gray-600: #4B5563; --gray-700: #374151; --gray-800: #252525;
    --shadow: 0 1px 3px rgba(0,0,0,0.06); --radius: 8px;
  }}
  * {{ margin: 0; padding: 0; box-sizing: border-box; }}
  body {{ font-family: 'Source Sans 3', 'Source Sans Pro', system-ui, sans-serif; background: var(--warm-50); color: var(--gray-800); line-height: 1.5; font-size: 16px; }}
  .header {{ background: white; color: var(--toast-dark); padding: 16px 32px; display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid var(--gray-200); }}
  .header-left {{ display: flex; align-items: center; gap: 16px; }}
  .header-left a {{ color: var(--toast-orange); text-decoration: none; font-size: 14px; font-weight: 600; }}
  .header-left a:hover {{ color: var(--toast-dark); }}
  .header-left h1 {{ font-size: 20px; font-weight: 500; color: var(--toast-dark); }}
  .header-right {{ font-size: 14px; color: var(--gray-500); }}
  .content {{ max-width: 1200px; margin: 0 auto; padding: 24px 32px; }}
  .scorecard {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; margin-bottom: 28px; }}
  .sc-card {{ background: white; border: 1px solid var(--gray-200); border-radius: var(--radius); padding: 20px 16px; text-align: center; }}
  .sc-card .label {{ font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px; color: var(--gray-500); font-weight: 600; }}
  .sc-card .value {{ font-size: 22px; font-weight: 600; margin: 4px 0; color: var(--toast-dark); }}
  .sc-card .sub {{ font-size: 12px; color: var(--gray-500); }}
  .section {{ background: white; border: 1px solid var(--gray-200); border-radius: var(--radius); padding: 20px; margin-bottom: 20px; box-shadow: var(--shadow); }}
  .section h2 {{ font-size: 16px; font-weight: 500; margin-bottom: 14px; color: var(--toast-dark); display: flex; align-items: center; gap: 8px; }}
  .two-col {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }}
  @media (max-width: 900px) {{ .two-col {{ grid-template-columns: 1fr; }} }}
  table {{ width: 100%; border-collapse: collapse; font-size: 14px; }}
  th {{ background: var(--gray-50); color: var(--gray-600); font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px; font-weight: 600; padding: 8px 12px; text-align: left; border-bottom: 2px solid var(--gray-200); }}
  td {{ padding: 8px 12px; border-top: 1px solid var(--gray-100); }}
  tr:hover {{ background: var(--warm-100); }}
  .badge {{ display: inline-block; padding: 3px 10px; border-radius: 999px; font-size: 11px; font-weight: 600; }}
  .badge-green {{ background: var(--green-bg); color: var(--green); }}
  .badge-yellow {{ background: var(--yellow-bg); color: var(--yellow); }}
  .badge-red {{ background: var(--red-bg); color: var(--red); }}
  .coaching {{ list-style: none; }}
  .coaching li {{ padding: 8px 0; border-bottom: 1px solid var(--gray-100); font-size: 14px; line-height: 1.6; }}
  .coaching li:last-child {{ border-bottom: none; }}
  .coaching-label {{ font-weight: 600; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 6px; padding: 4px 10px; border-radius: 4px; display: inline-block; }}
  .c-behavior {{ background: var(--blue-bg); color: var(--blue); }}
  .c-technique {{ background: var(--toast-orange-bg); color: var(--toast-orange); }}
  .c-results {{ background: var(--green-bg); color: var(--green); }}
  .footer {{ text-align: center; padding: 20px; font-size: 12px; color: var(--gray-400); }}
  @media print {{
    body {{ background: white; font-size: 11px; }}
    .header {{ padding: 8px 16px; border-bottom: 2px solid #000; }}
    .header-left a {{ display: none; }}
    .content {{ padding: 8px 16px; }}
    .scorecard {{ grid-template-columns: repeat(4, 1fr); gap: 6px; margin-bottom: 12px; }}
    .sc-card {{ padding: 8px 6px; border: 1px solid #ccc; }}
    .sc-card .label {{ font-size: 9px; }}
    .sc-card .value {{ font-size: 16px; }}
    .sc-card .sub {{ font-size: 9px; }}
    .section {{ padding: 10px; margin-bottom: 8px; box-shadow: none; border: 1px solid #ccc; page-break-inside: avoid; }}
    .section h2 {{ font-size: 13px; margin-bottom: 8px; }}
    .two-col {{ grid-template-columns: 1fr 1fr; gap: 8px; }}
    table {{ font-size: 11px; }}
    th {{ padding: 4px 6px; font-size: 9px; }}
    td {{ padding: 4px 6px; }}
    .coaching li {{ padding: 4px 0; font-size: 11px; line-height: 1.4; }}
    .badge {{ font-size: 9px; padding: 1px 6px; }}
    .footer {{ font-size: 9px; padding: 8px; }}
    @page {{ margin: 0.5in; size: letter; }}
  }}
</style>
</head>
<body>
<div class="header">
  <div class="header-left">
    <a href="../index.html">&larr; Back to Dashboard</a>
    <h1>{name}</h1>
  </div>
  <div class="header-right">{month_name} — Data as of {data_date}</div>
</div>
<div class="content">
""")

        # Rep Scorecard Cards
        rp.append('<div class="scorecard">\n')
        rp.append(f'<div class="sc-card"><div class="label">MTD ARR</div><div class="value">{fmt_money(rep_data["arr"])}</div><div class="sub">of {fmt_money(rep_data["quota"])}{delta("arr", "money")}</div></div>\n')
        rp.append(f'<div class="sc-card"><div class="label">Attainment</div><div class="value">{fmt_pct(rep_data["arr_pct"])}</div><div class="sub">to quota{delta("arr_pct", "pct")}</div></div>\n')
        rp.append(f'<div class="sc-card"><div class="label">Wins</div><div class="value">{rep_data["wins"]}</div><div class="sub">closed deals{delta("wins")}</div></div>\n')
        rp.append(f'<div class="sc-card"><div class="label">Demos</div><div class="value">{rep_data["demos"]}</div><div class="sub">held{delta("demos")}</div></div>\n')
        rp.append(f'<div class="sc-card"><div class="label">Opps</div><div class="value">{rep_data["opps"]}</div><div class="sub">created{delta("opps")}</div></div>\n')
        rp.append(f'<div class="sc-card"><div class="label">Avg ARR</div><div class="value">{fmt_money(rep_data["avg_arr_per_opp"])}</div><div class="sub">per opp{delta("avg_arr_per_opp", "money")}</div></div>\n')
        rp.append(f'<div class="sc-card"><div class="label">Demo:Win</div><div class="value">{fmt_pct(rep_data["demo_win"])}</div><div class="sub">conversion{delta("demo_win", "pct")}</div></div>\n')
        rp.append(f'<div class="sc-card"><div class="label">EC ARR</div><div class="value">{fmt_money(rep_data["ec_arr"])}</div><div class="sub">{rep_data["ec_units"]} units{delta("ec_arr", "money")}</div></div>\n')
        rp.append('</div>\n')

        # Sandler Coaching Section
        rp.append('<div class="section">\n')
        rp.append(f'  <h2><img src="../assets/Other_Award.svg" alt="" style="width:18px;height:18px;vertical-align:middle;opacity:0.7;"> 1:1 Coaching — Sandler Coach to Success</h2>\n')
        rp.append('  <div class="two-col">\n')

        # Left column: Behavior + Technique
        rp.append('    <div>\n')
        rp.append('      <div class="coaching-label c-behavior">Behavior (Activity)</div>\n')
        rp.append('      <ul class="coaching">\n')
        for item in behavior_items:
            rp.append(f'        <li>{item}</li>\n')
        rp.append('      </ul>\n')
        rp.append('      <div class="coaching-label c-technique" style="margin-top:16px">Technique (Conversion)</div>\n')
        rp.append('      <ul class="coaching">\n')
        for item in technique_items:
            rp.append(f'        <li>{item}</li>\n')
        rp.append('      </ul>\n')
        rp.append('    </div>\n')

        # Right column: Results + Suggested Questions
        rp.append('    <div>\n')
        rp.append('      <div class="coaching-label c-results">Results (Outcomes)</div>\n')
        rp.append('      <ul class="coaching">\n')
        for item in results_items:
            rp.append(f'        <li>{item}</li>\n')
        rp.append('      </ul>\n')

        # Sandler coaching questions based on data
        rp.append('      <div class="coaching-label" style="margin-top:16px;background:var(--gray-100);color:var(--gray-700)">Suggested Coaching Questions</div>\n')
        rp.append('      <ul class="coaching">\n')
        if pace_ratio < 0.8:
            rp.append('        <li>"Walk me through your top 3 deals this month — what needs to happen to close each one?"</li>\n')
        if demo_win > 0 and demo_win < 60:
            rp.append('        <li>"After your last demo, what was the prospect\'s commitment? Did you get a clear next step?"</li>\n')
        if stale_count_rep > 3:
            rp.append(f'        <li>"You have {stale_count_rep} deals sitting >14 days. Which ones are real and which should we close out?"</li>\n')
        if call_pace_pct < 80:
            rp.append('        <li>"What does your daily call block look like? Are you protecting that time?"</li>\n')
        if avg_arr > 0 and avg_arr < 1800:
            rp.append('        <li>"On your last few deals, did you explore EC/Payroll and all available product bundles?"</li>\n')
        rp.append('        <li>"What\'s one thing you want to improve this week, and how can I help?"</li>\n')
        rp.append('      </ul>\n')
        rp.append('    </div>\n')
        rp.append('  </div>\n')
        rp.append('</div>\n')

        # Activity Metrics
        rp.append('<div class="two-col">\n')
        rp.append('<div class="section">\n')
        rp.append(f'  <h2><img src="../assets/phone.svg" alt="" style="width:18px;height:18px;vertical-align:middle;opacity:0.7;"> Calls</h2>\n')
        rp.append('  <table>\n')
        rp.append(f'    <tr><td>Total Calls</td><td style="font-weight:600">{c.get("calls", "0")}</td></tr>\n')
        rp.append(f'    <tr><td>Avg/Day</td><td>{c.get("average_calls_logged_per_day", "0")}</td></tr>\n')
        rp.append(f'    <tr><td>Conversations</td><td>{c.get("conversations", "0")}</td></tr>\n')
        rp.append(f'    <tr><td>Conv Rate</td><td>{c.get("conversations_rate", "0")}%</td></tr>\n')
        rp.append(f'    <tr><td>Avg Duration</td><td>{c.get("average_calls_duration", "N/A")}</td></tr>\n')
        rp.append(f'    <tr><td>Voicemails</td><td>{c.get("voicemails", "0")}</td></tr>\n')
        rp.append('  </table>\n</div>\n')

        rp.append('<div class="section">\n')
        rp.append(f'  <h2><img src="../assets/email.svg" alt="" style="width:18px;height:18px;vertical-align:middle;opacity:0.7;"> Emails</h2>\n')
        rp.append('  <table>\n')
        rp.append(f'    <tr><td>Sent</td><td style="font-weight:600">{e.get("sent", "0")}</td></tr>\n')
        rp.append(f'    <tr><td>Opened</td><td>{e.get("opened", "0")} ({e.get("opened_rate", "0")}%)</td></tr>\n')
        rp.append(f'    <tr><td>Replied</td><td>{e.get("replied", "0")} ({e.get("replied_rate", "0")}%)</td></tr>\n')
        rp.append(f'    <tr><td>Positive Replies</td><td>{e.get("replied_positive", "0")}</td></tr>\n')
        rp.append(f'    <tr><td>Personalization</td><td>{e.get("personalized_rate", "0")}%</td></tr>\n')
        rp.append(f'    <tr><td>Clicked</td><td>{e.get("clicked", "0")} ({e.get("clicked_rate", "0")}%)</td></tr>\n')
        rp.append('  </table>\n</div>\n')
        rp.append('</div>\n')

        # Stale Deals (>14 days)
        if rep_stale:
            rp.append('<div class="section">\n')
            rp.append(f'  <h2><img src="../assets/percentage.svg" alt="" style="width:18px;height:18px;vertical-align:middle;opacity:0.7;"> Stale Deals ({len(rep_stale)} deals &gt; 14 days)</h2>\n')
            rp.append('  <table>\n')
            rp.append('    <thead><tr><th>Account</th><th>Opportunity</th><th>Stage</th><th>Days</th><th>ARR</th><th>Next Step</th></tr></thead>\n')
            rp.append('    <tbody>\n')
            for o in rep_stale[:15]:
                acct = o.get("Account Name", "")
                opp_name = o.get("Opportunity Name", "")[:40]
                stage = o.get("Stage", "")
                days = int(float(o.get("Stage Duration", "0") or 0))
                opp_arr = parse_money(o.get("Software (Annual)", "0"))
                next_step = (o.get("Next Step", "") or "")[:60]
                days_color = "var(--red)" if days > 30 else "var(--yellow)"
                rp.append(f'      <tr><td><strong>{acct}</strong></td><td style="font-size:12px">{opp_name}</td><td style="font-size:12px">{stage}</td><td style="color:{days_color};font-weight:600">{days}d</td><td>{fmt_money(opp_arr)}</td><td style="font-size:11px;color:var(--gray-500);max-width:200px;white-space:normal">{next_step}</td></tr>\n')
            rp.append('    </tbody>\n  </table>\n</div>\n')

        # Recent Wins
        if rep_wins:
            rp.append('<div class="section">\n')
            rp.append(f'  <h2><img src="../assets/dollar.svg" alt="" style="width:18px;height:18px;vertical-align:middle;opacity:0.7;"> Recent Wins ({len(rep_wins)} this month)</h2>\n')
            rp.append('  <table>\n')
            rp.append('    <thead><tr><th>Account</th><th>Type</th><th>ARR</th><th>Close Date</th></tr></thead>\n')
            rp.append('    <tbody>\n')
            for w in rep_wins[:10]:
                acct = w.get("Account Name", "")
                wtype = w.get("Type", "")
                w_arr = parse_money(w.get("Software (Annual)", "0"))
                close = w.get("Close Date", "")
                rp.append(f'      <tr><td><strong>{acct}</strong></td><td style="font-size:12px">{wtype}</td><td>{fmt_money(w_arr)}</td><td>{close}</td></tr>\n')
            rp.append('    </tbody>\n  </table>\n</div>\n')

        # Week-over-week trend table (if history exists)
        if len(all_history) > 1:
            rp.append('<div class="section">\n')
            rp.append(f'  <h2><img src="../assets/trending-up.svg" alt="" style="width:18px;height:18px;vertical-align:middle;opacity:0.7;"> Week-over-Week Trend</h2>\n')
            rp.append('  <table>\n')
            rp.append('    <thead><tr><th>Date</th><th>ARR</th><th>Wins</th><th>Demos</th><th>Opps</th><th>Calls</th><th>Emails</th><th>Demo:Win</th></tr></thead>\n')
            rp.append('    <tbody>\n')
            for h in reversed(all_history[-8:]):
                hr = h.get("reps", {}).get(name, {})
                if hr:
                    rp.append(f'      <tr><td>{h["date"]}</td><td>{fmt_money(hr.get("arr", 0))}</td><td>{hr.get("wins", 0)}</td><td>{hr.get("demos", 0)}</td><td>{hr.get("opps", 0)}</td><td>{hr.get("calls", 0)}</td><td>{hr.get("emails_sent", 0)}</td><td>{fmt_pct(hr.get("demo_win", 0))}</td></tr>\n')
            rp.append('    </tbody>\n  </table>\n</div>\n')

        rp.append(f'\n<div class="footer">Growth Sales Dashboard — {name} | Generated {generated}</div>\n</body></html>\n')

        rep_path = reps_dir / f"{slug}.html"
        rep_path.write_text("".join(rp))

    print(f"  Rep pages generated: {len(metrics)} pages in reps/")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print(f"Usage: {sys.argv[0]} /path/to/MM.DD.Data/")
        sys.exit(1)
    generate(sys.argv[1])
