#!/usr/bin/env python3
"""
Generate the dm.mellender.io sales dashboard HTML from daily CSV data.

Usage:
    python3 generate.py /path/to/MM.DD.Data/

Reads the 5 CSVs from the data folder and outputs index.html in this script's directory.
"""

import csv
import os
import sys
from datetime import datetime, date
from pathlib import Path


def read_csv(path):
    with open(path, newline="", encoding="utf-8-sig", errors="replace") as f:
        return list(csv.DictReader(f))


def parse_pct(s):
    """Parse '28.4%' -> 28.4, or return 0.0"""
    if not s:
        return 0.0
    return float(s.replace("%", "").replace(",", ""))


def parse_money(s):
    """Parse '$17,625' or '17625.00' -> 17625.0"""
    if not s:
        return 0.0
    return float(s.replace("$", "").replace(",", ""))


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

    # Figure out business days in month for pacing
    day_of_month = int(day_num)
    # Approximate business days in a month
    total_biz_days = 22
    expected_pct = (day_of_month / 30) * 100  # rough pacing expectation

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

    # Rep order from metrics
    reps = [r["Rep Name"] for r in metrics]

    # --- Compute KPI totals ---
    team_arr = sum(parse_money(r.get("Total Booked Saas ARR", "0")) for r in metrics)
    team_quota = sum(parse_money(r.get("Booked SaaS Quota (Xactly)", "0")) for r in metrics)
    team_attainment = (team_arr / team_quota * 100) if team_quota else 0
    team_wins_count = sum(int(r.get("Wins", "0") or 0) for r in metrics)

    # Pipeline from open opps
    total_pipeline = sum(parse_money(r.get("Software (Annual)", "0")) for r in opps)
    total_open_deals = len(opps)

    # Stale deals (>30 days old based on stage duration)
    stale_count = sum(1 for r in opps if float(r.get("Stage Duration", "0") or 0) > 30)

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
<style>
  :root {
    --toast-orange: #FF6600; --toast-orange-light: #FF8533;
    --toast-dark: #1A1A2E; --toast-dark-light: #2D2D44;
    --green: #22C55E; --green-bg: #F0FDF4;
    --yellow: #EAB308; --yellow-bg: #FEFCE8;
    --red: #EF4444; --red-bg: #FEF2F2;
    --blue: #3B82F6; --blue-bg: #EFF6FF;
    --gray-50: #F9FAFB; --gray-100: #F3F4F6; --gray-200: #E5E7EB;
    --gray-300: #D1D5DB; --gray-400: #9CA3AF; --gray-500: #6B7280;
    --gray-600: #4B5563; --gray-700: #374151; --gray-800: #1F2937;
    --shadow: 0 1px 3px rgba(0,0,0,0.1), 0 1px 2px rgba(0,0,0,0.06);
    --shadow-md: 0 4px 6px -1px rgba(0,0,0,0.1), 0 2px 4px -1px rgba(0,0,0,0.06);
  }
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: var(--gray-50); color: var(--gray-800); line-height: 1.5; }
  .header { background: linear-gradient(135deg, var(--toast-dark) 0%, var(--toast-dark-light) 100%); color: white; padding: 24px 32px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 12px; }
  .header-left h1 { font-size: 24px; font-weight: 700; }
  .header-left h1 span { color: var(--toast-orange); }
  .header-left p { font-size: 14px; color: var(--gray-400); margin-top: 2px; }
  .header-right { text-align: right; font-size: 13px; color: var(--gray-400); }
  .header-right .date { font-size: 16px; color: white; font-weight: 600; }
  .kpi-banner { display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 16px; padding: 20px 32px; background: white; border-bottom: 1px solid var(--gray-200); box-shadow: var(--shadow); }
  .kpi-card { text-align: center; padding: 12px; }
  .kpi-card .label { font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px; color: var(--gray-500); font-weight: 600; }
  .kpi-card .value { font-size: 28px; font-weight: 800; margin: 4px 0; color: var(--toast-dark); }
  .kpi-card .sub { font-size: 12px; color: var(--gray-500); }
  .section { padding: 24px 32px; }
  .section-title { font-size: 18px; font-weight: 700; margin-bottom: 16px; display: flex; align-items: center; gap: 8px; }
  table { width: 100%; border-collapse: collapse; background: white; border-radius: 8px; overflow: hidden; box-shadow: var(--shadow); margin-bottom: 8px; }
  th { background: var(--gray-100); font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px; color: var(--gray-600); font-weight: 600; padding: 10px 14px; text-align: left; }
  td { padding: 10px 14px; border-top: 1px solid var(--gray-100); font-size: 13px; }
  tr:hover { background: var(--gray-50); }
  .progress-bar { background: var(--gray-200); border-radius: 999px; height: 8px; width: 100px; display: inline-block; vertical-align: middle; }
  .progress-fill { height: 100%; border-radius: 999px; }
  .badge { display: inline-block; padding: 2px 8px; border-radius: 999px; font-size: 11px; font-weight: 600; }
  .badge-green { background: var(--green-bg); color: var(--green); }
  .badge-yellow { background: var(--yellow-bg); color: var(--yellow); }
  .badge-red { background: var(--red-bg); color: var(--red); }
  .badge-blue { background: var(--blue-bg); color: var(--blue); }
  .forecast-bar { display: flex; align-items: center; gap: 8px; margin: 8px 0; }
  .forecast-bar-track { flex: 1; background: var(--gray-200); border-radius: 4px; height: 24px; position: relative; }
  .forecast-bar-fill { height: 100%; border-radius: 4px; }
  .forecast-bar-label { font-size: 13px; font-weight: 600; min-width: 50px; text-align: right; }
  .forecast-bar-name { font-size: 14px; font-weight: 600; min-width: 130px; }
  .pace-line { position: absolute; top: -4px; bottom: -4px; width: 2px; background: var(--gray-700); z-index: 2; }
  .pace-line::after { content: "\\25BC"; position: absolute; top: -14px; left: -4px; font-size: 9px; color: var(--gray-700); }
  .two-col { display: grid; grid-template-columns: 1fr 1fr; gap: 24px; }
  @media (max-width: 900px) { .two-col { grid-template-columns: 1fr; } }
  .footer { text-align: center; padding: 16px; font-size: 12px; color: var(--gray-400); }
</style>
</head>
<body>
""")

    # Header
    html.append(f"""
<div class="header">
  <div class="header-left">
    <h1>&#x1F35E; <span>Toast</span> Growth Sales Dashboard</h1>
    <p>Josh Mellender — Growth Sales District Manager | Reports to Katie Patterson (RVP)</p>
  </div>
  <div class="header-right">
    <div class="date">{month_name} — MTD Performance</div>
    <div>Data as of {data_date}</div>
  </div>
</div>
""")

    # KPI Banner
    pacing_color = "var(--red)" if team_pacing_val < expected_pct else "var(--green)"
    pacing_label = "Behind pace" if team_pacing_val < expected_pct else "On track"
    html.append(f"""
<div class="kpi-banner">
  <div class="kpi-card"><div class="label">Team MTD ACV</div><div class="value">{fmt_money(team_arr)}</div><div class="sub">of {fmt_money(team_quota)} quota</div></div>
  <div class="kpi-card"><div class="label">Attainment</div><div class="value">{fmt_pct(team_attainment)}</div><div class="sub">Day {day_of_month}/{total_biz_days}</div></div>
  <div class="kpi-card"><div class="label">Team Pacing</div><div class="value" style="color:{pacing_color}">{fmt_pct(team_pacing_val)}</div><div class="sub">{pacing_label}</div></div>
  <div class="kpi-card"><div class="label">Closed Deals</div><div class="value">{team_wins_count}</div><div class="sub">MTD</div></div>
  <div class="kpi-card"><div class="label">Open Pipeline</div><div class="value">{fmt_money(total_pipeline)}</div><div class="sub">{total_open_deals} opps</div></div>
  <div class="kpi-card"><div class="label">Stale (&gt;30d)</div><div class="value" style="color: var(--red)">{stale_count}</div><div class="sub">stuck deals</div></div>
</div>
""")

    # ---- Team Overview Table ----
    html.append("""
<div class="section">
  <div class="section-title">&#x1F4CA; Team Overview — MTD ACV vs Quota</div>
  <table>
    <thead><tr><th>Rep</th><th>Role</th><th>Quota</th><th>MTD ACV</th><th>%</th><th>Progress</th><th>Wins</th><th>Pacing</th></tr></thead>
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
        badge = pacing_badge(pct, day_of_month, 30)
        html.append(f'<tr><td><strong>{name}</strong></td><td style="font-size:12px;color:var(--gray-500)">{role_display}</td><td>{fmt_money(quota)}</td><td><strong>{fmt_money(arr)}</strong></td><td>{fmt_pct(pct)}</td><td><div class="progress-bar"><div class="progress-fill" style="width:{bar_w:.1f}%;background:{color}"></div></div></td><td>{wins_count}</td><td>{badge}</td></tr>\n')

    html.append(f'<tr style="background:var(--gray-100);font-weight:700"><td>TEAM</td><td>{len(metrics)} AEs</td><td>{fmt_money(team_quota)}</td><td>{fmt_money(team_arr)}</td><td>{fmt_pct(team_attainment)}</td><td></td><td>{team_wins_count}</td><td></td></tr>\n')
    html.append("    </tbody></table>\n</div>\n")

    # ---- Self-Sourced + NBR Referrals (two-col) ----
    html.append("""
<div class="section">
  <div class="two-col">
    <div>
      <div class="section-title">&#x1F3AF; Self-Sourced Opps Created</div>
      <table>
        <thead><tr><th>Rep</th><th>Created</th><th>Goal</th><th>% to Goal</th></tr></thead>
        <tbody>
""")
    team_ss = 0
    team_ss_goal = 0
    for r in metrics:
        name = r["Rep Name"]
        ss = int(r.get("Self-Sourced Created Opps", "0") or 0)
        goal = int(r.get("Total SS Opp Goal", "0") or 0)
        pct = parse_pct(r.get("SS Opps Created % to Goal", "0"))
        color = pct_color(pct, 60, 40)
        team_ss += ss
        team_ss_goal += goal
        html.append(f'<tr><td><strong>{name}</strong></td><td>{ss}</td><td>{goal}</td><td style="color:{color};font-weight:600">{fmt_pct(pct)}</td></tr>\n')
    html.append(f'<tr style="background:var(--gray-100);font-weight:700"><td>TEAM</td><td>{team_ss}</td><td>{team_ss_goal}</td><td></td></tr>\n')
    html.append("        </tbody></table>\n    </div>\n")

    # NBR Referrals
    html.append("""    <div>
      <div class="section-title">&#x1F91D; NBR Referrals</div>
      <table>
        <thead><tr><th>Rep</th><th>Created</th><th>Closed Won</th></tr></thead>
        <tbody>
""")
    team_nbr_created = 0
    team_nbr_won = 0
    for r in metrics:
        name = r["Rep Name"]
        created = int(r.get("NB Referral Created Opps", "0") or 0)
        won = int(r.get("NB Referral Closed Won", "0") or 0)
        team_nbr_created += created
        team_nbr_won += won
        html.append(f'<tr><td><strong>{name}</strong></td><td>{created}</td><td>{won}</td></tr>\n')
    html.append(f'<tr style="background:var(--gray-100);font-weight:700"><td>TEAM</td><td>{team_nbr_created}</td><td>{team_nbr_won}</td></tr>\n')
    html.append("        </tbody></table>\n    </div>\n  </div>\n</div>\n")

    # ---- EC / Payroll Sales ----
    html.append("""
<div class="section">
  <div class="section-title">&#x1F4B3; Employee Cloud / Payroll (EC) Sales</div>
  <table>
    <thead><tr><th>Rep</th><th>EC Units</th><th>EC ARR</th><th>EC Goal</th><th>EC % to Goal</th></tr></thead>
    <tbody>
""")
    for r in metrics:
        name = r["Rep Name"]
        units = int(r.get("EC Units", "0") or 0)
        ec_arr = parse_money(r.get("EC ARR", "0"))
        ec_goal = parse_money(r.get("EC ARR Goal", "0"))
        ec_pct = parse_pct(r.get("EC ARR % to Goal", "0"))
        color = pct_color(ec_pct, 50, 25)
        html.append(f'<tr><td><strong>{name}</strong></td><td>{units}</td><td>{fmt_money(ec_arr)}</td><td>{fmt_money(ec_goal)}</td><td style="color:{color};font-weight:600">{fmt_pct(ec_pct)}</td></tr>\n')
    html.append("    </tbody></table>\n</div>\n")

    # ---- Pipeline Health ----
    html.append("""
<div class="section">
  <div class="section-title">&#x1F52E; Pipeline Health</div>
  <table>
    <thead><tr><th>Rep</th><th>Open Deals</th><th>Pipeline Value</th><th>Stale (&gt;30d)</th><th>Health</th></tr></thead>
    <tbody>
""")
    # Group open opps by rep
    rep_pipeline = {}
    for r in opps:
        owner = r.get("Opportunity Owner", "Unknown")
        if owner not in rep_pipeline:
            rep_pipeline[owner] = {"deals": 0, "value": 0.0, "stale": 0}
        rep_pipeline[owner]["deals"] += 1
        rep_pipeline[owner]["value"] += parse_money(r.get("Software (Annual)", "0"))
        if float(r.get("Stage Duration", "0") or 0) > 30:
            rep_pipeline[owner]["stale"] += 1

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

    # ---- Calls & Emails Activity (two-col) ----
    html.append("""
<div class="section">
  <div class="two-col">
    <div>
      <div class="section-title">&#x1F4DE; Calls Activity</div>
      <table>
        <thead><tr><th>Rep</th><th>Calls</th><th>Avg/Day</th><th>Conversations</th><th>Conv Rate</th><th>Avg Duration</th></tr></thead>
        <tbody>
""")
    calls_by_name = {r["user_name"]: r for r in calls}
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
      <div class="section-title">&#x1F4E7; Email Activity</div>
      <table>
        <thead><tr><th>Rep</th><th>Sent</th><th>Opened</th><th>Open Rate</th><th>Replied</th><th>Reply Rate</th></tr></thead>
        <tbody>
""")
    emails_by_name = {r["user_name"]: r for r in emails}
    for name in reps:
        e = emails_by_name.get(name, {})
        sent = e.get("sent", "0")
        opened = e.get("opened", "0")
        open_rate = e.get("opened_rate", "0")
        replied = e.get("replied", "0")
        reply_rate = e.get("replied_rate", "0")
        html.append(f'<tr><td><strong>{name}</strong></td><td>{sent}</td><td>{opened}</td><td>{open_rate}%</td><td>{replied}</td><td>{reply_rate}%</td></tr>\n')
    html.append("        </tbody></table>\n    </div>\n  </div>\n</div>\n")

    # ---- Recent Wins ----
    html.append("""
<div class="section">
  <div class="section-title">&#x1F3C6; Recent Wins</div>
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
        owner = w.get("Opportunity Owner", "")
        html.append(f'<tr><td><strong>{acct}</strong></td><td style="font-size:12px">{wtype}</td><td>{fmt_money(arr)}</td><td>{close}</td><td>{owner}</td></tr>\n')
    html.append("    </tbody></table>\n</div>\n")

    # ---- Forecast vs Quota bars ----
    pace_pct = (day_of_month / 30) * 100
    html.append("""
<div class="section">
  <div class="section-title">&#x1F4C8; Forecast vs Quota — Pacing</div>
""")
    for r in metrics:
        name = r["Rep Name"]
        quota = parse_money(r.get("Booked SaaS Quota (Xactly)", "0"))
        pct = parse_pct(r.get("ARR % to Goal (Xactly)", "0"))
        bar_w = min(pct, 100)
        color = bar_color(pct, expected_pct)
        html.append(f'<div class="forecast-bar"><div class="forecast-bar-name">{name}</div><div class="forecast-bar-track"><div class="forecast-bar-fill" style="width:{bar_w:.1f}%;background:{color}"></div><div class="pace-line" style="left:{pace_pct:.1f}%"></div></div><div class="forecast-bar-label">{fmt_money(quota)}</div><div class="forecast-bar-label">{fmt_pct(pct)}</div></div>\n')

    html.append(f'<div class="forecast-bar" style="margin-top:12px;padding-top:12px;border-top:2px solid var(--gray-300)"><div class="forecast-bar-name" style="font-weight:800">TEAM TOTAL</div><div class="forecast-bar-track"><div class="forecast-bar-fill" style="width:{min(team_attainment, 100):.1f}%;background:var(--toast-orange)"></div><div class="pace-line" style="left:{pace_pct:.1f}%"></div></div><div class="forecast-bar-label">{fmt_money(team_quota)}</div><div class="forecast-bar-label">{fmt_pct(team_attainment)}</div></div>\n')
    html.append(f'  <p style="margin-top:12px;font-size:13px;color:var(--gray-500)">&#x25BC; = Expected pace (Day {day_of_month}/{total_biz_days})</p>\n</div>\n')

    # Footer
    generated = datetime.now().strftime("%b %d, %Y")
    html.append(f'\n<div class="footer">Growth Sales Team Dashboard — Built for Josh Mellender | Data: Sigma + Xactly + SalesLoft | Generated {generated}</div>\n</body></html>\n')

    # Write output
    out_path = Path(__file__).parent / "index.html"
    out_path.write_text("".join(html))
    print(f"Dashboard generated: {out_path}")
    print(f"  Month: {month_name} | Data date: {data_date}")
    print(f"  Team ARR: {fmt_money(team_arr)} / {fmt_money(team_quota)} ({fmt_pct(team_attainment)})")
    print(f"  Wins: {team_wins_count} | Pipeline: {fmt_money(total_pipeline)} ({total_open_deals} opps)")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print(f"Usage: {sys.argv[0]} /path/to/MM.DD.Data/")
        sys.exit(1)
    generate(sys.argv[1])
