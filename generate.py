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


BADGE_ON_TRACK = '<span class="badge badge-green">On Track</span>'
BADGE_AT_RISK = '<span class="badge badge-red">At Risk</span>'


def icon(name, size=18):
    """Return an inline img tag for a Toast icon asset."""
    return f'<img src="assets/{name}.svg" alt="" style="width:{size}px;height:{size}px;vertical-align:middle;opacity:0.7;">'


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
    --toast-navy: #2B4FB9; --toast-navy-light: #3D63D4;
    --toast-dark: #2B2E35; --toast-dark-light: #3A3D45;
    --green: #22C55E; --green-bg: #F0FDF4;
    --yellow: #EAB308; --yellow-bg: #FEFCE8;
    --red: #EF4444; --red-bg: #FEF2F2;
    --blue: #3B82F6; --blue-bg: #EFF6FF;
    --warm-50: #F9F5F3; --warm-100: #F6F1EE;
    --gray-50: #F7FAFC; --gray-100: #F3F4F6; --gray-200: #E1E7EE;
    --gray-300: #D1D5DB; --gray-400: #9CA3AF; --gray-500: #6B7280;
    --gray-600: #4B5563; --gray-700: #374151; --gray-800: #252525;
    --shadow: 0 1px 3px rgba(0,0,0,0.08), 0 1px 2px rgba(0,0,0,0.04);
    --shadow-md: 0 4px 6px -1px rgba(0,0,0,0.08), 0 2px 4px -1px rgba(0,0,0,0.04);
    --radius: 12px;
  }
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body { font-family: 'Source Sans 3', 'Source Sans Pro', -apple-system, BlinkMacSystemFont, sans-serif; background: var(--warm-50); color: var(--gray-800); line-height: 1.5; }
  .header { background: var(--toast-dark); color: white; padding: 20px 32px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 12px; border-bottom: 3px solid var(--toast-orange); }
  .header-left { display: flex; align-items: center; gap: 16px; }
  .header-left .logo { width: 36px; height: 36px; }
  .header-left h1 { font-size: 22px; font-weight: 700; letter-spacing: -0.3px; }
  .header-left h1 span { color: var(--toast-orange); }
  .header-left p { font-size: 13px; color: var(--gray-400); margin-top: 2px; }
  .header-right { text-align: right; font-size: 13px; color: var(--gray-400); }
  .header-right .date { font-size: 15px; color: white; font-weight: 600; }
  .kpi-banner { display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 12px; padding: 20px 32px; background: white; border-bottom: 1px solid var(--gray-200); box-shadow: var(--shadow); }
  .kpi-card { text-align: center; padding: 14px 12px; border-radius: var(--radius); border: 1px solid var(--gray-200); }
  .kpi-card .label { font-size: 10px; text-transform: uppercase; letter-spacing: 0.6px; color: var(--gray-500); font-weight: 700; }
  .kpi-card .value { font-size: 26px; font-weight: 800; margin: 4px 0; color: var(--toast-dark); }
  .kpi-card .sub { font-size: 11px; color: var(--gray-500); }
  .section { padding: 24px 32px; }
  .section-title { font-size: 16px; font-weight: 700; margin-bottom: 14px; display: flex; align-items: center; gap: 8px; color: var(--toast-dark); letter-spacing: -0.2px; }
  .section-title .accent { color: var(--toast-orange); }
  table { width: 100%; border-collapse: collapse; background: white; border-radius: var(--radius); overflow: hidden; box-shadow: var(--shadow); margin-bottom: 8px; }
  th { background: var(--toast-dark); font-size: 10px; text-transform: uppercase; letter-spacing: 0.6px; color: rgba(255,255,255,0.85); font-weight: 700; padding: 10px 14px; text-align: left; }
  td { padding: 10px 14px; border-top: 1px solid var(--gray-100); font-size: 13px; }
  tr:hover { background: var(--warm-100); }
  .progress-bar { background: var(--gray-200); border-radius: 999px; height: 8px; width: 100px; display: inline-block; vertical-align: middle; }
  .progress-fill { height: 100%; border-radius: 999px; }
  .badge { display: inline-block; padding: 3px 10px; border-radius: 999px; font-size: 11px; font-weight: 700; }
  .badge-green { background: var(--green-bg); color: var(--green); }
  .badge-yellow { background: var(--yellow-bg); color: var(--yellow); }
  .badge-red { background: var(--red-bg); color: var(--red); }
  .badge-blue { background: var(--blue-bg); color: var(--blue); }
  .badge-orange { background: var(--toast-orange-bg); color: var(--toast-orange); }
  .forecast-bar { display: flex; align-items: center; gap: 8px; margin: 8px 0; }
  .forecast-bar-track { flex: 1; background: var(--gray-200); border-radius: 6px; height: 24px; position: relative; }
  .forecast-bar-fill { height: 100%; border-radius: 6px; }
  .forecast-bar-label { font-size: 13px; font-weight: 600; min-width: 50px; text-align: right; }
  .forecast-bar-name { font-size: 14px; font-weight: 600; min-width: 130px; }
  .pace-line { position: absolute; top: -4px; bottom: -4px; width: 2px; background: var(--gray-700); z-index: 2; }
  .pace-line::after { content: "\\25BC"; position: absolute; top: -14px; left: -4px; font-size: 9px; color: var(--gray-700); }
  .two-col { display: grid; grid-template-columns: 1fr 1fr; gap: 24px; }
  @media (max-width: 900px) { .two-col { grid-template-columns: 1fr; } }
  .footer { text-align: center; padding: 20px; font-size: 12px; color: var(--gray-400); border-top: 1px solid var(--gray-200); margin-top: 8px; }
</style>
</head>
<body>
""")

    # Header
    html.append(f"""
<div class="header">
  <div class="header-left">
    <img class="logo" src="assets/toast-logo_white.svg" alt="Toast" style="height:28px;width:auto;">
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

    # ---- MBO Tracker ----
    # Compute MBO metrics from data
    total_opps = sum(int(r.get("Opps", "0") or 0) for r in metrics)
    total_demos = sum(int(r.get("Demos", "0") or 0) for r in metrics)
    total_ss = sum(int(r.get("Self-Sourced Created Opps", "0") or 0) for r in metrics)
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
  <div class="section-title">{icon('Other_Award')} MBO Tracker — District Scorecard</div>
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

    html.append("""      <table>
        <thead><tr><th>Key Metrics</th><th>Goal</th><th>Actual</th><th>District % to Goal</th></tr></thead>
        <tbody>
""")
    html.append(f'<tr><td>Total Opps Created</td><td></td><td>{total_opps}</td><td></td></tr>\n')
    html.append(f'<tr><td>Opps Created Per Rep</td><td></td><td>{opps_per_rep:.1f}</td><td></td></tr>\n')
    avg_arr_color = pct_color(avg_arr_per_opp / 1800 * 100, 100, 80)
    html.append(f'<tr><td>Avg ARR</td><td>$1,800</td><td style="font-weight:600">{fmt_money(avg_arr_per_opp)}</td><td style="color:{avg_arr_color};font-weight:600">{fmt_pct(avg_arr_per_opp/1800*100)}</td></tr>\n')
    html.append(f'<tr><td>Opp:Win Conversion</td><td></td><td>{fmt_pct(team_opp_win)}</td><td></td></tr>\n')
    html.append(f'<tr><td>Opp:Demo Conversion</td><td></td><td>{fmt_pct(team_opp_demo)}</td><td></td></tr>\n')
    html.append(f'<tr><td>NBR Referral Opps (total)</td><td></td><td>{total_nbr}</td><td></td></tr>\n')
    html.append(f'<tr><td>SS Opps (total)</td><td></td><td>{total_ss}</td><td></td></tr>\n')
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

    # ---- Team Overview Table ----
    html.append("""
<div class="section">
  <div class="section-title">{icon('bar-chart')} Team Overview — MTD ACV vs Quota</div>
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
      <div class="section-title">{icon('star')} Self-Sourced Opps Created</div>
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
      <div class="section-title">{icon('restaurant-group')} NBR Referrals</div>
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
  <div class="section-title">{icon('pay-check')} Employee Cloud / Payroll (EC) Sales</div>
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
  <div class="section-title">{icon('analytics-check')} Enablement MBO Tracker <span style="font-size:12px;color:var(--gray-500);font-weight:400">(WorkRamp — as of {e_date})</span></div>
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

    # ---- Pipeline Health ----
    html.append("""
<div class="section">
  <div class="section-title">{icon('percentage')} Pipeline Health</div>
  <table>
    <thead><tr><th>Rep</th><th>Open Deals</th><th>Pipeline Value</th><th>Stale (&gt;30d)</th><th>Health</th></tr></thead>
    <tbody>
""")
    # Group open opps by rep (normalize names to match Key Metrics)
    rep_pipeline = {}
    for r in opps:
        owner = normalize_name(r.get("Opportunity Owner", "Unknown"))
        if owner not in reps:
            continue  # skip people no longer on the team
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
      <div class="section-title">{icon('phone')} Calls Activity</div>
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
      <div class="section-title">{icon('email')} Email Activity</div>
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
  <div class="section-title">{icon('dollar')} Recent Wins</div>
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
    pace_pct = (day_of_month / 30) * 100
    html.append("""
<div class="section">
  <div class="section-title">{icon('trending-up')} Forecast vs Quota — Pacing</div>
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
