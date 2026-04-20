#!/usr/bin/env python3
"""
Daily digest generator for dm.mellender.io dashboard.

Reads the latest historical snapshot and generates a summary suitable for
Slack (webhook) or email. Can be run via cron or GitHub Actions.

Usage:
    python3 digest.py                  # Print to stdout
    python3 digest.py --slack URL      # POST to Slack webhook
    python3 digest.py --email TO_ADDR  # Send via email (requires SMTP config)
"""

import json
import sys
import os
from pathlib import Path
from datetime import date


def load_latest_snapshots():
    history_dir = Path(__file__).parent / "data" / "history"
    files = sorted(history_dir.glob("*.json"))
    if not files:
        print("No history snapshots found.")
        sys.exit(1)
    with open(files[-1]) as f:
        latest = json.load(f)
    prev = None
    if len(files) >= 2:
        with open(files[-2]) as f:
            prev = json.load(f)
    return latest, prev


def fmt_money(v):
    if v >= 1000:
        return f"${v:,.0f}"
    return f"${v:,.2f}"


def fmt_pct(v):
    return f"{v:.1f}%"


def delta_str(cur, prev, fmt="num"):
    diff = cur - prev
    if diff == 0:
        return ""
    sign = "+" if diff > 0 else ""
    arrow = ":arrow_up:" if diff > 0 else ":arrow_down:"
    if fmt == "money":
        return f" {arrow} {sign}{fmt_money(abs(diff))}"
    elif fmt == "pct":
        return f" {arrow} {sign}{diff:.1f}pp"
    return f" {arrow} {sign}{diff}"


def generate_digest():
    latest, prev = load_latest_snapshots()
    snap_date = latest["date"]
    team_arr = latest["team_arr"]
    team_quota = latest["team_quota"]
    attainment = (team_arr / team_quota * 100) if team_quota else 0
    team_wins = latest["team_wins"]
    pipeline = latest["team_pipeline"]

    import calendar
    parts = snap_date.split("-")
    year, month, day = int(parts[0]), int(parts[1]), int(parts[2])
    total_biz = sum(1 for d in range(1, calendar.monthrange(year, month)[1] + 1) if calendar.weekday(year, month, d) < 5)
    biz_passed = sum(1 for d in range(1, day + 1) if calendar.weekday(year, month, d) < 5)
    expected_pct = (biz_passed / total_biz) * 100 if total_biz else 0
    pace_target = team_quota * expected_pct / 100
    pace_delta = team_arr - pace_target
    remaining_biz = total_biz - biz_passed

    lines = []
    lines.append(f":bar_chart: *Daily Dashboard Digest — {snap_date}*")
    lines.append(f"Working Day {biz_passed}/{total_biz} ({remaining_biz} remaining)")
    lines.append("")

    lines.append(f":moneybag: *Team ARR:* {fmt_money(team_arr)} / {fmt_money(team_quota)} ({fmt_pct(attainment)})")
    if pace_delta < 0:
        lines.append(f":warning: *Pace:* Behind by {fmt_money(-pace_delta)}")
    else:
        lines.append(f":white_check_mark: *Pace:* Ahead by {fmt_money(pace_delta)}")
    lines.append(f":trophy: *Wins:* {team_wins} | *Pipeline:* {fmt_money(pipeline)} ({latest['team_open_deals']} opps)")
    lines.append("")

    lines.append("*Rep Snapshot:*")
    reps = latest.get("reps", {})
    sorted_reps = sorted(reps.items(), key=lambda x: x[1].get("arr_pct", 0), reverse=True)
    for name, rd in sorted_reps:
        arr = rd.get("arr", 0)
        pct = rd.get("arr_pct", 0)
        wins = rd.get("wins", 0)
        quota = rd.get("quota", 0)
        pace_ratio = pct / expected_pct if expected_pct and quota > 0 else 0
        if quota == 0:
            emoji = ":seedling:"
        elif pct >= 100:
            emoji = ":star:"
        elif pace_ratio >= 1.0:
            emoji = ":white_check_mark:"
        elif pace_ratio >= 0.8:
            emoji = ":large_yellow_circle:"
        else:
            emoji = ":red_circle:"
        delta = ""
        if prev and name in prev.get("reps", {}):
            prev_arr = prev["reps"][name].get("arr", 0)
            arr_diff = arr - prev_arr
            if arr_diff != 0:
                delta = delta_str(arr, prev_arr, "money")
        lines.append(f"  {emoji} *{name}:* {fmt_money(arr)} ({fmt_pct(pct)}){delta} — {wins} wins")

    lines.append("")

    movers = []
    if prev:
        for name, rd in reps.items():
            prev_rd = prev.get("reps", {}).get(name, {})
            arr_diff = rd.get("arr", 0) - prev_rd.get("arr", 0)
            win_diff = rd.get("wins", 0) - prev_rd.get("wins", 0)
            if win_diff > 0 and arr_diff > 0:
                movers.append(f":tada: *{name}* closed {win_diff} new deal{'s' if win_diff > 1 else ''} (+{fmt_money(arr_diff)} ARR)")
        if movers:
            lines.append("*New Activity:*")
            lines.extend(movers)
            lines.append("")

    ec_total = sum(rd.get("ec_arr", 0) for rd in reps.values())
    ec_pct = (ec_total / team_arr * 100) if team_arr else 0
    ec_emoji = ":white_check_mark:" if ec_pct >= 36 else ":warning:"
    lines.append(f"{ec_emoji} *EC Attach:* {fmt_pct(ec_pct)} (target: 36%)")

    lines.append("")
    lines.append(f":link: <https://dm.mellender.io|View Full Dashboard>")

    return "\n".join(lines)


def send_slack(webhook_url, text):
    import urllib.request
    payload = json.dumps({"text": text}).encode("utf-8")
    req = urllib.request.Request(webhook_url, data=payload, headers={"Content-Type": "application/json"})
    urllib.request.urlopen(req)
    print("Slack message sent.")


def send_email(to_addr, text):
    import smtplib
    from email.mime.text import MIMEText
    smtp_host = os.environ.get("SMTP_HOST", "smtp.gmail.com")
    smtp_port = int(os.environ.get("SMTP_PORT", "587"))
    smtp_user = os.environ.get("SMTP_USER", "")
    smtp_pass = os.environ.get("SMTP_PASS", "")
    if not smtp_user:
        print("Set SMTP_USER and SMTP_PASS environment variables for email.")
        sys.exit(1)
    msg = MIMEText(text)
    msg["Subject"] = f"Dashboard Digest — {date.today().isoformat()}"
    msg["From"] = smtp_user
    msg["To"] = to_addr
    with smtplib.SMTP(smtp_host, smtp_port) as server:
        server.starttls()
        server.login(smtp_user, smtp_pass)
        server.send_message(msg)
    print(f"Email sent to {to_addr}.")


if __name__ == "__main__":
    text = generate_digest()

    if "--slack" in sys.argv:
        idx = sys.argv.index("--slack")
        url = sys.argv[idx + 1]
        send_slack(url, text)
    elif "--email" in sys.argv:
        idx = sys.argv.index("--email")
        addr = sys.argv[idx + 1]
        send_email(addr, text)
    else:
        print(text)
