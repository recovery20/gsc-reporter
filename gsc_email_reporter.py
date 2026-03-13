#!/usr/bin/env python3
"""
GSC Weekly Swing Reporter → Email (Outlook / Microsoft 365)
Uses the Anthropic API (with GSC MCP connector) to pull Search Console data,
detect week-over-week swings, and send a formatted HTML email.

Requirements:
    pip install anthropic

Environment variables:
    ANTHROPIC_API_KEY   - Your Anthropic API key (sk-ant-...)
    EMAIL_ADDRESS       - Your Outlook email address (you@yourdomain.com)
    EMAIL_PASSWORD      - Your Outlook password or App Password
    SWING_THRESHOLD     - Min % change to flag, default 0.20 (20%)
    TOP_N               - Max pages per property in email, default 5
"""

import os
import json
import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import anthropic

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
ANTHROPIC_API_KEY = os.environ["ANTHROPIC_API_KEY"]
EMAIL_ADDRESS     = os.environ["EMAIL_ADDRESS"]
EMAIL_PASSWORD    = os.environ["EMAIL_PASSWORD"]
SWING_THRESHOLD   = float(os.getenv("SWING_THRESHOLD", "0.20"))
TOP_N             = int(os.getenv("TOP_N", "5"))
MIN_CLICKS        = int(os.getenv("MIN_CLICKS_FILTER", "5"))

SMTP_SERVER = "smtp.office365.com"
SMTP_PORT   = 587

PROPERTIES = [
    "https://www.columbusrecoverycenter.com/",
    "https://www.palmerlakerecovery.com/",
    "sc-domain:recoveryindianapolis.com",
    "sc-domain:recoveryatlanta.com",
    "https://www.floridarehab.com/",
    "sc-domain:orlandorecovery.com",
    "https://www.southjerseyrecovery.com/",
    "sc-domain:recoverykansascity.com",
    "sc-domain:therecoveryvillage.com",
    "sc-domain:recoverysalem.com",
    "https://www.ridgefieldrecovery.com/",
]

# ─────────────────────────────────────────────
# DATE HELPERS
# ─────────────────────────────────────────────
def get_week_ranges():
    today = datetime.date.today()
    days_since_monday = today.weekday()
    last_sunday  = today - datetime.timedelta(days=days_since_monday + 1)
    last_monday  = last_sunday - datetime.timedelta(days=6)
    prev_sunday  = last_monday - datetime.timedelta(days=1)
    prev_monday  = prev_sunday - datetime.timedelta(days=6)
    return last_monday, last_sunday, prev_monday, prev_sunday

def fmt(d):
    return d.strftime("%Y-%m-%d")

def fmt_display(d):
    return d.strftime("%b %d, %Y")

# ─────────────────────────────────────────────
# FETCH DATA VIA ANTHROPIC + GSC MCP
# ─────────────────────────────────────────────
def fetch_swing_data(site_url, curr_start, curr_end, prev_start, prev_end):
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    prompt = f"""
You have access to Google Search Console data. Please compare two time periods for this property:

Property: {site_url}
Current week:  {fmt(curr_start)} to {fmt(curr_end)}
Previous week: {fmt(prev_start)} to {fmt(prev_end)}

Use the compare_time_periods tool with dimensions=["page"] to get page-level data.

Return ONLY a JSON array of pages that had a swing of {int(SWING_THRESHOLD*100)}% or more
in clicks, impressions, CTR, or position. Only include pages where at least one week had {MIN_CLICKS}+ clicks.

Format each item exactly like this:
{{
  "page": "https://example.com/page",
  "clicks_curr": 120,
  "clicks_prev": 80,
  "impressions_curr": 1500,
  "impressions_prev": 900,
  "ctr_curr": 0.08,
  "ctr_prev": 0.089,
  "position_curr": 4.2,
  "position_prev": 6.1
}}

Return only the raw JSON array, nothing else.
"""

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4000,
        mcp_servers=[
            {
                "type": "url",
                "url": "https://gsc.mcp.claude.com/mcp",
                "name": "search-console-analytics"
            }
        ],
        messages=[{"role": "user", "content": prompt}]
    )

    text = ""
    for block in response.content:
        if hasattr(block, "text"):
            text += block.text

    text = text.strip()
    if text.startswith("```"):
        text = text.split("```")[1]
        if text.startswith("json"):
            text = text[4:]
    text = text.strip()

    try:
        return json.loads(text)
    except json.JSONDecodeError:
        print(f"  Warning: Could not parse JSON for {site_url}")
        return []

# ─────────────────────────────────────────────
# METRIC HELPERS
# ─────────────────────────────────────────────
def pct_change(prev, curr):
    if prev == 0:
        return None
    return (curr - prev) / prev

def fmt_val(val, metric):
    if metric == "ctr":
        return f"{val*100:.2f}%"
    elif metric == "position":
        return f"{val:.1f}"
    else:
        return f"{int(val):,}"

def fmt_pct(pct):
    if pct is None:
        return "n/a"
    sign = "+" if pct >= 0 else ""
    return f"{sign}{pct*100:.1f}%"

def is_positive(pct, metric):
    """True = good direction (green), False = bad direction (red)."""
    if pct is None:
        return None
    if metric == "position":
        return pct < 0   # lower position = better
    return pct > 0

def short_url(url):
    for prefix in ["https://www.", "https://", "http://www.", "http://"]:
        if url.startswith(prefix):
            url = url[len(prefix):]
            break
    return url.rstrip("/")[:70]

# ─────────────────────────────────────────────
# HTML EMAIL BUILDER
# ─────────────────────────────────────────────
def build_html(all_swings, curr_start, curr_end, prev_start, prev_end):
    total = sum(len(v) for v in all_swings.values())
    threshold_pct = int(SWING_THRESHOLD * 100)

    html = f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
  body {{ font-family: Arial, sans-serif; background: #f4f6f9; margin: 0; padding: 20px; color: #222; }}
  .wrapper {{ max-width: 800px; margin: 0 auto; background: #fff; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }}
  .header {{ background: #1A56A0; color: #fff; padding: 28px 32px; }}
  .header h1 {{ margin: 0 0 6px 0; font-size: 22px; }}
  .header p {{ margin: 0; font-size: 13px; opacity: 0.8; }}
  .summary {{ background: #e8f0fb; padding: 14px 32px; font-size: 13px; color: #1A56A0; border-bottom: 1px solid #d0dff5; }}
  .summary span {{ margin-right: 28px; font-weight: bold; }}
  .property {{ border-bottom: 2px solid #eef0f3; padding: 20px 32px; }}
  .property-title {{ font-size: 16px; font-weight: bold; color: #1A56A0; margin: 0 0 4px 0; }}
  .date-range {{ font-size: 11px; color: #888; margin-bottom: 14px; }}
  .no-swings {{ color: #27ae60; font-size: 13px; font-style: italic; }}
  table.page-table {{ width: 100%; border-collapse: collapse; font-size: 12px; margin-bottom: 8px; }}
  table.page-table th {{ background: #f0f4fa; color: #555; text-align: left; padding: 7px 10px; border-bottom: 2px solid #dce3ef; }}
  table.page-table td {{ padding: 7px 10px; border-bottom: 1px solid #f0f0f0; vertical-align: top; }}
  table.page-table tr:hover td {{ background: #fafbff; }}
  .page-url {{ color: #1A56A0; text-decoration: none; font-weight: 500; word-break: break-all; }}
  .up {{ color: #1a7a3e; font-weight: bold; }}
  .down {{ color: #c0392b; font-weight: bold; }}
  .neutral {{ color: #888; }}
  .more {{ font-size: 11px; color: #888; font-style: italic; padding: 4px 0; }}
  .footer {{ background: #f9f9f9; padding: 16px 32px; font-size: 11px; color: #aaa; text-align: center; border-top: 1px solid #eee; }}
</style>
</head>
<body>
<div class="wrapper">
  <div class="header">
    <h1>📊 Weekly GSC Swing Report</h1>
    <p>{fmt_display(curr_start)} – {fmt_display(curr_end)} vs {fmt_display(prev_start)} – {fmt_display(prev_end)}</p>
  </div>
  <div class="summary">
    <span>Threshold: ±{threshold_pct}%</span>
    <span>Properties: {len(all_swings)}</span>
    <span>Pages flagged: {total}</span>
  </div>
"""

    metrics = [
        ("Clicks",      "clicks_curr",      "clicks_prev",      "clicks"),
        ("Impressions", "impressions_curr",  "impressions_prev", "impressions"),
        ("CTR",         "ctr_curr",          "ctr_prev",         "ctr"),
        ("Position",    "position_curr",     "position_prev",    "position"),
    ]

    for site_url, pages in all_swings.items():
        domain = (site_url
                  .replace("sc-domain:", "")
                  .replace("https://www.", "")
                  .replace("https://", "")
                  .strip("/"))

        html += f"""
  <div class="property">
    <div class="property-title">🌐 {domain}</div>
    <div class="date-range">Current: {fmt(curr_start)} to {fmt(curr_end)} &nbsp;|&nbsp; Previous: {fmt(prev_start)} to {fmt(prev_end)}</div>
"""

        if not pages:
            html += '    <div class="no-swings">✅ No major swings detected this week</div>\n'
        else:
            sorted_pages = sorted(pages,
                                  key=lambda p: abs(p.get("clicks_curr", 0) - p.get("clicks_prev", 0)),
                                  reverse=True)

            html += """    <table class="page-table">
      <tr>
        <th style="width:35%">Page</th>
        <th>Clicks</th>
        <th>Impressions</th>
        <th>CTR</th>
        <th>Position</th>
      </tr>
"""
            for page in sorted_pages[:TOP_N]:
                url = page["page"]
                html += f'      <tr>\n        <td><a class="page-url" href="{url}">{short_url(url)}</a></td>\n'

                for label, ck, pk, metric in metrics:
                    curr_val = page.get(ck, 0)
                    prev_val = page.get(pk, 0)
                    pct = pct_change(prev_val, curr_val)
                    pct_str = fmt_pct(pct)
                    good = is_positive(pct, metric)
                    css = "up" if good else ("down" if good is False else "neutral")
                    arrow = "▲" if (pct or 0) > 0 else ("▼" if (pct or 0) < 0 else "–")
                    html += (
                        f'        <td>'
                        f'{fmt_val(prev_val, metric)} → <strong>{fmt_val(curr_val, metric)}</strong><br>'
                        f'<span class="{css}">{arrow} {pct_str}</span>'
                        f'</td>\n'
                    )

                html += "      </tr>\n"

            html += "    </table>\n"

            if len(pages) > TOP_N:
                remaining = len(pages) - TOP_N
                html += f'    <div class="more">...and {remaining} more pages with swings</div>\n'

        html += "  </div>\n"

    html += """  <div class="footer">
    Generated automatically by GSC Weekly Swing Reporter &nbsp;•&nbsp; Powered by Anthropic + Google Search Console
  </div>
</div>
</body>
</html>
"""
    return html

# ─────────────────────────────────────────────
# EMAIL SENDER
# ─────────────────────────────────────────────
def send_email(html, curr_start, curr_end):
    subject = f"📊 GSC Weekly Swing Report — {fmt_display(curr_start)} to {fmt_display(curr_end)}"

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = EMAIL_ADDRESS
    msg["To"]      = EMAIL_ADDRESS
    msg.attach(MIMEText(html, "html"))

    print(f"Connecting to {SMTP_SERVER}:{SMTP_PORT}...")
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        server.sendmail(EMAIL_ADDRESS, EMAIL_ADDRESS, msg.as_string())

    print(f"✅ Email sent to {EMAIL_ADDRESS}")

# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
def main():
    print("🔍 GSC Weekly Swing Reporter starting...\n")

    curr_start, curr_end, prev_start, prev_end = get_week_ranges()
    print(f"📅 Current week:  {fmt(curr_start)} → {fmt(curr_end)}")
    print(f"📅 Previous week: {fmt(prev_start)} → {fmt(prev_end)}\n")

    all_swings = {}
    for site_url in PROPERTIES:
        print(f"Fetching: {site_url}")
        pages = fetch_swing_data(site_url, curr_start, curr_end, prev_start, prev_end)
        all_swings[site_url] = pages
        print(f"  → {len(pages)} pages with major swings")

    print("\n📧 Building email...")
    html = build_html(all_swings, curr_start, curr_end, prev_start, prev_end)

    print("📨 Sending email...")
    send_email(html, curr_start, curr_end)
    print("\nDone.")

if __name__ == "__main__":
    main()
