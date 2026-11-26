"""
Daily IPO tracker run:
  1. Scrape mainboard IPO list + GMP (investorgain.com, via Playwright).
  2. Scrape ~10 financial/KPI metrics per company (chittorgarh.com).
  3. Render docs/index.html (static site, published via GitHub Pages).
  4. Email an alert for any IPO whose subscription opens or closes today.
"""

import os
import sys
from datetime import date

from ipo_bot.scrape import (
    scrape_gmp_list,
    scrape_chittorgarh_links,
    match_chittorgarh_url,
    scrape_company_financials,
)
from ipo_bot.site import render_site
from ipo_bot.notify import send_email, describe_record_for_email

SITE_URL = os.environ.get("SITE_URL", "https://tushar-mishra999.github.io/ipo-bot/")


def build_records():
    gmp_rows = scrape_gmp_list()
    mainboard = [r for r in gmp_rows if r["is_mainboard"]]
    print(f"Found {len(mainboard)} mainboard IPOs (of {len(gmp_rows)} total).")

    chittorgarh_map = scrape_chittorgarh_links()

    records = []
    for row in mainboard:
        url = match_chittorgarh_url(row["name"], chittorgarh_map)
        metrics = []
        if url:
            try:
                fin = scrape_company_financials(url)
                metrics = fin["metrics"]
            except Exception as exc:
                print(f"Could not fetch financials for {row['name']}: {exc}", file=sys.stderr)
        else:
            print(f"No chittorgarh match for {row['name']}", file=sys.stderr)

        record = dict(row)
        record["metrics"] = metrics
        record["detail_url"] = url
        records.append(record)
    return records


def send_trigger_emails(records, today):
    for record in records:
        is_open_day = record.get("open_date") == today
        is_close_day = record.get("close_date") == today
        if not (is_open_day or is_close_day):
            continue

        if is_open_day and is_close_day:
            event = "opens and closes today"
        elif is_open_day:
            event = "subscription opens today"
        else:
            event = "subscription closes today"

        subject = f"IPO Alert: {record['name']} - {event}"
        body = describe_record_for_email(record, SITE_URL)
        send_email(subject, body)


def main():
    records = build_records()

    os.makedirs("docs", exist_ok=True)
    html_out = render_site(records)
    with open("docs/index.html", "w", encoding="utf-8") as f:
        f.write(html_out)
    print("Wrote docs/index.html")

    send_trigger_emails(records, date.today())


if __name__ == "__main__":
    main()
