"""Email sending for IPO open/close day alerts."""

import os
import sys
import smtplib
from email.mime.text import MIMEText


def get_env(name, required=True, default=None):
    value = os.environ.get(name, default)
    if required and not value:
        print(f"Missing required environment variable: {name}", file=sys.stderr)
        sys.exit(1)
    return value


def send_email(subject, body):
    smtp_host = get_env("SMTP_HOST", required=False, default="smtp.gmail.com")
    smtp_port = int(get_env("SMTP_PORT", required=False, default="587"))
    smtp_user = get_env("EMAIL_FROM")
    smtp_password = get_env("EMAIL_PASSWORD")
    to_addrs = [a.strip() for a in get_env("EMAIL_TO").split(",") if a.strip()]

    msg = MIMEText(body)
    msg["Subject"] = subject
    msg["From"] = smtp_user
    msg["To"] = ", ".join(to_addrs)

    with smtplib.SMTP(smtp_host, smtp_port) as server:
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.sendmail(smtp_user, to_addrs, msg.as_string())
    print(f"Sent email: {subject} -> {to_addrs}")


def describe_record_for_email(record, site_url):
    lines = [f"{record['name']}", ""]
    if record.get("gmp_value") is not None:
        lines.append(f"GMP: ₹{record['gmp_value']:g} ({record['gmp_pct']:g}%)")
    else:
        lines.append("GMP: not available yet")
    lines.append(f"Price band: {record.get('price') or '-'}")
    lines.append(f"Issue size: {record.get('issue_size') or '-'}")
    lines.append(f"Subscription: {record.get('open_date')} to {record.get('close_date')}")
    lines.append(f"Listing date: {record.get('listing_date') or '-'}")
    lines.append("")
    lines.append("Key financials / KPIs:")
    for k, v in record.get("metrics", []):
        lines.append(f" - {k}: {v}")
    lines.append("")
    lines.append(f"Full tracker: {site_url}")
    return "\n".join(lines)
