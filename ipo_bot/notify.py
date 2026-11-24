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
    smtp_user = os.environ.get("EMAIL_FROM")
    smtp_password = os.environ.get("EMAIL_PASSWORD")
    email_to_raw = os.environ.get("EMAIL_TO")
    if not (smtp_user and smtp_password and email_to_raw):
        print("Skipping email alert: EMAIL_FROM, EMAIL_PASSWORD, or EMAIL_TO not configured in secrets.", file=sys.stderr)
        return

    smtp_host = os.environ.get("SMTP_HOST", "smtp.gmail.com")
    smtp_port = int(os.environ.get("SMTP_PORT", "587"))
    to_addrs = [a.strip() for a in email_to_raw.split(",") if a.strip()]
    if not to_addrs:
        return

    msg = MIMEText(body)
    msg["Subject"] = subject
    msg["From"] = smtp_user
    msg["To"] = ", ".join(to_addrs)

    try:
        with smtplib.SMTP(smtp_host, smtp_port) as server:
            server.starttls()
            server.login(smtp_user, smtp_password)
            server.sendmail(smtp_user, to_addrs, msg.as_string())
        print(f"Sent email: {subject} -> {to_addrs}")
    except Exception as exc:
        print(f"Failed to send email alert: {exc}", file=sys.stderr)


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
