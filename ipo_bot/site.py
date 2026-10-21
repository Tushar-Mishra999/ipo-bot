"""Renders the static IPO tracker website into docs/index.html."""

import html
from datetime import date

PAGE_TEMPLATE = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Mainboard IPO Tracker</title>
<style>
  :root {{ color-scheme: light dark; }}
  body {{ font-family: -apple-system, Segoe UI, Roboto, sans-serif; max-width: 960px; margin: 0 auto; padding: 24px 16px; background: #0f1115; color: #e8e8e8; }}
  h1 {{ font-size: 1.5rem; margin-bottom: 4px; }}
  .updated {{ color: #9aa0a6; font-size: 0.85rem; margin-bottom: 24px; }}
  .section-title {{ font-size: 1.1rem; margin: 32px 0 12px; border-bottom: 1px solid #333; padding-bottom: 6px; }}
  .grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(300px, 1fr)); gap: 16px; }}
  .card {{ background: #1a1d24; border: 1px solid #2a2d35; border-radius: 10px; padding: 16px; }}
  .card h3 {{ margin: 0 0 8px; font-size: 1.05rem; }}
  .badge {{ display: inline-block; font-size: 0.7rem; padding: 2px 8px; border-radius: 999px; margin-left: 6px; }}
  .badge.open {{ background: #1b4332; color: #74c69d; }}
  .badge.upcoming {{ background: #3a3d1f; color: #d4d84a; }}
  .badge.closed {{ background: #3a2323; color: #e07a7a; }}
  .row {{ display: flex; justify-content: space-between; font-size: 0.85rem; padding: 3px 0; border-bottom: 1px dashed #2a2d35; }}
  .gmp {{ font-weight: 600; color: #74c69d; }}
  .gmp.neg {{ color: #e07a7a; }}
  details {{ margin-top: 10px; }}
  summary {{ cursor: pointer; font-size: 0.85rem; color: #9aa0a6; }}
  .metrics-table {{ width: 100%; margin-top: 8px; font-size: 0.8rem; border-collapse: collapse; }}
  .metrics-table td {{ padding: 3px 4px; border-bottom: 1px solid #22252c; }}
  footer {{ margin-top: 40px; color: #666; font-size: 0.75rem; }}
</style>
</head>
<body>
  <h1>Mainboard IPO Tracker</h1>
  <div class="updated">Last updated: {updated}</div>
  {sections}
  <footer>Data: investorgain.com (GMP) &amp; chittorgarh.com (financials). Mainboard IPOs only.</footer>
</body>
</html>
"""


def _badge(status):
    s = (status or "").upper()
    if s.startswith("O") or s == "CT":
        return '<span class="badge open">OPEN</span>'
    if s.startswith("U"):
        return '<span class="badge upcoming">UPCOMING</span>'
    return '<span class="badge closed">CLOSED</span>'


def _fmt_date(d):
    return d.strftime("%d %b") if d else "-"


def _card(record):
    name = html.escape(record["name"])
    gmp_value = record.get("gmp_value")
    gmp_pct = record.get("gmp_pct")
    gmp_html = ""
    if gmp_value is not None:
        cls = "gmp neg" if gmp_value < 0 else "gmp"
        gmp_html = f'<span class="{cls}">₹{gmp_value:g} ({gmp_pct:g}%)</span>'
    else:
        gmp_html = '<span class="gmp">-</span>'

    metrics_rows = "".join(
        f"<tr><td>{html.escape(k)}</td><td>{html.escape(str(v))}</td></tr>"
        for k, v in record.get("metrics", [])
    )
    metrics_block = (
        f'<details><summary>Financials &amp; KPIs ({len(record.get("metrics", []))})</summary>'
        f'<table class="metrics-table">{metrics_rows}</table></details>'
        if record.get("metrics") else ""
    )

    return f"""<div class="card">
      <h3>{name}{_badge(record.get('status'))}</h3>
      <div class="row"><span>GMP</span>{gmp_html}</div>
      <div class="row"><span>Price</span><span>{html.escape(record.get('price') or '-')}</span></div>
      <div class="row"><span>Issue size</span><span>{html.escape(record.get('issue_size') or '-')}</span></div>
      <div class="row"><span>Open - Close</span><span>{_fmt_date(record.get('open_date'))} - {_fmt_date(record.get('close_date'))}</span></div>
      <div class="row"><span>Listing</span><span>{_fmt_date(record.get('listing_date'))}</span></div>
      {metrics_block}
    </div>"""


def render_site(records):
    def sort_key(r):
        return r.get("open_date") or date.max

    open_now = [r for r in records if (r.get("status") or "").upper().startswith(("O", "CT"))]
    upcoming = [r for r in records if (r.get("status") or "").upper().startswith("U")]
    recent = [r for r in records if r not in open_now and r not in upcoming]

    sections = []
    for title, group in [("Open now", open_now), ("Upcoming", upcoming), ("Recently closed / listed", recent)]:
        group = sorted(group, key=sort_key)
        if not group:
            continue
        cards = "\n".join(_card(r) for r in group)
        sections.append(f'<div class="section-title">{title}</div><div class="grid">{cards}</div>')

    return PAGE_TEMPLATE.format(
        updated=date.today().strftime("%d %B %Y"),
        sections="\n".join(sections) if sections else "<p>No mainboard IPO data right now.</p>",
    )
