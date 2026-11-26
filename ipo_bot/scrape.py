"""
Scraping for mainboard IPOs: list + GMP from investorgain.com (JS-rendered,
needs Playwright), and per-company financials from chittorgarh.com (plain
server-rendered HTML, no browser needed).
"""

import re
import difflib
from datetime import date, datetime

import requests
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright

GMP_URL = "https://www.investorgain.com/report/live-ipo-gmp/331/ipo/"
CHITTORGARH_LIST_URL = "https://www.chittorgarh.com/ipo/"
HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; ipo-bot/1.0)"}

STATUS_CODES = ["CT", "LT", "U", "O", "C", "L"]  # longest first for matching
MONTHS = {
    "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6,
    "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12,
}


def _split_name_status(raw):
    """'Jio Platforms IPOO' -> ('Jio Platforms', 'IPO', 'O')"""
    for marker in ("BSE SME", "NSE SME", "IPO"):
        if marker in raw:
            idx = raw.index(marker)
            name = raw[:idx].strip()
            rest = raw[idx + len(marker):].strip()
            for code in STATUS_CODES:
                if rest == code:
                    return name, marker, code
            return name, marker, rest or None
    return raw.strip(), None, None


def _parse_money(text):
    """'₹30 (20.00%)' -> (30.0, 20.0); '₹-- (0.00%)' -> (None, 0.0)"""
    m = re.search(r"₹([\d,.\-]+|--)\s*\(([\-\d.]+)%\)", text)
    if not m:
        return None, None
    raw_val, pct = m.group(1), float(m.group(2))
    if raw_val == "--":
        return None, pct
    return float(raw_val.replace(",", "")), pct


def _parse_date_cell(text, today):
    """'27-Oct\\nGMP: 52' -> date(...). Infers year from `today`."""
    first_line = text.split("\n")[0].strip()
    m = re.match(r"(\d{1,2})-([A-Za-z]{3})", first_line)
    if not m:
        return None
    day, mon_str = int(m.group(1)), m.group(2)
    month = MONTHS.get(mon_str)
    if not month:
        return None
    year = today.year
    candidate = date(year, month, day)
    # IPO list is always near-term; if this date looks >6 months in the
    # past, it's probably next year (year-boundary case), and vice versa.
    if (today - candidate).days > 180:
        candidate = date(year + 1, month, day)
    elif (candidate - today).days > 180:
        candidate = date(year - 1, month, day)
    return candidate


def scrape_gmp_list(today=None):
    """Returns list of dicts for ALL IPOs (mainboard + SME) from investorgain."""
    today = today or date.today()
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(user_agent=HEADERS["User-Agent"])
        page.goto(GMP_URL, wait_until="networkidle", timeout=60000)
        page.wait_for_selector("table tbody tr", timeout=30000)
        rows = page.query_selector_all("table tbody tr")
        raw_rows = []
        for row in rows:
            cells = [c.inner_text() for c in row.query_selector_all("td")]
            if len(cells) >= 11:
                raw_rows.append(cells)
        browser.close()

    results = []
    for cells in raw_rows:
        name, marker, status = _split_name_status(cells[0])
        gmp_value, gmp_pct = _parse_money(cells[1])
        results.append({
            "name": name,
            "is_mainboard": marker == "IPO",
            "status": status,
            "gmp_value": gmp_value,
            "gmp_pct": gmp_pct,
            "price": cells[4],
            "issue_size": cells[5],
            "lot": cells[6],
            "open_date": _parse_date_cell(cells[7], today),
            "close_date": _parse_date_cell(cells[8], today),
            "listing_date": _parse_date_cell(cells[10], today),
        })
    return results


def scrape_chittorgarh_links():
    """Returns {normalized_name: {'name': ..., 'url': ...}} for mainboard IPOs."""
    resp = requests.get(CHITTORGARH_LIST_URL, headers=HEADERS, timeout=20)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")

    mapping = {}
    for heading in soup.find_all(["h2", "h3"]):
        if "Mainboard IPO" not in heading.get_text():
            continue
        table = heading.find_next("table")
        if not table:
            continue
        for a in table.find_all("a", href=True):
            name = a.get("title", a.get_text(strip=True)).strip()
            if not name:
                continue
            url = a["href"]
            if url.startswith("/"):
                url = "https://www.chittorgarh.com" + url
            mapping[_normalize(name)] = {"name": name, "url": url}
        break
    return mapping


def _normalize(name):
    return re.sub(r"[^a-z0-9]", "", name.lower())


def match_chittorgarh_url(name, mapping):
    norm = _normalize(name)
    if norm in mapping:
        return mapping[norm]["url"]
    for key, val in mapping.items():
        if norm in key or key in norm:
            return val["url"]
    close = difflib.get_close_matches(norm, mapping.keys(), n=1, cutoff=0.6)
    if close:
        return mapping[close[0]]["url"]
    return None


def scrape_company_financials(url):
    """Pulls ~10 financial/KPI metrics from a chittorgarh.com IPO page."""
    resp = requests.get(url, headers=HEADERS, timeout=20)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")

    def find_table(keyword):
        for heading in soup.find_all(["h2", "h3"]):
            if keyword.lower() in heading.get_text(strip=True).lower():
                return heading.find_next("table")
        return None

    def table_rows(table, max_cols=2):
        if not table:
            return []
        out = []
        for tr in table.find_all("tr"):
            cells = [c.get_text(strip=True) for c in tr.find_all(["td", "th"])]
            if len(cells) >= 2:
                out.append((cells[0], cells[1]))  # label, most-recent-period value
        return out

    metrics = []

    kpi_table = find_table("Key Performance Indicator")
    for label, value in table_rows(kpi_table):
        if label and label.upper() != "KPI":
            metrics.append((label, value))

    fin_table = find_table("Company Financials")
    for label, value in table_rows(fin_table):
        if label and label.lower() != "period ended":
            metrics.append((f"{label} (₹ Lakh)", value))

    details_table = soup.find("table")
    details = {}
    for label, value in table_rows(details_table):
        details[label] = value

    return {"metrics": metrics[:12], "details": details}
