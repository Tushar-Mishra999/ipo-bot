"""Best-effort market snapshots for the daily static dashboard."""

import re
from datetime import datetime
from zoneinfo import ZoneInfo

import requests
from bs4 import BeautifulSoup

HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; ipo-bot/1.0)"}


def _index_quote(symbol, label):
    url = f"https://www.google.com/finance/quote/{symbol}?hl=en"
    response = requests.get(url, headers=HEADERS, timeout=15)
    response.raise_for_status()
    text = BeautifulSoup(response.text, "html.parser").get_text(" ", strip=True)
    match = re.search(rf"{re.escape(label)}\s+([\d,]+\.\d{{2}}).*?\b([A-Z][a-z]{{2}} \d{{1,2}}, \d{{1,2}}:\d{{2}}:\d{{2}}\s*[AP]M GMT\+5:30)", text)
    if not match:
        raise ValueError(f"Could not read {label} quote")
    return {"value": float(match.group(1).replace(",", "")), "as_of": match.group(2), "source": "Google Finance", "url": url}


def _bullion_quotes():
    url = "https://ibjarates.com/"
    response = requests.get(url, headers=HEADERS, timeout=15)
    response.raise_for_status()
    table = BeautifulSoup(response.text, "html.parser").find("table", id="TodayRatesTableDataYes")
    if table is None:
        raise ValueError("Could not find IBJA rates")
    found = {}
    for row in table.find_all("tr"):
        cells = [cell.get_text(" ", strip=True) for cell in row.find_all(["th", "td"])]
        if not cells:
            continue
        for name, key, unit in (("Gold 999", "gold", "₹/10 g"), ("Silver 999", "silver", "₹/kg")):
            if cells[0] != name:
                continue
            values = [float(value.replace(",", "")) for value in cells[1:3] if re.fullmatch(r"[\d,]+(?:\.\d+)?", value)]
            if values:
                found[key] = {"value": values[-1], "unit": unit, "session": "PM" if len(values) > 1 else "AM", "source": "IBJA", "url": url}
    if len(found) != 2:
        raise ValueError("Could not read both IBJA rates")
    return found


def market_snapshot():
    snapshot = {"fetched_at": datetime.now(ZoneInfo("Asia/Kolkata")).isoformat(timespec="minutes"), "quotes": {}}
    for key, symbol, label in (("nifty", "NIFTY_50:INDEXNSE", "NIFTY 50"), ("sensex", "SENSEX:INDEXBOM", "BSE SENSEX")):
        try:
            snapshot["quotes"][key] = _index_quote(symbol, label)
        except (requests.RequestException, ValueError) as exc:
            print(f"Market quote unavailable for {key}: {exc}")
    try:
        snapshot["quotes"].update(_bullion_quotes())
    except (requests.RequestException, ValueError) as exc:
        print(f"Bullion quotes unavailable: {exc}")
    return snapshot
