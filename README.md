# Mainboard IPO Tracker

A daily-refreshing website of all open/upcoming/recently-closed **mainboard**
IPOs (SME IPOs excluded), each with GMP (grey market premium) and ~10 company
financial/KPI metrics, plus an email alert on the day a company's IPO
subscription **opens** and again the day it **closes**.

Runs entirely on **GitHub Actions' free scheduled workflows + GitHub Pages** —
no server needed.

## How it works

1. Once a day, a GitHub Actions workflow scrapes:
   - [investorgain.com](https://www.investorgain.com/report/live-ipo-gmp/331/ipo/) for the mainboard IPO list, GMP, and open/close/listing dates.
   - [chittorgarh.com](https://www.chittorgarh.com/ipo/) for each company's financial statements and KPIs (RoNW, PAT margin, EBITDA margin, NAV, P/B, revenue, profit, net worth, etc.)
2. It builds a static website (`docs/index.html`) and publishes it to **GitHub Pages**.
3. For any IPO whose subscription opens or closes *today*, it emails a summary (GMP + financials) to the configured recipients.

## Setup (one-time)

### 1. Repository secrets
Settings → Secrets and variables → Actions → New repository secret:

| Secret | Value |
|---|---|
| `EMAIL_FROM` | sending Gmail address |
| `EMAIL_PASSWORD` | Gmail app password |
| `EMAIL_TO` | comma-separated recipient list, e.g. `a@x.com,b@y.com` |

(Optional: `SMTP_HOST`, `SMTP_PORT` if not using Gmail.)

### 2. Repository variable (optional)
Settings → Secrets and variables → Actions → Variables tab:
- `SITE_URL` — the published Pages URL, used as the link in alert emails. Defaults to `https://<owner>.github.io/<repo>/` if not set.

### 3. Enable GitHub Pages
Settings → Pages → Build and deployment → Source → **GitHub Actions**.

### 4. Run it
Actions tab → "IPO Tracker" → Run workflow (or wait for the daily 03:00 UTC / 08:30 AM IST run).

## Notes
- "Small companies" are excluded by only scraping the mainboard IPO list — SME/NSE Emerge/BSE SME platform IPOs are never included.
- GMP scraping uses a headless browser (Playwright/Chromium) since investorgain.com renders its table client-side; this adds ~30-60s per run, still well within GitHub's free Actions minutes for one run a day.
- Scraping depends on the structure of two third-party sites; if either changes its layout, the relevant scraper in `ipo_bot/scrape.py` may need updating.
