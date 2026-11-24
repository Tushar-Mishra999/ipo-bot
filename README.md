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
4. **Interactive 5-Pillar Fundamental Analyzer**: Each "Open now" card includes an **Analyze** upload button accepting a Screener.in-formatted `.xlsx` model. The browser computes 5 pillar health scores, forensic risk, override caps, radar/bar charts, and rule-based red flags & strengths in real time.

## 5-Pillar Fundamental Analyzer

The web dashboard and Python engine (`ipo_bot/analyzer.py`) evaluate companies across 5 core dimensions (0 to 100):

| Pillar | Weight | Key Metrics & Benchmarks (Bad &rarr; Good) |
|---|---|---|
| **1. Earnings Quality** | 25% | Cumulative CFO/PAT (0.5 &rarr; 1.0), Accruals Ratio (+0.10 &rarr; -0.05) |
| **2. Forensic Risk** | 20% | Beneish M-Score (-1.78 &rarr; -2.50), 2-Year Debtor Days Change (+15 &rarr; 0 days), Other Income % of PBT (30% &rarr; 5%) |
| **3. Financial Strength** | 20% | Altman Z'' Score (1.1 &rarr; 2.6), Interest Coverage (1.5x &rarr; 5.0x), Debt to Equity (1.5 &rarr; 0.3) |
| **4. Profitability & Growth** | 20% | Sales CAGR (5% &rarr; 20%), PAT CAGR (5% &rarr; 20%), ROE (10% &rarr; 20%), EBITDA Margin Trend (-2% &rarr; +2%) |
| **5. Cash & Capital Allocation** | 15% | FCF Proxy relative to PAT (0.20 &rarr; 0.80), Dividend Payout vs CFO (80% &rarr; 20%), Cash vs Borrowings (0.20 &rarr; 1.0) |

- **Linear Interpolation:** Every metric is mapped to 0–100 between bad and good thresholds defined in [`ipo_bot/analyzer_config.json`](file:///C:/Users/Tushar%20Mishra/Desktop/ipo-bot/ipo_bot/analyzer_config.json).
- **Composite Score:** `0.25*EQ + 0.20*FR + 0.20*FS + 0.20*PG + 0.15*CA` (re-normalized if data is missing, with coverage % shown).
- **Rule-Based Verdicts:**
  - $\ge 75$: **Invest**
  - $55$ to $74$: **Watchlist**
  - $< 55$: **Avoid**
- **Forensic & Solvency Override Caps:**
  - Beneish M-Score $> -1.78$
  - Altman Z'' $< 1.1$
  - Cumulative CFO/PAT $< 0.5$
  - *1 cap active*: Limits verdict to **Watchlist** (even if composite $\ge 75$).
  - *2+ caps active*: Forces verdict to **Avoid**.
- **Visuals & Sentiment:** Displays composite score, SVG radar chart, bar chart with coverage %, top 3 strengths, top 3 red flags, and displays GMP separately as a market sentiment signal.

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
