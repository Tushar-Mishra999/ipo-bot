# IPO Tracker

A daily-refreshing website of open/upcoming/recently-closed **mainboard and SME**
IPOs, each with GMP (grey market premium) and ~10 company
financial/KPI metrics, plus an email alert on the day a company's IPO
subscription **opens** and again the day it **closes**.

Runs entirely on **GitHub Actions' free scheduled workflows + GitHub Pages** —
no server needed.

## How it works

1. Once a day, a GitHub Actions workflow scrapes:
   - [investorgain.com](https://www.investorgain.com/report/live-ipo-gmp/331/ipo/) for mainboard and SME IPO lists, GMP, and open/close/listing dates.
   - [chittorgarh.com](https://www.chittorgarh.com/ipo/) for each company's financial statements and KPIs (RoNW, PAT margin, EBITDA margin, NAV, P/B, revenue, profit, net worth, etc.)
2. It builds a static website (`docs/index.html`) and publishes it to **GitHub Pages**.
   The build also fetches Nifty 50 and Sensex snapshots from Google Finance and
   Gold 999 (₹/10 g) and Silver 999 (₹/kg) benchmark rates from IBJA. The market
   strip shows the source and fetch time; quotes can be delayed, and missing
   quotes appear as unavailable rather than reusing old values.
3. For any IPO whose subscription opens or closes *today*, it emails a summary (GMP + financials) to the configured recipients.

The dashboard includes a **Refresh data** control. The Render service starts the
GitHub workflow and the browser monitors its progress. Visitors do not need a GitHub
token. Set `GITHUB_REFRESH_TOKEN` once in the Render service environment to a
fine-grained GitHub token limited to this repository with Actions read/write access.
Keep it out of the browser and repository. Manual refreshes are limited to one start
every ten minutes and do not send duplicate IPO alert emails. The daily schedule
continues to run at 03:00 UTC.

For an open IPO, choose **Analyze** and upload its Screener.in Excel export. The browser
reads the `Data Sheet` tab and calculates five pillar scores, data coverage, a composite
score, and a verdict. Missing metrics are skipped. The uploaded workbook is processed
locally in the browser. The complete analysis appears after the AI explanation returns.
Scoring thresholds and weights are in
`ipo_bot/analyzer_config.json`. The info tooltips are defined in
`ipo_bot/metric_info.json`; their numeric Good and Watch levels are read from
the scoring config at display time.
Holding horizon, safety priority, and income need sliders adjust only the
composite weights and the extra safety cap; pillar scores and metric values stay
fixed. Short-horizon views show available GMP and the subscription window as
context, outside the score. The IPO feed does not provide subscription multiples.

The optional AI explanation service lives in `ai-service/`. It uses the same
`@google/genai` Vertex AI pattern as `project-s` and needs
`GOOGLE_CLOUD_PROJECT` plus Cloud Run's attached service account. For local
development, it can use `GOOGLE_APPLICATION_CREDENTIALS_JSON`. Set the
dashboard workflow variable `ANALYSIS_API_URL` to its deployed HTTPS URL.
The service-account key must stay out of the repository and browser code.

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
- SME IPOs are included and labeled in the dashboard. Chittorgarh's mainboard listing may not provide matching financial KPIs for every SME; those fields remain empty when unavailable.
- GMP scraping uses a headless browser (Playwright/Chromium) since investorgain.com renders its table client-side; this adds ~30-60s per run, still well within GitHub's free Actions minutes for one run a day.
- Scraping depends on the structure of two third-party sites; if either changes its layout, the relevant scraper in `ipo_bot/scrape.py` may need updating.
