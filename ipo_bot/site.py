"""Renders the static IPO tracker website with interactive Fundamental Analysis into docs/index.html."""

import html
import json
import os
from datetime import date

CONFIG_PATH = os.path.join(os.path.dirname(__file__), "analyzer_config.json")


def _load_analyzer_config_json() -> str:
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return f.read()
    return "{}"


PAGE_TEMPLATE = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Mainboard IPO Tracker &amp; Fundamental Analyzer</title>
<!-- SheetJS for client-side Excel parsing -->
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<style>
  :root {
    color-scheme: light dark;
    --bg-main: #0f1115;
    --card-bg: #1a1d24;
    --border-color: #2a2d35;
    --text-primary: #e8e8e8;
    --text-muted: #9aa0a6;
    --accent-green: #2ecc71;
    --accent-red: #e74c3c;
    --accent-yellow: #f1c40f;
    --accent-blue: #3498db;
    --btn-bg: #222733;
    --btn-hover: #2d3446;
  }
  * { box-sizing: border-box; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, Segoe UI, Roboto, Helvetica, Arial, sans-serif;
    max-width: 1040px;
    margin: 0 auto;
    padding: 24px 16px;
    background: var(--bg-main);
    color: var(--text-primary);
    line-height: 1.5;
  }
  header { margin-bottom: 24px; }
  h1 { font-size: 1.6rem; margin: 0 0 4px; display: flex; align-items: center; gap: 10px; }
  .tagline { color: var(--text-muted); font-size: 0.9rem; margin-bottom: 8px; }
  .updated { color: #7a828e; font-size: 0.8rem; }
  .section-title {
    font-size: 1.15rem;
    margin: 36px 0 14px;
    border-bottom: 1px solid var(--border-color);
    padding-bottom: 8px;
    display: flex;
    justify-content: space-between;
    align-items: baseline;
  }
  .section-badge { font-size: 0.75rem; color: var(--text-muted); font-weight: normal; }
  .grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(320px, 1fr)); gap: 18px; }
  .card {
    background: var(--card-bg);
    border: 1px solid var(--border-color);
    border-radius: 12px;
    padding: 18px;
    display: flex;
    flex-direction: column;
    justify-content: space-between;
    transition: transform 0.15s ease, border-color 0.15s ease;
  }
  .card:hover { border-color: #3e4452; transform: translateY(-2px); }
  .card h3 { margin: 0 0 10px; font-size: 1.1rem; display: flex; justify-content: space-between; align-items: center; }
  .badge { display: inline-block; font-size: 0.7rem; padding: 2px 8px; border-radius: 999px; font-weight: 600; text-transform: uppercase; }
  .badge.open { background: #133a28; color: #52b788; border: 1px solid #2d6a4f; }
  .badge.upcoming { background: #353818; color: #d4d84a; border: 1px solid #5a5f25; }
  .badge.closed { background: #381f21; color: #e07a7a; border: 1px solid #5c2f32; }
  
  .card-rows { margin-bottom: 12px; }
  .row { display: flex; justify-content: space-between; font-size: 0.85rem; padding: 4px 0; border-bottom: 1px dashed #242832; }
  .gmp { font-weight: 600; color: #52b788; }
  .gmp.neg { color: #e07a7a; }

  /* Analyze Actions on Card */
  .analyze-box {
    margin-top: 14px;
    padding-top: 12px;
    border-top: 1px solid #272b36;
  }
  .upload-btn-label {
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 8px;
    background: #1c2e42;
    color: #8ac0ff;
    border: 1px solid #2c4a6b;
    border-radius: 8px;
    padding: 8px 14px;
    font-size: 0.85rem;
    font-weight: 600;
    cursor: pointer;
    transition: background 0.15s ease, border-color 0.15s ease;
    text-align: center;
    width: 100%;
  }
  .upload-btn-label:hover { background: #243c57; border-color: #4371a3; }
  .upload-input { display: none; }
  .sample-link {
    display: block;
    text-align: center;
    font-size: 0.72rem;
    color: var(--text-muted);
    margin-top: 6px;
    text-decoration: underline;
    cursor: pointer;
  }
  .sample-link:hover { color: #8ac0ff; }

  details { margin-top: 10px; }
  summary { cursor: pointer; font-size: 0.82rem; color: var(--text-muted); }
  .metrics-table { width: 100%; margin-top: 8px; font-size: 0.78rem; border-collapse: collapse; }
  .metrics-table td { padding: 3px 4px; border-bottom: 1px solid #22252c; }

  /* Modal Styling */
  .modal-overlay {
    display: none;
    position: fixed;
    top: 0; left: 0; right: 0; bottom: 0;
    background: rgba(8, 10, 14, 0.82);
    backdrop-filter: blur(4px);
    z-index: 1000;
    overflow-y: auto;
    padding: 24px 16px;
  }
  .modal-overlay.active { display: flex; justify-content: center; align-items: flex-start; }
  .modal-content {
    background: #161920;
    border: 1px solid #2d323f;
    border-radius: 14px;
    max-width: 820px;
    width: 100%;
    padding: 28px;
    box-shadow: 0 16px 40px rgba(0,0,0,0.6);
    position: relative;
    margin: 20px auto;
  }
  .modal-close {
    position: absolute;
    top: 16px;
    right: 18px;
    background: none;
    border: none;
    color: #9aa0a6;
    font-size: 1.5rem;
    cursor: pointer;
    line-height: 1;
  }
  .modal-close:hover { color: #ffffff; }

  /* Modal Header & Hero */
  .analysis-header {
    border-bottom: 1px solid #282d39;
    padding-bottom: 16px;
    margin-bottom: 20px;
  }
  .analysis-title { font-size: 1.4rem; font-weight: 700; margin: 0 0 6px; }
  .analysis-subtitle { font-size: 0.85rem; color: var(--text-muted); }
  
  .sentiment-banner {
    background: #111e2e;
    border: 1px solid #1e3857;
    border-radius: 8px;
    padding: 10px 14px;
    margin: 14px 0;
    font-size: 0.85rem;
    display: flex;
    justify-content: space-between;
    align-items: center;
  }
  .sentiment-label { color: #8bbcef; font-weight: 600; }
  .sentiment-note { font-size: 0.75rem; color: #7f99b5; }

  .score-hero {
    display: flex;
    align-items: center;
    justify-content: space-between;
    background: #1a1e27;
    border: 1px solid #2a3140;
    border-radius: 12px;
    padding: 20px 24px;
    margin-bottom: 20px;
    flex-wrap: wrap;
    gap: 16px;
  }
  .score-box { display: flex; flex-direction: column; }
  .score-val { font-size: 2.6rem; font-weight: 800; line-height: 1; }
  .score-sub { font-size: 0.8rem; color: var(--text-muted); margin-top: 4px; }
  
  .verdict-badge {
    font-size: 1.15rem;
    font-weight: 700;
    padding: 8px 18px;
    border-radius: 8px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
  }
  .verdict-invest { background: #1b4332; color: #4ade80; border: 1px solid #2d6a4f; }
  .verdict-watchlist { background: #423512; color: #facc15; border: 1px solid #715b1a; }
  .verdict-avoid { background: #431b1b; color: #f87171; border: 1px solid #752929; }

  .override-alert {
    background: #3e1f18;
    border: 1px solid #6b352b;
    border-radius: 8px;
    padding: 12px 16px;
    margin-bottom: 20px;
    font-size: 0.85rem;
    color: #ff9d8c;
    display: flex;
    gap: 10px;
    align-items: flex-start;
  }

  /* Charts Layout */
  .charts-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 20px;
    margin-bottom: 24px;
  }
  @media (max-width: 700px) {
    .charts-grid { grid-template-columns: 1fr; }
  }
  .chart-card {
    background: #14171f;
    border: 1px solid #232733;
    border-radius: 10px;
    padding: 16px;
  }
  .chart-title { font-size: 0.95rem; font-weight: 600; margin-bottom: 12px; color: #c4cad4; }

  /* Bar Chart */
  .bar-item { margin-bottom: 12px; }
  .bar-header { display: flex; justify-content: space-between; font-size: 0.8rem; margin-bottom: 4px; }
  .bar-name { font-weight: 500; }
  .bar-score { font-weight: 700; }
  .bar-cov { font-size: 0.72rem; color: var(--text-muted); margin-left: 6px; }
  .bar-track { height: 10px; background: #232733; border-radius: 5px; overflow: hidden; }
  .bar-fill { height: 100%; border-radius: 5px; transition: width 0.4s ease; }

  /* Radar Chart SVG */
  .radar-container { display: flex; justify-content: center; align-items: center; min-height: 220px; }

  /* Insights / Red Flags */
  .insights-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 16px;
    margin-bottom: 24px;
  }
  @media (max-width: 700px) {
    .insights-grid { grid-template-columns: 1fr; }
  }
  .insight-col {
    background: #14171f;
    border: 1px solid #232733;
    border-radius: 10px;
    padding: 16px;
  }
  .insight-col h4 { margin: 0 0 10px; font-size: 0.95rem; display: flex; align-items: center; gap: 6px; }
  .insight-col.strengths h4 { color: #4ade80; }
  .insight-col.redflags h4 { color: #f87171; }
  .insight-list { margin: 0; padding-left: 18px; font-size: 0.82rem; }
  .insight-list li { margin-bottom: 8px; color: #d0d4dc; }

  /* Metrics Breakdown Collapsible */
  .breakdown-details {
    background: #14171f;
    border: 1px solid #232733;
    border-radius: 10px;
    padding: 12px 16px;
  }
  .breakdown-table { width: 100%; border-collapse: collapse; font-size: 0.8rem; margin-top: 10px; }
  .breakdown-table th, .breakdown-table td { padding: 6px 8px; text-align: left; border-bottom: 1px solid #202430; }
  .breakdown-table th { color: #8e95a5; font-weight: 600; font-size: 0.75rem; text-transform: uppercase; }
  .breakdown-score-badge {
    display: inline-block;
    padding: 1px 6px;
    border-radius: 4px;
    font-weight: 600;
    font-size: 0.75rem;
  }
  
  footer { margin-top: 48px; color: #666; font-size: 0.8rem; text-align: center; }
</style>
</head>
<body>
  <header>
    <h1>📈 Mainboard IPO Tracker &amp; Fundamental Analyzer</h1>
    <div class="tagline">Daily mainboard IPO monitoring with automated 5-pillar fundamental health and forensic scoring.</div>
    <div class="updated">Last updated: __UPDATED__</div>
  </header>

  __SECTIONS__

  <!-- Analysis Modal -->
  <div id="analysisModal" class="modal-overlay">
    <div class="modal-content">
      <button class="modal-close" onclick="closeModal()">&times;</button>
      
      <div class="analysis-header">
        <h2 id="modalCompanyTitle" class="analysis-title">Company Fundamental Analysis</h2>
        <div id="modalFileSub" class="analysis-subtitle">Parsed from uploaded financial model (.xlsx)</div>
      </div>

      <!-- GMP Sentiment Callout (shown separately from score) -->
      <div class="sentiment-banner">
        <div>
          <span class="sentiment-label">GMP Sentiment Signal: </span>
          <span id="modalGmpText">-</span>
        </div>
        <div class="sentiment-note">⚠️ Note: GMP is a market sentiment signal and is not part of the fundamental score.</div>
      </div>

      <!-- Score Hero -->
      <div class="score-hero">
        <div class="score-box">
          <div class="score-val"><span id="modalCompositeScore">--</span> <span style="font-size:1.1rem; color:#8e95a5; font-weight:400;">/ 100</span></div>
          <div class="score-sub">5-Pillar Weighted Fundamental Score &bull; <span id="modalCoverage">100%</span> Coverage</div>
        </div>
        <div>
          <div id="modalVerdictBadge" class="verdict-badge verdict-watchlist">WATCHLIST</div>
        </div>
      </div>

      <!-- Override Alert (if triggered) -->
      <div id="modalOverrideAlert" class="override-alert" style="display:none;">
        <span style="font-size:1.2rem;">⚠️</span>
        <div id="modalOverrideText"></div>
      </div>

      <!-- Visual Charts: Bar and Radar -->
      <div class="charts-grid">
        <div class="chart-card">
          <div class="chart-title">5 Pillar Scores</div>
          <div id="barChartContainer"></div>
        </div>
        <div class="chart-card">
          <div class="chart-title">Pillar Balance (Radar)</div>
          <div class="radar-container" id="radarChartContainer"></div>
        </div>
      </div>

      <!-- Top Strengths & Red Flags -->
      <div class="insights-grid">
        <div class="insight-col strengths">
          <h4>✅ Top 3 Strengths</h4>
          <ul id="modalStrengthsList" class="insight-list"></ul>
        </div>
        <div class="insight-col redflags">
          <h4>🚩 Top 3 Red Flags</h4>
          <ul id="modalRedFlagsList" class="insight-list"></ul>
        </div>
      </div>

      <!-- Full Breakdown Details -->
      <details class="breakdown-details">
        <summary><strong>View All 15 Metric Scores &amp; Thresholds</strong></summary>
        <table class="breakdown-table">
          <thead>
            <tr>
              <th>Pillar / Metric</th>
              <th>Company Value</th>
              <th>Bad &rarr; Good Threshold</th>
              <th>Score (0-100)</th>
            </tr>
          </thead>
          <tbody id="modalBreakdownBody"></tbody>
        </table>
      </details>

    </div>
  </div>

  <footer>
    IPO Data: investorgain.com (GMP) &amp; chittorgarh.com (financials) &bull; Mainboard IPOs only.<br>
    Fundamental Analysis: Screener.in standard model parser &amp; 5-pillar scoring framework.
  </footer>

  <script>
    // Embedded Analyzer Config
    const ANALYZER_CONFIG = __CONFIG_JSON__;

    function closeModal() {
      document.getElementById('analysisModal').classList.remove('active');
    }

    document.addEventListener('keydown', function(e) {
      if (e.key === 'Escape') closeModal();
    });

    // Helper: Linear Interpolation
    function interpolateScore(value, bad, good, higherIsBetter) {
      if (value === null || value === undefined || isNaN(value)) return null;
      let score;
      if (higherIsBetter) {
        if (good === bad) return value >= good ? 100 : 0;
        score = ((value - bad) / (good - bad)) * 100;
      } else {
        if (bad === good) return value <= good ? 100 : 0;
        score = ((bad - value) / (bad - good)) * 100;
      }
      return Math.max(0, Math.min(100, score));
    }

    function formatVal(val, fmt) {
      if (val === null || val === undefined || isNaN(val)) return 'N/A';
      if (fmt === 'pct') return (val * 100).toFixed(1) + '%';
      if (fmt === 'ratio') return val.toFixed(2) + 'x';
      if (fmt === 'score') return val.toFixed(2);
      if (fmt === 'days') return (val >= 0 ? '+' : '') + val.toFixed(1) + ' days';
      return val.toFixed(2);
    }

    function parseScreenerWorkbook(wb) {
      let sheet = wb.Sheets['Data Sheet'] || wb.Sheets[wb.SheetNames[0]];
      let jsonRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
      
      // Locate date row (usually row with 'Report Date')
      let dateRowIdx = -1;
      for (let r = 0; r < Math.min(jsonRows.length, 30); r++) {
        let row = jsonRows[r];
        if (row && row[0] && String(row[0]).toLowerCase().includes('report date')) {
          dateRowIdx = r;
          break;
        }
      }
      if (dateRowIdx === -1) dateRowIdx = 15; // 0-based for row 16

      let dateRow = jsonRows[dateRowIdx] || [];
      let dateCols = [];
      let years = [];
      for (let c = 1; c < dateRow.length; c++) {
        let v = dateRow[c];
        if (v !== null && v !== undefined && String(v).trim() !== '') {
          dateCols.push(c);
          years.push(String(v).split(' ')[0]);
        }
      }

      // Build row index by label
      let rowLookup = {};
      for (let r = 0; r < jsonRows.length; r++) {
        let lbl = jsonRows[r] ? jsonRows[r][0] : null;
        if (lbl) {
          let clean = String(lbl).trim().toLowerCase();
          if (!(clean in rowLookup)) rowLookup[clean] = r;
        }
      }

      function getSeries(aliases) {
        for (let a of aliases) {
          let norm = a.trim().toLowerCase();
          for (let key in rowLookup) {
            if (key.includes(norm)) {
              let r = rowLookup[key];
              let row = jsonRows[r];
              return dateCols.map(c => {
                let val = row[c];
                if (val === null || val === undefined || val === '') return null;
                let num = parseFloat(val);
                return isNaN(num) ? null : num;
              });
            }
          }
        }
        return dateCols.map(() => null);
      }

      let companyName = (jsonRows[0] && jsonRows[0][1]) ? jsonRows[0][1] : 'IPO Company';
      
      let sales = getSeries(['sales', 'revenue']);
      let raw_mat = getSeries(['raw material cost']);
      let change_inv = getSeries(['change in inventory']);
      let power_fuel = getSeries(['power and fuel']);
      let other_mfr = getSeries(['other mfr. exp', 'manufacturing']);
      let emp_cost = getSeries(['employee cost', 'employee']);
      let sga = getSeries(['selling and admin', 'administrative', 'sga']);
      let other_exp = getSeries(['other expenses']);
      let other_income = getSeries(['other income']);
      let depr = getSeries(['depreciation']);
      let interest = getSeries(['interest', 'finance costs']);
      let pbt = getSeries(['profit before tax', 'pbt']);
      let net_profit = getSeries(['net profit', 'pat']);
      let dividends = getSeries(['dividend amount', 'dividend']);

      let eq_cap = getSeries(['equity share capital', 'share capital']);
      let reserves = getSeries(['reserves']);
      let borrowings = getSeries(['borrowings', 'total debt']);
      let other_liab = getSeries(['other liabilities']);
      let tot_assets = getSeries(['total']);

      let net_block = getSeries(['net block', 'fixed assets']);
      let cwip = getSeries(['capital work in progress', 'cwip']);
      let receivables = getSeries(['receivables', 'debtors']);
      let inventory = getSeries(['inventory']);
      let cash = getSeries(['cash & bank', 'cash']);
      let cfo = getSeries(['cash from operating activ', 'cash from operations', 'cfo']);

      // Calculate OPM per period
      let opm = [];
      for (let i = 0; i < dateCols.length; i++) {
        let s = sales[i];
        if (s && s > 0) {
          let rm = raw_mat[i] || 0;
          let inv = change_inv[i] || 0;
          let o_mfr = other_mfr[i] || 0;
          let pf = power_fuel[i] || 0;
          let emp = emp_cost[i] || 0;
          let sg = sga[i] || 0;
          let oe = other_exp[i] || 0;
          let expenses = rm + pf + o_mfr + emp + sg + oe - inv;
          opm.push((s - expenses) / s);
        } else {
          opm.push(null);
        }
      }

      return {
        companyName, years, sales, raw_mat, change_inv, emp_cost, sga,
        other_income, depr, interest, pbt, net_profit, dividends,
        eq_cap, reserves, borrowings, other_liab, tot_assets,
        net_block, cwip, receivables, inventory, cash, cfo, opm
      };
    }

    function computeMetrics(fin) {
      let n = fin.years.length;
      let t = n - 1;
      let m = {};

      let sumArr = (arr) => arr.filter(v => v !== null).reduce((a, b) => a + b, 0);

      // 1. Cumulative CFO / PAT
      let cumCFO = sumArr(fin.cfo);
      let cumPAT = sumArr(fin.net_profit);
      m.cum_cfo_pat = (cumPAT && cumPAT !== 0) ? (cumCFO / cumPAT) : null;

      // 2. Accruals ratio: (PAT - CFO) / Total Assets
      let ta = fin.tot_assets[t];
      if (ta && ta > 0 && fin.net_profit[t] !== null && fin.cfo[t] !== null) {
        m.accruals_ratio = (fin.net_profit[t] - fin.cfo[t]) / ta;
      } else {
        m.accruals_ratio = null;
      }

      // 3. Beneish M-Score
      if (n >= 2) {
        let t1 = n - 2;
        let s_t = fin.sales[t], s_t1 = fin.sales[t1];
        let rec_t = fin.receivables[t], rec_t1 = fin.receivables[t1];
        let dsri = (rec_t && rec_t1 && s_t && s_t1 && s_t > 0 && s_t1 > 0)
          ? ((rec_t / s_t) / (rec_t1 / s_t1)) : 1.0;

        let cogs_t = (fin.raw_mat[t] || 0) + (fin.change_inv[t] || 0);
        let cogs_t1 = (fin.raw_mat[t1] || 0) + (fin.change_inv[t1] || 0);
        let gm_t = (s_t && s_t > 0) ? ((s_t - cogs_t) / s_t) : 1.0;
        let gm_t1 = (s_t1 && s_t1 > 0) ? ((s_t1 - cogs_t1) / s_t1) : 1.0;
        let gmi = (gm_t && gm_t !== 0) ? (gm_t1 / gm_t) : 1.0;

        let nb_t = fin.net_block[t], nb_t1 = fin.net_block[t1];
        let ta_t = fin.tot_assets[t], ta_t1 = fin.tot_assets[t1];
        let aq_t = (nb_t !== null && ta_t && ta_t > 0) ? (1.0 - nb_t / ta_t) : 0.5;
        let aq_t1 = (nb_t1 !== null && ta_t1 && ta_t1 > 0) ? (1.0 - nb_t1 / ta_t1) : 0.5;
        let aqi = (aq_t1 && aq_t1 !== 0) ? (aq_t / aq_t1) : 1.0;

        let sgi = (s_t && s_t1 && s_t1 > 0) ? (s_t / s_t1) : 1.0;

        let depr_t = fin.depr[t], depr_t1 = fin.depr[t1];
        let dr_t = (depr_t !== null && nb_t !== null && (nb_t + depr_t) > 0) ? (depr_t / (nb_t + depr_t)) : 0.1;
        let dr_t1 = (depr_t1 !== null && nb_t1 !== null && (nb_t1 + depr_t1) > 0) ? (depr_t1 / (nb_t1 + depr_t1)) : 0.1;
        let depi = (dr_t && dr_t !== 0) ? (dr_t1 / dr_t) : 1.0;

        let sg_t = (fin.emp_cost[t] || 0) + (fin.sga[t] || 0);
        let sg_t1 = (fin.emp_cost[t1] || 0) + (fin.sga[t1] || 0);
        let sgai = (s_t && s_t1 && sg_t1 && s_t > 0 && s_t1 > 0) ? ((sg_t / s_t) / (sg_t1 / s_t1)) : 1.0;

        let b_t = fin.borrowings[t], b_t1 = fin.borrowings[t1];
        let lv_t = (b_t !== null && ta_t && ta_t > 0) ? (b_t / ta_t) : 0.2;
        let lv_t1 = (b_t1 !== null && ta_t1 && ta_t1 > 0) ? (b_t1 / ta_t1) : 0.2;
        let lvgi = (lv_t1 && lv_t1 !== 0) ? (lv_t / lv_t1) : 1.0;

        let tata = (ta_t && ta_t > 0) ? (((fin.net_profit[t] || 0) - (fin.cfo[t] || 0)) / ta_t) : 0.0;

        m.beneish_m_score = -4.84 + 0.920*dsri + 0.528*gmi + 0.404*aqi + 0.892*sgi + 0.115*depi - 0.172*sgai + 4.037*tata + 0.0327*lvgi;
      } else {
        m.beneish_m_score = null;
      }

      // 4. Change in receivable days over 2 years
      let refIdx = n >= 3 ? (n - 3) : 0;
      if (n >= 2 && fin.sales[t] > 0 && fin.sales[refIdx] > 0 && fin.receivables[t] !== null && fin.receivables[refIdx] !== null) {
        let d_now = (fin.receivables[t] / fin.sales[t]) * 365.0;
        let d_past = (fin.receivables[refIdx] / fin.sales[refIdx]) * 365.0;
        m.chg_receivable_days_2y = d_now - d_past;
      } else {
        m.chg_receivable_days_2y = null;
      }

      // 5. Other income as % of PBT
      if (fin.other_income[t] !== null && fin.pbt[t] !== null) {
        m.other_income_share_pbt = fin.pbt[t] > 0 ? (fin.other_income[t] / fin.pbt[t]) : 1.0;
      } else {
        m.other_income_share_pbt = null;
      }

      // 6. Altman Z'' Score
      let eqVal = (fin.eq_cap[t] || 0) + (fin.reserves[t] || 0);
      let totLiab = (fin.borrowings[t] || 0) + (fin.other_liab[t] || 0);
      let ebit = (fin.pbt[t] || 0) + (fin.interest[t] || 0);
      let caVal = (fin.receivables[t] || 0) + (fin.inventory[t] || 0) + (fin.cash[t] || 0);
      let wc = caVal - (fin.other_liab[t] || 0);
      if (ta && ta > 0) {
        let x1 = wc / ta;
        let x2 = (fin.reserves[t] || 0) / ta;
        let x3 = ebit / ta;
        let x4 = totLiab > 0 ? (eqVal / totLiab) : 2.0;
        m.altman_z_score = 6.56*x1 + 3.26*x2 + 6.72*x3 + 1.05*x4;
      } else {
        m.altman_z_score = null;
      }

      // 7. Interest Cover
      let intVal = fin.interest[t] || 0;
      if (intVal > 0) {
        m.interest_cover = ebit / intVal;
      } else if (intVal === 0) {
        m.interest_cover = 50.0;
      } else {
        m.interest_cover = null;
      }

      // 8. Debt / Equity
      if (eqVal > 0 && fin.borrowings[t] !== null) {
        m.debt_to_equity = (fin.borrowings[t] || 0) / eqVal;
      } else {
        m.debt_to_equity = null;
      }

      // 9 & 10. Sales & PAT CAGR
      let y_cagr = n - 1;
      if (y_cagr >= 1 && fin.sales[0] > 0 && fin.sales[t] > 0) {
        m.sales_cagr = Math.pow(fin.sales[t] / fin.sales[0], 1.0 / y_cagr) - 1.0;
      } else {
        m.sales_cagr = null;
      }

      if (y_cagr >= 1 && fin.net_profit[0] > 0 && fin.net_profit[t] > 0) {
        m.pat_cagr = Math.pow(fin.net_profit[t] / fin.net_profit[0], 1.0 / y_cagr) - 1.0;
      } else if (y_cagr >= 1 && fin.net_profit[0] > 0 && fin.net_profit[t] <= 0) {
        m.pat_cagr = -0.5;
      } else {
        m.pat_cagr = null;
      }

      // 11. ROE
      if (eqVal > 0 && fin.net_profit[t] !== null) {
        m.roe = fin.net_profit[t] / eqVal;
      } else {
        m.roe = null;
      }

      // 12. EBITDA Margin Trend
      let opmArr = fin.opm;
      if (opmArr.length >= 3 && opmArr[t] !== null && opmArr[t-2] !== null) {
        m.ebitda_margin_trend = opmArr[t] - opmArr[t-2];
      } else if (opmArr.length >= 2 && opmArr[t] !== null && opmArr[0] !== null) {
        m.ebitda_margin_trend = opmArr[t] - opmArr[0];
      } else {
        m.ebitda_margin_trend = null;
      }

      // 13. FCF Proxy / PAT
      let fcfSum = 0, patSum = 0, countFcf = 0;
      for (let i = 1; i < n; i++) {
        if (fin.cfo[i] !== null && fin.net_profit[i] !== null && fin.net_block[i] !== null && fin.net_block[i-1] !== null) {
          let d_nb = fin.net_block[i] - fin.net_block[i-1];
          let d_cwip = (fin.cwip[i] || 0) - (fin.cwip[i-1] || 0);
          fcfSum += (fin.cfo[i] - (d_nb + d_cwip));
          patSum += fin.net_profit[i];
          countFcf++;
        }
      }
      m.fcf_to_pat = (countFcf > 0 && patSum > 0) ? (fcfSum / patSum) : null;

      // 14. Dividend Payout vs Cash
      let divSum = sumArr(fin.dividends);
      m.div_payout_vs_cfo = (cumCFO > 0) ? (divSum / cumCFO) : (divSum === 0 ? 0 : 1.0);

      // 15. Cash vs Borrowings
      let b_last = fin.borrowings[t];
      let cash_last = fin.cash[t] || 0;
      if (b_last !== null && b_last !== undefined) {
        m.cash_vs_borrowings = b_last > 0 ? (cash_last / b_last) : 5.0;
      } else {
        m.cash_vs_borrowings = null;
      }

      return m;
    }

    function analyzeWorkbook(wb, gmpValue, gmpPct) {
      let fin = parseScreenerWorkbook(wb);
      let metrics = computeMetrics(fin);

      let pillarResults = {};
      let strengths = [];
      let redFlags = [];

      for (let pKey in ANALYZER_CONFIG.pillars) {
        let pConf = ANALYZER_CONFIG.pillars[pKey];
        let metricConf = pConf.metrics;
        let scores = [];
        let details = {};

        for (let mKey in metricConf) {
          let conf = metricConf[mKey];
          let val = metrics[mKey];
          let s = interpolateScore(val, conf.bad, conf.good, conf.higher_is_better);
          if (s !== null) scores.push(s);

          details[mKey] = {
            label: conf.label,
            value: val,
            formatted: formatVal(val, conf.format),
            score: s !== null ? Math.round(s * 10) / 10 : null,
            bad: conf.bad,
            good: conf.good,
            format: conf.format,
            higher_is_better: conf.higher_is_better
          };

          if (val !== null && s !== null) {
            if (s >= 70 && conf.strength_template) {
              let txt = conf.strength_template.replace(/\\{value:[^\\}]+\\}(?:x| days)?/g, formatVal(val, conf.format));
              strengths.push({ score: s, text: txt });
            } else if (s <= 45 && conf.red_flag_template) {
              let txt = conf.red_flag_template.replace(/\\{value:[^\\}]+\\}(?:x| days)?/g, formatVal(val, conf.format));
              redFlags.push({ score: s, text: txt });
            }
          }
        }

        let totalM = Object.keys(metricConf).length;
        let cov = (scores.length / totalM) * 100;
        let pScore = scores.length > 0 ? (scores.reduce((a,b)=>a+b,0) / scores.length) : null;

        pillarResults[pKey] = {
          name: pConf.name,
          weight: pConf.weight,
          score: pScore !== null ? Math.round(pScore * 10) / 10 : null,
          coverage: Math.round(cov * 10) / 10,
          metrics: details
        };
      }

      // Composite Score
      let wSum = 0, wTotal = 0, covSum = 0;
      for (let pKey in pillarResults) {
        let p = pillarResults[pKey];
        covSum += p.coverage * p.weight;
        if (p.score !== null) {
          wSum += p.score * p.weight;
          wTotal += p.weight;
        }
      }
      let compositeScore = wTotal > 0 ? Math.round((wSum / wTotal) * 10) / 10 : 0;
      let overallCov = Math.round(covSum * 10) / 10;

      // Base Verdict
      let baseVerdict = 'Avoid';
      if (compositeScore >= ANALYZER_CONFIG.verdict_thresholds.invest) {
        baseVerdict = 'Invest';
      } else if (compositeScore >= ANALYZER_CONFIG.verdict_thresholds.watchlist) {
        baseVerdict = 'Watchlist';
      }

      // Overrides
      let overrides = [];
      let ov = ANALYZER_CONFIG.overrides;
      if (metrics.beneish_m_score !== null && metrics.beneish_m_score > ov.beneish_max) {
        overrides.push(`Beneish M-Score of ${metrics.beneish_m_score.toFixed(2)} exceeds ${ov.beneish_max.toFixed(2)} (earnings manipulation risk)`);
      }
      if (metrics.altman_z_score !== null && metrics.altman_z_score < ov.altman_min) {
        overrides.push(`Altman Z'' of ${metrics.altman_z_score.toFixed(2)} is below ${ov.altman_min.toFixed(2)} (financial distress risk)`);
      }
      if (metrics.cum_cfo_pat !== null && metrics.cum_cfo_pat < ov.cum_cfo_pat_min) {
        overrides.push(`Cumulative CFO/PAT of ${metrics.cum_cfo_pat.toFixed(2)}x is under ${ov.cum_cfo_pat_min.toFixed(2)}x (poor cash conversion)`);
      }

      let finalVerdict = baseVerdict;
      let overrideNote = '';
      if (overrides.length >= 2) {
        finalVerdict = 'Avoid';
        overrideNote = `Multiple override caps triggered (${overrides.length}). Verdict capped at AVOID.`;
      } else if (overrides.length === 1) {
        if (baseVerdict === 'Invest') {
          finalVerdict = 'Watchlist';
          overrideNote = `Override cap active: ${overrides[0]}. Verdict capped at WATCHLIST.`;
        }
      }

      strengths.sort((a,b) => b.score - a.score);
      redFlags.sort((a,b) => a.score - b.score);

      return {
        companyName: fin.companyName,
        compositeScore,
        baseVerdict,
        verdict: finalVerdict,
        overallCoverage: overallCov,
        overrides,
        overrideNote,
        pillarResults,
        topStrengths: strengths.slice(0, 3).map(s => s.text),
        topRedFlags: redFlags.slice(0, 3).map(r => r.text),
        gmpValue,
        gmpPct
      };
    }

    function renderAnalysisModal(data) {
      document.getElementById('modalCompanyTitle').textContent = data.companyName + ' &bull; Fundamental Analysis';
      document.getElementById('modalFileSub').textContent = 'Automated 5-pillar fundamental analysis & forensic audit';

      // GMP Signal
      let gmpText = '-';
      if (data.gmpValue !== null && data.gmpValue !== undefined && !isNaN(data.gmpValue)) {
        let sentiment = data.gmpPct >= 20 ? 'Strong Bullish' : (data.gmpPct > 0 ? 'Mildly Bullish' : (data.gmpPct === 0 ? 'Neutral' : 'Bearish / Discount'));
        let sign = data.gmpValue >= 0 ? '+' : '';
        gmpText = `₹${data.gmpValue} (${sign}${data.gmpPct}%) &bull; Sentiment: ${sentiment}`;
      }
      document.getElementById('modalGmpText').innerHTML = gmpText;

      // Score and Verdict
      document.getElementById('modalCompositeScore').textContent = data.compositeScore.toFixed(1);
      document.getElementById('modalCoverage').textContent = data.overallCoverage.toFixed(0) + '%';
      
      let vBadge = document.getElementById('modalVerdictBadge');
      vBadge.textContent = data.verdict;
      vBadge.className = 'verdict-badge verdict-' + data.verdict.toLowerCase();

      // Override Alert
      let ovBox = document.getElementById('modalOverrideAlert');
      let ovText = document.getElementById('modalOverrideText');
      if (data.overrideNote || data.overrides.length > 0) {
        ovBox.style.display = 'flex';
        ovText.innerHTML = `<strong>Forensic / Solvency Override Alert:</strong><br>${data.overrideNote || data.overrides.join('<br>')}`;
      } else {
        ovBox.style.display = 'none';
      }

      // Bar Chart
      let barHtml = '';
      let radarData = [];
      for (let pKey in data.pillarResults) {
        let p = data.pillarResults[pKey];
        let scoreVal = p.score !== null ? p.score : 0;
        let color = scoreVal >= 70 ? 'var(--accent-green)' : (scoreVal >= 50 ? 'var(--accent-yellow)' : 'var(--accent-red)');
        radarData.push({ label: p.name, score: scoreVal });

        barHtml += `
          <div class="bar-item">
            <div class="bar-header">
              <span class="bar-name">${p.name} <span class="bar-cov">(${p.coverage}% cov)</span></span>
              <span class="bar-score" style="color:${color}">${p.score !== null ? p.score.toFixed(1) : 'N/A'}</span>
            </div>
            <div class="bar-track">
              <div class="bar-fill" style="width:${scoreVal}%; background:${color};"></div>
            </div>
          </div>
        `;
      }
      document.getElementById('barChartContainer').innerHTML = barHtml;

      // Radar Chart (SVG)
      document.getElementById('radarChartContainer').innerHTML = generateRadarSvg(radarData);

      // Strengths & Red Flags
      let sHtml = data.topStrengths.length > 0 
        ? data.topStrengths.map(s => `<li>${s}</li>`).join('') 
        : '<li>No prominent positive anomalies detected.</li>';
      document.getElementById('modalStrengthsList').innerHTML = sHtml;

      let rfHtml = data.topRedFlags.length > 0 
        ? data.topRedFlags.map(r => `<li>${r}</li>`).join('') 
        : '<li>No major forensic or solvency red flags flagged.</li>';
      document.getElementById('modalRedFlagsList').innerHTML = rfHtml;

      // Full Breakdown Table
      let bHtml = '';
      for (let pKey in data.pillarResults) {
        let p = data.pillarResults[pKey];
        bHtml += `<tr style="background:#1b1f29;"><td colspan="4"><strong>${p.name} (${p.score !== null ? p.score.toFixed(1) : 'N/A'})</strong></td></tr>`;
        for (let mKey in p.metrics) {
          let m = p.metrics[mKey];
          let mScore = m.score !== null ? m.score : 'N/A';
          let mColor = m.score >= 70 ? '#4ade80' : (m.score >= 50 ? '#facc15' : '#f87171');
          let badGood = `${formatVal(m.bad, m.format)} &rarr; ${formatVal(m.good, m.format)}`;
          bHtml += `
            <tr>
              <td>${m.label}</td>
              <td><strong>${m.formatted}</strong></td>
              <td style="color:#8e95a5;">${badGood}</td>
              <td><span class="breakdown-score-badge" style="color:${mColor}; border:1px solid ${mColor};">${mScore}</span></td>
            </tr>
          `;
        }
      }
      document.getElementById('modalBreakdownBody').innerHTML = bHtml;

      // Show Modal
      document.getElementById('analysisModal').classList.add('active');
    }

    function generateRadarSvg(items) {
      const size = 220;
      const center = size / 2;
      const radius = 80;
      const numSides = items.length;
      
      let gridPolys = '';
      for (let level of [0.25, 0.5, 0.75, 1.0]) {
        let pts = [];
        for (let i = 0; i < numSides; i++) {
          let angle = (Math.PI * 2 / numSides) * i - Math.PI / 2;
          let x = center + Math.cos(angle) * (radius * level);
          let y = center + Math.sin(angle) * (radius * level);
          pts.push(`${x.toFixed(1)},${y.toFixed(1)}`);
        }
        gridPolys += `<polygon points="${pts.join(' ')}" fill="none" stroke="#2c3240" stroke-width="1"/>`;
      }

      // Data polygon
      let dataPts = [];
      let labelNodes = '';
      for (let i = 0; i < numSides; i++) {
        let angle = (Math.PI * 2 / numSides) * i - Math.PI / 2;
        let scoreNorm = Math.max(0, Math.min(100, items[i].score)) / 100.0;
        let x = center + Math.cos(angle) * (radius * scoreNorm);
        let y = center + Math.sin(angle) * (radius * scoreNorm);
        dataPts.push(`${x.toFixed(1)},${y.toFixed(1)}`);

        // Label position slightly outside radius
        let lx = center + Math.cos(angle) * (radius + 20);
        let ly = center + Math.sin(angle) * (radius + 15);
        let anchor = Math.abs(Math.cos(angle)) < 0.2 ? 'middle' : (Math.cos(angle) > 0 ? 'start' : 'end');
        let shortName = items[i].label.split(' ')[0];
        labelNodes += `<text x="${lx.toFixed(1)}" y="${ly.toFixed(1)}" fill="#8e95a5" font-size="9" text-anchor="${anchor}">${shortName} (${items[i].score.toFixed(0)})</text>`;
      }

      let dataPoly = `<polygon points="${dataPts.join(' ')}" fill="rgba(52, 152, 219, 0.35)" stroke="#3498db" stroke-width="2"/>`;

      return `
        <svg width="${size}" height="${size}" viewBox="0 0 ${size} ${size}">
          ${gridPolys}
          ${dataPoly}
          ${labelNodes}
        </svg>
      `;
    }

    // File Upload Handler
    function handleFileUpload(input, gmpVal, gmpPct) {
      if (!input.files || input.files.length === 0) return;
      let file = input.files[0];
      let reader = new FileReader();
      reader.onload = function(e) {
        let data = new Uint8Array(e.target.result);
        let workbook = XLSX.read(data, { type: 'array' });
        let result = analyzeWorkbook(workbook, gmpVal, gmpPct);
        renderAnalysisModal(result);
      };
      reader.readAsArrayBuffer(file);
    }

    // Quick Sample Loader for instant testing
    function loadSampleTemplate(sampleFileName, gmpVal, gmpPct) {
      fetch(sampleFileName)
        .then(res => res.arrayBuffer())
        .then(ab => {
          let workbook = XLSX.read(new Uint8Array(ab), { type: 'array' });
          let result = analyzeWorkbook(workbook, gmpVal, gmpPct);
          renderAnalysisModal(result);
        })
        .catch(err => {
          alert('Could not fetch sample template: ' + err.message);
        });
    }
  </script>
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


def _card(record, is_open=False, card_idx=0):
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

    # Analyze Feature on "Open now" cards (or any card where is_open=True)
    analyze_block = ""
    if is_open:
        gmp_val_arg = f"{gmp_value}" if gmp_value is not None else "null"
        gmp_pct_arg = f"{gmp_pct}" if gmp_pct is not None else "null"
        input_id = f"upload_xlsx_{card_idx}"
        analyze_block = f"""
        <div class="analyze-box">
          <label class="upload-btn-label" for="{input_id}">
            📊 <span>Analyze (.xlsx)</span>
          </label>
          <input type="file" id="{input_id}" class="upload-input" accept=".xlsx"
                 onchange="handleFileUpload(this, {gmp_val_arg}, {gmp_pct_arg})">
          <span class="sample-link" onclick="loadSampleTemplate('sample_data/Crystal_Crop_Template.xlsx', {gmp_val_arg}, {gmp_pct_arg})">
            💡 Or load sample Excel model
          </span>
        </div>
        """

    return f"""<div class="card">
      <div>
        <h3>{name}{_badge(record.get('status'))}</h3>
        <div class="card-rows">
          <div class="row"><span>GMP</span>{gmp_html}</div>
          <div class="row"><span>Price</span><span>{html.escape(record.get('price') or '-')}</span></div>
          <div class="row"><span>Issue size</span><span>{html.escape(record.get('issue_size') or '-')}</span></div>
          <div class="row"><span>Open - Close</span><span>{_fmt_date(record.get('open_date'))} - {_fmt_date(record.get('close_date'))}</span></div>
          <div class="row"><span>Listing</span><span>{_fmt_date(record.get('listing_date'))}</span></div>
        </div>
        {metrics_block}
      </div>
      {analyze_block}
    </div>"""


def render_site(records):
    def sort_key(r):
        return r.get("open_date") or date.max

    open_now = [r for r in records if (r.get("status") or "").upper().startswith(("O", "CT"))]
    upcoming = [r for r in records if (r.get("status") or "").upper().startswith("U")]
    recent = [r for r in records if r not in open_now and r not in upcoming]

    sections = []
    card_counter = 0

    for title, group in [("Open now", open_now), ("Upcoming", upcoming), ("Recently closed / listed", recent)]:
        group = sorted(group, key=sort_key)
        if not group:
            continue
        is_open_section = (title == "Open now")
        card_htmls = []
        for r in group:
            card_counter += 1
            card_htmls.append(_card(r, is_open=is_open_section, card_idx=card_counter))
        
        cards = "\n".join(card_htmls)
        section_note = '<span class="section-badge">Upload .xlsx for instant 5-pillar fundamental audit</span>' if is_open_section else ""
        sections.append(f'<div class="section-title"><span>{title}</span>{section_note}</div><div class="grid">{cards}</div>')

    config_json_str = _load_analyzer_config_json()

    return (
        PAGE_TEMPLATE
        .replace("__UPDATED__", date.today().strftime("%d %B %Y"))
        .replace("__SECTIONS__", "\n".join(sections) if sections else "<p>No mainboard IPO data right now.</p>")
        .replace("__CONFIG_JSON__", config_json_str)
    )
