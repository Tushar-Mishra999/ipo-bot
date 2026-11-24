"""
Fundamental Financial Analyzer for IPOs.
Parses Screener.in-formatted financial Excel files (.xlsx) and calculates:
  - 5 Pillar Scores (0 to 100) with linear interpolation and coverage %
  - Composite Score (0 to 100)
  - Rule-based Verdict (Invest / Watchlist / Avoid)
  - Forensic & Solvency Override Caps
  - Top 3 Red Flags & Top 3 Strengths
"""

import json
import math
import os
from typing import Dict, Any, List, Optional, Tuple

import openpyxl

CONFIG_PATH = os.path.join(os.path.dirname(__file__), "analyzer_config.json")


def load_config() -> Dict[str, Any]:
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def _safe_float(val: Any) -> Optional[float]:
    if val is None or val == "":
        return None
    try:
        f = float(val)
        return None if (math.isnan(f) or math.isinf(f)) else f
    except (ValueError, TypeError):
        return None


def parse_screener_excel(file_path_or_stream) -> Dict[str, Any]:
    """
    Parses a Screener-style Excel model into structured financial time series.
    Looks primarily in 'Data Sheet' with fallback to financial statement sheets.
    """
    wb = openpyxl.load_workbook(file_path_or_stream, data_only=True)
    
    # Try Data Sheet first
    if "Data Sheet" in wb.sheetnames:
        sheet = wb["Data Sheet"]
    else:
        sheet = wb.active

    # Find the report dates row (typically row 16 in Data Sheet)
    date_cols: List[int] = []
    years: List[str] = []
    
    # Scan rows 10-20 to locate 'Report Date'
    date_row = None
    for r in range(1, min(sheet.max_row + 1, 30)):
        cell_v = sheet.cell(r, 1).value
        if cell_v and "report date" in str(cell_v).lower():
            date_row = r
            break

    if date_row is None:
        # Fallback to row 16 if not found
        date_row = 16

    for c in range(2, sheet.max_column + 1):
        v = sheet.cell(date_row, c).value
        if v is not None and str(v).strip():
            # Could be datetime or string year
            date_cols.append(c)
            years.append(str(v).split()[0])

    if not date_cols:
        # If still no dates found, inspect row 3 in active sheet
        for c in range(2, min(sheet.max_column + 1, 15)):
            v = sheet.cell(3, c).value
            if v is not None:
                date_cols.append(c)
                years.append(str(v).strip())

    # Build row index lookup by normalized label
    row_lookup: Dict[str, int] = {}
    for r in range(1, sheet.max_row + 1):
        lbl = sheet.cell(r, 1).value
        if lbl:
            clean_lbl = str(lbl).strip().lower()
            if clean_lbl not in row_lookup:
                row_lookup[clean_lbl] = r

    def get_series(aliases: List[str]) -> List[Optional[float]]:
        for alias in aliases:
            norm_alias = alias.strip().lower()
            # Exact match first
            if norm_alias in row_lookup:
                r = row_lookup[norm_alias]
                return [_safe_float(sheet.cell(r, c).value) for c in date_cols]
            # Substring match
            for key, r in row_lookup.items():
                if norm_alias in key:
                    return [_safe_float(sheet.cell(r, c).value) for c in date_cols]
        return [None] * len(date_cols)

    company_name = sheet.cell(1, 2).value or sheet.cell(1, 1).value or "IPO Company"

    sales = get_series(["sales", "revenue"])
    raw_material = get_series(["raw material cost"])
    change_inv = get_series(["change in inventory"])
    power_fuel = get_series(["power and fuel"])
    other_mfr = get_series(["other mfr. exp", "manufacturing"])
    employee_cost = get_series(["employee cost", "employee"])
    selling_admin = get_series(["selling and admin", "sga", "administrative"])
    other_expenses = get_series(["other expenses"])
    other_income = get_series(["other income"])
    depreciation = get_series(["depreciation"])
    interest = get_series(["interest", "finance costs"])
    pbt = get_series(["profit before tax", "pbt"])
    tax = get_series(["tax"])
    net_profit = get_series(["net profit", "pat"])
    dividends = get_series(["dividend amount", "dividend"])

    equity_cap = get_series(["equity share capital", "share capital"])
    reserves = get_series(["reserves"])
    borrowings = get_series(["borrowings", "total debt"])
    other_liab = get_series(["other liabilities"])
    tot_liab_assets = get_series(["total"])

    net_block = get_series(["net block", "fixed assets", "property, plant"])
    cwip = get_series(["capital work in progress", "cwip"])
    investments = get_series(["investments"])
    other_assets = get_series(["other assets"])
    receivables = get_series(["receivables", "debtors"])
    inventory = get_series(["inventory"])
    cash = get_series(["cash & bank", "cash"])

    cfo = get_series(["cash from operating activ", "cash from operations", "cfo"])
    cfi = get_series(["cash from investing activ", "cfi"])
    cff = get_series(["cash from financing activ", "cff"])

    # Calculate Operating Profit and OPM per period
    op_profit: List[Optional[float]] = []
    opm: List[Optional[float]] = []
    for i in range(len(date_cols)):
        s = sales[i]
        if s is not None and s > 0:
            # sum operating expenses with negative change in inventory
            rm = raw_material[i] or 0.0
            inv = change_inv[i] or 0.0
            o_mfr = other_mfr[i] or 0.0
            pf = power_fuel[i] or 0.0
            emp = employee_cost[i] or 0.0
            sga = selling_admin[i] or 0.0
            oe = other_expenses[i] or 0.0
            # Screener formula: RM + Power + OtherMfr + Emp + SGA + OtherExp - ChangeInInventory
            expenses = rm + pf + o_mfr + emp + sga + oe - inv
            op = s - expenses
            op_profit.append(op)
            opm.append(op / s)
        else:
            op_profit.append(None)
            opm.append(None)

    return {
        "company_name": str(company_name).strip(),
        "years": years,
        "sales": sales,
        "raw_material": raw_material,
        "change_inv": change_inv,
        "power_fuel": power_fuel,
        "other_mfr": other_mfr,
        "employee_cost": employee_cost,
        "selling_admin": selling_admin,
        "other_expenses": other_expenses,
        "other_income": other_income,
        "depreciation": depreciation,
        "interest": interest,
        "pbt": pbt,
        "tax": tax,
        "net_profit": net_profit,
        "dividends": dividends,
        "equity_cap": equity_cap,
        "reserves": reserves,
        "borrowings": borrowings,
        "other_liab": other_liab,
        "tot_liab_assets": tot_liab_assets,
        "net_block": net_block,
        "cwip": cwip,
        "investments": investments,
        "other_assets": other_assets,
        "receivables": receivables,
        "inventory": inventory,
        "cash": cash,
        "cfo": cfo,
        "cfi": cfi,
        "cff": cff,
        "op_profit": op_profit,
        "opm": opm,
    }


def compute_metrics(fin: Dict[str, Any]) -> Dict[str, Optional[float]]:
    """Calculates all 15 core financial metrics for the 5 pillars."""
    metrics: Dict[str, Optional[float]] = {}
    years = fin["years"]
    n = len(years)
    if n == 0:
        return metrics

    def clean_sum(series: List[Optional[float]]) -> Optional[float]:
        vals = [v for v in series if v is not None]
        return sum(vals) if vals else None

    # --- Pillar 1: Earnings Quality ---
    # 1. Cumulative CFO / PAT
    cum_cfo = clean_sum(fin["cfo"])
    cum_pat = clean_sum(fin["net_profit"])
    if cum_cfo is not None and cum_pat is not None and cum_pat != 0:
        metrics["cum_cfo_pat"] = cum_cfo / cum_pat
    else:
        metrics["cum_cfo_pat"] = None

    # 2. Accruals ratio: (PAT - CFO) / Total Assets (latest year)
    t = n - 1
    pat_t = fin["net_profit"][t]
    cfo_t = fin["cfo"][t]
    tot_assets_t = fin["tot_liab_assets"][t]
    if tot_assets_t is None or tot_assets_t <= 0:
        # Try to construct total assets
        nb = fin["net_block"][t] or 0.0
        cw = fin["cwip"][t] or 0.0
        inv = fin["investments"][t] or 0.0
        oa = fin["other_assets"][t] or 0.0
        tot_assets_t = nb + cw + inv + oa

    if pat_t is not None and cfo_t is not None and tot_assets_t and tot_assets_t > 0:
        metrics["accruals_ratio"] = (pat_t - cfo_t) / tot_assets_t
    else:
        metrics["accruals_ratio"] = None

    # --- Pillar 2: Forensic Risk ---
    # 1. Beneish M-Score
    # Needs year t and t-1
    if n >= 2:
        t1 = n - 2
        rec_t, rec_t1 = fin["receivables"][t], fin["receivables"][t1]
        s_t, s_t1 = fin["sales"][t], fin["sales"][t1]
        cfo_t_val = fin["cfo"][t] or 0.0
        pat_t_val = fin["net_profit"][t] or 0.0
        depr_t, depr_t1 = fin["depreciation"][t], fin["depreciation"][t1]
        nb_t, nb_t1 = fin["net_block"][t], fin["net_block"][t1]
        ta_t, ta_t1 = fin["tot_liab_assets"][t], fin["tot_liab_assets"][t1]
        borrow_t, borrow_t1 = fin["borrowings"][t], fin["borrowings"][t1]
        emp_t, emp_t1 = fin["employee_cost"][t] or 0.0, fin["employee_cost"][t1] or 0.0
        sga_t, sga_t1 = fin["selling_admin"][t] or 0.0, fin["selling_admin"][t1] or 0.0

        dsri = ((rec_t / s_t) / (rec_t1 / s_t1)) if (rec_t and rec_t1 and s_t and s_t1 and s_t > 0 and s_t1 > 0) else 1.0

        # GMI: Gross margin = (Sales - COGS) / Sales. COGS = RM + ChangeInInv
        rm_t = fin["raw_material"][t] or 0.0
        rm_t1 = fin["raw_material"][t1] or 0.0
        cogs_t = rm_t + (fin["change_inv"][t] or 0.0)
        cogs_t1 = rm_t1 + (fin["change_inv"][t1] or 0.0)
        gm_t = (s_t - cogs_t) / s_t if (s_t and s_t > 0) else 1.0
        gm_t1 = (s_t1 - cogs_t1) / s_t1 if (s_t1 and s_t1 > 0) else 1.0
        gmi = (gm_t1 / gm_t) if (gm_t and gm_t != 0) else 1.0

        # AQI
        aq_t = 1.0 - (nb_t / ta_t) if (nb_t is not None and ta_t and ta_t > 0) else 0.5
        aq_t1 = 1.0 - (nb_t1 / ta_t1) if (nb_t1 is not None and ta_t1 and ta_t1 > 0) else 0.5
        aqi = (aq_t / aq_t1) if (aq_t1 and aq_t1 != 0) else 1.0

        # SGI
        sgi = (s_t / s_t1) if (s_t and s_t1 and s_t1 > 0) else 1.0

        # DEPI
        dr_t = (depr_t / (nb_t + depr_t)) if (depr_t is not None and nb_t is not None and (nb_t + depr_t) > 0) else 0.1
        dr_t1 = (depr_t1 / (nb_t1 + depr_t1)) if (depr_t1 is not None and nb_t1 is not None and (nb_t1 + depr_t1) > 0) else 0.1
        depi = (dr_t1 / dr_t) if (dr_t and dr_t != 0) else 1.0

        # SGAI
        sg_exp_t = emp_t + sga_t
        sg_exp_t1 = emp_t1 + sga_t1
        sgai = ((sg_exp_t / s_t) / (sg_exp_t1 / s_t1)) if (s_t and s_t1 and sg_exp_t1 and s_t > 0 and s_t1 > 0) else 1.0

        # LVGI
        lv_t = (borrow_t / ta_t) if (borrow_t is not None and ta_t and ta_t > 0) else 0.2
        lv_t1 = (borrow_t1 / ta_t1) if (borrow_t1 is not None and ta_t1 and ta_t1 > 0) else 0.2
        lvgi = (lv_t / lv_t1) if (lv_t1 and lv_t1 != 0) else 1.0

        # TATA
        tata = ((pat_t_val - cfo_t_val) / ta_t) if (ta_t and ta_t > 0) else 0.0

        m_score = (
            -4.84
            + 0.920 * dsri
            + 0.528 * gmi
            + 0.404 * aqi
            + 0.892 * sgi
            + 0.115 * depi
            - 0.172 * sgai
            + 4.037 * tata
            + 0.0327 * lvgi
        )
        metrics["beneish_m_score"] = m_score
    else:
        metrics["beneish_m_score"] = None

    # 2. Change in receivable days over 2 years: (DebtorDays_t - DebtorDays_t-2)
    ref_idx = n - 3 if n >= 3 else 0
    if n >= 2:
        rec_now = fin["receivables"][t]
        s_now = fin["sales"][t]
        rec_past = fin["receivables"][ref_idx]
        s_past = fin["sales"][ref_idx]
        if rec_now is not None and s_now and s_now > 0 and rec_past is not None and s_past and s_past > 0:
            days_now = (rec_now / s_now) * 365.0
            days_past = (rec_past / s_past) * 365.0
            metrics["chg_receivable_days_2y"] = days_now - days_past
        else:
            metrics["chg_receivable_days_2y"] = None
    else:
        metrics["chg_receivable_days_2y"] = None

    # 3. Other income as share of PBT
    oi_t = fin["other_income"][t]
    pbt_t = fin["pbt"][t]
    if oi_t is not None and pbt_t is not None:
        if pbt_t > 0:
            metrics["other_income_share_pbt"] = oi_t / pbt_t
        else:
            # If PBT is negative or zero, core business is unprofitable
            metrics["other_income_share_pbt"] = 1.0
    else:
        metrics["other_income_share_pbt"] = None

    # --- Pillar 3: Financial Strength ---
    # 1. Altman Z'' Score = 6.56*X1 + 3.26*X2 + 6.72*X3 + 1.05*X4
    tot_assets = fin["tot_liab_assets"][t]
    other_liab_val = fin["other_liab"][t] or 0.0
    borrow_val = fin["borrowings"][t] or 0.0
    tot_liab_val = borrow_val + other_liab_val
    eq_cap = fin["equity_cap"][t] or 0.0
    res_val = fin["reserves"][t] or 0.0
    equity_val = eq_cap + res_val
    pbt_val = fin["pbt"][t] or 0.0
    interest_val = fin["interest"][t] or 0.0
    ebit_val = pbt_val + interest_val

    # Working Capital = Current Assets - Current Liabilities
    rec_val = fin["receivables"][t] or 0.0
    inv_val = fin["inventory"][t] or 0.0
    cash_val = fin["cash"][t] or 0.0
    ca_val = rec_val + inv_val + cash_val
    wc_val = ca_val - other_liab_val

    if tot_assets and tot_assets > 0:
        x1 = wc_val / tot_assets
        x2 = res_val / tot_assets
        x3 = ebit_val / tot_assets
        x4 = (equity_val / tot_liab_val) if tot_liab_val > 0 else 2.0
        metrics["altman_z_score"] = 6.56 * x1 + 3.26 * x2 + 6.72 * x3 + 1.05 * x4
    else:
        metrics["altman_z_score"] = None

    # 2. Interest Coverage Ratio = EBIT / Interest
    if interest_val > 0:
        metrics["interest_cover"] = ebit_val / interest_val
    elif interest_val == 0:
        metrics["interest_cover"] = 50.0  # Zero interest expense = very safe
    else:
        metrics["interest_cover"] = None

    # 3. Debt to Equity = Borrowings / Equity
    if equity_val > 0 and borrow_val is not None:
        metrics["debt_to_equity"] = borrow_val / equity_val
    else:
        metrics["debt_to_equity"] = None

    # --- Pillar 4: Profitability and Growth ---
    # 1. Sales CAGR
    s_first = fin["sales"][0]
    s_last = fin["sales"][t]
    years_cagr = n - 1
    if s_first and s_last and s_first > 0 and s_last > 0 and years_cagr >= 1:
        metrics["sales_cagr"] = (s_last / s_first) ** (1.0 / years_cagr) - 1.0
    else:
        metrics["sales_cagr"] = None

    # 2. PAT CAGR
    p_first = fin["net_profit"][0]
    p_last = fin["net_profit"][t]
    if p_first and p_last and p_first > 0 and p_last > 0 and years_cagr >= 1:
        metrics["pat_cagr"] = (p_last / p_first) ** (1.0 / years_cagr) - 1.0
    elif p_first and p_last and p_first > 0 and p_last <= 0:
        metrics["pat_cagr"] = -0.5
    else:
        metrics["pat_cagr"] = None

    # 3. ROE = Latest PAT / Latest Equity
    if equity_val > 0 and p_last is not None:
        metrics["roe"] = p_last / equity_val
    else:
        metrics["roe"] = None

    # 4. EBITDA Margin Trend over 2 years (or first vs last)
    opm_series = fin["opm"]
    if len(opm_series) >= 3 and opm_series[t] is not None and opm_series[t - 2] is not None:
        metrics["ebitda_margin_trend"] = opm_series[t] - opm_series[t - 2]
    elif len(opm_series) >= 2 and opm_series[t] is not None and opm_series[0] is not None:
        metrics["ebitda_margin_trend"] = opm_series[t] - opm_series[0]
    else:
        metrics["ebitda_margin_trend"] = None

    # --- Pillar 5: Cash and Capital Allocation ---
    # 1. FCF Proxy relative to PAT: Cumulative (CFO - Delta NetBlock - Delta CWIP) / Cumulative PAT
    fcf_list: List[float] = []
    pat_fcf_list: List[float] = []
    for i in range(1, n):
        cfo_i = fin["cfo"][i]
        pat_i = fin["net_profit"][i]
        nb_i, nb_prev = fin["net_block"][i], fin["net_block"][i - 1]
        cw_i, cw_prev = fin["cwip"][i], fin["cwip"][i - 1]
        if cfo_i is not None and pat_i is not None and nb_i is not None and nb_prev is not None:
            delta_nb = nb_i - nb_prev
            delta_cw = (cw_i or 0.0) - (cw_prev or 0.0)
            capex = delta_nb + delta_cw
            fcf = cfo_i - capex
            fcf_list.append(fcf)
            pat_fcf_list.append(pat_i)

    if fcf_list and sum(pat_fcf_list) > 0:
        metrics["fcf_to_pat"] = sum(fcf_list) / sum(pat_fcf_list)
    else:
        metrics["fcf_to_pat"] = None

    # 2. Dividend Payout vs Cash Generated: Cumulative Dividends / Cumulative CFO
    cum_div = clean_sum(fin["dividends"]) or 0.0
    if cum_cfo and cum_cfo > 0:
        metrics["div_payout_vs_cfo"] = cum_div / cum_cfo
    elif cum_div == 0:
        metrics["div_payout_vs_cfo"] = 0.0
    else:
        metrics["div_payout_vs_cfo"] = 1.0

    # 3. Cash vs Borrowings
    latest_cash = fin["cash"][t] or 0.0
    latest_borrow = fin["borrowings"][t]
    if latest_borrow is not None:
        if latest_borrow > 0:
            metrics["cash_vs_borrowings"] = latest_cash / latest_borrow
        else:
            # Debt-free company
            metrics["cash_vs_borrowings"] = 5.0
    else:
        metrics["cash_vs_borrowings"] = None

    return metrics


def interpolate_score(value: Optional[float], bad: float, good: float, higher_is_better: bool) -> Optional[float]:
    """Linearly maps a metric value between bad (0) and good (100), clamped to [0, 100]."""
    if value is None or math.isnan(value) or math.isinf(value):
        return None

    if higher_is_better:
        if good == bad:
            return 100.0 if value >= good else 0.0
        score = (value - bad) / (good - bad) * 100.0
    else:
        if bad == good:
            return 100.0 if value <= good else 0.0
        score = (bad - value) / (bad - good) * 100.0

    return max(0.0, min(100.0, score))


def format_metric_value(val: Optional[float], fmt: str) -> str:
    if val is None:
        return "N/A"
    if fmt == "pct":
        return f"{val:+.1%}" if fmt == "pct" and val < 0 else f"{val:.1%}"
    if fmt == "ratio":
        return f"{val:.2f}x"
    if fmt == "score":
        return f"{val:.2f}"
    if fmt == "days":
        return f"{val:+.1f} days"
    return f"{val:.2f}"


def score_company(fin: Dict[str, Any], config: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """Computes all pillar scores, composite score, overrides, red flags, and strengths."""
    if config is None:
        config = load_config()

    metrics = compute_metrics(fin)

    pillar_results: Dict[str, Any] = {}
    metric_scores: Dict[str, Any] = {}

    insights_strengths: List[Tuple[float, str]] = []
    insights_red_flags: List[Tuple[float, str]] = []

    for pillar_key, pillar_conf in config["pillars"].items():
        p_name = pillar_conf["name"]
        p_weight = pillar_conf["weight"]
        p_metrics_conf = pillar_conf["metrics"]

        scores_in_pillar: List[float] = []
        metric_details: Dict[str, Any] = {}

        for m_key, m_conf in p_metrics_conf.items():
            val = metrics.get(m_key)
            bad = m_conf["bad"]
            good = m_conf["good"]
            higher_is_better = m_conf["higher_is_better"]
            fmt = m_conf.get("format", "score")

            m_score = interpolate_score(val, bad, good, higher_is_better)
            if m_score is not None:
                scores_in_pillar.append(m_score)

            metric_details[m_key] = {
                "label": m_conf["label"],
                "value": val,
                "formatted_value": format_metric_value(val, fmt),
                "score": round(m_score, 1) if m_score is not None else None,
                "bad": bad,
                "good": good,
                "higher_is_better": higher_is_better,
            }

            # Gather candidates for strengths and red flags
            if val is not None and m_score is not None:
                try:
                    if m_score >= 70.0 and "strength_template" in m_conf:
                        txt = m_conf["strength_template"].format(value=val)
                        insights_strengths.append((m_score, txt))
                    elif m_score <= 45.0 and "red_flag_template" in m_conf:
                        txt = m_conf["red_flag_template"].format(value=val)
                        insights_red_flags.append((m_score, txt))
                except Exception:
                    pass

        total_metrics_count = len(p_metrics_conf)
        avail_count = len(scores_in_pillar)
        coverage_pct = (avail_count / total_metrics_count * 100.0) if total_metrics_count > 0 else 0.0

        if scores_in_pillar:
            p_score = sum(scores_in_pillar) / len(scores_in_pillar)
        else:
            p_score = None

        pillar_results[pillar_key] = {
            "name": p_name,
            "weight": p_weight,
            "score": round(p_score, 1) if p_score is not None else None,
            "coverage_pct": round(coverage_pct, 1),
            "metrics": metric_details,
        }

    # Re-normalize composite score across available pillars
    weighted_sum = 0.0
    total_avail_weight = 0.0
    pillar_coverage_sum = 0.0

    for p_key, res in pillar_results.items():
        w = res["weight"]
        pillar_coverage_sum += res["coverage_pct"] * w
        if res["score"] is not None:
            weighted_sum += res["score"] * w
            total_avail_weight += w

    if total_avail_weight > 0:
        composite_score = weighted_sum / total_avail_weight
    else:
        composite_score = 0.0

    composite_score = round(composite_score, 1)
    overall_coverage = round(pillar_coverage_sum, 1)

    # Base Verdict
    verdict_conf = config["verdict_thresholds"]
    if composite_score >= verdict_conf["invest"]:
        base_verdict = "Invest"
    elif composite_score >= verdict_conf["watchlist"]:
        base_verdict = "Watchlist"
    else:
        base_verdict = "Avoid"

    # Evaluate Overrides
    overrides_triggered: List[str] = []
    overrides_conf = config["overrides"]

    # 1. Beneish above -1.78
    beneish_val = metrics.get("beneish_m_score")
    if beneish_val is not None and beneish_val > overrides_conf["beneish_max"]:
        overrides_triggered.append(
            f"Beneish M-Score of {beneish_val:.2f} is higher than {overrides_conf['beneish_max']:.2f} (earnings manipulation risk)"
        )

    # 2. Altman Z'' below 1.1
    altman_val = metrics.get("altman_z_score")
    if altman_val is not None and altman_val < overrides_conf["altman_min"]:
        overrides_triggered.append(
            f"Altman Z'' score of {altman_val:.2f} is below {overrides_conf['altman_min']:.2f} (financial distress risk)"
        )

    # 3. Cumulative CFO/PAT below 0.5
    cum_cfo_pat_val = metrics.get("cum_cfo_pat")
    if cum_cfo_pat_val is not None and cum_cfo_pat_val < overrides_conf["cum_cfo_pat_min"]:
        overrides_triggered.append(
            f"Cumulative CFO/PAT of {cum_cfo_pat_val:.2f}x is below {overrides_conf['cum_cfo_pat_min']:.2f}x (weak cash conversion)"
        )

    # Override rules:
    # 1 override caps verdict at Watchlist
    # 2+ overrides caps verdict at Avoid
    final_verdict = base_verdict
    override_note = ""

    if len(overrides_triggered) >= 2:
        final_verdict = "Avoid"
        override_note = f"Multiple forensic/solvency override caps triggered ({len(overrides_triggered)}). Verdict capped at Avoid."
    elif len(overrides_triggered) == 1:
        if base_verdict == "Invest":
            final_verdict = "Watchlist"
            override_note = f"Override cap triggered: {overrides_triggered[0]}. Verdict capped at Watchlist."

    # Top 3 Strengths (highest scores first)
    insights_strengths.sort(key=lambda x: x[0], reverse=True)
    top_strengths = [text for _, text in insights_strengths[:3]]

    # Top 3 Red Flags (lowest scores first)
    insights_red_flags.sort(key=lambda x: x[0])
    top_red_flags = [text for _, text in insights_red_flags[:3]]

    return {
        "company_name": fin.get("company_name", "IPO Company"),
        "years": fin.get("years", []),
        "composite_score": composite_score,
        "base_verdict": base_verdict,
        "verdict": final_verdict,
        "overall_coverage_pct": overall_coverage,
        "overrides_triggered": overrides_triggered,
        "override_note": override_note,
        "pillars": pillar_results,
        "top_strengths": top_strengths,
        "top_red_flags": top_red_flags,
    }


def analyze_file(file_path: str) -> Dict[str, Any]:
    fin = parse_screener_excel(file_path)
    return score_company(fin)
