"""Monthly pivots and the forward estimates built on them: spend by category,
Fidelity income by account and by fund, the future-month estimation, and the
wealth summary data block.

Extracted from spending_summary.py on 2026-07-19 — that file had grown to 4,020
lines. Behaviour is unchanged: the code below is the original, moved verbatim.
"""
import csv
import io
import os

import pandas as pd

from .constants import ACCOUNT_OWNER, CATEGORIES, ESTIMATES_AS_OF, FAMILY_ORDER
from .loaders import load_history, load_income_history, load_spend_history


def build_spending_pivot(amex_df, bar_df, anchors):
    combined = pd.concat([amex_df, bar_df])
    pivot = combined.pivot_table(
        index="category", columns="month", values="spend",
        aggfunc="sum", fill_value=0
    )
    months = sorted(pivot.columns)
    pivot = pivot.reindex(columns=months, fill_value=0)
    pivot["Total"] = pivot[months].sum(axis=1)
    pivot = pivot.reindex(CATEGORIES + ["Salary", "Fidelity"], fill_value=0)

    # Inject the pinned spend history — never overwrite with live data
    HIST_CUTOFF = anchors.hist_cutoff
    spend_hist = load_spend_history()
    for cat, hist_vals in spend_hist.items():
        for period, val in hist_vals.items():
            if period < HIST_CUTOFF:
                if period not in pivot.columns:
                    pivot[period] = 0
                pivot.loc[cat, period] = val
    # Re-sort columns
    all_cols = sorted([c for c in pivot.columns if c != "Total"])
    pivot = pivot.reindex(columns=all_cols + ["Total"], fill_value=0)
    pivot["Total"] = pivot[all_cols].sum(axis=1)
    months = all_cols
    return pivot, months


def build_fidelity_pivot(fid_df, anchors):
    if fid_df.empty:
        pivot = pd.DataFrame()
    else:
        fid_df = fid_df[fid_df["account"].isin(ACCOUNT_OWNER)].copy()
        pivot = fid_df.pivot_table(
            index="account", columns="month", values="amount",
            aggfunc="sum", fill_value=0
        )

    # Inject the pinned income history — never overwrite with live data
    HIST_CUTOFF = anchors.hist_cutoff
    inc_hist = load_income_history()
    for acc, hist_vals in inc_hist.items():
        for period, val in hist_vals.items():
            if period < HIST_CUTOFF:
                if acc not in pivot.index:
                    pivot.loc[acc] = 0
                if period not in pivot.columns:
                    pivot[period] = 0
                pivot.loc[acc, period] = val

    if pivot.empty:
        return pd.DataFrame(), []

    months = sorted([c for c in pivot.columns if c != "Total"])
    pivot = pivot.reindex(columns=months, fill_value=0)
    pivot["Total"] = pivot[months].sum(axis=1)
    pivot = pivot.sort_values("Total", ascending=False)
    return pivot, months


def estimate_future_months(pivot, actual_months, future_months, anchors,
                           skip_funds=None):
    """
    Estimate the months that aren't complete actuals.
    - The PARTIAL month (the one the broker snapshot lands inside): scaled up
      from its own raw value by how much of the month the snapshot covers.
    - Every other estimated month — future months and any GAP month between the
      pinned history and the transaction window: median of actuals.
    - University Fees: Feb actual + Sep estimate only.
    - Tax: no extrapolation.
    - Stocks: use confirmed dividend payment months from market data.
    """
    import numpy as np

    result = pivot.copy()
    PARTIAL = anchors.partial_month
    partial_scale = anchors.partial_scale

    # Confirmed 2026 dividend payment months for stocks
    # Sold stocks get empty list (no future payments)
    # Amounts estimated from actuals; for stocks with no 2026 actuals, use 0 (unknown amount)
    STOCK_DIV_MONTHS = {
        # Currently held
        "AVIVA, ORD GBP0.328947368 (AV.)":          [10],        # Oct (May already in actuals)
        "AUTOTRADER GROUP PLC, ORD GBP0.01 (AUTO)":  [9],         # Sep interim (~Aug ex-div)
        "RELX PLC, ORD GBP0.1444 (REL)":             [6, 9],      # Jun final + Sep interim
        "THE SAGE GROUP PLC, GBP0.01051948 (SGE)":   [6],         # Jun interim (~May ex-div)
        "WEIR GROUP, ORD GBP0.125 (WEIR)":           [6, 11],     # Jun + Nov
        "VODAFONE GROUP, ORD USD0.2095238 (VOD)":    [7],         # Jul final FY26
        # Sold — no future payments
        "BP, ORD USD0.25 (BP.)":                     [],
        "SHELL PLC, ORD EUR0.07 (SHEL)":             [],
        "LEGAL & GENERAL GROUP, ORD GBP0.025 (LGEN)": [],
        "ITV, ORD GBP0.10 (ITV)":                   [],
        "SAP SE, ORD NPV (SAP)":                     [],
        # Funds with confirmed quarterly payment schedule
        # M&G Global Floating Rate HY: quarterly, pay dates 31 Mar/30 Jun/30 Sep/31 Dec
        # (Mar already in actuals; remaining: Jun, Sep, Dec)
        "M&G Global Floating Rate High Yield Fund Sterling I-H Inc": [6, 9, 12],
    }

    skip_funds = skip_funds or set()

    for idx in result.index:
        # Skip funds that were injected with precise pence-based estimates
        if idx in skip_funds:
            continue
        if idx == "University Fees":
            feb_val = 0
            for m in actual_months:
                if m.month == 2:
                    feb_val = float(result.loc[idx, m])
            for m in future_months:
                if m == PARTIAL:
                    raw = float(result.loc[idx, m]) if m in result.columns else 0
                    result.loc[idx, m] = round(raw / partial_scale) if raw > 0 else 0
                else:
                    result.loc[idx, m] = round(feb_val) if m.month == 9 else 0
            result.loc[idx, "Total"] = round(sum(float(result.loc[idx, m]) for m in actual_months + future_months))
            continue

        if idx == "Tax":
            for m in future_months:
                result.loc[idx, m] = 0
            result.loc[idx, "Total"] = round(sum(float(result.loc[idx, m]) for m in actual_months))
            continue

        if idx == "Salary":
            # Salary is a fixed payment — never scale up for a partial month.
            # Keep whatever has actually been paid in the partial month; if the
            # payslip hasn't landed yet (snapshot taken before pay day) fall back
            # to the median rather than reporting a month with no salary at all.
            actuals_vals = [float(result.loc[idx, m]) for m in actual_months if m in result.columns]
            monthly_salary = round(np.median(actuals_vals)) if actuals_vals else 0
            for m in future_months:
                raw = float(result.loc[idx, m]) if m in result.columns else 0
                if m == PARTIAL and raw > 0:
                    continue  # already paid this month — keep the actual
                result.loc[idx, m] = monthly_salary
            result.loc[idx, "Total"] = round(sum(float(result.loc[idx, m]) for m in actual_months + future_months))
            continue

        if idx == "Holidays":
            ANNUAL_BUDGET = 12000
            actual_total = sum(float(result.loc[idx, m]) for m in actual_months)
            remaining = max(0, ANNUAL_BUDGET - actual_total)
            per_month = round(remaining / len(future_months)) if future_months else 0
            for m in future_months:
                result.loc[idx, m] = per_month
            result.loc[idx, "Total"] = round(actual_total + remaining)
            continue

        # ── Stock / equity dividends ──────────────────────────────────────────
        # Schedule: { fund_name: [(payment_month, pence_per_share), ...] }
        # Pence-per-share from confirmed 2026 announcements (dividenddata.co.uk / company IR)
        # Per-account cash estimate = units_held * pence / 100
        # If no units data available, falls back to historical mean payment
        STOCK_DIV_SCHEDULE = {
            # Semi-annual. Final: pay May 14 (already in actuals). Interim: pay Oct 16 (ex Aug 28)
            "AVIVA, ORD GBP0.328947368 (AV.)":         [(10, 13.1)],
            # Semi-annual. Interim: pay Jan 26 (already in actuals). Final: ~Sep 26 (ex ~Aug 27)
            "AUTOTRADER GROUP PLC, ORD GBP0.01 (AUTO)": [(9,  7.1)],
            # Semi-annual. Final: pay Jun 18 (ex May 7). Interim: ~Sep 11 (ex ~Aug 7)
            "RELX PLC, ORD GBP0.1444 (REL)":            [(6, 48.0), (9, 19.5)],
            # Semi-annual. Final: pay Feb 10 (already in actuals). Interim: ~Jun 27 (ex ~May 29)
            "THE SAGE GROUP PLC, GBP0.01051948 (SGE)":  [(6,  7.5)],
            # Semi-annual. Final: pay May 29 (already in actuals). Interim: ~Nov 4 (ex ~Oct 2)
            "WEIR GROUP, ORD GBP0.125 (WEIR)":          [(11, 19.6)],
            # Semi-annual. Interim: pay Feb 5 (already in actuals). Final: pay Jul 31 (ex Jun 5)
            "VODAFONE GROUP, ORD USD0.2095238 (VOD)":   [(7,  1.9)],
            # Semi-annual. Final: pay Jun 4 (ex Apr 23 — already past). Interim: pay Sep 25 (ex Aug 20)
            "LEGAL & GENERAL GROUP, ORD GBP0.025 (LGEN)": [(6, 15.36), (9, 6.12)],
            # Sold — no future payments
            "BP, ORD USD0.25 (BP.)":                     [],
            "SHELL PLC, ORD EUR0.07 (SHEL)":             [],
            "ITV, ORD GBP0.10 (ITV)":                    [],
            "SAP SE, ORD NPV (SAP)":                     [],
            # Funds with confirmed quarterly schedule — pay months only, no pence-per-share
            # M&G GFRHYF: quarterly pay dates 31 Mar/30 Jun/30 Sep/31 Dec (ex-div ~1 month before)
            "M&G Global Floating Rate High Yield Fund Sterling I-H Inc": [(6, None), (9, None), (12, None)],
        }

        if idx in STOCK_DIV_SCHEDULE:
            schedule = STOCK_DIV_SCHEDULE[idx]
            # For each future payment, try to compute from units × pence, else use mean of actuals

            # Gather actual payments for fallback
            actual_vals = [float(result.loc[idx, m]) for m in actual_months if float(result.loc[idx, m]) > 0]
            if not actual_vals:
                actual_vals = [float(result.loc[idx, m]) for m in future_months if m in result.columns and float(result.loc[idx, m]) > 0]

            for m in future_months:
                if m == PARTIAL:
                    raw = float(result.loc[idx, m]) if m in result.columns else 0
                    if raw > 0:
                        # A dividend is a discrete payment, not a monthly rate —
                        # one already banked stands as it is. Scaling it by the
                        # partial-month fraction would invent cash.
                        continue
                # Check if this month has a scheduled payment
                pay_entry = next(((mo, ppm) for mo, ppm in schedule if mo == m.month), None)
                if pay_entry:
                    pay_month, ppm = pay_entry
                    if actual_vals:
                        per_payment = round(float(np.mean(actual_vals)))
                    else:
                        per_payment = 0
                    result.loc[idx, m] = per_payment
                else:
                    result.loc[idx, m] = 0
            result.loc[idx, "Total"] = round(sum(float(result.loc[idx, m]) for m in actual_months + future_months))
            continue

        # Any unscheduled stock: zero future months
        idx_upper = str(idx).upper()
        is_stock = any(p in idx_upper for p in [", ORD ", "ORD GBP", "ORD USD", "ORD EUR", "ORD NPV"])
        if is_stock:
            for m in future_months:
                # Keep a payment already banked in the partial month
                if m == PARTIAL and m in result.columns and float(result.loc[idx, m]) > 0:
                    continue
                result.loc[idx, m] = 0
            result.loc[idx, "Total"] = round(sum(float(result.loc[idx, m])
                                                 for m in actual_months + future_months))
            continue

        # Default: median of actuals for regular monthly/recurring income
        actuals = [float(result.loc[idx, m]) if m in result.columns else 0 for m in actual_months]
        non_zero = [v for v in actuals if v > 0]
        median_est = float(np.median(non_zero)) if non_zero else 0

        for m in future_months:
            if m == PARTIAL:
                raw = float(result.loc[idx, m]) if m in result.columns else 0
                result.loc[idx, m] = round(raw / partial_scale) if raw > 0 else round(median_est)
            else:
                result.loc[idx, m] = round(median_est)

    result["Total"] = result[actual_months + future_months].sum(axis=1)
    return result

def build_account_fund_pivot(fid_df):
    """
    Returns OrderedDict: person -> { account -> DataFrame(fund x month) }
    Filtered to family accounts only, grouped by person in FAMILY_ORDER,
    accounts within each person sorted by total desc, funds alphabetically.
    """
    if fid_df.empty:
        return {}, []

    months = sorted(fid_df["month"].unique())

    # Filter to family accounts only
    fid_df = fid_df[fid_df["account"].isin(ACCOUNT_OWNER)].copy()

    result = {}
    for person in FAMILY_ORDER:
        person_accounts = [acc for acc, owner in ACCOUNT_OWNER.items() if owner == person]
        # Sort by total desc
        acc_totals = (
            fid_df[fid_df["account"].isin(person_accounts)]
            .groupby("account")["amount"].sum()
            .reindex(person_accounts, fill_value=0)
            .sort_values(ascending=False)
        )
        person_data = {}
        for acc in acc_totals.index:
            sub = fid_df[fid_df["account"] == acc]
            if sub.empty:
                continue
            pivot = sub.pivot_table(
                index="fund", columns="month", values="amount",
                aggfunc="sum", fill_value=0
            )
            pivot = pivot.reindex(columns=months, fill_value=0)
            pivot["Total"] = pivot[months].sum(axis=1)
            pivot = pivot.sort_index()
            person_data[acc] = pivot
        if person_data:
            result[person] = person_data

    return result, months

def build_summary_data(account_summary_path, all_months, anchors):
    """Build summary data. History is embedded in load_history() — no CSV needed."""
    FIDELITY_ACCS = {
        'AW10032966','SANX002282','2000001606','SANQ000468',
        'SANX002936','AW10261123','SANX002617','2000001604',
        'AW10580794','AS10303823','AG10131710'
    }

    # ── Fidelity total from AccountSummary ────────────────────────────────────
    fidelity_total = 0
    fidelity_by_acc = {}
    if account_summary_path and os.path.exists(account_summary_path):
        with open(account_summary_path, encoding="utf-8-sig") as f:
            content = f.read()
        lines = content.replace("\r", "").split("\n")
        header_idx = max(i for i, l in enumerate(lines) if l.startswith("Type,Holdings,Account number"))
        reader = csv.DictReader(io.StringIO("\n".join(lines[header_idx:])))
        for row in reader:
            if row.get("Type", "").strip() != "Account": continue
            acc = row.get("Account number", "").strip()
            if acc not in FIDELITY_ACCS: continue
            try:
                v = float(row.get("Value (£)", "0").replace(",", ""))
            except (ValueError, TypeError):
                v = 0
            fidelity_by_acc[acc] = v
            fidelity_total += v

    # ── Index of the month the pinned estimates below were measured ───────────
    # This one IS a fixed month on purpose: the pension/house/car figures were
    # taken as at ESTIMATES_AS_OF and the projection runs out from there, so it
    # stays correct as the calendar advances. It is not a live-data anchor.
    est_idx = next((i for i, m in enumerate(all_months) if m == ESTIMATES_AS_OF), 0)

    def project_monthly(base_value, annual_pct=None, monthly_add=None):
        """Project the ESTIMATES_AS_OF value forward and backward across the year."""
        vals = {}
        for i, m in enumerate(all_months):
            offset = i - est_idx  # months from ESTIMATES_AS_OF
            if annual_pct is not None:
                # Compound monthly: (1 + annual_pct) ^ (offset/12)
                vals[m] = round(base_value * ((1 + annual_pct) ** (offset / 12)))
            elif monthly_add is not None:
                vals[m] = round(max(0, base_value + offset * monthly_add))
        return vals

    # ── PS Arriva (Defined Benefit) ────────────────────────────────────────────
    arriva = project_monthly(67786, annual_pct=0.05)

    # ── Expleo Scottish Widows ────────────────────────────────────────────────
    expleo_sw = project_monthly(12000, monthly_add=6000)

    # ── Susan's pensions (5%/year growth from the ESTIMATES_AS_OF figures) ────
    susan_pensions = {
        "Capita (RMSPS) – to 2018 (65 yrs)":    project_monthly(489475, annual_pct=0.05),
        "RMPP 2012–2023 (65 years)":             project_monthly(229010, annual_pct=0.05),
        "Collective Pension (2023+)":            project_monthly(43859,  annual_pct=0.05),
        "Cash Balance":                          project_monthly(160046, annual_pct=0.05),
        "AVC Bonus Plan (Scottish Widows)":      project_monthly(88174,  annual_pct=0.05),
    }

    # Susan's Fidelity SIPP (2000001604)
    # Jan–Apr: use 'Susan Fidelity Pension' history row
    # May+: start from AccountSummary value, updated with income growth in write_excel
    susan_fidelity_sipp = fidelity_by_acc.get("2000001604", 0)
    susan_fidelity_sipp_vals = {m: susan_fidelity_sipp for m in all_months}

    # ── Historical data (hardcoded — no file needed) ──────────────────────────
    history = load_history()

    # Override pre-cutoff months with history values for Susan Fidelity Pension
    susan_fid_hist = history.get("Susan Fidelity Pension", {})
    for m in all_months:
        if m < anchors.wealth_cutoff and m in susan_fid_hist:
            susan_fidelity_sipp_vals[m] = susan_fid_hist[m]

    # ── Liam Fidelity (AS10303823 + AG10131710) ───────────────────────────────
    liam_fid_total = sum(fidelity_by_acc.get(a, 0) for a in ["AS10303823", "AG10131710"])

    # ── Jayne Fidelity (SANX002936) ───────────────────────────────────────────
    jayne_fid_total = fidelity_by_acc.get("SANX002936", 0)

    # ── Barclays balance by month ──────────────────────────────────────────────
    # Pinned balances from manual/previous tracking, ending with the last one
    # confirmed. Every month with no pinned balance holds at that last figure,
    # so this needs no calendar anchor — extend the dict when a newer balance
    # is confirmed.
    barclays_by_month = {
        pd.Period("2026-01", "M"): 6000,
        pd.Period("2026-02", "M"): 6000,
        pd.Period("2026-03", "M"): 5500,
        pd.Period("2026-04", "M"): 5500,
        pd.Period("2026-05", "M"): 13000,
    }
    barclays_by_month[pd.Period("2026-06", "M")] = round(6916.87)   # confirmed
    last_known = barclays_by_month[max(barclays_by_month)]
    for m in all_months:
        if m not in barclays_by_month:
            barclays_by_month[m] = last_known

    return {
        "fidelity_total": fidelity_total,
        "fidelity_by_acc": fidelity_by_acc,
        "arriva": arriva,
        "expleo_sw": expleo_sw,
        "susan_pensions": susan_pensions,
        "susan_fidelity_sipp": susan_fidelity_sipp_vals,
        "barclays_by_month": barclays_by_month,
        "liam_fid_total": liam_fid_total,
        "jayne_fid_total": jayne_fid_total,
        "estimates_idx": est_idx,
        "_account_summary_path": account_summary_path,
    }
