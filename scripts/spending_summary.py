#!/usr/bin/env python3
"""
Spending Summary Generator
--------------------------
Usage:
    python spending_summary.py                                    # uses defaults
    python spending_summary.py amex.csv barclays.csv fidelity.csv
    python spending_summary.py amex.csv barclays.csv fidelity.csv output.xlsx

Reads Amex (activity.csv), Barclays (data.csv) and Fidelity
(TransactionHistory.csv) and produces a categorised monthly spending summary
plus a Fidelity income-by-account section as a formatted Excel file.
"""

import sys
import os
from types import SimpleNamespace
import pandas as pd
from openpyxl import Workbook

# The reading/categorising/pivoting half of this report lives in the `spending`
# package (split out 2026-07-19 — this file was 4,020 lines). What stays here is
# the Excel writer and the CLI. Imported by name rather than `import *` so every
# cross-module use is greppable.
from spending.constants import ACCOUNT_LABELS
from spending.anchors import resolve_anchors
from spending.loaders import load_amex, load_barclays, load_fidelity_income, load_history
from spending.holdings import build_acc_holdings, build_holdings
from spending.pivots import (build_account_fund_pivot, build_fidelity_pivot,
                             build_spending_pivot, build_summary_data,
                             estimate_future_months)
from spending.sheet.style import HIST_MAP, make_col_fill, make_val_font
from spending.sheet.sheet_assets import write_asset_rows
from spending.sheet.sheet_spending import write_spending
from spending.sheet.sheet_income import write_income
from spending.sheet.sheet_totals import write_totals
from spending.sheet.sheet_targets import write_targets
from spending.sheet.sheet_finish import finish_workbook

# Windows' default console codepage (cp1252) can't encode characters like
# "→" used in console-only status prints below — reconfigure to UTF-8 so a
# cosmetic console print can't crash the script after the workbook is
# already written (same class of bug fixed in verify_pipeline.py).
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass











# ── Excel styles helper ────────────────────────────────────────────────────────








# ── Write Excel ────────────────────────────────────────────────────────────────
def write_excel(spend_pivot, actual_months, future_months, fid_pivot,
                acc_fund_map, holdings, summary_data, acc_holdings, anchors,
                output_path, reimbursements=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "Wealth Summary"

    all_months = actual_months + future_months
    spend_months = all_months
    fid_months   = all_months

    spend_month_labels = [m.strftime("%b %Y") for m in all_months]
    fid_month_labels   = [m.strftime("%b %Y") for m in all_months]
    actual_set = set(actual_months)  # months with real CSV data (Jan–Apr)

    # ── History boundary ───────────────────────────────────────────────────────
    # Previous Data owns everything BEFORE this month; live files own this month
    # and everything after. It sits one month past the end of load_history(), so
    # it moves by itself when that pinned series is extended.
    HISTORY_CUTOFF = anchors.wealth_cutoff

    def is_history_month(m):
        """True if this month belongs to Previous Data (read-only)."""
        return m < HISTORY_CUTOFF

    def live_and_hist_safe(live_dict, hist_key_str):
        """Merge values with strict boundary:
        - Months BEFORE HISTORY_CUTOFF: Previous Data only (never overwritten by live)
        - Months FROM HISTORY_CUTOFF: live CSV files only
        """
        merged = {}
        hist_series = history.get(HIST_MAP.get(hist_key_str, ""), {})
        for m in all_months:
            if is_history_month(m):
                # History owns this month — ignore live value
                if m in hist_series:
                    merged[m] = hist_series[m]
                # If no history value exists, leave blank (don't backfill from live)
            else:
                # Live owns this month
                v = live_dict.get(m)
                if v:
                    merged[m] = v
        return merged

    def projection_with_hist_override(proj_dict, hist_key_str):
        """For formula-projected series (Expleo SW, Susan pensions, Cars):
        - Use hardcoded history value if it exists for a month (actuals take priority)
        - Otherwise use the live projection (covers all months incl. June+)
        """
        merged = {}
        hist_series = history.get(HIST_MAP.get(hist_key_str, ""), {})
        for m in all_months:
            if m in hist_series:
                merged[m] = hist_series[m]
            elif m in proj_dict:
                merged[m] = proj_dict[m]
        return merged

    col_fill = make_col_fill(actual_set)
    val_font = make_val_font(actual_set)

    n_sum_cols = 3 + len(all_months)  # Label | blank | blank | month cols (live)

    # ── Historic months (from Previous_Data CSV — used for calcs, NOT displayed) ─
    history = load_history()  # hardcoded — no file needed

    # Everything the phases hand between each other. Deliberately a plain
    # namespace rather than a dataclass: the phase bodies are the ORIGINAL code
    # moved verbatim, and their generated unpack/repack lines read and write these
    # attributes by name (see spending/sheet/).
    ctx = SimpleNamespace(
        wb=wb, ws=ws, output_path=output_path,
        # inputs
        spend_pivot=spend_pivot, fid_pivot=fid_pivot, acc_fund_map=acc_fund_map,
        holdings=holdings, summary_data=summary_data, acc_holdings=acc_holdings,
        anchors=anchors, reimbursements=reimbursements,
        actual_months=actual_months, future_months=future_months,
        # derived month sets
        all_months=all_months, spend_months=spend_months, fid_months=fid_months,
        spend_month_labels=spend_month_labels, fid_month_labels=fid_month_labels,
        actual_set=actual_set, n_sum_cols=n_sum_cols,
        # history boundary + the closures over it
        history=history, HISTORY_CUTOFF=HISTORY_CUTOFF,
        is_history_month=is_history_month,
        live_and_hist_safe=live_and_hist_safe,
        projection_with_hist_override=projection_with_hist_override,
        col_fill=col_fill, val_font=val_font,
    )

    # Order matters: each phase depends on row numbers the previous ones set.
    write_asset_rows(ctx)     # summary table, accounts, pensions, assets, calculations
    write_spending(ctx)       # Section 1 + reimbursements
    write_income(ctx)         # Sections 2 and 3
    write_totals(ctx)         # cross-section totals, once the row numbers are known
    write_targets(ctx)        # risk metrics + targets tables
    finish_workbook(ctx)      # widths, split onto the Targets tab, note row, save




# ── Main ───────────────────────────────────────────────────────────────────────
def main():
    args = sys.argv[1:]
    amex_path    = args[0] if len(args) > 0 else "activity.csv"
    bar_path     = args[1] if len(args) > 1 else "data.csv"
    fid_path     = args[2] if len(args) > 2 else "TransactionHistory.csv"
    summary_path = args[3] if len(args) > 3 else "AccountSummary.csv"
    output_path  = args[4] if len(args) > 4 else "spending_summary.xlsx"
    # Optional: a pending-orders export (every row has Completion date ==
    # 'Pending') — net units from pending Buy/Sell orders are layered on top
    # of the settled holdings figure. Not required; omit to skip.
    pending_path = args[5] if len(args) > 5 else None

    # Partial inputs are allowed (user request 2026-07-18): build with whatever
    # sources are present. Only bail if NONE of the three are — nothing to do.
    def _have(p):
        return bool(p) and os.path.exists(p)
    if not any(_have(p) for p in (amex_path, bar_path, fid_path)):
        print("Error: none of the spending sources found (Amex / Barclays / Fidelity).")
        sys.exit(1)
    missing = [name for name, p in (('Amex', amex_path), ('Barclays', bar_path),
                                    ('Fidelity', fid_path)) if not _have(p)]
    if missing:
        print(f"  Partial inputs — building without: {', '.join(missing)}")

    print(f"  Amex:     {amex_path if _have(amex_path) else '(not provided)'}")
    print(f"  Barclays: {bar_path if _have(bar_path) else '(not provided)'}")
    print(f"  Fidelity: {fid_path if _have(fid_path) else '(not provided)'}")
    if pending_path:
        print(f"  Pending:  {pending_path}")
    print(f"  Output:   {output_path}\n")

    # Every month boundary this run depends on, derived from the inputs.
    anchors = resolve_anchors(summary_path)
    print(f"Anchors: {anchors.describe()}")
    for _w in anchors.warnings():
        print(f"  WARNING: {_w}")
    print()

    print("Loading Amex...")
    amex_df = load_amex(amex_path)
    print(f"  {len(amex_df)} transactions\n")

    print("Loading Barclays...")
    bar_df = load_barclays(bar_path)
    print(f"  {len(bar_df)} transactions\n")

    print("Loading Fidelity income...")
    fid_df = load_fidelity_income(fid_path, anchors.year)
    print(f"  {len(fid_df)} income entries\n")

    print("Building holdings...")
    settled_holdings = build_holdings(fid_path, summary_path)
    holdings = build_holdings(fid_path, summary_path, pending_path)
    src = "AccountSummary.csv" if os.path.exists(summary_path) else "transaction history"
    print(f"  {len(holdings)} positions (from {src})")
    if pending_path and os.path.exists(pending_path):
        changed = sorted(k for k in set(settled_holdings) | set(holdings)
                          if round(settled_holdings.get(k, 0), 2) != round(holdings.get(k, 0), 2))
        if changed:
            print(f"  {len(changed)} position(s) adjusted by pending orders:")
            for acc, fund in changed:
                before = settled_holdings.get((acc, fund), 0)
                after = holdings.get((acc, fund), 0)
                print(f"    {acc} — {fund}: {before:,.2f} -> {after:,.2f}")
        else:
            print("  No pending Buy/Sell orders affected current holdings")
    print()

    print("Building pivots...")
    spend_pivot, spend_months       = build_spending_pivot(amex_df, bar_df, anchors)
    fid_pivot, fid_months           = build_fidelity_pivot(fid_df, anchors)
    acc_fund_map, acc_fund_months   = build_account_fund_pivot(fid_df)

    # Always show the full reporting year — the transaction files only cover a
    # 60-day window, the pinned history covers the early months, and everything
    # in between or beyond is estimated.
    all_months = list(anchors.months)

    # Split the year into complete actuals vs months that must be estimated.
    # Done PER PIVOT, from that pivot's own coverage: the spending sources and the
    # Fidelity export cover different months, and a bank export that is missing
    # entirely must not make its months read as actual zeroes (that once left the
    # salary row blank for a month and dragged the median down for every other).
    spend_actual, spend_future = anchors.split_months(spend_months)
    fid_actual, fid_future = anchors.split_months(fid_months)
    actual_months, future_months = anchors.split_months(
        sorted(set(spend_months) | set(fid_months)))

    spend_months = fid_months = acc_fund_months = all_months
    spend_pivot = spend_pivot.reindex(columns=all_months + ["Total"], fill_value=0)
    if not fid_pivot.empty:
        fid_pivot = fid_pivot.reindex(columns=all_months + ["Total"], fill_value=0)
        fid_pivot["Total"] = fid_pivot[all_months].sum(axis=1)
    for person in acc_fund_map:
        for acc in acc_fund_map[person]:
            df = acc_fund_map[person][acc]
            df = df.reindex(columns=all_months + ["Total"], fill_value=0)
            df["Total"] = df[all_months].sum(axis=1)
            acc_fund_map[person][acc] = df

    def _months(ms):
        return ", ".join(str(m) for m in ms) if ms else "none"
    print(f"  Anchors:  {anchors.describe()}")
    print(f"  Spend   — actual: {_months(spend_actual)} | estimated: {_months(spend_future)}")
    print(f"  Income  — actual: {_months(fid_actual)} | estimated: {_months(fid_future)}")

    # Full year Jan–Dec (actuals + future, always 12 months)
    full_months = all_months

    # ── Inject missing stocks BEFORE estimation ───────────────────────────────
    # Stocks held (from AccountSummary) but with no 2026 income transactions
    # Uses confirmed dividend pence-per-share × units for accurate estimates
    STOCK_INJECT = {
        "LEGAL & GENERAL GROUP,ORD GBP0.025(LGEN)": [
            (6, 15.36),   # Final: pay Jun 4 (ex Apr 23)
            (9, 6.12),    # Interim: pay Sep 25 (ex Aug 20)
        ],
    }
    FUND_NAME_MAP = {
        "LEGAL & GENERAL GROUP,ORD GBP0.025(LGEN)": "LEGAL & GENERAL GROUP, ORD GBP0.025 (LGEN)",
    }
    # Track injected fund names so estimate_future_months can skip them
    injected_funds = set()

    for person in acc_fund_map:
        for acc in list(acc_fund_map[person].keys()):
            for summary_name, payments in STOCK_INJECT.items():
                income_name = FUND_NAME_MAP.get(summary_name, summary_name)
                fund_df = acc_fund_map[person][acc]
                if income_name in fund_df.index and fund_df.loc[income_name].sum() > 0:
                    continue
                units = holdings.get((acc, summary_name), 0)
                if units <= 0:
                    continue
                new_row = {m: 0 for m in full_months + ["Total"]}
                for pay_month, ppm in payments:
                    for m in full_months:
                        if m.month == pay_month and m.year == anchors.year:
                            new_row[m] = round(units * ppm / 100)
                new_row["Total"] = sum(new_row[m] for m in full_months)
                new_df = pd.DataFrame([new_row], index=[income_name])
                new_df.index.name = "fund"
                fund_df = fund_df.reindex(columns=full_months + ["Total"], fill_value=0)
                acc_fund_map[person][acc] = pd.concat([fund_df, new_df]).sort_index()
                injected_funds.add(income_name)
                print(f"  Injected {income_name} → {acc} ({units:,.0f} units, "
                      f"Jun=£{round(units*payments[0][1]/100):,}, "
                      f"Sep=£{round(units*payments[1][1]/100):,})")

    # Apply estimates to all pivots (injected funds are skipped)
    spend_pivot    = estimate_future_months(spend_pivot, spend_actual, spend_future, anchors)
    if not fid_pivot.empty:
        fid_pivot  = estimate_future_months(fid_pivot, fid_actual, fid_future, anchors)
    for person in acc_fund_map:
        for acc in acc_fund_map[person]:
            acc_fund_map[person][acc] = estimate_future_months(
                acc_fund_map[person][acc], fid_actual, fid_future, anchors,
                skip_funds=injected_funds)

    # ── Partial-month provisional fallback ────────────────────────────────────
    # The transaction export only covers the last 60 days, and the snapshot is
    # usually taken mid-month, so a monthly distribution due later in the partial
    # month isn't in the file yet. Carry the previous month's figure as a
    # placeholder — it is replaced by the real one on the next export.
    partial = anchors.partial_month
    prev = (partial - 1) if partial else None
    provisional_funds = set()
    if partial:
        for person in acc_fund_map:
            for acc, fund_df in acc_fund_map[person].items():
                if partial in fund_df.columns and prev in fund_df.columns:
                    for fund in fund_df.index:
                        if fund in injected_funds:
                            continue
                        cur_val = fund_df.loc[fund, partial]
                        prev_val = fund_df.loc[fund, prev]
                        if (pd.isna(cur_val) or cur_val == 0) and prev_val and prev_val > 0:
                            fund_df.loc[fund, partial] = prev_val
                            provisional_funds.add((acc, fund))
    if provisional_funds:
        print(f"  {partial}: provisional estimates on {len(provisional_funds)} fund rows "
              f"(carried from {prev})")

    # Rebuild fid_pivot account totals after estimation + injection
    for person in acc_fund_map:
        for acc in acc_fund_map[person]:
            fund_df = acc_fund_map[person][acc]
            acc_total = fund_df[full_months].sum()
            if acc in fid_pivot.index:
                for m in full_months:
                    # Only overwrite months the pinned income history doesn't own
                    if m >= anchors.hist_cutoff:
                        fid_pivot.loc[acc, m] = acc_total[m]
                # Always update Total
                fid_pivot.loc[acc, "Total"] = fid_pivot.loc[acc, full_months].sum()

    all_months = full_months
    spend_months = fid_months = acc_fund_months = all_months

    print(f"  Fidelity: {len(fid_pivot)} accounts\n")

    print("Writing Excel...")
    # The sheet shades months as actual vs estimated. Everything up to and
    # including the snapshot month is shown as actual (the partial month has real
    # data in it, just not all of it); everything after is projection.
    actual_for_excel = [m for m in anchors.months if m <= anchors.data_month]
    future_for_excel = [m for m in anchors.months if m > anchors.data_month]
    full_months = actual_for_excel + future_for_excel  # the whole year, in order

    summary_data = build_summary_data(summary_path, full_months, anchors)
    acc_holdings = build_acc_holdings(summary_path, anchors, fid_path, inc_income_df=fid_df)
    # Expense Reimbursements: Royal Mail or Expleo credits under £1,000 (excl. salary)
    REIMBURSEMENT_KEYWORDS = ["ROYAL MAIL", "EXPLEO"]

    # One-time backfill from expenses.csv (Jan-Jun 2026 full year feed) — these
    # months are no longer covered by the live data.csv (60-day window), so
    # hardcode them here. Going forward, data.csv provides ongoing updates.
    REIMBURSEMENT_BACKFILL = [
        {"date": pd.Timestamp("2026-02-04"), "amount": 36.80,
         "memo": "ROYAL MAIL            \t1612 2000284071 K BGC"},
        {"date": pd.Timestamp("2026-03-06"), "amount": 68.24,
         "memo": "EXPLEO UK LIMITED     \t1160304622 BGC\t"},
        {"date": pd.Timestamp("2026-03-20"), "amount": 163.03,
         "memo": "EXPLEO UK LIMITED     \t1161123617 BGC\t"},
        {"date": pd.Timestamp("2026-04-07"), "amount": 342.97,
         "memo": "EXPLEO UK LIMITED     \t1162170846 BGC\t"},
    ]

    reimbursements = []
    seen_keys = set()
    for entry in REIMBURSEMENT_BACKFILL:
        key = (entry["date"].date(), round(entry["amount"], 2))
        if key not in seen_keys:
            seen_keys.add(key)
            reimbursements.append({
                "date": entry["date"],
                "month": entry["date"].to_period("M"),
                "amount": entry["amount"],
                "memo": entry["memo"].strip()
            })

    # Live extraction from data.csv — covers ongoing/future updates
    try:
        bar_df_r = pd.read_csv("data.csv")
        bar_df_r["Amount"] = pd.to_numeric(bar_df_r["Amount"], errors="coerce").fillna(0)
        bar_df_r["Date"] = pd.to_datetime(bar_df_r["Date"], dayfirst=True, errors="coerce")
        bar_df_r = bar_df_r.dropna(subset=["Date"])
        for _, r in bar_df_r.iterrows():
            memo = str(r.get("Memo", "")).upper()
            amt = r["Amount"]
            if amt > 0 and amt < 1000 and any(k in memo for k in REIMBURSEMENT_KEYWORDS):
                if "EUKPT" in memo:
                    continue
                key = (r["Date"].date(), round(amt, 2))
                if key not in seen_keys:
                    seen_keys.add(key)
                    reimbursements.append({
                        "date": r["Date"],
                        "month": r["Date"].to_period("M"),
                        "amount": amt,
                        "memo": str(r.get("Memo", "")).strip()
                    })
    except Exception:
        pass

    write_excel(spend_pivot, actual_for_excel, future_for_excel, fid_pivot,
                acc_fund_map, holdings, summary_data, acc_holdings, anchors,
                output_path, reimbursements=reimbursements)
    print(f"Done → {output_path}\n")

    # Console preview of Fidelity section
    if not fid_pivot.empty:
        col_w, num_w = 32, 10
        hdr = f"{'Account':<{col_w}} {'Total':>{num_w}}" + "".join(
            f"  {m.strftime('%b %y'):>{num_w}}" for m in fid_months)
        print("Fidelity Income by Account")
        print("-" * len(hdr))
        print(hdr)
        print("-" * len(hdr))
        for acc in fid_pivot.index:
            label = ACCOUNT_LABELS.get(acc, acc)
            total = int(round(fid_pivot.loc[acc, "Total"]))
            monthly = [int(round(fid_pivot.loc[acc, m])) for m in fid_months]
            vals = f"{total:>{num_w},}" + "".join(f"  {v:>{num_w},}" for v in monthly)
            print(f"{label:<{col_w}} {vals}")
        gtotal = int(round(fid_pivot["Total"].sum()))
        gmonthly = [int(round(sum(fid_pivot[m]))) for m in fid_months]
        print("-" * len(hdr))
        gvals = f"{gtotal:>{num_w},}" + "".join(f"  {v:>{num_w},}" for v in gmonthly)
        print(f"{'TOTAL':<{col_w}} {gvals}")


if __name__ == "__main__":
    main()
