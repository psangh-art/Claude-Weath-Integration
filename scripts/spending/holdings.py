"""What is held, and what each holding was worth month by month — settled
holdings plus pending orders, and the per-account monthly valuation series.

Extracted from spending_summary.py on 2026-07-19 — that file had grown to 4,020
lines. Behaviour is unchanged: the code below is the original, moved verbatim.
"""
import csv
import io
import os

import pandas as pd

from .constants import ACCOUNT_OWNER


def apply_pending_holdings(holdings: dict, pending_path: str) -> dict:
    """
    Adds net units from pending (not-yet-settled) Buy/Sell orders on top of a
    holdings dict. AccountSummary and completed-transaction history only
    reflect *settled* trades, so a pending sell/buy that hasn't cleared yet
    wouldn't otherwise show up until the next export. A pending export has
    Completion date == 'Pending' on every row and Status values of
    Priced/Ordered rather than Completed — it's read here without a Status
    filter since none of its rows are ever 'Completed'.

    Returns a new dict; does not mutate the input.
    """
    if not pending_path or not os.path.exists(pending_path):
        return holdings

    FAMILY_ACCS = set(ACCOUNT_OWNER.keys())
    with open(pending_path, encoding="utf-8-sig") as f:
        content = f.read()
    lines = content.replace("\r", "").split("\n")
    try:
        start = next(i for i, l in enumerate(lines) if l.startswith("Order date"))
    except StopIteration:
        return holdings

    result = dict(holdings)
    reader = csv.DictReader(io.StringIO("\n".join(lines[start:])))
    for row in reader:
        t = row.get("Transaction type", "").strip()
        if t not in ("Buy", "Sell"):
            continue
        acc = row.get("Account Number", "").strip()
        if acc not in FAMILY_ACCS:
            continue
        inv = row.get("Investments", "").strip().strip('"')
        if not inv or inv == "Cash":
            continue
        try:
            qty = float(row.get("Quantity", "0") or 0)
        except (ValueError, TypeError):
            continue
        delta = qty if t == "Buy" else -qty
        result[(acc, inv)] = result.get((acc, inv), 0) + delta

    return {k: v for k, v in result.items() if v > 0.001}


def build_holdings(fidelity_path: str, account_summary_path: str = None, pending_path: str = None) -> dict:
    """
    Returns dict: (account, fund_name) -> net units held.

    If account_summary_path is provided (AccountSummary CSV export), uses that
    as the authoritative source — it reflects exact holdings at export date.
    Falls back to calculating net units from transaction history otherwise.

    If pending_path is provided, pending Buy/Sell orders are layered on top
    via apply_pending_holdings — see that function for why.
    """
    FAMILY_ACCS = set(ACCOUNT_OWNER.keys())
    holdings = None

    # ── Primary: AccountSummary CSV ───────────────────────────────────────────
    if account_summary_path and os.path.exists(account_summary_path):
        with open(account_summary_path, encoding="utf-8-sig") as f:
            content = f.read()
        lines = content.replace("\r", "").split("\n")

        # Find the last "Type,Holdings,Account number" header — that's the detail section
        header_idx = None
        for i, l in enumerate(lines):
            if l.startswith("Type,Holdings,Account number"):
                header_idx = i

        if header_idx is not None:
            reader = csv.DictReader(iter(lines[header_idx:]))
            holdings = {}
            for row in reader:
                if row.get("Type", "").strip() != "Asset":
                    continue
                acc = row.get("Account number", "").strip()
                if acc not in FAMILY_ACCS:
                    continue
                fund = row.get("Holdings", "").strip().strip('"')
                try:
                    qty = float(row.get("Quantity", "0") or 0)
                except (ValueError, TypeError):
                    continue
                if qty > 0:
                    holdings[(acc, fund)] = qty

    # ── Fallback: calculate from transaction history ───────────────────────────
    if holdings is None:
        with open(fidelity_path) as f:
            content = f.read()
        lines = content.replace("\r", "").split("\n")
        start = next(i for i, l in enumerate(lines) if l.startswith("Order date"))
        reader = csv.DictReader(io.StringIO("\n".join(lines[start:])))
        rows = [r for r in reader if r.get("Status", "").strip() == "Completed"]

        from collections import defaultdict
        raw_holdings = defaultdict(float)
        for r in rows:
            t = r["Transaction type"].strip()
            acc = r["Account Number"].strip()
            inv = r["Investments"].strip().strip('"')
            try:
                qty = float(r["Quantity"])
            except (ValueError, KeyError):
                continue
            if t == "Buy":
                raw_holdings[(acc, inv)] += qty
            elif t == "Sell":
                raw_holdings[(acc, inv)] -= qty
        holdings = {k: v for k, v in raw_holdings.items() if v > 0.001}

    return apply_pending_holdings(holdings, pending_path)

def _read_account_summary_rows(account_summary_path):
    """Yield all rows from AccountSummary CSV as dicts."""
    if not account_summary_path or not os.path.exists(account_summary_path):
        return
    with open(account_summary_path, encoding="utf-8-sig") as f:
        content = f.read()
    lines = content.replace("\r", "").split("\n")
    try:
        hi = max(i for i, l in enumerate(lines) if l.startswith("Type,Holdings,Account number"))
        import io as _io_helper
        for row in csv.DictReader(_io_helper.StringIO("\n".join(lines[hi:]))):
            yield row
    except (ValueError, StopIteration):
        return


def build_acc_holdings(account_summary_path, anchors, fidelity_path=None, inc_income_df=None):
    """
    Returns dict: { acc: { fund: {'value', 'units', 'price_anchor', 'monthly_values'} } }
    Only accumulation (Acc/ACC) funds in family accounts.
    monthly_values: { pd.Period -> value } based on cumulative units × price.

    Units and price are anchored at the broker snapshot's own month
    (anchors.data_month) and walked backwards/forwards from there.
    """
    FAMILY_ACCS = set(ACCOUNT_OWNER.keys())
    result = {}
    if not account_summary_path or not os.path.exists(account_summary_path):
        return result

    with open(account_summary_path, encoding="utf-8-sig") as f:
        content = f.read()
    lines = content.replace("\r", "").split("\n")
    header_idx = max(i for i, l in enumerate(lines) if l.startswith("Type,Holdings,Account number"))
    reader = csv.DictReader(io.StringIO("\n".join(lines[header_idx:])))
    for row in reader:
        if row.get("Type", "").strip() != "Asset": continue
        acc = row.get("Account number", "").strip()
        if acc not in FAMILY_ACCS: continue
        fund = row.get("Holdings", "").strip().replace('"', '')
        if " Acc" not in fund and " ACC" not in fund: continue
        try:
            val = float(row.get("Value (£)", "0").replace(",", ""))
            qty = float(row.get("Quantity", "0").replace(",", ""))
        except (ValueError, TypeError):
            continue
        if acc not in result: result[acc] = {}
        price_anchor = val / qty if qty > 0 else 0
        result[acc][fund] = {"value": val, "units": qty, "price_anchor": price_anchor, "monthly_values": {}}

    if not fidelity_path or not os.path.exists(fidelity_path):
        return result

    # ── Build price history and unit history from transaction CSV ──────────────
    with open(fidelity_path) as f:
        fid_lines = f.read().replace("\r", "").split("\n")
    fid_start = next(i for i, l in enumerate(fid_lines) if l.startswith("Order date"))
    fid_reader = csv.DictReader(io.StringIO("\n".join(fid_lines[fid_start:])))
    fid_rows = [r for r in fid_reader if r.get("Status", "").strip() == "Completed"]

    # Price observations: (fund, period) -> price
    from collections import defaultdict
    price_obs = defaultdict(dict)   # fund -> {period: price}
    unit_changes = defaultdict(lambda: defaultdict(float))  # (acc, fund) -> {period: net_units}

    for r in fid_rows:
        t = r["Transaction type"].strip()
        if t not in ("Buy", "Sell"): continue
        acc = r["Account Number"].strip()
        inv = r["Investments"].strip().replace('"', '')
        if " Acc" not in inv and " ACC" not in inv: continue
        ds = (r.get("Completion date") or r["Order date"]).strip()
        if not ds or ds.lower() == "pending": ds = r["Order date"].strip()
        try:
            dt = pd.to_datetime(ds, dayfirst=True)
            period = dt.to_period("M")
            price = float(r["Price per unit"]) if r.get("Price per unit") else 0
            qty = float(r["Quantity"]) if r.get("Quantity") else 0
        except (ValueError, TypeError):
            continue
        if price > 0:
            price_obs[inv][period] = price
        if acc in FAMILY_ACCS:
            if t == "Buy":
                unit_changes[(acc, inv)][period] += qty
            else:
                unit_changes[(acc, inv)][period] -= qty

    # ── Build monthly value for each (acc, fund) ────────────────────────────────
    ANCHOR = anchors.data_month          # month the broker snapshot was taken
    YEAR_START = anchors.jan
    all_periods = list(anchors.months)

    # Annual yield overrides for funds where transaction history is sparse
    # For Acc funds, this represents total return (income reinvested into NAV)
    # Sources: Fidelity key statistics pages
    FUND_ANNUAL_YIELD = {
        "Aegon High Yield Bond B Acc":                    0.0732,  # 7.32% distribution yield
        "Schroder High Yield Opportunities Fund Z Acc":   0.0769,  # 7.69% distribution yield
        "WS Guinness Global Energy Fund I Acc":           0.0235,  # 2.35% historic yield
    }

    def interpolate_price(fund, period, price_anchor):
        """Get price for a given period using observations + interpolation/extrapolation."""
        obs = dict(price_obs.get(fund, {}))
        obs[ANCHOR] = price_anchor   # snapshot month is ground truth
        periods_sorted = sorted(obs.keys())
        if not periods_sorted:
            return price_anchor
        if period in obs:
            return obs[period]
        before = [p for p in periods_sorted if p <= period]
        after  = [p for p in periods_sorted if p >= period]
        if before and after:
            p0, p1 = before[-1], after[0]
            if p0 == p1: return obs[p0]
            w = (period - p0).n / max((p1 - p0).n, 1)
            return obs[p0] + w * (obs[p1] - obs[p0])
        elif before:
            p0 = before[-1]
            # Always use known annual yield for forward extrapolation — never use
            # short-term price trends which can be negative
            annual = FUND_ANNUAL_YIELD.get(fund, 0.05)
            monthly_growth = max(0.0, (1 + annual) ** (1/12) - 1)
            months_fwd = (period - p0).n
            return obs[p0] * ((1 + monthly_growth) ** months_fwd)
        else:
            # Only future observations — extrapolate backward using annual yield
            p0 = after[0]
            annual = FUND_ANNUAL_YIELD.get(fund, 0.05)
            monthly_growth = max(0.0, (1 + annual) ** (1/12) - 1)
            months_back = (p0 - period).n
            return obs[p0] / ((1 + monthly_growth) ** months_back)

    for acc, funds in result.items():
        for fund, data in funds.items():
            price_anchor = data["price_anchor"]
            units_anchor = data["units"]

            # Work backwards/forwards from the snapshot month to get units each month
            monthly_vals = {}
            monthly_units = {}   # units held at START of each month

            # Go forward from the snapshot month
            running = units_anchor
            for p in sorted(all_periods):
                if p >= ANCHOR:
                    delta = unit_changes.get((acc, fund), {}).get(p, 0)
                    if p > ANCHOR:
                        running += delta
                    price = interpolate_price(fund, p, price_anchor)
                    monthly_vals[p] = round(running * price)
                    monthly_units[p] = running
            # Go backward from the snapshot month
            running = units_anchor
            # Find earliest buy date for this fund/account from unit_changes
            fund_unit_changes = unit_changes.get((acc, fund), {})
            # If there are no buy transactions, the fund was held before our data window
            # Use the start of the year (fund was already held)
            buys = [p for p, d in fund_unit_changes.items() if d > 0]
            earliest_buy = min(buys) if buys else YEAR_START

            for p in sorted([pp for pp in all_periods if pp < ANCHOR], reverse=True):
                delta = fund_unit_changes.get(p, 0)
                running -= delta  # undo this month's purchase
                if running <= 0 and p < earliest_buy:
                    # Fund genuinely not held before this point
                    monthly_vals[p] = 0
                    monthly_units[p] = 0
                else:
                    price = interpolate_price(fund, p, price_anchor)
                    monthly_vals[p] = round(running * price)
                    monthly_units[p] = running

            # Appreciation = Inc dividend ppu × Acc units held at start of month
            # ppu derived from: total Inc income paid / total Inc units held that month
            # For months with no Inc data, fall back to interpolated price change
            price_appreciation = {}

            # Build Inc ppu lookup: fund_base -> month -> ppu (pence per unit)
            inc_ppu = {}  # { (fund_base, period) -> ppu }
            INC_EQUIV = {
                "Aegon High Yield Bond B Acc":                  "Aegon High Yield Bond B Inc",
                "Schroder High Yield Opportunities Fund Z Acc": "Schroder High Yield Opportunities Fund Z Inc",
            }
            inc_fund_name = INC_EQUIV.get(fund)
            if inc_fund_name and inc_income_df is not None and not inc_income_df.empty:
                inc_df = inc_income_df[inc_income_df["fund"] == inc_fund_name]
                # Total Inc units across all accounts (from AccountSummary)
                total_inc_units = sum(
                    float(str(row.get("Quantity","0")).replace(",",""))
                    for row in _read_account_summary_rows(account_summary_path)
                    if row.get("Holdings","").strip() == inc_fund_name
                    and row.get("Type","").strip() == "Asset"
                )
                if total_inc_units > 0:
                    monthly_inc = inc_df.groupby("month")["amount"].sum()
                    for m_period, total_inc_amt in monthly_inc.items():
                        ppu = total_inc_amt / total_inc_units  # in £
                        inc_ppu[(fund, m_period)] = ppu

            sorted_periods = sorted(monthly_vals.keys())
            for i, p in enumerate(sorted_periods):
                if i == 0:
                    price_appreciation[p] = 0
                else:
                    p_prev = sorted_periods[i - 1]
                    units_start = monthly_units.get(p_prev, 0)
                    # Use Inc ppu if available for this period
                    ppu = inc_ppu.get((fund, p))
                    if ppu is not None and units_start > 0:
                        price_appreciation[p] = round(units_start * ppu)
                    else:
                        # Fallback: interpolated price change × units — floor at 0
                        price_now  = interpolate_price(fund, p, price_anchor)
                        price_prev = interpolate_price(fund, p_prev, price_anchor)
                        price_appreciation[p] = max(0, round(units_start * (price_now - price_prev)))

            data["monthly_values"] = monthly_vals
            data["price_appreciation"] = price_appreciation

    return result
