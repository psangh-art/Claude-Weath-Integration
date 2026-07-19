"""Investment Dashboard — data layer (Phase 1).

Reads the master workbook (Stocks_Buy_Strategy.xlsx) + history.db and emits the
JSON the dashboard front end renders: overview.json, portfolio.json, historic.json.

KEY CONSTRAINT (why this file computes rather than reads): the workbook is written
by openpyxl (the pipeline), so every FORMULA cell — all the %/P&L columns, and every
stock's googlefinance() current price — has NO cached value and reads as None locally.
Only literals survive: identity, Holdings (£), Alert Low/High/Source, and the History
transaction inputs (Qty, Buy/Sell price, Fees, Cost, Proceeds). So we take those
literals, pull the freshest CURRENT PRICE per ticker from history.db (the captured
TradingView last price), and derive every percentage ourselves. Never trust a formula
cell's value here.
"""
import os
import json
import sqlite3
import datetime
import openpyxl

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(SCRIPT_DIR)
WORKBOOK = os.path.join(os.path.expanduser('~'), 'Downloads', 'Stocks_Buy_Strategy.xlsx')
HISTORY_DB = os.path.join(REPO, 'data', 'history.db')
OUT_DIR = os.path.join(SCRIPT_DIR, 'dashboard_app', 'data')

# --- Investments sheet columns (1-indexed; header row 2, data from row 4) ---
I_NAME, I_TICKER, I_HOLDINGS, I_CURPRICE = 3, 4, 5, 10
I_ALERT_LOW, I_ALERT_LOW_SRC, I_ALERT_HIGH = 13, 14, 16
I_TV = 37
I_TYPE = 39   # 'Type' — Short Term / Long Term (added 2026-07-17)
# --- Income Funds columns (header row 4, data from row 5) ---
F_NAME, F_CURVAL, F_DIVYLD, F_MONTHLY_REV, F_ANNUAL_REV = 1, 3, 6, 9, 10
# --- History columns (header row 1) ---
H_INV, H_ACCT, H_WRAP, H_BUY, H_SELL = 1, 2, 3, 4, 5
H_COST, H_PROCEEDS = 10, 12
# --- Wealth Summary: month headers on row 3 from col D; investable account rows ---
WS_MONTH_ROW = 3
WS_INVEST_ROWS = [5, 6, 7, 8, 9, 12]     # Investment Account(s) + ISAs + Junior ISA
WS_CASH_ROWS = [10, 11, 13]              # Fidelity Cash Accounts
# The whole 'Fidelity accounts' block. Total Portfolio Value is this block's total,
# so the dashboard agrees with the Wealth Summary tab and the Finance Google Sheet
# (user decision 2026-07-19) — pensions/SIPPs sit in their own blocks below it and
# are deliberately NOT in the portfolio figure.
WS_FIDELITY_ROWS = WS_INVEST_ROWS + WS_CASH_ROWS
WS_MONTHLY_INCREASE_ROW = 46            # 'Monthly Investment Increase' (literal)
# --- Stocks of Interest section-table columns (A..Q) ---
SOI_STOCK, SOI_TICKER, SOI_PATTERN, SOI_ALOW, SOI_AHIGH = 1, 2, 3, 5, 7
SOI_CHARTNOTE, SOI_RATING, SOI_HOLDINGS, SOI_TARGET, SOI_NOTES, SOI_UPDATED, SOI_TV = 11, 12, 13, 14, 15, 16, 17
# Dashboard watchlist = strictly the 'AT LOWER BOUNDARY — within 5% of alert low'
# section band of Stocks of Interest (user decision 2026-07-17). Only this band.
SOI_WATCHLIST_BAND = 'AT LOWER BOUNDARY'
# --- Base Data columns (header row 2) ---
BD_TICKER, BD_NAME, BD_PE, BD_DIV_PENCE, BD_DIVYLD, BD_EXDIV = 1, 2, 6, 8, 9, 11


def _num(v):
    return v if isinstance(v, (int, float)) else None


def _latest_prices():
    """ticker(upper) -> captured last price from the most recent history.db run."""
    prices = {}
    if not os.path.exists(HISTORY_DB):
        return prices
    con = sqlite3.connect(HISTORY_DB)
    row = con.execute('SELECT MAX(run_id) FROM chart_snapshots').fetchone()
    if row and row[0] is not None:
        for tkr, price in con.execute(
                'SELECT ticker, price FROM chart_snapshots WHERE run_id=? AND price IS NOT NULL',
                (row[0],)):
            if tkr:
                prices[str(tkr).strip().upper()] = price
    con.close()
    return prices


def _price_changes():
    """ticker(upper) -> {'change', 'change_pct'} day-over-day: latest captured
    price vs the most recent captured price from an EARLIER calendar day. Same-day
    re-runs don't count as 'the last change' — we want the daily move. Returns {}
    when there's no prior-day run to compare against."""
    changes = {}
    if not os.path.exists(HISTORY_DB):
        return changes
    con = sqlite3.connect(HISTORY_DB)
    try:
        row = con.execute('SELECT MAX(run_id) FROM chart_snapshots').fetchone()
        if not row or row[0] is None:
            return changes
        latest_run = row[0]
        # Calendar day of the latest run (price_checked_at is an ISO string).
        d = con.execute('SELECT MAX(substr(price_checked_at,1,10)) FROM chart_snapshots '
                        'WHERE run_id=?', (latest_run,)).fetchone()
        latest_day = d[0] if d else None
        latest = {}
        for tkr, price in con.execute(
                'SELECT ticker, price FROM chart_snapshots WHERE run_id=? AND price IS NOT NULL',
                (latest_run,)):
            if tkr:
                latest[str(tkr).strip().upper()] = price
        # Most recent run strictly before latest_day (a genuine prior trading day).
        prow = con.execute(
            'SELECT MAX(run_id) FROM chart_snapshots WHERE substr(price_checked_at,1,10) < ?',
            (latest_day,)).fetchone()
        if not prow or prow[0] is None:
            return changes
        prev = {}
        for tkr, price in con.execute(
                'SELECT ticker, price FROM chart_snapshots WHERE run_id=? AND price IS NOT NULL',
                (prow[0],)):
            if tkr:
                prev[str(tkr).strip().upper()] = price
    finally:
        con.close()
    for tkr, now_p in latest.items():
        old_p = prev.get(tkr)
        if isinstance(old_p, (int, float)) and old_p and isinstance(now_p, (int, float)):
            changes[tkr] = {'change': round(now_p - old_p, 2),
                            'change_pct': round((now_p - old_p) / old_p * 100.0, 2)}
    return changes


def _price_for(ticker, xlsx_literal, latest):
    """Prefer the captured price; fall back to a literal xlsx price (commodities)."""
    if ticker:
        p = latest.get(str(ticker).strip().upper())
        if isinstance(p, (int, float)):
            return p
    return _num(xlsx_literal)


ACCOUNT_SUMMARY = os.path.join(os.path.expanduser('~'), 'Downloads', 'AccountSummary.csv')
# --- AccountSummary.csv 'View all account details' columns (0-indexed) ---
AS_TYPE, AS_NAME, AS_ACCTNO, AS_PRODUCT, AS_HOLDER = 0, 1, 2, 3, 4
AS_PRICE, AS_QTY, AS_VALUE, AS_BOOKCOST, AS_GAIN = 7, 9, 10, 13, 14
# Broker ticker -> the Investments-sheet ticker the chart/alerts are keyed on.
AS_TICKER_ALIASES = {'AV.': 'AV', 'SPLT': 'PLAT', 'SPDM': 'PALL', 'BT.A': 'BT.A'}
# Fund holdings that belong in the Investments table, not Income Funds (no ticker
# in the broker name, so matched by name prefix) -> Investments ticker.
AS_FUND_AS_INVESTMENT = {'WS GUINNESS GLOBAL ENERGY': 'GUINNESS'}
# The dashboard models the FAMILY's investments only (user decision 2026-07-19).
# The Fidelity export covers accounts held for wider relatives too (Dorothy Wall,
# Olive Elizabeth Sangha, Freda Hibbert) — those are administered here but are not
# part of the family portfolio, and including them inflated every total. Matched on
# the account holder's first name, which is how the Wealth Summary labels them.
FAMILY_HOLDERS = {'PAUL', 'SUSAN', 'LIAM', 'JAYNE'}


def _as_num(s):
    try:
        return float(str(s).replace(',', '').replace('+', '').strip())
    except (TypeError, ValueError):
        return None


def _fidelity_positions(path=ACCOUNT_SUMMARY, funds=False):
    """Open positions from the Fidelity AccountSummary export, one row per
    holding-per-account (the 'View all account details' section — the section
    above it is an aggregate across accounts and would double-count).

    This is the authoritative list of what is actually HELD, and the only source
    of Account / Wrapper / book cost. Returns [] if the export isn't present.

    funds=False returns the equity/ETC/Investments-table holdings; funds=True
    returns the income-fund holdings instead (same section, complementary split).
    """
    if not os.path.exists(path):
        return []
    import csv
    rows = []
    in_detail = False
    with open(path, encoding='utf-8-sig', newline='') as fh:
        for rec in csv.reader(fh):
            if not rec:
                continue
            head = (rec[0] or '').strip()
            if head.lower().startswith('view all account details'):
                in_detail = True
                continue
            if not in_detail or head != 'Asset':
                continue
            name = (rec[AS_NAME] or '').strip()
            if not name or name.lower() == 'cash':
                continue
            ticker = None
            display = name
            if name.endswith(')') and '(' in name:
                ticker = name[name.rfind('(') + 1:-1].strip().upper()
                display = name[:name.rfind('(')].split(',')[0].strip().title()
            else:
                up = name.upper()
                for pref, tk in AS_FUND_AS_INVESTMENT.items():
                    if up.startswith(pref):
                        ticker = tk
                        break
            if funds != (ticker is None):
                continue       # wrong side of the equity / income-fund split
            ticker = AS_TICKER_ALIASES.get(ticker, ticker) if ticker else None
            holder = (rec[AS_HOLDER] or '').strip().split()[0] if rec[AS_HOLDER] else None
            if (holder or '').upper() not in FAMILY_HOLDERS:
                continue   # not a family account — see FAMILY_HOLDERS
            qty = _as_num(rec[AS_QTY])
            cost = _as_num(rec[AS_BOOKCOST])
            rows.append({
                'name': display, 'ticker': ticker,
                'account_no': (rec[AS_ACCTNO] or '').strip() or None,
                'account': holder, 'wrapper': (rec[AS_PRODUCT] or '').strip() or None,
                'quantity': qty,
                'holdings': _as_num(rec[AS_VALUE]),
                'book_cost': cost,
                'buy_price': round(cost / qty * 100.0, 2) if (cost and qty) else None,
                'pl_today': _as_num(rec[AS_GAIN]),
                'broker_price': _as_num(rec[AS_PRICE]),
            })
    return rows


DOWNLOADS = os.path.join(os.path.expanduser('~'), 'Downloads')
# --- TransactionHistory.csv columns (0-indexed, header 'Order date,...') ---
TX_ORDER, TX_COMPLETE, TX_TYPE, TX_WRAPPER, TX_ACCTNO, TX_SOURCE, TX_AMOUNT = 0, 1, 2, 4, 5, 6, 7


def _tx_date(s):
    for fmt in ('%d %b %Y', '%d/%m/%Y', '%Y-%m-%d'):
        try:
            return datetime.datetime.strptime(str(s).strip(), fmt).date()
        except (TypeError, ValueError):
            continue
    return None


def _fund_income_events(downloads=DOWNLOADS):
    """(fund name upper, account number) -> the most recent 'Income Received' the
    Fidelity transaction export shows for that holding.

    This is the ONLY source of real dividend dates for the income funds — the
    workbook has none, and the funds aren't in Base Data (no ticker). Fidelity's
    Order date IS the ex-dividend date and Completion date the payment date.
    The export only covers the last ~30 days, so this is 'the latest payment',
    not a history. Nothing is inferred: a fund with no row gets no dates.
    """
    import csv
    events = {}
    for fn in sorted(os.listdir(downloads)) if os.path.isdir(downloads) else []:
        low = fn.lower()
        if not (low.startswith('transactionhistory') or low.startswith('transactions')):
            continue
        if not low.endswith('.csv'):
            continue
        try:
            with open(os.path.join(downloads, fn), encoding='utf-8-sig', newline='') as fh:
                for rec in csv.reader(fh):
                    if len(rec) <= TX_AMOUNT or (rec[TX_TYPE] or '').strip() != 'Income Received':
                        continue
                    src = (rec[TX_SOURCE] or '').strip()
                    acct = (rec[TX_ACCTNO] or '').strip()
                    if not src:
                        continue
                    paid = _tx_date(rec[TX_COMPLETE])
                    key = (src.upper(), acct)
                    prev = events.get(key)
                    if prev and prev['paid_date'] and paid and prev['paid_date'] >= paid.isoformat():
                        continue
                    ex = _tx_date(rec[TX_ORDER])
                    events[key] = {
                        'amount': _as_num(rec[TX_AMOUNT]),
                        'ex_div_date': ex.isoformat() if ex else None,
                        'paid_date': paid.isoformat() if paid else None,
                        'wrapper': (rec[TX_WRAPPER] or '').strip() or None,
                    }
        except OSError:
            continue   # a file mid-download / locked by Excel — skip, never fail
    return events


# Chart symbols priced in USD (per ounce / barrel / MMBtu) rather than UK pence.
# Every LSE equity and ETC the pipeline captures is quoted in pence, so pence is
# the default and only these are the exception — the dashboard labels the unit
# rather than putting a misleading '£' in front of a pence figure.
USD_PRICED_TICKERS = {'PLAT', 'PALL', 'GOLD', 'SILVER', 'COPP', 'NATGAS', 'UKOIL', 'BRENT'}


def _price_unit(ticker):
    tk = str(ticker).strip().upper() if ticker else ''
    return '$' if tk in USD_PRICED_TICKERS else 'p'


def _tv_url(cell):
    """Extract the chart URL from a =HYPERLINK("url","label") formula cell."""
    if isinstance(cell, str) and cell.upper().startswith('=HYPERLINK('):
        try:
            return cell.split('"')[1]
        except IndexError:
            return None
    return None


def _read_soi_band(soi, soif, band_substring, base, latest, changes=None):
    """Read one 'Stocks of Interest' section-table band (e.g. 'AT LOWER BOUNDARY')
    into dashboard watchlist row dicts. Returns [] if that band isn't found on the
    sheet (never invents rows) — membership within a band is still hand-curated
    (see refresh_soi_sections.py); this only reads the rows already placed there."""
    rows = []
    band_row = None
    for r in range(1, soi.max_row + 1):
        a = soi.cell(r, SOI_STOCK).value
        if isinstance(a, str) and band_substring in a.upper():
            band_row = r
            break
    if band_row is None:
        return rows
    r = band_row + 2  # skip the band row and the column-header row
    while r <= soi.max_row:
        stock = soi.cell(r, SOI_STOCK).value
        if not stock or (isinstance(stock, str) and stock.strip().upper() == 'STOCK'):
            break  # blank row or the next section's column header -> end of section
        ticker = soi.cell(r, SOI_TICKER).value
        tk = str(ticker).strip().upper() if ticker else None
        low = _num(soi.cell(r, SOI_ALOW).value)
        high = _num(soi.cell(r, SOI_AHIGH).value)
        price = _price_for(ticker, None, latest)
        prox = ((price - low) / low * 100.0) if (price and low) else None
        upside = ((high - low) / low * 100.0) if (high and low) else None
        bd_row = base.get(tk, {})
        updated = soi.cell(r, SOI_UPDATED).value
        chg = (changes or {}).get(tk, {})
        rows.append({
            'stock': stock, 'ticker': ticker,
            'pattern': soi.cell(r, SOI_PATTERN).value,
            'proximity_pct': round(prox, 2) if prox is not None else None,
            'alert_low': low, 'current': price, 'alert_high': high,
            'price_unit': _price_unit(ticker),
            'change': chg.get('change'), 'change_pct': chg.get('change_pct'),
            'upside_pct': round(upside, 2) if upside is not None else None,
            'pe': bd_row.get('pe'), 'div_yield_pct': bd_row.get('div_yield_pct'),
            'chart_note': soi.cell(r, SOI_CHARTNOTE).value,
            'analyst_rating': soi.cell(r, SOI_RATING).value,
            'holdings': _num(soi.cell(r, SOI_HOLDINGS).value),
            'target_value': _num(soi.cell(r, SOI_TARGET).value),
            'notes': soi.cell(r, SOI_NOTES).value,
            'last_updated': updated.strftime('%Y-%m-%d') if isinstance(updated, datetime.datetime) else updated,
            'chart_url': _tv_url(soif.cell(r, SOI_TV).value),
        })
        r += 1
    return rows


def build(workbook=WORKBOOK):
    wb = openpyxl.load_workbook(workbook, data_only=True)
    wbf = openpyxl.load_workbook(workbook, data_only=False)  # for HYPERLINK formulas
    latest = _latest_prices()
    changes = _price_changes()   # day-over-day price move per ticker
    now = datetime.datetime.now()

    # ---------------- Portfolio (holdings tables) ----------------
    inv = wb['Investments']
    invf = wbf['Investments']
    # Index the Investments sheet by ticker — it owns the alert levels, Type and
    # TradingView link, but NOT what is actually held (its Holdings column goes
    # stale and misses positions, e.g. Centrica). The broker export owns holdings.
    inv_by_ticker = {}
    for r in range(4, inv.max_row + 1):
        tk = inv.cell(r, I_TICKER).value
        if tk:
            inv_by_ticker.setdefault(str(tk).strip().upper(), r)

    def _inv_row(ticker):
        return inv_by_ticker.get(str(ticker).strip().upper()) if ticker else None

    positions = _fidelity_positions()
    if not positions:
        # No broker export in Downloads — fall back to the workbook's own Holdings.
        for r in range(4, inv.max_row + 1):
            hv = _num(inv.cell(r, I_HOLDINGS).value)
            if not inv.cell(r, I_NAME).value or not hv or hv <= 0:
                continue
            positions.append({
                'name': inv.cell(r, I_NAME).value,
                'ticker': inv.cell(r, I_TICKER).value,
                'account': None, 'wrapper': None, 'quantity': None,
                'holdings': hv, 'book_cost': None, 'buy_price': None,
                'pl_today': None, 'broker_price': None,
            })

    investments = []
    inv_value_total = 0.0
    short_term_value = long_term_value = 0.0
    alert_below = alert_near = alert_above = 0
    for pos in positions:
        ticker = pos['ticker']
        r = _inv_row(ticker)
        holdings = pos['holdings'] or 0.0
        name = inv.cell(r, I_NAME).value if r else pos['name']
        price = (_price_for(ticker, inv.cell(r, I_CURPRICE).value if r else None, latest)
                 or pos['broker_price'])
        low = _num(inv.cell(r, I_ALERT_LOW).value) if r else None
        high = _num(inv.cell(r, I_ALERT_HIGH).value) if r else None
        diff_low = ((price - low) / low * 100.0) if (price and low) else None
        gap_low = ((price - low) / price * 100.0) if (price and low) else None
        inv_value_total += holdings
        itype = inv.cell(r, I_TYPE).value if r else None
        if isinstance(itype, str) and itype.strip().lower().startswith('long'):
            long_term_value += holdings
        else:
            short_term_value += holdings
        if diff_low is not None:
            if diff_low <= 0:
                alert_below += 1
            elif diff_low <= 5:
                alert_near += 1
            else:
                alert_above += 1
        investments.append({
            'name': name, 'ticker': ticker,
            'type': itype,                       # Short Term / Long Term
            'account': pos['account'], 'wrapper': pos['wrapper'],
            'holdings': round(holdings, 2),
            'quantity': pos['quantity'],
            'buy_price': pos['buy_price'],       # pence, = book cost / quantity
            'book_cost': pos['book_cost'],
            'pl_today': pos['pl_today'],         # £ profit/loss if sold today
            'pl_today_pct': (round(pos['pl_today'] / pos['book_cost'] * 100.0, 2)
                             if (pos['pl_today'] is not None and pos['book_cost']) else None),
            'current_price': price,
            'price_unit': _price_unit(ticker),   # unit of current_price/alert levels
            'gap_to_low_pct': round(gap_low, 2) if gap_low is not None else None,
            'alert_low': low, 'alert_low_source': inv.cell(r, I_ALERT_LOW_SRC).value if r else None,
            'diff_to_low_pct': round(diff_low, 2) if diff_low is not None else None,
            'alert_high': high,
            'chart_url': _tv_url(invf.cell(r, I_TV).value) if r else None,
        })
    investments.sort(key=lambda x: (-(x['holdings'] or 0)))

    # ---------------- Income Funds ----------------
    inf = wb['Income Funds']
    income_funds = []
    funds_value_total = 0.0
    funds_monthly_income = 0.0
    for r in range(5, inf.max_row + 1):
        fname = inf.cell(r, F_NAME).value
        # The fund list ends at the 'TOTAL (Family)' row; a second per-person
        # sub-table follows it. Stop there so neither is summed into the holdings.
        if isinstance(fname, str) and fname.strip().upper().startswith('TOTAL'):
            break
        curval = _num(inf.cell(r, F_CURVAL).value)
        if not fname or curval is None or curval <= 0:
            continue
        monthly_rev = _num(inf.cell(r, F_MONTHLY_REV).value) or 0.0
        annual_rev = _num(inf.cell(r, F_ANNUAL_REV).value) or 0.0
        yld = _num(inf.cell(r, F_DIVYLD).value)
        funds_value_total += curval
        funds_monthly_income += monthly_rev
        income_funds.append({
            'name': fname, 'account': None, 'wrapper': None,
            'holdings': round(curval, 2),
            'div_yield_pct': round(yld * 100.0, 2) if yld is not None else None,
            'monthly_income': round(monthly_rev, 2),
            'annual_income': round(annual_rev, 2),   # 2026 annual dividend (run-rate)
        })
    funds_annual_income = round(sum(f['annual_income'] for f in income_funds), 2)

    # ---------------- Income fund POSITIONS (per fund, per account) ----------------
    # income_funds above stays as the sheet's family-level roll-up (the Overview
    # metrics, Relevant News and the Historic dividends table all read it). This is
    # the Portfolio screen's per-account view: broker holdings + the real dividend
    # dates from the transaction export.
    income_events = _fund_income_events()
    income_positions = []
    for pos in _fidelity_positions(funds=True):
        ev = income_events.get((pos['name'].upper(), pos['account_no'] or '')) or {}
        income_positions.append({
            'name': pos['name'],
            'account': pos['account'], 'wrapper': pos['wrapper'],
            'holdings': pos['holdings'], 'book_cost': pos['book_cost'],
            'pl_today': pos['pl_today'],
            'pl_today_pct': (round(pos['pl_today'] / pos['book_cost'] * 100.0, 2)
                             if (pos['pl_today'] is not None and pos['book_cost']) else None),
            # Yield is DERIVED from this holding's own last payment (these funds
            # distribute monthly), not matched to the sheet by name — the fund names
            # differ enough between broker and sheet that token matching paired AXA
            # and IFSL with other managers' funds.
            'div_yield_pct': (round(ev['amount'] * 12.0 / pos['holdings'] * 100.0, 2)
                              if (ev.get('amount') and pos['holdings']) else None),
            'last_income': ev.get('amount'),
            'ex_div_date': ev.get('ex_div_date'),
            'payment_date': ev.get('paid_date'),
        })
    income_positions.sort(key=lambda x: -(x['holdings'] or 0))
    income_positions_total = {
        'holdings': round(sum(p['holdings'] or 0 for p in income_positions), 2),
        'book_cost': round(sum(p['book_cost'] or 0 for p in income_positions), 2),
        'pl_today': round(sum(p['pl_today'] or 0 for p in income_positions), 2),
        'last_income': round(sum(p['last_income'] or 0 for p in income_positions), 2),
        'with_dates': sum(1 for p in income_positions if p['payment_date']),
        'count': len(income_positions),
    }

    # ---------------- Wealth Summary (monthly trend, cash, MoM gain) ----------------
    ws = wb['Wealth Summary']
    months = []
    _MON = {m: i for i, m in enumerate(
        ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'], start=1)}
    for c in range(4, ws.max_column + 1):
        label = ws.cell(WS_MONTH_ROW, c).value
        if not label:
            continue
        # Only ACTUALS: the Wealth Summary carries forward-projected future months,
        # which would draw a misleading flat line past today. Cap at the current month.
        parts = str(label).split()
        if len(parts) == 2 and parts[0] in _MON and parts[1].isdigit():
            yr, mo = int(parts[1]), _MON[parts[0]]
            if (yr, mo) > (now.year, now.month):
                continue
        invest_sum = sum(_num(ws.cell(row, c).value) or 0.0 for row in WS_FIDELITY_ROWS)
        if invest_sum <= 0:
            continue  # drop empty / future months with no data
        months.append({'col': c, 'label': str(label), 'value': round(invest_sum, 2)})
    value_over_time = {'labels': [m['label'] for m in months],
                       'portfolio': [m['value'] for m in months]}
    last_col = months[-1]['col'] if months else 4
    cash_available = sum(_num(ws.cell(row, last_col).value) or 0.0 for row in WS_CASH_ROWS)
    # Month-over-month change in the investable total (transparent, and consistent
    # with the trend chart). Includes contributions — flagged in the metric caveat.
    gain_last_month = round(months[-1]['value'] - months[-2]['value'], 2) if len(months) >= 2 else None
    gain_last_month_pct = (round(gain_last_month / months[-2]['value'] * 100.0, 2)
                           if gain_last_month is not None and months[-2]['value'] else None)

    # Total Portfolio Value = the Wealth Summary's 'Fidelity accounts' block for the
    # latest actual month — the same series the trend chart plots, so the headline and
    # the sparkline can never disagree. (It is NOT holdings + income funds: that summed
    # positions across pension accounts too and came out ~£1M higher than the sheet.)
    portfolio_value = round(months[-1]['value'], 2) if months else round(
        inv_value_total + funds_value_total, 2)

    # ---------------- Fidelity accounts, last FULL month ----------------
    # 'Last full month' = the month before the current one; today's month is still
    # accruing, so comparing it to the previous month understates the increase.
    import re as _re
    accounts, accounts_month, accounts_prev_month = [], None, None
    if len(months) >= 2:
        full = months[-2] if str(months[-1]['label']).startswith(
            now.strftime('%b')) else months[-1]
        idx = months.index(full)
        prev = months[idx - 1] if idx >= 1 else None
        accounts_month = full['label']
        accounts_prev_month = prev['label'] if prev else None
        for row in WS_FIDELITY_ROWS:
            label = ws.cell(row, 1).value
            if not label:
                continue
            m = _re.match(r'^\s*(.*?)\s*\(([^()]*)\)\s*\(([^()]*)\)\s*$', str(label))
            wrapper, holder, acct_no = (m.group(1), m.group(2), m.group(3)) if m else (
                str(label).strip(), None, None)
            val = _num(ws.cell(row, full['col']).value)
            pval = _num(ws.cell(row, prev['col']).value) if prev else None
            if val is None:
                continue
            accounts.append({
                'account': holder, 'wrapper': wrapper, 'account_no': acct_no,
                'value': round(val, 2),
                'prev_value': round(pval, 2) if pval is not None else None,
                'increase': round(val - pval, 2) if pval is not None else None,
                'increase_pct': (round((val - pval) / pval * 100.0, 2)
                                 if (pval not in (None, 0)) else None),
            })
        accounts.sort(key=lambda a: -(a['value'] or 0))
    # monthly_dividend is completed below, once Base Data dividend yields are read,
    # so it can include SHARE dividends (income + accumulation), not just income funds.

    # ---------------- Historic (completed sales) ----------------
    hist = wb['History']
    sold = []
    profit_2026 = 0.0
    wins = 0
    for r in range(2, hist.max_row + 1):
        sell = hist.cell(r, H_SELL).value
        if not isinstance(sell, datetime.datetime):
            continue  # section dividers / open positions
        cost = _num(hist.cell(r, H_COST).value)
        proceeds = _num(hist.cell(r, H_PROCEEDS).value)
        profit = (proceeds - cost) if (cost is not None and proceeds is not None) else None
        buy = hist.cell(r, H_BUY).value
        days = (sell - buy).days if isinstance(buy, datetime.datetime) else None
        pnl_pct = (profit / cost * 100.0) if (profit is not None and cost) else None
        if profit is not None:
            if sell.year == 2026:
                profit_2026 += profit
            if profit > 0:
                wins += 1
        sold.append({
            'name': hist.cell(r, H_INV).value,
            'account': hist.cell(r, H_ACCT).value,
            'wrapper': hist.cell(r, H_WRAP).value,
            'buy_date': buy.strftime('%Y-%m-%d') if isinstance(buy, datetime.datetime) else None,
            'sell_date': sell.strftime('%Y-%m-%d'),
            'profit': round(profit, 2) if profit is not None else None,
            'pnl_pct': round(pnl_pct, 2) if pnl_pct is not None else None,
            'days_held': days,
        })
    win_rate = round(wins / len(sold) * 100.0, 1) if sold else None

    # ---------------- Watchlist (Stocks of Interest, within 5% of alert low) ----------
    # Base Data (literal, from HL.co.uk) gives P/E + Div Yield by ticker — the section
    # table's own P/E/Div-Yield cells are googlefinance/VLOOKUP formulas (empty here).
    bd = wb['Base Data']
    base = {}
    for r in range(3, bd.max_row + 1):
        tkr = bd.cell(r, BD_TICKER).value
        if tkr:
            exdiv = bd.cell(r, BD_EXDIV).value
            base[str(tkr).strip().upper()] = {
                'pe': _num(bd.cell(r, BD_PE).value),
                'div_yield_pct': (round(_num(bd.cell(r, BD_DIVYLD).value) * 100.0, 2)
                                  if _num(bd.cell(r, BD_DIVYLD).value) is not None else None),
                'div_pence': _num(bd.cell(r, BD_DIV_PENCE).value),
                'ex_div': (exdiv.strftime('%Y-%m-%d') if isinstance(exdiv, datetime.datetime) else exdiv)}
    # Monthly SHARE dividends (user request 2026-07-18): add the ACTUAL share income
    # + accumulation dividends that spending_summary computes and exports to
    # data/spending_dividends.json each pipeline run, on top of the income-fund revenue
    # (which stays sourced from the master's Income Funds tab — no double count).
    # Falls back to 0 until that file exists (e.g. a run without the Fidelity exports).
    share_income_monthly = share_accum_monthly = 0.0
    try:
        with open(os.path.join(REPO, 'data', 'spending_dividends.json'), encoding='utf-8') as _f:
            _dv = json.load(_f)
        share_income_monthly = float(_dv.get('share_income_monthly') or 0)
        share_accum_monthly = float(_dv.get('share_accumulation_monthly') or 0)
    except (OSError, ValueError):
        pass
    monthly_dividend = round(funds_monthly_income + share_income_monthly + share_accum_monthly, 2)

    soi = wb['Stocks of Interest']
    soif = wbf['Stocks of Interest']
    # Strictly the 'AT LOWER BOUNDARY — within 5% of alert low' band (user decision
    # 2026-07-17): the Watchlist is the buy-zone list, not the whole Stocks-of-Interest
    # ladder. The other bands stay out.
    watchlist = _read_soi_band(soi, soif, SOI_WATCHLIST_BAND, base, latest, changes)

    overview = {
        'generated_at': now.isoformat(timespec='seconds'),
        'currency': 'GBP',
        'metrics': {
            'accumulation_income': {'value': share_accum_monthly,
                                    'caveat': 'Monthly reinvested income from accumulation funds '
                                              '(latest month of the Acc funds’ price-appreciation).'},
            'portfolio_value': {'value': portfolio_value,
                                'caveat': "Wealth Summary 'Fidelity accounts' total for "
                                          + (months[-1]['label'] if months else 'the latest month')
                                          + ' — matches the sheet and the Finance Google Sheet.'},
            'gain_last_month': {'value': gain_last_month, 'pct': gain_last_month_pct,
                                'caveat': 'Month-on-month investment-account change (includes contributions).'},
            'trading_profit_2026': {'value': round(profit_2026, 2),
                                    'sells': sum(1 for s in sold if s['sell_date'][:4] == '2026')},
            'monthly_dividend': {'value': monthly_dividend,
                                 'caveat': f'Income funds £{funds_monthly_income:,.0f}/mo + share income '
                                           f'£{share_income_monthly:,.0f}/mo + share accumulation '
                                           f'£{share_accum_monthly:,.0f}/mo.'},
            'cash_available': {'value': round(cash_available, 2),
                               'pct_of_portfolio': round(cash_available / portfolio_value * 100.0, 2) if portfolio_value else None,
                               'caveat': 'Fidelity cash accounts only; uninvested account cash added in Phase 1.5.'},
        },
        'value_over_time': value_over_time,
        'accounts': {'month': accounts_month, 'prev_month': accounts_prev_month,
                     'rows': accounts,
                     'total': round(sum(a['value'] or 0 for a in accounts), 2),
                     'total_increase': round(sum(a['increase'] or 0 for a in accounts), 2)},
        'alert_status': {'below': alert_below, 'near': alert_near,
                         'above': alert_above, 'total': alert_below + alert_near + alert_above},
    }
    portfolio = {'generated_at': now.isoformat(timespec='seconds'),
                 'investments': investments, 'income_funds': income_funds,
                 'income_positions': income_positions,
                 'income_positions_total': income_positions_total,
                 'income_summary': {'year': 2026,
                                    'total_annual_income': funds_annual_income,
                                    'total_monthly_income': monthly_dividend,
                                    'fund_count': len(income_funds)}}
    historic = {'generated_at': now.isoformat(timespec='seconds'),
                'sold': sold,
                'summary': {'total_profit_2026': round(profit_2026, 2),
                            'count': len(sold), 'win_rate': win_rate}}
    watchlist_payload = {'generated_at': now.isoformat(timespec='seconds'),
                         'criterion': 'Within 5% of alert low (Stocks of Interest — at lower boundary)',
                         'rows': watchlist}

    # ---------------- Targets (income + allocation by Type) ----------------
    targets = {
        'generated_at': now.isoformat(timespec='seconds'),
        'income_per_month': monthly_dividend,
        'annual_income': funds_annual_income,
        'allocation': [
            {'label': 'Income Funds', 'value': round(funds_value_total, 2)},
            {'label': 'Short-Term (Shares)', 'value': round(short_term_value, 2)},
            {'label': 'Long-Term', 'value': round(long_term_value, 2)},
            {'label': 'Cash', 'value': round(cash_available, 2)},
        ],
    }
    targets['total'] = round(sum(a['value'] for a in targets['allocation']), 2)

    # ---------------- Relevant News (ex-div dates + amounts for holdings) ------------
    news = []
    for iv in investments:                     # equity/commodity holdings
        b = base.get(str(iv['ticker']).strip().upper()) if iv['ticker'] else None
        if b and b.get('ex_div'):
            news.append({'name': iv['name'], 'ticker': iv['ticker'], 'event': 'Ex-dividend',
                         'date': b['ex_div'], 'amount_pence': b.get('div_pence'),
                         'div_yield_pct': b.get('div_yield_pct'), 'holding': iv['holdings']})
    for f in income_funds:                      # income funds pay monthly
        news.append({'name': f['name'], 'ticker': None, 'event': 'Monthly income',
                     'date': None, 'amount_pounds': f['monthly_income'],
                     'div_yield_pct': f['div_yield_pct'], 'holding': f['holdings']})
    news.sort(key=lambda n: (n['date'] is None, n['date'] or ''))
    news_payload = {'generated_at': now.isoformat(timespec='seconds'), 'rows': news}

    # ---------------- Activity (things to do, with links) ----------------
    SHEET_URL = 'https://docs.google.com/spreadsheets/d/1UjAz_QUuh86_e6yq8QJf2veI8IpkRCyVfWaK6maqiyc/'
    tv_by_ticker = {}
    for r in range(4, inv.max_row + 1):
        tkr = inv.cell(r, I_TICKER).value
        url = _tv_url(invf.cell(r, I_TV).value)
        if tkr and url:
            tv_by_ticker[str(tkr).strip().upper()] = url
    unmarked, inherited = [], []
    cr_path = os.path.join(SCRIPT_DIR, 'channel_results_tmp.json')
    if os.path.exists(cr_path):
        for rr in json.load(open(cr_path, encoding='utf-8')):
            reason = (rr.get('reason') or '').lower()
            t = rr.get('ticker')
            if 'no channel or trend line found near price' in reason:
                unmarked.append({'ticker': t, 'chart_url': tv_by_ticker.get(str(t).strip().upper())})
            elif 'axis' in reason:
                inherited.append(t)
    below_tickers = [w['ticker'] for w in watchlist if (w.get('proximity_pct') is not None and w['proximity_pct'] <= 0)]

    def act(category, priority, title, detail, link=None, link_label=None):
        return {'category': category, 'priority': priority, 'title': title,
                'detail': detail, 'link': link, 'link_label': link_label}

    actions = [
        act('Charts', 'high', f'{len(unmarked)} charts have no trend lines drawn',
            'Mark up the channel / trend lines in TradingView so they feed alert levels. See the list below.'),
        act('Charts', 'medium', 'Price extended beyond the trend line / channel',
            'New pattern (added 2026-07-17): the drawn line stops before today. None detected this run — recently refreshed.'),
        act('Charts', 'medium', f'{len(inherited)} charts on inherited levels',
            'Axis read failed this run, so the Alert Low/High are carried over — redraw or re-check these in TradingView.'),
        act('Buys', 'high', f'{len(below_tickers)} stocks at or below alert low',
            'Buy candidates: ' + (', '.join(str(t) for t in below_tickers) or 'none') + '. Review the Watchlist.',
            '#watchlist', 'Open Watchlist'),
        act('Data', 'high', 'Review the figures I changed',
            'Confirm the July Wealth Summary fix (Joint £791,992) and WS Guinness value (£111,655) look right.'),
        act('Data', 'medium', 'Refresh bank exports',
            'Amex (activity.csv) + Barclays (data.csv) are not in Downloads, so spending/wealth tabs go stale past ~6 weeks.'),
        act('Data', 'low', 'Sync the workbook to the Finance Google Sheet', 'Push the latest master workbook to Google Sheets.',
            SHEET_URL, 'Open sheet'),
        act('Code', 'medium', 'Answer any open questions in Claude Code',
            'I may be waiting on a decision — check the Claude Code session.'),
        act('Code', 'low', 'Commit & push pending changes',
            'Lock in the latest dashboard/data changes when you are happy with them.'),
    ]
    activity = {'generated_at': now.isoformat(timespec='seconds'),
                'actions': actions, 'charts_to_markup': unmarked}

    return {'overview': overview, 'portfolio': portfolio, 'historic': historic,
            'watchlist': watchlist_payload, 'targets': targets, 'news': news_payload,
            'activity': activity}


def write(out_dir=OUT_DIR, workbook=WORKBOOK):
    os.makedirs(out_dir, exist_ok=True)
    data = build(workbook)
    for name, payload in data.items():
        with open(os.path.join(out_dir, name + '.json'), 'w', encoding='utf-8') as f:
            json.dump(payload, f, indent=2, default=str)
    return data


if __name__ == '__main__':
    d = write()
    ov = d['overview']['metrics']
    print('Wrote dashboard JSON ->', OUT_DIR)
    print(f"  Portfolio value : £{ov['portfolio_value']['value']:,.0f}")
    print(f"  Gain last month : £{(ov['gain_last_month']['value'] or 0):,.0f}")
    print(f"  2026 sells profit: £{ov['trading_profit_2026']['value']:,.0f} ({ov['trading_profit_2026']['sells']} sells)")
    print(f"  Monthly dividend: £{ov['monthly_dividend']['value']:,.0f}")
    print(f"  Cash available  : £{ov['cash_available']['value']:,.0f}")
    print(f"  Holdings: {len(d['portfolio']['investments'])} investments, {len(d['portfolio']['income_funds'])} income funds")
    print(f"  Historic sales  : {d['historic']['summary']['count']} (win rate {d['historic']['summary']['win_rate']}%)")
    print(f"  Alert status    : {d['overview']['alert_status']}")
