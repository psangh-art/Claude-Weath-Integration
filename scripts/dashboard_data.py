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
    investments = []
    inv_value_total = 0.0
    short_term_value = long_term_value = 0.0
    alert_below = alert_near = alert_above = 0
    for r in range(4, inv.max_row + 1):
        holdings = _num(inv.cell(r, I_HOLDINGS).value)
        name = inv.cell(r, I_NAME).value
        if not name or not holdings or holdings <= 0:
            continue  # skip macro reference rows (FTSE/NDX) and blanks
        ticker = inv.cell(r, I_TICKER).value
        price = _price_for(ticker, inv.cell(r, I_CURPRICE).value, latest)
        low = _num(inv.cell(r, I_ALERT_LOW).value)
        high = _num(inv.cell(r, I_ALERT_HIGH).value)
        diff_low = ((price - low) / low * 100.0) if (price and low) else None
        gap_low = ((price - low) / price * 100.0) if (price and low) else None
        inv_value_total += holdings
        itype = inv.cell(r, I_TYPE).value
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
            'type': inv.cell(r, I_TYPE).value,   # Short Term / Long Term
            'account': None, 'wrapper': None,  # Phase 1.5: plumb from Fidelity holdings
            'holdings': round(holdings, 2),
            'current_price': price,
            'gap_to_low_pct': round(gap_low, 2) if gap_low is not None else None,
            'alert_low': low, 'alert_low_source': inv.cell(r, I_ALERT_LOW_SRC).value,
            'diff_to_low_pct': round(diff_low, 2) if diff_low is not None else None,
            'alert_high': high,
            'chart_url': _tv_url(invf.cell(r, I_TV).value),
        })

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
        invest_sum = sum(_num(ws.cell(row, c).value) or 0.0 for row in WS_INVEST_ROWS)
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

    portfolio_value = round(inv_value_total + funds_value_total, 2)
    monthly_dividend = round(funds_monthly_income, 2)  # equities added in Phase 1.5

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
            'portfolio_value': {'value': portfolio_value},
            'gain_last_month': {'value': gain_last_month, 'pct': gain_last_month_pct,
                                'caveat': 'Month-on-month investment-account change (includes contributions).'},
            'trading_profit_2026': {'value': round(profit_2026, 2),
                                    'sells': sum(1 for s in sold if s['sell_date'][:4] == '2026')},
            'monthly_dividend': {'value': monthly_dividend,
                                 'caveat': 'Income-fund revenue; equity dividends added in Phase 1.5.'},
            'cash_available': {'value': round(cash_available, 2),
                               'pct_of_portfolio': round(cash_available / portfolio_value * 100.0, 2) if portfolio_value else None,
                               'caveat': 'Fidelity cash accounts only; uninvested account cash added in Phase 1.5.'},
        },
        'value_over_time': value_over_time,
        'alert_status': {'below': alert_below, 'near': alert_near,
                         'above': alert_above, 'total': alert_below + alert_near + alert_above},
    }
    portfolio = {'generated_at': now.isoformat(timespec='seconds'),
                 'investments': investments, 'income_funds': income_funds,
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
