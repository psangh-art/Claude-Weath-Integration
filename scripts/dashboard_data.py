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
# --- Income Funds columns (header row 4, data from row 5) ---
F_NAME, F_CURVAL, F_DIVYLD, F_MONTHLY_REV = 1, 3, 6, 9
# --- History columns (header row 1) ---
H_INV, H_ACCT, H_WRAP, H_BUY, H_SELL = 1, 2, 3, 4, 5
H_COST, H_PROCEEDS = 10, 12
# --- Wealth Summary: month headers on row 3 from col D; investable account rows ---
WS_MONTH_ROW = 3
WS_INVEST_ROWS = [5, 6, 7, 8, 9, 12]     # Investment Account(s) + ISAs + Junior ISA
WS_CASH_ROWS = [10, 11, 13]              # Fidelity Cash Accounts
WS_MONTHLY_INCREASE_ROW = 46            # 'Monthly Investment Increase' (literal)


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


def build(workbook=WORKBOOK):
    wb = openpyxl.load_workbook(workbook, data_only=True)
    wbf = openpyxl.load_workbook(workbook, data_only=False)  # for HYPERLINK formulas
    latest = _latest_prices()
    now = datetime.datetime.now()

    # ---------------- Portfolio (holdings tables) ----------------
    inv = wb['Investments']
    invf = wbf['Investments']
    investments = []
    inv_value_total = 0.0
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
        if diff_low is not None:
            if diff_low <= 0:
                alert_below += 1
            elif diff_low <= 5:
                alert_near += 1
            else:
                alert_above += 1
        investments.append({
            'name': name, 'ticker': ticker,
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
        yld = _num(inf.cell(r, F_DIVYLD).value)
        funds_value_total += curval
        funds_monthly_income += monthly_rev
        income_funds.append({
            'name': fname, 'account': None, 'wrapper': None,
            'holdings': round(curval, 2),
            'div_yield_pct': round(yld * 100.0, 2) if yld is not None else None,
            'monthly_income': round(monthly_rev, 2),
        })

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
                 'investments': investments, 'income_funds': income_funds}
    historic = {'generated_at': now.isoformat(timespec='seconds'),
                'sold': sold,
                'summary': {'total_profit_2026': round(profit_2026, 2),
                            'count': len(sold), 'win_rate': win_rate}}
    return {'overview': overview, 'portfolio': portfolio, 'historic': historic}


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
