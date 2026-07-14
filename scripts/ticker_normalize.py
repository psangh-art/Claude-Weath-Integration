#!/usr/bin/env python3
"""Shared ticker normalization between TradingView's export naming and the master
'Stocks Buy Strategy.xlsx' sheet, per the rules in Claude_Code_Handoff_Instructions.md
sections 5-7. Used by build_layout_excel.py (Google Finance column) and
update_master_sheet.py (row matching).
"""
import re

# Master sheet ticker -> TradingView export name (handoff doc section 5)
COMMODITY_MASTER_TO_TV = {
    'GOLD': 'GOLD',
    'SLVR': 'SILVER',
    'COPP': 'COPPER',
    'OIL': 'USOIL',
    'UKOIL': 'BRENT',  # master row 'Brent Oil' is UKOIL; the TV chart is 'Crude Oil Brent Cash'
    'PLAT': 'PLATINUM',
    'PALL': 'PALLADIUM',
    'NDX': 'NASDAQ',
}
TV_TO_COMMODITY_MASTER = {v: k for k, v in COMMODITY_MASTER_TO_TV.items()}

# TradingView continuous-futures symbols carry a '1!'/'2!' suffix (e.g. COPPER1!)
_FUTURES_SUFFIX_RE = re.compile(r'\d+!$')

# GOOGLEFINANCE can no longer price commodities AT ALL — verified live 2026-07-11:
# TVC:GOLD returns #N/A, and so do CURRENCY:XAUUSD/XAGUSD/XPTUSD/XPDUSD (Google
# dropped metals support), while equities (LON:*/GOOG) and real FX (GBPUSD) still
# work. The handoff doc section 6's "confirmed-reliable TVC:" claim is falsified.
# Commodity rows get their Current Price written as a VALUE from the TradingView
# capture instead (see update_master_sheet.py), so no commodity gets a formula.
RELIABLE_GOOGLEFINANCE_COMMODITIES = set()

# Macro/FX symbols that must never be forced into an equity row (handoff doc section 5)
MACRO_EXCLUDE = {'JP10Y', 'US10Y', 'GBPUSD'}

_CLASS_SUFFIX_RE = re.compile(r'^([A-Z]+)\.([A-Z])$')


def normalize(ticker, symbol=None):
    """Classify and normalize a raw TradingView ticker/symbol.

    Returns a dict:
      kind: 'commodity' | 'macro_excluded' | 'equity' | None (no ticker given)
      tv_ticker: the original ticker as given
      master_ticker: the form expected in Stocks Buy Strategy.xlsx's Ticker column,
                     or None if this should never be matched to an equity row
      google_finance_ticker: e.g. "LON:BP" or "TVC:GOLD", or None if no reliable
                     GOOGLEFINANCE pattern exists for this instrument
      google_finance_formula: a ready `=googlefinance(...)` string, or None
    """
    if not ticker:
        return None
    t = ticker.strip()
    upper = _FUTURES_SUFFIX_RE.sub('', t.upper())

    if upper in TV_TO_COMMODITY_MASTER:
        master = TV_TO_COMMODITY_MASTER[upper]
        reliable = upper in RELIABLE_GOOGLEFINANCE_COMMODITIES
        gf_ticker = f"TVC:{upper}" if reliable else None
        return {
            'kind': 'commodity',
            'tv_ticker': t,
            'master_ticker': master,
            'google_finance_ticker': gf_ticker,
            'google_finance_formula': f'=googlefinance("{gf_ticker}","price")' if gf_ticker else None,
        }

    combined = f"{(symbol or '')}:{t}".upper()
    if upper in MACRO_EXCLUDE or any(m in combined for m in MACRO_EXCLUDE):
        return {
            'kind': 'macro_excluded',
            'tv_ticker': t,
            'master_ticker': None,
            'google_finance_ticker': None,
            'google_finance_formula': None,
        }

    m = _CLASS_SUFFIX_RE.match(upper)
    if m:
        master_ticker = f"{m.group(1)}-{m.group(2)}"
        gf_ticker = f"LON:{upper}"
        return {
            'kind': 'equity',
            'tv_ticker': t,
            'master_ticker': master_ticker,
            'google_finance_ticker': gf_ticker,
            'google_finance_formula': f'=googlefinance("{gf_ticker}","price")',
        }

    if upper.endswith('.') and len(upper) > 1:
        stripped = upper[:-1]
        gf_ticker = f"LON:{stripped}"
        return {
            'kind': 'equity',
            'tv_ticker': t,
            'master_ticker': stripped,
            'google_finance_ticker': gf_ticker,
            'google_finance_formula': f'=googlefinance("{gf_ticker}","price")',
        }

    gf_ticker = f"LON:{upper}"
    return {
        'kind': 'equity',
        'tv_ticker': t,
        'master_ticker': upper,
        'google_finance_ticker': gf_ticker,
        'google_finance_formula': f'=googlefinance("{gf_ticker}","price")',
    }


def master_tickers_match(a, b):
    """Compare two master-sheet-form tickers for equality, normalizing case/whitespace
    and the dash/dot class-suffix variants (BT-A == BT.A)."""
    if not a or not b:
        return False
    norm_a = a.strip().upper().replace('.', '-')
    norm_b = b.strip().upper().replace('.', '-')
    return norm_a == norm_b
