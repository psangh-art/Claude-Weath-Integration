#!/usr/bin/env python3
"""Refresh the 'Stocks Trading Below Alert Low' table at the TOP of the
'Stocks of Interest' sheet in Stocks_Buy_Strategy.xlsx: every ticker whose live
TradingView price (captured during the chart-export step) has dropped below its
Alert Low, sorted worst-gap-first.

History: this table used to live on its own 'Below Alert Low' first-tab sheet
(that was the safe option at the time — openpyxl's insert_rows() had been shown
to corrupt Stocks of Interest's ~121 formulas and 34 merged bands, RELX fix
2026-07-10). On 2026-07-11 the user asked for it to be the top table of Stocks
of Interest instead, so restructure_soi_stats_2026-07-11.py rebuilt that sheet
once with rows 1-40 RESERVED for this table (all pre-existing content, formulas
and merges were shifted to row 41+). This script now rewrites ONLY that reserved
block, never inserting or deleting rows — so the corruption risk that motivated
the separate sheet never applies. Do not write below row 40 here, and do not
add manual content above row 41 in that sheet.

Usage: python add_below_alert_sheet.py <master.xlsx> <below_alert_rows.json>
  rows: [{"ticker","share_name","price","alert_low","alert_high","gap_pct",
          "holdings","target_value","checked_at"}, ...], worst gap first.
"""
import sys
import json
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill

SHEET_NAME = 'Stocks of Interest'
RESERVED_BLOCK = 40      # rows 1..40 belong to this table; row 41+ is the
                         # pre-existing Stocks of Interest content — never touch it
DATA_START_ROW = 5
MAX_DATA_ROWS = 34       # rows 5..38; 39-40 stay blank as a separator
HEADER_ROW = 4
HEADER_FILL = PatternFill(fill_type='solid', fgColor='FFF2CC')
GAP_BAD_FONT = Font(color='FFCC0000', bold=True)

HEADER = ['Ticker', 'Share Name', 'Current Price', 'Alert Low', 'Gap %', 'Alert High',
          'Holdings (£)', 'Target Value (£)', 'Price Checked At']


def refresh_block(ws, rows):
    if len(rows) > MAX_DATA_ROWS:
        print(f'WARNING: {len(rows)} below-alert rows but only {MAX_DATA_ROWS} fit in the '
              f'reserved block — writing the {MAX_DATA_ROWS} worst; '
              f'{len(rows) - MAX_DATA_ROWS} omitted.', file=sys.stderr)
        rows = rows[:MAX_DATA_ROWS]

    ws.cell(row=1, column=1, value='Stocks Trading Below Alert Low').font = Font(bold=True, size=14)
    ws.cell(row=2, column=1, value=(
        f'Auto-generated {datetime.now().date().isoformat()} from tradingview_layouts.xlsx live '
        'price capture cross-checked against Alert Low. Rebuilt by the pipeline into rows '
        f'1-{RESERVED_BLOCK} of this sheet each run — do not add manual content above row '
        f'{RESERVED_BLOCK + 1}.'
    ))

    for c, h in enumerate(HEADER, 1):
        cell = ws.cell(row=HEADER_ROW, column=c, value=h)
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL

    # Clear the whole data region first so rows from a previous (longer) run
    # can't linger below this run's shorter table.
    for r in range(DATA_START_ROW, DATA_START_ROW + MAX_DATA_ROWS):
        for c in range(1, len(HEADER) + 1):
            cell = ws.cell(row=r, column=c)
            cell.value = None
            cell.font = Font()
            cell.fill = PatternFill()
            cell.number_format = 'General'

    for i, r in enumerate(rows):
        row_n = DATA_START_ROW + i
        ws.cell(row=row_n, column=1, value=r['ticker']).font = Font(bold=True)
        ws.cell(row=row_n, column=2, value=r['share_name'])
        ws.cell(row=row_n, column=3, value=round(r['price'], 2))
        ws.cell(row=row_n, column=4, value=r['alert_low'])
        gap_cell = ws.cell(row=row_n, column=5, value=round(r['gap_pct'], 1))
        gap_cell.number_format = '0.0"%"'
        if r['gap_pct'] <= -10:
            gap_cell.font = GAP_BAD_FONT
        ws.cell(row=row_n, column=6, value=r['alert_high'])
        # £ amounts display as whole pounds (user rule, 2026-07-11)
        holdings_cell = ws.cell(row=row_n, column=7, value=r['holdings'])
        holdings_cell.number_format = '#,##0'
        target_cell = ws.cell(row=row_n, column=8, value=r['target_value'])
        target_cell.number_format = '#,##0'
        ws.cell(row=row_n, column=9, value=r['checked_at'])

    return len(rows)


def main():
    master_path, rows_path = sys.argv[1], sys.argv[2]
    with open(rows_path, 'r', encoding='utf-8') as f:
        rows = json.load(f)

    wb = openpyxl.load_workbook(master_path, data_only=False)
    ws = wb[SHEET_NAME]
    count = refresh_block(ws, rows)
    wb.save(master_path)
    print(f'Wrote {count} below-alert rows into the reserved top block of "{SHEET_NAME}" in {master_path}')


if __name__ == '__main__':
    main()
