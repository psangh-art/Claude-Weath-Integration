#!/usr/bin/env python3
"""Add/refresh a 'Below Alert Low' sheet in Stocks_Buy_Strategy.xlsx: every ticker
whose live TradingView price (captured during the chart-export step) has dropped
below its Alert Low, sorted worst-gap-first. Placed as the FIRST sheet in the
workbook (tab order), not inserted into 'Stocks of Interest' itself — that sheet
has 121 formula cells and dozens of merged section-header bands, and openpyxl's
insert_rows() has already been shown (during the RELX History fix, 2026-07-10) to
silently corrupt both when rows are inserted mid-sheet. A dedicated sheet avoids
that risk entirely while still satisfying "show this at the top" (first tab).

Usage: python add_below_alert_sheet.py <master.xlsx> <below_alert_rows.json>
  rows: [{"ticker","share_name","price","alert_low","alert_high","gap_pct",
          "holdings","target_value","checked_at"}, ...], worst gap first.
"""
import sys
import json
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

SHEET_NAME = 'Below Alert Low'
HEADER_FILL = PatternFill(fill_type='solid', fgColor='FFF2CC')
GAP_BAD_FONT = Font(color='FFCC0000', bold=True)


def build_sheet(wb, rows):
    if SHEET_NAME in wb.sheetnames:
        del wb[SHEET_NAME]
    ws = wb.create_sheet(SHEET_NAME, 0)  # index 0 = first tab

    ws.cell(row=1, column=1, value='Stocks Trading Below Alert Low').font = Font(bold=True, size=14)
    ws.cell(row=2, column=1, value=(
        f'Auto-generated {datetime.now().date().isoformat()} from tradingview_layouts.xlsx live '
        'price capture cross-checked against Alert Low. Re-run the full pipeline to refresh — '
        'this sheet is fully rebuilt each time, not manually maintained.'
    ))

    header = ['Ticker', 'Share Name', 'Current Price', 'Alert Low', 'Gap %', 'Alert High',
               'Holdings (£)', 'Target Value (£)', 'Price Checked At']
    for c, h in enumerate(header, 1):
        cell = ws.cell(row=4, column=c, value=h)
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL

    for i, r in enumerate(rows):
        row_n = 5 + i
        ws.cell(row=row_n, column=1, value=r['ticker']).font = Font(bold=True)
        ws.cell(row=row_n, column=2, value=r['share_name'])
        ws.cell(row=row_n, column=3, value=round(r['price'], 2))
        ws.cell(row=row_n, column=4, value=r['alert_low'])
        gap_cell = ws.cell(row=row_n, column=5, value=round(r['gap_pct'], 1))
        gap_cell.number_format = '0.0"%"'
        if r['gap_pct'] <= -10:
            gap_cell.font = GAP_BAD_FONT
        ws.cell(row=row_n, column=6, value=r['alert_high'])
        ws.cell(row=row_n, column=7, value=r['holdings'])
        ws.cell(row=row_n, column=8, value=r['target_value'])
        ws.cell(row=row_n, column=9, value=r['checked_at'])

    widths = [10, 24, 14, 12, 10, 12, 14, 16, 24]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    return len(rows)


def main():
    master_path, rows_path = sys.argv[1], sys.argv[2]
    with open(rows_path, 'r', encoding='utf-8') as f:
        rows = json.load(f)

    wb = openpyxl.load_workbook(master_path, data_only=False)
    count = build_sheet(wb, rows)
    wb.save(master_path)
    print(f'Wrote {count} rows to "{SHEET_NAME}" sheet (first tab) in {master_path}')


if __name__ == '__main__':
    main()
