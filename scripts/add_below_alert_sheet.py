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
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

SHEET_NAME = 'Stocks of Interest'
RESERVED_BLOCK = 40      # rows 1..40 belong to this table; row 41+ is the
                         # pre-existing Stocks of Interest content — never touch it
TITLE_ROW = 1
NOTE_ROW = 2
BAND_ROW = 3             # coloured section band, like the section headers below
HEADER_ROW = 4
DATA_START_ROW = 5
MAX_DATA_ROWS = 34       # rows 5..38; 39-40 stay blank as a separator
LAST_COL = 10            # this table occupies columns A..J

# Per-row click-through to the stock's TradingView layout, matching the
# Investments sheet's "TradingView" column (=HYPERLINK(chart/<chartId>/)).
TV_LAYOUT_URL = 'https://www.tradingview.com/chart/{chart_id}/'
TV_LINK_LABEL = '📊 Layout'

# Palette + typography lifted from the section tables lower down this same sheet
# (rows 41+), so the below-alert table reads as part of the same document rather
# than a foreign block. Verified against the live workbook 2026-07-12.
FONT_NAME = 'Arial'
WHITE = 'FFFFFFFF'
NAVY = 'FF1F3864'          # title band + ticker text (identity colour)
SUBHEAD_FILL = 'FF2E5077'  # subtitle/note band + column-header band
BAND_FILL = 'FFC00000'     # section band — red extends the existing green/amber/
                           # blue priority ladder: "below alert low" is the most
                           # urgent state (change here if a different colour is wanted)
DATA_FILL = 'FFFDF0F0'     # pale-red data tint, matching the per-section pale tints
                           # (green FFF0FFF4 / amber FFFFFDF0 / watchlist FFF7F9FC)
GAP_BAD = 'FFCC0000'       # red-bold gap when a stock is >10% below its alert low

_THIN = Side(style='thin')
BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_TOP_LEFT = Alignment(horizontal='left', vertical='top', wrap_text=False)
_TOP_LEFT_WRAP = Alignment(horizontal='left', vertical='top', wrap_text=True)
_TOP_RIGHT = Alignment(horizontal='right', vertical='top', wrap_text=False)

# Columns reordered to match the section tables below: Stock name first (wide col
# A), Ticker second (narrow col B). Numeric columns are right-aligned with the
# same formats the tables below use (#,##0.00 for prices, #,##0 for whole-pound £).
HEADER = ['Stock', 'Ticker', 'Current Price', 'Alert Low', 'Gap %', 'Alert High',
          'Holdings (£)', 'Target Value (£)', 'Price Checked At', 'TradingView']
_NUM_FMT = {3: '#,##0.00', 4: '#,##0.00', 5: '0.0"%"', 6: '#,##0.00', 7: '#,##0', 8: '#,##0'}


def _band(ws, row, text, fill, size, bold, italic=False):
    """Render a full-width (A..J) coloured band on `row`, merged, bordered, with
    `text` in the top-left cell — the same treatment the section headers use."""
    # Unmerge any existing band on this row FIRST (a previous run may have used a
    # different width), so the cells below are writable rather than read-only
    # MergedCells — then re-merge to the current width. Keeps this idempotent.
    for m in list(ws.merged_cells.ranges):
        if m.min_row == row and m.max_row == row:
            ws.unmerge_cells(str(m))
    for c in range(1, LAST_COL + 1):
        cell = ws.cell(row=row, column=c)
        cell.value = None
        cell.fill = PatternFill(fill_type='solid', fgColor=fill)
        cell.border = BORDER
        cell.alignment = _TOP_LEFT
        cell.font = Font(name=FONT_NAME, size=size, bold=bold, italic=italic, color=WHITE)
    ws.cell(row=row, column=1, value=text)
    ws.merge_cells(f'A{row}:{chr(64 + LAST_COL)}{row}')
    ws.row_dimensions[row].height = 15


def refresh_block(ws, rows):
    if len(rows) > MAX_DATA_ROWS:
        print(f'WARNING: {len(rows)} below-alert rows but only {MAX_DATA_ROWS} fit in the '
              f'reserved block — writing the {MAX_DATA_ROWS} worst; '
              f'{len(rows) - MAX_DATA_ROWS} omitted.', file=sys.stderr)
        rows = rows[:MAX_DATA_ROWS]

    # Title / note / section bands (rows 1-3), mirroring the section-table header
    # stack below (navy title, FF2E5077 italic subtitle, coloured section band).
    _band(ws, TITLE_ROW, 'Stocks Trading Below Alert Low', NAVY, size=13, bold=True)
    _band(ws, NOTE_ROW, (
        f'Auto-generated {datetime.now().date().isoformat()} from tradingview_layouts.xlsx live '
        'price capture cross-checked against Alert Low. Rebuilt by the pipeline into rows '
        f'1-{RESERVED_BLOCK} of this sheet each run — do not add manual content above row '
        f'{RESERVED_BLOCK + 1}.'
    ), SUBHEAD_FILL, size=8, bold=False, italic=True)
    _band(ws, BAND_ROW, '🔴  BELOW ALERT LOW — price has fallen through the alert level  (Action required)',
          BAND_FILL, size=10, bold=True)

    # Column header row (FF2E5077 band, bold white, wrapped) — matches row 44 below.
    for c, h in enumerate(HEADER, 1):
        cell = ws.cell(row=HEADER_ROW, column=c, value=h)
        cell.font = Font(name=FONT_NAME, size=8, bold=True, color=WHITE)
        cell.fill = PatternFill(fill_type='solid', fgColor=SUBHEAD_FILL)
        cell.border = BORDER
        cell.alignment = _TOP_LEFT_WRAP
    ws.row_dimensions[HEADER_ROW].height = 15

    # Reset the whole data region first so a shorter run leaves no styled ghosts.
    for r in range(DATA_START_ROW, DATA_START_ROW + MAX_DATA_ROWS):
        for c in range(1, LAST_COL + 1):
            cell = ws.cell(row=r, column=c)
            cell.value = None
            cell.font = Font()
            cell.fill = PatternFill()
            cell.border = Border()
            cell.alignment = Alignment()
            cell.number_format = 'General'
        ws.row_dimensions[r].height = None

    for i, r in enumerate(rows):
        row_n = DATA_START_ROW + i
        # base data-row styling on every cell (pale-red fill, thin border, Arial 8)
        for c in range(1, LAST_COL + 1):
            cell = ws.cell(row=row_n, column=c)
            cell.fill = PatternFill(fill_type='solid', fgColor=DATA_FILL)
            cell.border = BORDER
            cell.font = Font(name=FONT_NAME, size=8)
            cell.alignment = _TOP_LEFT
        ws.row_dimensions[row_n].height = 15

        name_cell = ws.cell(row=row_n, column=1, value=r['share_name'])
        name_cell.font = Font(name=FONT_NAME, size=9, bold=True, color='FF000000')
        ticker_cell = ws.cell(row=row_n, column=2, value=r['ticker'])
        ticker_cell.font = Font(name=FONT_NAME, size=8, bold=True, color=NAVY)

        ws.cell(row=row_n, column=3, value=round(r['price'], 2))
        ws.cell(row=row_n, column=4, value=r['alert_low'])
        gap_cell = ws.cell(row=row_n, column=5, value=round(r['gap_pct'], 1))
        if r['gap_pct'] <= -10:
            gap_cell.font = Font(name=FONT_NAME, size=8, bold=True, color=GAP_BAD)
        ws.cell(row=row_n, column=6, value=r['alert_high'])
        # Holdings/Target £ display as whole pounds (user rule, 2026-07-11).
        ws.cell(row=row_n, column=7, value=r['holdings'])
        ws.cell(row=row_n, column=8, value=r['target_value'])
        ws.cell(row=row_n, column=9, value=r['checked_at'])

        # Click-through to this stock's TradingView layout (blank if no chart).
        chart_id = r.get('chart_id')
        tv_cell = ws.cell(row=row_n, column=10)
        if chart_id:
            url = TV_LAYOUT_URL.format(chart_id=chart_id)
            tv_cell.value = f'=HYPERLINK("{url}","{TV_LINK_LABEL}")'
            tv_cell.font = Font(name=FONT_NAME, size=8, color=NAVY, underline='single')
        tv_cell.alignment = _TOP_LEFT

        for c, fmt in _NUM_FMT.items():
            cell = ws.cell(row=row_n, column=c)
            cell.number_format = fmt
            cell.alignment = _TOP_RIGHT

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
