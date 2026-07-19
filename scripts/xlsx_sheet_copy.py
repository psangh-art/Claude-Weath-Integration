"""Shared openpyxl helpers for this repo's workbook writers.

Cross-workbook worksheet copying, used by spending_summary.py (preserving manual
tabs across rebuilds) and integrate_spending_tabs.py (mirroring the spending-summary
tabs into Stocks_Buy_Strategy.xlsx): openpyxl has no cross-workbook
copy_worksheet, so this is a cell-by-cell copy of values, styles, merges,
dimensions, freeze panes and tab colour.

Plus the two primitives every sheet writer here needs — `copy_cell_style` (match
the neighbouring cell's format, a standing rule for this workbook) and
`offset_formula` (re-point a row-relative formula when a row moves, the fix for the
recurring "formula still points at its old row" class of bug)."""
import re
from copy import copy as _copy


def offset_formula(formula, offset):
    """Add `offset` to the row number of every unquoted, unqualified A1-style
    reference. Text inside double-quoted string literals is left untouched, as
    are absolute-column-only ranges ($C:$I) and bare numeric arguments.

    openpyxl does NOT adjust formula references when rows/columns move, so any
    script that relocates a block has to do this itself — a row that moves while
    its formulas keep pointing at the old row is silently wrong, not an error."""
    parts = formula.split('"')
    for i in range(0, len(parts), 2):  # even indexes are outside quotes
        parts[i] = re.sub(
            r"(?<![A-Za-z0-9:$!])([A-Z]{1,3})(\d{1,5})",
            lambda m: m.group(1) + str(int(m.group(2)) + offset),
            parts[i],
        )
    return '"'.join(parts)


def copy_cell_style(src, dst):
    """Copy one cell's visual style (font/fill/border/alignment/number format/
    protection) to another. openpyxl style objects are shared references, so each
    one MUST be copied — assigning src.font directly makes a later edit of either
    cell change both. Several scripts had grown their own inline version of this
    six-line block; this is the one copy.

    Silently tolerant: a style openpyxl can't reconstruct (rare, from a hand-edited
    workbook) should not abort a sheet write over a cosmetic attribute."""
    if not getattr(src, 'has_style', False):
        return
    try:
        dst.font = _copy(src.font)
        dst.fill = _copy(src.fill)
        dst.border = _copy(src.border)
        dst.alignment = _copy(src.alignment)
        dst.number_format = src.number_format
        dst.protection = _copy(src.protection)
    except Exception:
        pass


def copy_sheet_into(src_ws, dst_ws):
    """Cell-by-cell copy of values + styles + merges + dimensions between
    workbooks (openpyxl has no cross-workbook copy_worksheet)."""
    for row in src_ws.iter_rows():
        for cell in row:
            if cell.__class__.__name__ == 'MergedCell':
                continue
            dst = dst_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            copy_cell_style(cell, dst)
    for m in src_ws.merged_cells.ranges:
        dst_ws.merge_cells(str(m))
    for col, dim in src_ws.column_dimensions.items():
        dst_ws.column_dimensions[col].width = dim.width
    for r, dim in src_ws.row_dimensions.items():
        dst_ws.row_dimensions[r].height = dim.height
    if src_ws.freeze_panes:
        dst_ws.freeze_panes = src_ws.freeze_panes
    if src_ws.sheet_properties.tabColor:
        dst_ws.sheet_properties.tabColor = _copy(src_ws.sheet_properties.tabColor)


def replace_sheet(dst_wb, name, src_ws):
    """Replace (or create) the sheet called `name` in dst_wb with a copy of
    src_ws, keeping its tab position if it already existed. Returns 'replaced'
    or 'added'."""
    if name in dst_wb.sheetnames:
        idx = dst_wb.sheetnames.index(name)
        del dst_wb[name]
        dst_ws = dst_wb.create_sheet(name, idx)
        outcome = 'replaced'
    else:
        dst_ws = dst_wb.create_sheet(name)
        outcome = 'added'
    copy_sheet_into(src_ws, dst_ws)
    return outcome
