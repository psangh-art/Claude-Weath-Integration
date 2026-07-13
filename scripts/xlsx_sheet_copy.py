"""Cross-workbook worksheet copying, shared by spending_summary.py (preserving
manual tabs across rebuilds) and integrate_spending_tabs.py (mirroring the
spending-summary tabs into Stocks_Buy_Strategy.xlsx). openpyxl has no
cross-workbook copy_worksheet, so this is a cell-by-cell copy of values,
styles, merges, dimensions, freeze panes and tab colour."""
from copy import copy as _copy


def copy_sheet_into(src_ws, dst_ws):
    """Cell-by-cell copy of values + styles + merges + dimensions between
    workbooks (openpyxl has no cross-workbook copy_worksheet)."""
    for row in src_ws.iter_rows():
        for cell in row:
            if cell.__class__.__name__ == 'MergedCell':
                continue
            dst = dst_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                dst.font = _copy(cell.font)
                dst.fill = _copy(cell.fill)
                dst.border = _copy(cell.border)
                dst.alignment = _copy(cell.alignment)
                dst.number_format = cell.number_format
                dst.protection = _copy(cell.protection)
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
