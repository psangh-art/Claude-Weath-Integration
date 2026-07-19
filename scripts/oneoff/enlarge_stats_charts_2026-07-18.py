"""One-off (2026-07-18, user-requested): make the Stats-sheet charts 50% bigger and
keep them on a single page.

The three charts are TwoCellAnchor (sized by the cells they span, ~8 cols x 16 rows
each, laid out 3-across by arrange_stats_charts_2026-07-12.py). This scales each
chart's cell span by 1.5, re-grids them 3-across without overlap, and sets the sheet
to fit a single page (landscape, fitToWidth=1, fitToHeight=1) so the bigger charts
still print/view on one page. Stats is a hand-maintained master tab (not regenerated
by the pipeline), so this persists. Backs the workbook up first; re-runnable
(idempotent once the charts are already at the scaled size + grid).

Usage: python enlarge_stats_charts_2026-07-18.py [workbook.xlsx]
"""
import datetime
import os
import shutil
import sys

import openpyxl
from openpyxl.worksheet.properties import PageSetupProperties

SCALE = 1.5
PER_ROW = 3
GUTTER = 1


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        os.path.expanduser("~"), "Downloads", "Stocks_Buy_Strategy.xlsx")
    if not os.path.exists(path):
        print("Not found:", path); sys.exit(1)

    wb = openpyxl.load_workbook(path)
    ws = wb["Stats"]
    charts = list(ws._charts)
    if not charts:
        print("No charts on Stats — nothing to do."); return

    # Target span = current span x 1.5 (use the largest chart so the grid is uniform).
    spans = [(c.anchor.to.col - c.anchor._from.col, c.anchor.to.row - c.anchor._from.row)
             for c in charts]
    base_w = max(w for w, _ in spans)
    base_h = max(h for _, h in spans)
    new_w = max(round(base_w * SCALE), base_w + 1)
    new_h = max(round(base_h * SCALE), base_h + 1)

    # Idempotency: if already at the scaled grid, don't re-back-up/re-save needlessly.
    already = all((c.anchor.to.col - c.anchor._from.col) == new_w and
                  (c.anchor.to.row - c.anchor._from.row) == new_h for c in charts)
    if already:
        print(f"Stats charts already {new_w}x{new_h} cells — nothing to do.")
        return

    bak = path + ".bak-stats-enlarge-" + datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    shutil.copyfile(path, bak)
    print("Backed up ->", os.path.basename(bak))

    cell_w, cell_h = new_w + GUTTER, new_h + GUTTER
    for i, ch in enumerate(charts):
        row_i, col_i = divmod(i, PER_ROW)
        col0, row0 = col_i * cell_w, row_i * cell_h
        a = ch.anchor
        a._from.col, a._from.colOff = col0, 0
        a._from.row, a._from.rowOff = row0, 0
        a.to.col, a.to.colOff = col0 + new_w, 0
        a.to.row, a.to.rowOff = row0 + new_h, 0

    # Fit everything onto a single page.
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    if ws.sheet_properties.pageSetUpPr is None:
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    else:
        ws.sheet_properties.pageSetUpPr.fitToPage = True

    wb.save(path)
    print(f"Enlarged {len(charts)} Stats charts to {new_w}x{new_h} cells "
          f"({int((SCALE-1)*100)}% bigger), 3-across, fit to one page.")


if __name__ == "__main__":
    main()
