#!/usr/bin/env python3
"""One-off (2026-07-12, user-requested): arrange the Stats sheet's charts in a
tabular grid — 3 across the top, continuing downward — instead of the scattered
positions they inherited from their original History-sheet anchors (rows 60-106,
where two of them even overlapped).

Each chart keeps its own width/height; grid cells are sized to the largest
chart plus a one-cell gutter, so this stays correct if charts are added later.

Usage: python arrange_stats_charts_2026-07-12.py <workbook.xlsx>
"""
import sys

import openpyxl

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

PER_ROW = 3
GUTTER = 1  # blank cells between charts


def main():
    path = sys.argv[1]
    wb = openpyxl.load_workbook(path)
    ws = wb['Stats']
    charts = list(ws._charts)
    if not charts:
        print('No charts on Stats — nothing to do.')
        return

    sizes = []
    for ch in charts:
        f, t = ch.anchor._from, ch.anchor.to
        sizes.append((t.col - f.col, t.row - f.row))
    cell_w = max(w for w, _ in sizes) + GUTTER
    cell_h = max(h for _, h in sizes) + GUTTER

    for i, ch in enumerate(charts):
        w, h = sizes[i]
        col0 = (i % PER_ROW) * cell_w
        row0 = (i // PER_ROW) * cell_h
        a = ch.anchor
        a._from.col, a._from.row = col0, row0
        a._from.colOff = a._from.rowOff = 0
        a.to.col, a.to.row = col0 + w, row0 + h
        a.to.colOff = a.to.rowOff = 0
        print(f'  chart {i + 1}: -> from (col {col0}, row {row0}) to (col {col0 + w}, row {row0 + h})')

    wb.save(path)
    print(f'{len(charts)} chart(s) arranged {PER_ROW}-across from the top. Saved -> {path}')


if __name__ == '__main__':
    main()
