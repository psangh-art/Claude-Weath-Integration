#!/usr/bin/env python3
"""Build tradingview_layouts.xlsx from a manifest JSON produced by export-layouts-excel.js.
One row per individual chart (not per layout) — a layout with 6 panes produces 6 rows,
each with its own cropped screenshot rather than one squeezed multi-pane grid image.

Manifest format: [{"id": int, "chartId": str, "name": str, "ticker": str|None,
"description": str|None, "screenshot": str|None, "error": str|None}, ...]
Usage: python build_layout_excel.py <manifest.json> <output.xlsx>
"""
import sys
import json
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage

MAX_DISPLAY_WIDTH = 700

def main():
    manifest_path, out_path = sys.argv[1], sys.argv[2]
    with open(manifest_path, 'r', encoding='utf-8') as f:
        rows = json.load(f)

    wb = Workbook()
    ws = wb.active
    ws.title = 'Charts'
    ws.append(['Layout ID', 'Chart ID', 'Layout Name', 'Symbol', 'Company', 'Screenshot'])
    for cell in ws[1]:
        cell.font = cell.font.copy(bold=True)
    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 32
    ws.column_dimensions['D'].width = 16
    ws.column_dimensions['E'].width = 32
    ws.column_dimensions['F'].width = 100

    for i, row in enumerate(rows):
        r = i + 2
        if row.get('screenshot'):
            with PILImage.open(row['screenshot']) as im:
                w, h = im.size
            scale = min(1.0, MAX_DISPLAY_WIDTH / w)
            disp_w, disp_h = w * scale, h * scale

            img = XLImage(row['screenshot'])
            img.width, img.height = disp_w, disp_h
            ws.add_image(img, f"F{r}")
            screenshot_value = row['screenshot'].split('\\')[-1].split('/')[-1]
            ws.row_dimensions[r].height = disp_h * 0.75
        else:
            screenshot_value = f"FAILED - {row.get('error', 'unknown error')}"

        ws.cell(row=r, column=1, value=row['id'])
        ws.cell(row=r, column=2, value=row['chartId'])
        c3 = ws.cell(row=r, column=3, value=row['name'])
        c3.font = c3.font.copy(bold=True)
        ws.cell(row=r, column=4, value=row.get('ticker'))
        ws.cell(row=r, column=5, value=row.get('description'))
        ws.cell(row=r, column=6, value=screenshot_value)

    wb.save(out_path)
    print(f"Saved {len(rows)} rows to {out_path}")

if __name__ == '__main__':
    main()
