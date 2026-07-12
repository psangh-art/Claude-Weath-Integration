#!/usr/bin/env python3
"""Build a PowerPoint review deck from the latest pipeline run: sorted by layout
then chart, ONE SLIDE PER INVESTMENT, showing the captured chart image alongside
everything the pipeline knows about it (live price, master-sheet holdings/alert
levels, channel read, TradingView alerts) — with loud red/amber flags wherever a
chart or alert is MISSING, so the whole run can be eyeballed in PowerPoint Online.

Data sources (all produced by run_full_pipeline.js):
  scripts/layout_manifest_tmp.json    charts + screenshots + live prices
  scripts/alerts_manifest_tmp.json    live TradingView alerts
  scripts/channel_results_tmp.json    OCR channel-boundary reads
  ~/Downloads/Stocks_Buy_Strategy.xlsx  master sheet ('Investments' tab)

Usage: python build_review_deck.py [out.pptx]
  (default out: ~/Downloads/Investment_Review_Deck.pptx)
"""
import json
import os
import sys
from datetime import datetime

import openpyxl
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from PIL import Image

from ticker_normalize import normalize, master_tickers_match

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DOWNLOADS = os.path.join(os.path.expanduser('~'), 'Downloads')
MASTER_PATH = os.path.join(DOWNLOADS, 'Stocks_Buy_Strategy.xlsx')

# Master 'Investments' columns (same map as update_master_sheet.py)
COL_SHARE_NAME, COL_TICKER, COL_HOLDINGS, COL_TARGET = 2, 3, 4, 6
COL_ALERT_LOW, COL_ALERT_LOW_SOURCE, COL_ALERT_HIGH = 12, 13, 15

NAVY = RGBColor(0x1F, 0x38, 0x64)
RED = RGBColor(0xCC, 0x00, 0x00)
AMBER = RGBColor(0xB9, 0x77, 0x0E)
GREEN = RGBColor(0x27, 0x62, 0x21)
GREY = RGBColor(0x60, 0x60, 0x60)

SLIDE_W, SLIDE_H = Inches(13.333), Inches(7.5)


def load_json(name):
    path = os.path.join(SCRIPT_DIR, name)
    if not os.path.exists(path):
        return None
    with open(path, encoding='utf-8') as f:
        return json.load(f)


def load_master_index():
    """ticker(str, master form) -> dict of the master row's review-relevant cells."""
    if not os.path.exists(MASTER_PATH):
        return {}
    wb = openpyxl.load_workbook(MASTER_PATH, data_only=False)
    ws = wb['Investments']
    index = {}
    for r in range(1, ws.max_row + 1):
        t = ws.cell(row=r, column=COL_TICKER).value
        if not isinstance(t, str) or not t.strip() or t.strip().upper() == 'TICKER':
            continue  # skip blanks and the header row
        index[t.strip().upper()] = {
            'row': r,
            'share_name': ws.cell(row=r, column=COL_SHARE_NAME).value,
            'holdings': ws.cell(row=r, column=COL_HOLDINGS).value,
            'target': ws.cell(row=r, column=COL_TARGET).value,
            'alert_low': ws.cell(row=r, column=COL_ALERT_LOW).value,
            'alert_low_source': ws.cell(row=r, column=COL_ALERT_LOW_SOURCE).value,
            'alert_high': ws.cell(row=r, column=COL_ALERT_HIGH).value,
        }
    return index


def find_master(index, chart_ticker):
    norm = normalize(chart_ticker)
    if not norm or not norm['master_ticker']:
        return None, None
    for key, row in index.items():
        if master_tickers_match(key, norm['master_ticker']):
            return key, row
    return norm['master_ticker'], None


def alerts_for(alerts, chart_ticker):
    """TradingView alerts whose symbol's post-colon part matches this chart ticker."""
    out = []
    want = chart_ticker.strip().upper()
    for a in alerts or []:
        sym = (a.get('symbol') or '').split(':')[-1].strip().upper()
        if sym == want or sym.rstrip('.') == want.rstrip('.'):
            out.append(a)
    return out


def fmt_pounds(v):
    if isinstance(v, (int, float)):
        return f'£{v:,.0f}'
    return '—'


def fmt_num(v):
    if isinstance(v, (int, float)):
        return f'{v:,.2f}'.rstrip('0').rstrip('.')
    return '—'


def add_text(slide, x, y, w, h, lines, align=PP_ALIGN.LEFT):
    """lines: list of (text, size_pt, bold, colour) tuples, one paragraph each."""
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    first = True
    for text, size, bold, colour in lines:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        run = p.add_run()
        run.text = text
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = colour
    return box


def add_flag_banner(slide, y, text):
    box = add_text(slide, Inches(8.55), y, Inches(4.5), Inches(0.35),
                   [(text, 13, True, RGBColor(0xFF, 0xFF, 0xFF))])
    box.fill.solid()
    box.fill.fore_color.rgb = RED
    return y + Inches(0.42)


def add_picture_fitted(slide, path, x, y, max_w, max_h):
    with Image.open(path) as im:
        pw, ph = im.size
    scale = min(max_w / pw, max_h / ph)
    w, h = int(pw * scale), int(ph * scale)
    slide.shapes.add_picture(path, x, y + Emu(int((max_h - h) / 2)), width=Emu(w), height=Emu(h))


def chart_slide(prs, chart, layout_name, master_key, master, channel, tv_alerts):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank

    ticker = chart.get('ticker') or '?'
    desc = chart.get('description') or ''
    add_text(slide, Inches(0.35), Inches(0.15), Inches(9.5), Inches(0.55),
             [(f'{ticker} — {desc}', 24, True, NAVY)])
    add_text(slide, Inches(10.0), Inches(0.22), Inches(3.0), Inches(0.4),
             [(layout_name, 12, False, GREY)], align=PP_ALIGN.RIGHT)

    # Chart image (left, large) or a loud MISSING placeholder
    img = chart.get('screenshot')
    img_x, img_y = Inches(0.35), Inches(0.85)
    img_w, img_h = Inches(8.0), Inches(6.3)
    if img and os.path.exists(img):
        add_picture_fitted(slide, img, img_x, img_y, img_w, img_h)
    else:
        box = add_text(slide, img_x, Inches(3.2), img_w, Inches(1.2),
                       [('⚠ MISSING CHART', 36, True, RGBColor(0xFF, 0xFF, 0xFF)),
                        (chart.get('error') or 'no screenshot captured this run', 14, False,
                         RGBColor(0xFF, 0xFF, 0xFF))],
                       align=PP_ALIGN.CENTER)
        box.fill.solid()
        box.fill.fore_color.rgb = RED

    # Info column (right)
    x, w = Inches(8.55), Inches(4.5)
    y = Inches(0.85)

    if master is None:
        y = add_flag_banner(slide, y, '⚠ NOT IN MASTER SHEET')
    has_tv_alert = bool(tv_alerts)
    has_master_alert = master is not None and isinstance(master.get('alert_low'), (int, float))
    if not has_tv_alert and not has_master_alert:
        y = add_flag_banner(slide, y, '⚠ NO ALERTS (TradingView or master)')

    lines = [(f"Live price: {fmt_num(chart.get('price'))}", 15, True, NAVY),
             (f"  captured {chart.get('priceCheckedAt') or '—'}", 10, False, GREY)]

    if master:
        lines += [
            ('Master sheet (Investments)', 13, True, NAVY),
            (f"  Share name: {master.get('share_name') or '—'}", 12, False, GREY),
            (f"  Holdings: {fmt_pounds(master.get('holdings'))}    "
             f"Target: {fmt_pounds(master.get('target'))}", 12, False, GREY),
            (f"  Alert Low: {fmt_num(master.get('alert_low'))} "
             f"({master.get('alert_low_source') or 'no source'})", 12,
             not has_master_alert, RED if not has_master_alert else GREEN),
            (f"  Alert High: {fmt_num(master.get('alert_high'))}", 12, False, GREY),
        ]

    lines.append(('Channel read (OCR)', 13, True, NAVY))
    if channel and channel.get('kind') not in (None, 'rejected'):
        lines.append((f"  {channel['kind']}: lower {fmt_num(channel.get('lower'))}, "
                      f"upper {fmt_num(channel.get('upper'))}", 12, False, GREEN))
    elif channel:
        lines.append((f"  rejected: {channel.get('reason') or 'no reliable read'}", 12, False, AMBER))
    else:
        lines.append(('  not attempted this run', 12, False, GREY))

    lines.append((f'TradingView alerts ({len(tv_alerts)})', 13, True,
                  NAVY if has_tv_alert else RED))
    for a in tv_alerts[:6]:
        status = 'active' if a.get('active') else 'INACTIVE'
        lines.append((f"  {fmt_num(a.get('targetPrice'))} {a.get('conditionType') or ''} "
                      f"[{status}] exp {str(a.get('expiration') or '—')[:10]}", 11, False,
                      GREY if a.get('active') else AMBER))
    if len(tv_alerts) > 6:
        lines.append((f'  … and {len(tv_alerts) - 6} more', 11, False, GREY))

    add_text(slide, x, y, w, Inches(7.3) - y, lines)


def section_slide(prs, title, subtitle=''):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = add_text(slide, Inches(0.8), Inches(3.0), Inches(11.7), Inches(1.5),
                   [(title, 40, True, RGBColor(0xFF, 0xFF, 0xFF))] +
                   ([(subtitle, 16, False, RGBColor(0xD0, 0xD8, 0xE8))] if subtitle else []),
                   align=PP_ALIGN.CENTER)
    box.fill.solid()
    box.fill.fore_color.rgb = NAVY


def main():
    out_path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(DOWNLOADS, 'Investment_Review_Deck.pptx')

    charts = load_json('layout_manifest_tmp.json') or []
    alerts = load_json('alerts_manifest_tmp.json') or []
    channels = {c['ticker']: c for c in (load_json('channel_results_tmp.json') or []) if c.get('ticker')}
    master_index = load_master_index()

    prs = Presentation()
    prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H

    # ── summary page ────────────────────────────────────────────────────────
    missing_charts = [c['ticker'] for c in charts
                      if not (c.get('screenshot') and os.path.exists(c['screenshot']))]
    no_alert_tickers = []
    seen = set()
    for c in charts:
        t = c.get('ticker')
        if not t or t in seen:
            continue
        seen.add(t)
        _, master = find_master(master_index, t)
        if not alerts_for(alerts, t) and not (master and isinstance(master.get('alert_low'), (int, float))):
            no_alert_tickers.append(t)
    charted_master = set()
    for c in charts:
        key, m = find_master(master_index, c.get('ticker') or '')
        if m:
            charted_master.add(key)
    unchartered = sorted(k for k in master_index if k not in charted_master)
    # A missing chart matters most where money is actually held; the rest of the
    # master sheet is watchlist coverage.
    unchartered_held = [k for k in unchartered
                        if isinstance(master_index[k].get('holdings'), (int, float))
                        and master_index[k]['holdings'] > 0]
    unchartered_watch = [k for k in unchartered if k not in set(unchartered_held)]

    layout_names = []
    for c in charts:
        if c['name'] not in layout_names:
            layout_names.append(c['name'])

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_text(slide, Inches(0.5), Inches(0.4), Inches(12.3), Inches(0.7),
             [(f'Investment Review Deck — {datetime.now():%Y-%m-%d %H:%M}', 28, True, NAVY)])
    summary = [
        (f'{len(layout_names)} layouts, {len(charts)} charts, {len(alerts)} TradingView alerts, '
         f'{len(master_index)} master-sheet rows', 16, False, GREY),
        ('', 12, False, GREY),
        (f'Charts missing a screenshot: {len(missing_charts)}'
         + (f"  ({', '.join(missing_charts)})" if missing_charts else ' ✓'),
         16, bool(missing_charts), RED if missing_charts else GREEN),
        (f'Charted tickers with NO alert anywhere: {len(no_alert_tickers)}'
         + (f"  ({', '.join(no_alert_tickers)})" if no_alert_tickers else ' ✓'),
         16, bool(no_alert_tickers), RED if no_alert_tickers else GREEN),
        (f'HELD investments (holdings > £0) with NO chart: {len(unchartered_held)}'
         + (f"  ({', '.join(unchartered_held)})" if unchartered_held else ' ✓'),
         16, bool(unchartered_held), RED if unchartered_held else GREEN),
        (f'Watchlist-only master rows with no chart: {len(unchartered_watch)} '
         '(full list on the appendix slide at the end)', 14, False, GREY),
    ]
    add_text(slide, Inches(0.5), Inches(1.4), Inches(12.3), Inches(5.5), summary)

    # ── one slide per chart, grouped by layout in capture order ────────────
    current_layout = None
    for chart in charts:
        if chart['name'] != current_layout:
            current_layout = chart['name']
            n = sum(1 for c in charts if c['name'] == current_layout)
            section_slide(prs, current_layout, f'{n} chart(s)')
        _, master = find_master(master_index, chart.get('ticker') or '')
        chart_slide(prs, chart, current_layout, None, master,
                    channels.get(chart.get('ticker')), alerts_for(alerts, chart.get('ticker') or ''))

    # ── appendix: master rows without a chart ───────────────────────────────
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_text(slide, Inches(0.5), Inches(0.3), Inches(12.3), Inches(0.6),
             [('Appendix — master-sheet rows with no TradingView chart', 22, True, NAVY)])
    appendix = []
    if unchartered_held:
        appendix.append(('Held investments (needs a chart adding):', 15, True, RED))
        for k in unchartered_held:
            m = master_index[k]
            appendix.append((f"  {k} — {m.get('share_name') or ''} "
                             f"(holdings {fmt_pounds(m.get('holdings'))})", 13, False, GREY))
        appendix.append(('', 10, False, GREY))
    appendix.append((f'Watchlist-only ({len(unchartered_watch)}):', 15, True, NAVY))
    appendix.append((', '.join(unchartered_watch) or '—', 10, False, GREY))
    add_text(slide, Inches(0.5), Inches(1.1), Inches(12.3), Inches(6.1), appendix)

    prs.save(out_path)
    n_slides = len(prs.slides._sldIdLst)
    print(f'Deck: {n_slides} slides -> {out_path}')
    print(f'Missing charts: {len(missing_charts)}; no-alert tickers: {len(no_alert_tickers)}; '
          f'held-with-no-chart: {len(unchartered_held)}; watchlist-no-chart: {len(unchartered_watch)}')


if __name__ == '__main__':
    main()
