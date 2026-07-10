#!/usr/bin/env python3
"""Apply channel-detection results from a TradingView chart export into the master
'Stocks Buy Strategy.xlsx' sheet, following the rules in
Claude_Code_Handoff_Instructions.md sections 3-4-5-9. Never writes a value it isn't
confident in — silence (leaving a row untouched) is always preferable to a wrong
number in a live trading sheet.

Usage:
  python update_master_sheet.py <master_in.xlsx> <charts_manifest.json> \
      <channel_results.json> <master_out.xlsx> <feedback_md_path>

charts_manifest.json: Charts-sheet rows from export-layouts-excel.js
  [{"id", "chartId", "name", "ticker", "description", "screenshot", "error"}, ...]
channel_results.json: output of `channel_detect.py --batch`
  [{"ticker", "screenshot", "lower", "upper", "x_frac", "reason"}, ...]
"""
import sys
import json
import re
from datetime import date
from copy import copy

import openpyxl

from ticker_normalize import normalize, master_tickers_match

SHEET_NAME = 'Stocks Buy Strategy'
COL_CHART = 1
COL_TICKER = 3
COL_ALERT_LOW = 12
COL_ALERT_LOW_SOURCE = 13
COL_ALERT_HIGH = 15
COL_CLAUDE_NOTES = 31
HEADER_ROW = 2

CHART_YES_FONT = 'FF276221'
CHART_YES_FILL = 'FFC6EFCE'
CHART_NO_FONT = 'FF666666'
CHART_NO_FILL = 'FFF2F2F2'

REFRESH_NOISE_THRESHOLD = 0.03  # don't rewrite if new Alert Low is within 3% of existing


def load_json(path):
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def build_master_index(ws):
    """Map normalized master ticker -> row number, skipping section-header rows
    (rows with no Chart Yes/No value in column A aren't real ticker rows)."""
    index = {}
    for r in range(HEADER_ROW + 1, ws.max_row + 1):
        chart_val = ws.cell(row=r, column=COL_CHART).value
        ticker = ws.cell(row=r, column=COL_TICKER).value
        if chart_val not in ('Yes', 'No') or not ticker:
            continue
        index[str(ticker).strip().upper().replace('.', '-')] = r
    return index


def set_chart_flag(ws, row, is_yes):
    cell = ws.cell(row=row, column=COL_CHART)
    cell.value = 'Yes' if is_yes else 'No'
    font = copy(cell.font)
    font.color = openpyxl.styles.colors.Color(rgb=CHART_YES_FONT if is_yes else CHART_NO_FONT)
    cell.font = font
    fill = openpyxl.styles.PatternFill(fill_type='solid', fgColor=CHART_YES_FILL if is_yes else CHART_NO_FILL)
    cell.fill = fill


def append_note(ws, row, note):
    cell = ws.cell(row=row, column=COL_CLAUDE_NOTES)
    existing = cell.value
    cell.value = f"{existing} | {note}" if existing else note


def process(master_ws, charts, channel_by_ticker):
    """Returns (applied, rejected, skipped_manual, skipped_noise, unmatched) lists of
    dicts describing what happened for each charted ticker, for the feedback log."""
    index = build_master_index(master_ws)
    today = date.today().isoformat()

    applied, rejected, skipped_manual, skipped_noise, unmatched = [], [], [], [], []
    seen_tickers = set()

    for row in charts:
        raw_ticker = row.get('ticker')
        if not raw_ticker:
            continue
        norm = normalize(raw_ticker, row.get('description'))
        if norm is None or norm['kind'] == 'macro_excluded' or norm['master_ticker'] is None:
            continue
        master_ticker = norm['master_ticker']
        if master_ticker in seen_tickers:
            continue  # same ticker may appear in multiple layouts; only process once per run
        seen_tickers.add(master_ticker)

        company = row.get('description') or raw_ticker
        detection = channel_by_ticker.get(raw_ticker) or channel_by_ticker.get(master_ticker)

        master_row = None
        for key, r in index.items():
            if master_tickers_match(key, master_ticker):
                master_row = r
                break

        if master_row is None:
            unmatched.append({'ticker': master_ticker, 'company': company, 'reason': 'no existing row in master sheet'})
            continue

        if detection is None:
            unmatched.append({'ticker': master_ticker, 'company': company, 'reason': 'not yet attempted (no channel-detection result for this run)'})
            continue

        if detection.get('reason'):
            rejected.append({'ticker': master_ticker, 'company': company, 'reason': detection['reason']})
            continue

        existing_source = master_ws.cell(row=master_row, column=COL_ALERT_LOW_SOURCE).value
        if existing_source == 'Manual':
            skipped_manual.append({'ticker': master_ticker, 'company': company})
            continue

        lower, upper = detection['lower'], detection['upper']
        alert_low = round(lower * 1.05, 2)
        alert_high = upper

        existing_low = master_ws.cell(row=master_row, column=COL_ALERT_LOW).value
        if existing_source == 'Auto' and isinstance(existing_low, (int, float)) and existing_low:
            pct_change = abs(alert_low - existing_low) / existing_low
            if pct_change <= REFRESH_NOISE_THRESHOLD:
                skipped_noise.append({'ticker': master_ticker, 'company': company, 'existing_low': existing_low, 'new_low': alert_low})
                continue

        master_ws.cell(row=master_row, column=COL_ALERT_LOW, value=alert_low)
        master_ws.cell(row=master_row, column=COL_ALERT_LOW_SOURCE, value='Auto')
        master_ws.cell(row=master_row, column=COL_ALERT_HIGH, value=alert_high)
        set_chart_flag(master_ws, master_row, True)
        append_note(master_ws, master_row,
                    f"Auto: Alert Low {alert_low}, Alert High {alert_high} ({today})")

        applied.append({'ticker': master_ticker, 'company': company, 'lower': lower, 'upper': upper,
                         'alert_low': alert_low, 'alert_high': alert_high})

    return applied, rejected, skipped_manual, skipped_noise, unmatched


_TRACKER_HEADING_RE = re.compile(r'^##.*Coverage Tracker.*$', re.MULTILINE)
_TABLE_ROW_RE = re.compile(r'^\|(.+)\|\s*$')


def _parse_tracker_table(block):
    """Parse the '| Ticker | Company | Status | Why | Last checked |' table in the
    Coverage Tracker block into a dict keyed by ticker. Returns (rows_dict, header_lines)
    or (None, None) if the table isn't in the expected shape."""
    lines = block.splitlines()
    table_start = None
    for i, line in enumerate(lines):
        if _TABLE_ROW_RE.match(line) and 'Ticker' in line:
            table_start = i
            break
    if table_start is None:
        return None, None
    header_lines = lines[table_start:table_start + 2]  # header + separator
    rows = {}
    row_order = []
    i = table_start + 2
    while i < len(lines) and _TABLE_ROW_RE.match(lines[i]):
        cells = [c.strip() for c in lines[i].strip().strip('|').split('|')]
        if len(cells) >= 5:
            ticker = cells[0]
            rows[ticker] = {'company': cells[1], 'status': cells[2], 'why': cells[3], 'last_checked': cells[4]}
            row_order.append(ticker)
        i += 1
    return {'rows': rows, 'order': row_order, 'header': header_lines, 'start': table_start,
            'end': i, 'lines': lines}, None


def _update_coverage_tracker(content, applied, rejected, unmatched, today):
    """Best-effort update of the Coverage Tracker table: remove resolved tickers (move
    them to a 'Resolved since last update' bullet list) and add/refresh unresolved ones.
    Returns the updated content, or the original content unchanged if the table isn't
    in a shape this can safely parse (never guess at document structure)."""
    heading_match = _TRACKER_HEADING_RE.search(content)
    if not heading_match:
        return content, False

    # Coverage Tracker block runs from its heading to the next '---' after it.
    rest = content[heading_match.start():]
    end_match = re.search(r'\n---\n', rest)
    if not end_match:
        return content, False
    block = rest[:end_match.start()]

    parsed, _ = _parse_tracker_table(block)
    if parsed is None:
        return content, False

    rows, order, lines = parsed['rows'], parsed['order'], parsed['lines']

    resolved_bullets = []
    for a in applied:
        if a['ticker'] in rows:
            del rows[a['ticker']]
            order.remove(a['ticker'])
            resolved_bullets.append(f"- **{a['ticker']}** ✅ — Applied: Alert Low {a['alert_low']}, Alert High {a['alert_high']}.")

    for item, reason_key in [(r, 'reason') for r in rejected] + [(u, 'reason') for u in unmatched]:
        ticker = item['ticker']
        why = item[reason_key]
        if ticker in rows:
            rows[ticker]['why'] = why
            rows[ticker]['last_checked'] = today
        else:
            rows[ticker] = {'company': item.get('company', ''), 'status': '⚪ Not yet attempted', 'why': why, 'last_checked': today}
            order.append(ticker)

    new_table_lines = list(parsed['header'])
    for ticker in order:
        r = rows[ticker]
        new_table_lines.append(f"| {ticker} | {r['company']} | {r['status']} | {r['why']} | {r['last_checked']} |")

    updated_lines = lines[:parsed['start']] + new_table_lines + lines[parsed['end']:]

    # Merge resolved bullets into the existing "Resolved since last update:" list, if any.
    updated_block = "\n".join(updated_lines)
    if resolved_bullets:
        resolved_marker = re.search(r'\*\*Resolved since last update:\*\*\n((?:- .*\n?)*)', updated_block)
        if resolved_marker:
            insert_at = resolved_marker.end(1)
            addition = "\n".join(resolved_bullets) + "\n"
            updated_block = updated_block[:insert_at] + addition + updated_block[insert_at:]
        else:
            updated_block = updated_block.rstrip('\n') + "\n\n**Resolved since last update:**\n" + "\n".join(resolved_bullets) + "\n"

    new_content = (content[:heading_match.start()] + updated_block.rstrip('\n') + "\n\n"
                   + rest[end_match.start():].lstrip('\n'))
    return new_content, True


def update_feedback_md(feedback_path, applied, rejected, skipped_manual, skipped_noise, unmatched):
    today = date.today().isoformat()
    lines_new = []
    lines_new.append(f"## {today} — Automated batch import ({len(applied) + len(rejected) + len(unmatched)} tickers)\n")
    lines_new.append("\n**Source:** `tradingview_layouts.xlsx`, built + processed automatically by the tradingview-mcp export pipeline.\n")
    lines_new.append("**Method:** OCR axis read + colour-line channel detection (see `channel_detect.py`), applied via `update_master_sheet.py`.\n\n")

    if applied:
        plural = '' if len(applied) == 1 else 's'
        lines_new.append(f"### ✅ Applied — {len(applied)} ticker{plural} updated with fresh Alert Low/High\n\n")
        lines_new.append("| Ticker | Lower | Upper | Alert Low | Alert High |\n|---|---|---|---|---|\n")
        for a in applied:
            lines_new.append(f"| {a['ticker']} | {a['lower']} | {a['upper']} | {a['alert_low']} | {a['alert_high']} |\n")
        lines_new.append("\n")

    if skipped_noise:
        lines_new.append(f"{len(skipped_noise)} re-read within {int(REFRESH_NOISE_THRESHOLD*100)}% of already-current values — left untouched to avoid noise: "
                          + ", ".join(s['ticker'] for s in skipped_noise) + "\n\n")

    if rejected:
        lines_new.append(f"### ❌ Rejected — nothing written, for cause\n\n")
        lines_new.append("| Ticker | Reason |\n|---|---|\n")
        for r in rejected:
            lines_new.append(f"| {r['ticker']} | {r['reason']} |\n")
        lines_new.append("\n")

    if skipped_manual:
        lines_new.append("### ⚠️ Skipped — Alert Low Source is \"Manual\", not overwritten\n\n")
        lines_new.append(", ".join(s['ticker'] for s in skipped_manual) + "\n\n")

    if unmatched:
        lines_new.append("### ⚠️ Unmatched / not attempted\n\n")
        lines_new.append("| Ticker | Company | Reason |\n|---|---|---|\n")
        for u in unmatched:
            lines_new.append(f"| {u['ticker']} | {u['company']} | {u['reason']} |\n")
        lines_new.append("\n")

    lines_new.append("---\n\n")
    new_section = "".join(lines_new)

    try:
        with open(feedback_path, 'r', encoding='utf-8') as f:
            content = f.read()
    except FileNotFoundError:
        content = "# Chart Import Findings\n\n---\n\n## \U0001F4CB Coverage Tracker\n\n---\n\n"

    content, tracker_updated = _update_coverage_tracker(content, applied, rejected, unmatched, today)
    if not tracker_updated:
        print("WARNING: could not parse the Coverage Tracker table in the expected shape — "
              "left it untouched. Update it manually for this session.", file=sys.stderr)

    # Insert the new dated entry after the Coverage Tracker block (its own trailing
    # '---'), i.e. right before the previously-newest dated entry — never above the
    # tracker, which must stay pinned at the top of the file.
    inserted = False
    heading_match = _TRACKER_HEADING_RE.search(content)
    if heading_match:
        end_match = re.search(r'\n---\n', content[heading_match.start():])
        if end_match:
            idx = heading_match.start() + end_match.end()
            content = content[:idx] + "\n" + new_section + content[idx:]
            inserted = True
    if not inserted:
        marker = re.search(r'\n---\n', content)
        if marker:
            idx = marker.end()
            content = content[:idx] + "\n" + new_section + content[idx:]
            inserted = True

    if not inserted:
        content = content + "\n" + new_section

    with open(feedback_path, 'w', encoding='utf-8') as f:
        f.write(content)

    return feedback_path


def main():
    master_in, charts_path, channel_path, master_out, feedback_path = sys.argv[1:6]

    charts = load_json(charts_path)
    channel_results = load_json(channel_path)
    channel_by_ticker = {c['ticker']: c for c in channel_results if c.get('ticker')}

    wb = openpyxl.load_workbook(master_in, data_only=False)
    ws = wb[SHEET_NAME]

    applied, rejected, skipped_manual, skipped_noise, unmatched = process(ws, charts, channel_by_ticker)

    wb.save(master_out)
    update_feedback_md(feedback_path, applied, rejected, skipped_manual, skipped_noise, unmatched)

    # Machine-readable summary alongside the human-readable feedback log, so
    # verify_pipeline.py can report on this run without re-parsing markdown.
    import os
    result_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'master_update_result_tmp.json')
    with open(result_path, 'w', encoding='utf-8') as f:
        json.dump({
            'timestamp': date.today().isoformat(),
            'master_out': master_out,
            'applied': applied,
            'rejected': rejected,
            'skipped_manual': skipped_manual,
            'skipped_noise': skipped_noise,
            'unmatched': unmatched,
        }, f, indent=2)

    print(f"Applied: {len(applied)}, Rejected: {len(rejected)}, Manual-skipped: {len(skipped_manual)}, "
          f"Noise-skipped: {len(skipped_noise)}, Unmatched: {len(unmatched)}")
    print(f"Saved -> {master_out}")
    print(f"Feedback log updated -> {feedback_path}")
    print(f"Summary written -> {result_path}")


if __name__ == '__main__':
    main()
