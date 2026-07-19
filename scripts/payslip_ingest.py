#!/usr/bin/env python3
"""Payslip PDF -> a row on the workbook's 'Payslip Summary' tab.

Owned by the Investment Dashboard's Payslips screen, NOT the pipeline (user decision
2026-07-19: "we don't need to include payslip uploads in the pipeline — we'll load it
via this new payslips screen"). The pipeline never calls this; the dashboard's
POST /api/payslips/upload does.

Two stages, deliberately separate so a payslip whose layout we don't recognise fails
loudly instead of writing a wrong number into a live financial sheet:

  extract_fields(pdf)  — pull the figures out of the PDF text
  append_row(fields)   — write them onto the tab, in the right tax-year band

Usage:
    python payslip_ingest.py <payslip.pdf> [--dry-run]
    python payslip_ingest.py <payslip.pdf> --dump      # print the raw text and stop
"""
import os
import re
import sys
import json
import datetime
import shutil

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side

WORKBOOK = os.path.join(os.path.expanduser('~'), 'Downloads', 'Stocks_Buy_Strategy.xlsx')
SHEET = 'Payslip Summary'
HEADER_ROW = 3
REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ARCHIVE = os.path.join(REPO, 'data', 'payslips')

# Column -> the labels a payslip might use for it, tried in order. Kept as data so a
# new payroll layout is a one-line addition rather than a code change.
FIELD_LABELS = {
    'gross':       ['gross pay', 'total gross pay', 'gross salary', 'total payments'],
    'pension_ee':  ['pension ee', 'employee pension', 'pension (employee)',
                    'salary sacrifice', 'pension sacrifice'],
    'pension_er':  ['pension er', 'employer pension', 'pension (employer)',
                    'employers pension', "employer's pension"],
    'bonus':       ['bonus'],
    'taxable_pay': ['taxable pay', 'taxable gross'],
    'paye':        ['paye', 'paye tax', 'income tax', 'tax paid'],
    'ni':          ['national insurance', 'ni contribution', 'nic', 'employee ni'],
    'net':         ['net pay', 'total net pay', 'payment amount'],
}
MONEY = re.compile(r'-?[\d,]+\.\d{2}')
DATE_PATTERNS = [
    (re.compile(r'\b(\d{2})[/-](\d{2})[/-](\d{4})\b'), '%d/%m/%Y'),
    (re.compile(r'\b(\d{4})-(\d{2})-(\d{2})\b'), '%Y-%m-%d'),
]


class PayslipParseError(Exception):
    """The PDF didn't yield the fields we need — never guess, ask the human."""


def pdf_text(path):
    try:
        import pdfplumber
    except ImportError as e:
        raise PayslipParseError(
            'pdfplumber is not installed — run: python -m pip install pdfplumber') from e
    out = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            out.append(page.extract_text() or '')
    text = '\n'.join(out)
    if not text.strip():
        raise PayslipParseError(
            'No text in this PDF — it looks like a scan/image. A text payslip is needed '
            '(or the file has to go through OCR first).')
    return text


def _money_after(line, label):
    """The first money-looking number on the line, after the label."""
    idx = line.lower().find(label)
    if idx < 0:
        return None
    m = MONEY.search(line, idx + len(label))
    return float(m.group(0).replace(',', '')) if m else None


def extract_fields(path):
    """Best-effort field extraction. Returns (fields, text, missing)."""
    text = pdf_text(path)
    lines = [l.strip() for l in text.splitlines() if l.strip()]

    fields = {}
    for key, labels in FIELD_LABELS.items():
        for label in labels:
            for line in lines:
                if label in line.lower():
                    val = _money_after(line, label)
                    if val is not None:
                        fields[key] = val
                        break
            if key in fields:
                break

    # Pay date: the latest date on the page that isn't in the future
    today = datetime.date.today()
    dates = []
    for pat, fmt in DATE_PATTERNS:
        for m in pat.finditer(text):
            try:
                d = datetime.datetime.strptime(m.group(0), fmt).date()
            except ValueError:
                continue
            if datetime.date(2000, 1, 1) <= d <= today:
                dates.append(d)
    if dates:
        fields['pay_date'] = max(dates).isoformat()

    if fields.get('pension_ee') and fields.get('pension_er'):
        fields['pension_total'] = round(fields['pension_ee'] + fields['pension_er'], 2)
    if fields.get('paye') is not None and fields.get('ni') is not None:
        fields['deductions'] = round(fields['paye'] + fields['ni']
                                     + (fields.get('pension_ee') or 0), 2)
    if fields.get('pay_date'):
        fields['tax_year'] = tax_year_of(fields['pay_date'])

    missing = [k for k in ('pay_date', 'gross', 'net') if not fields.get(k)]
    return fields, text, missing


def tax_year_of(iso_date):
    """UK tax year containing a date — 6 April to 5 April."""
    d = datetime.date.fromisoformat(iso_date)
    start = d.year if (d.month, d.day) >= (4, 6) else d.year - 1
    return f'{start}/{str(start + 1)[2:]}'


COLS = ['pay_date', 'tax_year', 'gross', 'pension_ee', 'pension_er', 'pension_total',
        'bonus', 'taxable_pay', 'paye', 'ni', 'deductions', 'net']


def _row_style(ws, row, template_row):
    """Copy the styling of an existing payslip row — a new row must match the format
    of the ones beside it, never openpyxl's defaults (CLAUDE.md rule, 2026-07-16)."""
    for col in range(1, len(COLS) + 1):
        src = ws.cell(row=template_row, column=col)
        dst = ws.cell(row=row, column=col)
        dst.font = Font(name=src.font.name, size=src.font.size, bold=src.font.bold,
                        color=src.font.color)
        dst.alignment = Alignment(horizontal=src.alignment.horizontal,
                                  vertical=src.alignment.vertical)
        dst.number_format = src.number_format
        if src.fill and src.fill.fgColor and src.fill.fgColor.rgb:
            dst.fill = src.fill.copy()
        dst.border = Border(bottom=Side(style='thin', color='D9E1F2'))


def append_row(fields, workbook=WORKBOOK, dry_run=False):
    """Insert the payslip into its tax-year band, keeping pay dates in order."""
    if not os.path.exists(workbook):
        raise PayslipParseError(f'Workbook not found: {workbook}')
    wb = openpyxl.load_workbook(workbook)
    if SHEET not in wb.sheetnames:
        raise PayslipParseError(f"'{SHEET}' tab not found in {os.path.basename(workbook)}")
    ws = wb[SHEET]

    pay_date = datetime.date.fromisoformat(fields['pay_date'])
    tax_year = fields.get('tax_year') or tax_year_of(fields['pay_date'])

    # Walk the sheet: find this tax year's band and where the date belongs in it.
    band_start = band_end = None
    in_band = False
    last_data_row = None
    insert_at = None
    for r in range(HEADER_ROW + 1, ws.max_row + 2):
        label = str(ws.cell(row=r, column=1).value or '').strip()
        if label.lower().startswith('tax year'):
            if in_band:
                break
            in_band = tax_year in label
            if in_band:
                band_start = r
            continue
        if not in_band:
            continue
        if label.upper().startswith('TOTAL'):
            band_end = r
            break
        existing = ws.cell(row=r, column=1).value
        d = existing.date() if isinstance(existing, datetime.datetime) else None
        if d is None:
            try:
                d = datetime.datetime.strptime(str(existing).strip(), '%d/%m/%Y').date()
            except (ValueError, TypeError):
                continue
        last_data_row = r
        if d == pay_date:
            raise PayslipParseError(
                f'A payslip dated {pay_date:%d/%m/%Y} is already on the sheet (row {r}).')
        if d > pay_date and insert_at is None:
            insert_at = r

    if band_start is None:
        raise PayslipParseError(
            f"No 'Tax Year {tax_year}' band on the {SHEET} tab — add the band (and its "
            f"TOTAL row) by hand once, then re-upload.")
    target = insert_at or band_end or (last_data_row + 1 if last_data_row else band_start + 1)
    template = last_data_row or band_start + 1

    if dry_run:
        wb.close()
        return {'row': target, 'tax_year': tax_year, 'written': False}

    ws.insert_rows(target)
    ws.cell(row=target, column=1, value=pay_date.strftime('%d/%m/%Y'))
    ws.cell(row=target, column=2, value=tax_year)
    for i, key in enumerate(COLS[2:], start=3):
        val = fields.get(key)
        if val is not None:
            ws.cell(row=target, column=i, value=val)
    _row_style(ws, target, template + (1 if template >= target else 0))

    backup = workbook + '.bak-' + datetime.datetime.now().strftime('%Y%m%d-%H%M%S')
    shutil.copyfile(workbook, backup)
    wb.save(workbook)
    wb.close()
    return {'row': target, 'tax_year': tax_year, 'written': True, 'backup': backup}


def archive_pdf(path, fields):
    """Keep the source PDF — the sheet row is a summary, the payslip is the record."""
    os.makedirs(ARCHIVE, exist_ok=True)
    stamp = fields.get('pay_date') or datetime.date.today().isoformat()
    dst = os.path.join(ARCHIVE, f'payslip_{stamp}.pdf')
    shutil.copyfile(path, dst)
    return dst


def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    flags = {a for a in sys.argv[1:] if a.startswith('--')}
    if not args:
        print(__doc__)
        return 2
    path = args[0]

    try:
        if '--dump' in flags:
            print(pdf_text(path))
            return 0
        fields, text, missing = extract_fields(path)
        if missing:
            raise PayslipParseError(
                'Could not read ' + ', '.join(missing) + ' from this payslip. '
                'Run with --dump to see the text, then add the right labels to '
                'FIELD_LABELS in payslip_ingest.py.')
        result = append_row(fields, dry_run='--dry-run' in flags)
        if not result['written']:
            result['archived'] = None
        else:
            result['archived'] = archive_pdf(path, fields)
        print(json.dumps({'ok': True, 'fields': fields, **result}, indent=2))
        return 0
    except PayslipParseError as e:
        print(json.dumps({'ok': False, 'error': str(e)}, indent=2), file=sys.stderr)
        return 1


if __name__ == '__main__':
    sys.exit(main())
