#!/usr/bin/env python3
"""Stage 1 of the unified pipeline: report on the input files in ~/Downloads.

Rules (user policy 2026-07-13):
  - The bank/broker exports — Amex activity.csv, Barclays data.csv, Fidelity
    AccountSummary.csv and the Fidelity historic transaction export — are
    OPTIONAL. The pipeline runs without them (the spending-summary stage is
    skipped); what matters is knowing HOW OLD the data is. Each input reports
    an as-of date: the file's mtime when it's sitting in Downloads, or the
    recorded ingestion date (data/ingestion_state.json, written by
    consume_input_files.py) once a successful run has consumed it.
  - An input whose data is older than STALE_DAYS (6 weeks) — or that has never
    been seen at all — is flagged stale: time to download a fresh export.
  - Only the master workbook (Stocks_Buy_Strategy.xlsx) is REQUIRED: without
    it there is nothing to update, so the run halts.

CLI usage: python preflight_check.py [downloads_dir]
  Prints a JSON report to stdout. Exit 1 only if the master workbook is
  missing; 0 otherwise (even with every optional input absent).

Report shape (consumed by pipeline_app_server.js and the front end):
  ok             bool — master workbook present, run can proceed
  spending_ready bool — all three spending inputs + AccountSummary present
  found          {key: path|null} — back-compat map
  missing        [{key, expected, why}] — REQUIRED files only
  files          {key: {label, present, path, as_of, age_days, stale, note}}
"""
import sys
import os
import json
from datetime import datetime

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)
from fidelity_file_classifier import find_latest

STALE_DAYS = 42  # 6 weeks — beyond this the data needs a fresh export
INGESTION_STATE = os.path.join(os.path.dirname(SCRIPT_DIR), 'data', 'ingestion_state.json')


def load_ingestion_state():
    try:
        with open(INGESTION_STATE, encoding='utf-8') as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def file_entry(key, label, path, state, note=None):
    """Build one input's report entry: present/as_of/age/stale."""
    present = bool(path) and os.path.isfile(path)
    as_of = None
    if present:
        as_of = datetime.fromtimestamp(os.path.getmtime(path))
    elif key in state and state[key].get('as_of'):
        try:
            as_of = datetime.fromisoformat(state[key]['as_of'])
        except ValueError:
            as_of = None
    age_days = (datetime.now() - as_of).days if as_of else None
    return {
        'label': label,
        'present': present,
        'path': path if present else None,
        'as_of': as_of.isoformat(timespec='seconds') if as_of else None,
        'age_days': age_days,
        'stale': age_days is None or age_days > STALE_DAYS,
        'note': note or '',
    }


def main():
    from config import downloads_dir as _cfg_downloads
    downloads_dir = sys.argv[1] if len(sys.argv) > 1 else _cfg_downloads()
    state = load_ingestion_state()

    files = {}

    def optional(key, label, filename):
        path = os.path.join(downloads_dir, filename)
        files[key] = file_entry(key, label, path if os.path.isfile(path) else None, state)

    optional('amex', 'Amex spending export', 'activity.csv')
    optional('barclays', 'Barclays spending export', 'data.csv')
    optional('fidelity_account_summary', 'Fidelity AccountSummary', 'AccountSummary.csv')

    fidelity = find_latest(downloads_dir)
    files['fidelity_historic'] = file_entry(
        'fidelity_historic', 'Fidelity transaction history',
        fidelity['historic']['path'] if fidelity['historic'] else None, state)
    files['fidelity_pending'] = file_entry(
        'fidelity_pending', 'Fidelity pending transactions',
        fidelity['pending']['path'] if fidelity['pending'] else None, state,
        note='' if fidelity['pending'] else
        'Optional — holdings will use settled positions only.')
    files['fidelity_pending']['stale'] = False  # truly optional, never red-flagged

    # The one genuinely required file: no master workbook, nothing to update.
    master_path = os.path.join(downloads_dir, 'Stocks_Buy_Strategy.xlsx')
    master_ok = os.path.isfile(master_path)
    files['master_workbook'] = file_entry(
        'master_workbook', 'Master trading workbook',
        master_path if master_ok else None, state)
    files['master_workbook']['stale'] = False  # living workbook — age is normal

    missing = [] if master_ok else [{
        'key': 'master_workbook',
        'expected': master_path,
        'why': 'Master trading workbook — needed by the master-sheet-update stage',
    }]

    spending_keys = ('amex', 'barclays', 'fidelity_account_summary', 'fidelity_historic')
    report = {
        'ok': master_ok,
        # Build the spending summary on PARTIAL inputs (user request 2026-07-18):
        # ready if ANY spending source is present, not only when all four are.
        # spending_summary.py tolerates the missing ones.
        'spending_ready': any(files[k]['present'] for k in spending_keys),
        'stale_days': STALE_DAYS,
        'found': {k: v['path'] for k, v in files.items()},
        'missing': missing,
        'files': files,
    }
    print(json.dumps(report, indent=2))
    sys.exit(0 if report['ok'] else 1)


if __name__ == "__main__":
    main()
