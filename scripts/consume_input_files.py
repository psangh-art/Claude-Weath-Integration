#!/usr/bin/env python3
"""Post-run consumption of the pipeline's input files (added 2026-07-12; changed
2026-07-18 to KEEP the last copy instead of deleting): once a run has completed
successfully, the newest Amex/Barclays/Fidelity export of each type is MOVED to
`~/Downloads/old_pipeline/` under a canonical name, so exactly ONE (the latest)
version of each is retained there — Amex `activity.csv`, Barclays `data.csv`,
Fidelity `AccountSummary.csv`, the historic `TransactionHistory.csv`, and the
pending `TransactionHistory_pending.csv`. A prior copy already in old_pipeline is
replaced (only the last version is kept, including for the historic file). Every
OTHER version left in Downloads (the browser's "name (1).csv" duplicates and any
"Delete " copies) is sent to the Recycle Bin.

Nothing is hard-deleted: replaced/duplicate files go to the Windows RECYCLE BIN
(SHFileOperation, FOF_ALLOWUNDO), and the kept copies sit in old_pipeline — so a
misidentified file is always recoverable.

Consumed families (same matching as preflight_check.py / fidelity_file_classifier):
  - Amex:      activity.csv           (+ " (N)" duplicates, + "Delete " copies)
  - Barclays:  data.csv               (same variants)
  - Fidelity:  AccountSummary.csv     (same variants)
  - Fidelity:  TransactionHistory*.csv / transactions*.csv (ALL of them)

NEVER touched: Stocks_Buy_Strategy.xlsx (the living master workbook), its
backups, or anything else in Downloads.

Run this ONLY after a fully successful pipeline run — a failed run must keep its
inputs so it can be re-run. pipeline_app_server.js calls it with --apply after
all stages succeed; without --apply it just prints what it would remove.

Usage: python consume_input_files.py [downloads_dir] [--apply]
"""
import json
import os
import re
import sys
from datetime import datetime

# Recycle-bin delete lives in config.py — one implementation of the SHFileOperation
# call for the whole repo (this file used to carry a second, identical copy).
from config import recycle_to_bin

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

# "Delete " prefix = copies already flagged by cleanup_downloads.py; " (N)" = the
# browser's duplicate-download suffix. Both are "other versions of the file".
FAMILY_RES = [
    ('amex', re.compile(r'^(Delete )?activity( \(\d+\))?\.csv$', re.IGNORECASE)),
    ('barclays', re.compile(r'^(Delete )?data( \(\d+\))?\.csv$', re.IGNORECASE)),
    ('fidelity_summary', re.compile(r'^(Delete )?AccountSummary( \(\d+\))?\.csv$', re.IGNORECASE)),
    ('fidelity_transactions', re.compile(r'^(Delete )?(TransactionHistory|transactions).*\.csv$', re.IGNORECASE)),
]

# Once a file is recycled its mtime is gone, so the data's as-of date is
# recorded here first — preflight_check.py reads it back to show how old each
# feed's data is (and to flag it red past the 6-week mark). Keys match
# preflight_check.py's report keys, not the family names above.
INGESTION_STATE = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data', 'ingestion_state.json')
FAMILY_TO_STATE_KEY = {
    'amex': 'amex',
    'barclays': 'barclays',
    'fidelity_summary': 'fidelity_account_summary',
    'fidelity_transactions': 'fidelity_historic',
}


def record_ingestion(hits):
    """Record each consumed family's newest file mtime as its data as-of date."""
    state = {}
    try:
        with open(INGESTION_STATE, encoding='utf-8') as f:
            state = json.load(f)
    except (OSError, ValueError):
        pass
    for h in hits:
        key = FAMILY_TO_STATE_KEY.get(h['family'])
        if not key or 'mtime' not in h:
            continue
        as_of = datetime.fromtimestamp(h['mtime']).isoformat(timespec='seconds')
        prev = state.get(key, {}).get('as_of')
        if not prev or as_of > prev:
            state[key] = {'as_of': as_of,
                          'recorded_at': datetime.now().isoformat(timespec='seconds')}
    os.makedirs(os.path.dirname(INGESTION_STATE), exist_ok=True)
    with open(INGESTION_STATE, 'w', encoding='utf-8') as f:
        json.dump(state, f, indent=2)

def find_consumables(downloads_dir):
    hits = []
    for name in sorted(os.listdir(downloads_dir)):
        path = os.path.join(downloads_dir, name)
        if not os.path.isfile(path):
            continue
        for family, rx in FAMILY_RES:
            if rx.match(name):
                hits.append({'family': family, 'name': name, 'path': path,
                             'mtime': os.path.getmtime(path)})
                break
    return hits


# Where the LAST version of each consumed input is kept (user request 2026-07-18):
# a subfolder of Downloads, one file per type, so the most recent Amex/Fidelity
# export is always retrievable without cluttering Downloads or accumulating history.
OLD_DIR_NAME = 'old_pipeline'
# Canonical destination name per kept file, so old_pipeline holds exactly ONE of each.
CANON_NAME = {
    'amex': 'activity.csv',
    'barclays': 'data.csv',
    'fidelity_summary': 'AccountSummary.csv',
    'fidelity_historic': 'TransactionHistory.csv',
    'fidelity_pending': 'TransactionHistory_pending.csv',
}


def newest_per_family(downloads_dir):
    """Newest consumed file per family; transactions split historic vs pending via
    the content classifier so each is kept separately (only ONE version each)."""
    by = {}
    for h in find_consumables(downloads_dir):
        by.setdefault(h['family'], []).append(h)
    keep = {}
    for fam in ('amex', 'barclays', 'fidelity_summary'):
        if by.get(fam):
            keep[fam] = max(by[fam], key=lambda h: h['mtime'])
    try:
        from fidelity_file_classifier import find_latest
        fid = find_latest(downloads_dir)
        for kind, fam in (('historic', 'fidelity_historic'), ('pending', 'fidelity_pending')):
            info = fid.get(kind)
            if info and info.get('path') and os.path.exists(info['path']):
                keep[fam] = {'family': fam, 'name': os.path.basename(info['path']),
                             'path': info['path'], 'mtime': os.path.getmtime(info['path'])}
    except Exception as e:
        print(f'  (classifier unavailable, transactions grouped as-is: {e})')
    return by, keep


def main():
    args = [a for a in sys.argv[1:] if a != '--apply']
    apply = '--apply' in sys.argv
    from config import downloads_dir as _cfg_downloads
    downloads_dir = args[0] if args else _cfg_downloads()

    by, keep = newest_per_family(downloads_dir)
    all_hits = [h for hs in by.values() for h in hs]
    if not all_hits:
        print('No consumed input files found in Downloads — nothing to do.')
        return

    kept_paths = {os.path.abspath(h['path']) for h in keep.values()}
    to_recycle = [h for h in all_hits if os.path.abspath(h['path']) not in kept_paths]

    old_dir = os.path.join(downloads_dir, OLD_DIR_NAME)
    for fam, h in keep.items():
        print(f"  KEEP  [{fam}] {h['name']} -> {OLD_DIR_NAME}/{CANON_NAME[fam]}")
    for h in to_recycle:
        print(f"  RECYCLE [{h['family']}] {h['name']}")
    if not apply:
        print(f'DRY RUN: {len(keep)} file(s) would move to {OLD_DIR_NAME}/, '
              f'{len(to_recycle)} duplicate(s) to the Recycle Bin. Re-run with --apply.')
        return

    # Record data-as-of dates before anything moves (mtimes still on the originals).
    record_ingestion([h for h in keep.values() if 'mtime' in h])

    os.makedirs(old_dir, exist_ok=True)
    moved = 0
    for fam, h in keep.items():
        dest = os.path.join(old_dir, CANON_NAME[fam])
        try:
            if os.path.exists(dest):        # keep only the last version — recycle the old one
                recycle_to_bin([dest])
            os.replace(h['path'], dest)     # same-drive move; overwrites if recycle missed it
            moved += 1
        except OSError as e:
            print(f'  SKIPPED move of {h["name"]}: {e}', file=sys.stderr)

    if to_recycle:
        recycle_to_bin([h['path'] for h in to_recycle])
    remaining = [h['name'] for h in to_recycle if os.path.exists(h['path'])]
    if not remaining:
        print(f'Kept {moved} latest input(s) in {OLD_DIR_NAME}/; '
              f'recycled {len(to_recycle)} duplicate(s).')
    else:
        print(f'WARNING: recycle left files behind; still present: {remaining}',
              file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()
