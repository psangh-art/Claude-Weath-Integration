#!/usr/bin/env python3
"""Post-run consumption of the pipeline's input files (added 2026-07-12,
user-directed): once a run has completed successfully, the bank/broker exports
it consumed are REMOVED from ~/Downloads — including every other version of the
same file (Windows "name (1).csv" duplicates and copies previously renamed with
the "Delete " prefix by cleanup_downloads.py).

Files are sent to the Windows RECYCLE BIN (SHFileOperation with FOF_ALLOWUNDO),
not hard-deleted, so a misidentified file is always recoverable — this is the
deliberate safety floor under the "delete used inputs" policy.

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
import ctypes
import json
import os
import re
import sys
from datetime import datetime

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

FOF_SILENT = 0x0004
FOF_NOCONFIRMATION = 0x0010
FOF_ALLOWUNDO = 0x0040
FOF_NOERRORUI = 0x0400
FO_DELETE = 0x0003


class SHFILEOPSTRUCTW(ctypes.Structure):
    _fields_ = [
        ('hwnd', ctypes.c_void_p),
        ('wFunc', ctypes.c_uint),
        ('pFrom', ctypes.c_wchar_p),
        ('pTo', ctypes.c_wchar_p),
        ('fFlags', ctypes.c_ushort),
        ('fAnyOperationsAborted', ctypes.c_int),
        ('hNameMappings', ctypes.c_void_p),
        ('lpszProgressTitle', ctypes.c_wchar_p),
    ]


def recycle(paths):
    """Send `paths` to the Recycle Bin in one operation. Returns 0 on success."""
    src = '\0'.join(os.path.abspath(p) for p in paths) + '\0\0'
    op = SHFILEOPSTRUCTW()
    op.wFunc = FO_DELETE
    op.pFrom = src
    op.fFlags = FOF_ALLOWUNDO | FOF_NOCONFIRMATION | FOF_SILENT | FOF_NOERRORUI
    return ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op))


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


def main():
    args = [a for a in sys.argv[1:] if a != '--apply']
    apply = '--apply' in sys.argv
    from config import downloads_dir as _cfg_downloads
    downloads_dir = args[0] if args else _cfg_downloads()

    hits = find_consumables(downloads_dir)
    if not hits:
        print('No consumed input files found in Downloads — nothing to remove.')
        return

    for h in hits:
        print(f"  [{h['family']}] {h['name']}")
    if not apply:
        print(f'DRY RUN: {len(hits)} file(s) would be sent to the Recycle Bin. '
              'Re-run with --apply to remove them.')
        return

    record_ingestion(hits)  # before recycling, while mtimes still exist
    rc = recycle([h['path'] for h in hits])
    remaining = [h['name'] for h in hits if os.path.exists(h['path'])]
    if rc == 0 and not remaining:
        print(f'Removed {len(hits)} used input file(s) to the Recycle Bin.')
    else:
        print(f'WARNING: recycle returned {rc}; still present: {remaining or "none"}',
              file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()
