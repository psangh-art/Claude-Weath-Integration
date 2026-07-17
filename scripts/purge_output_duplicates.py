#!/usr/bin/env python3
"""Ensure only ONE version of each pipeline output file exists after a run
(user request 2026-07-17).

The browser / Office / OneDrive-sync path drops "` (N)`" duplicate copies next
to the canonical output files in Downloads (e.g. `Stocks_Buy_Strategy (1).xlsx`,
`Investment_Review_Deck (3).pptx`). This step recycles every such duplicate for
the known output files, leaving exactly the canonical one.

Unlike cleanup_downloads.py (which is deliberately rename-only), this actually
removes the duplicates — but only to the Recycle Bin (config.recycle_to_bin,
FOF_ALLOWUNDO), so it is always reversible and never touches the canonical file.

Run as the final pipeline step; safe to run standalone and idempotent (a second
run finds nothing to recycle).
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from config import CFG, downloads_dir, purge_old_versions

# The config keys whose value is a canonical output filename in Downloads.
OUTPUT_KEYS = [
    'masterWorkbook', 'layoutsWorkbook', 'feedbackMd',
    'reviewDeckPptx', 'architecturePptx', 'alertRulesPptx',
    'spendingSummaryXlsx',
]


def main():
    d = downloads_dir()
    total = 0
    for key in OUTPUT_KEYS:
        name = CFG.get(key)
        if not name:
            continue
        recycled = purge_old_versions(os.path.join(d, name))
        if recycled:
            total += len(recycled)
            print('  %s -> recycled %d duplicate(s): %s'
                  % (name, len(recycled), ', '.join(os.path.basename(p) for p in recycled)))
    if total:
        print('Output duplicates: %d "(N)" copy(ies) sent to the Recycle Bin; one version of each output remains.' % total)
    else:
        print('Output duplicates: none found — one version of each output already.')


if __name__ == '__main__':
    main()
