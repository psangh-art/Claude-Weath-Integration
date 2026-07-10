#!/usr/bin/env python3
"""Stage 1 of the unified pipeline: verify every REQUIRED input file is present in
~/Downloads before any stage that depends on it runs. This exists specifically so a
missing data file (e.g. no fresh Fidelity export) halts the run with a clear "supply
this file and re-run" message, instead of a later stage failing confusingly deep into
a multi-minute run — see CLAUDE.md for the "silence over guessing / an unresolved row
is fine, a wrong number is not" philosophy this follows.

Required:
  - Amex export       (activity.csv)
  - Barclays export    (data.csv)
  - Fidelity AccountSummary (AccountSummary.csv)
  - Fidelity historic transaction export (classified by content, not filename —
    see fidelity_file_classifier.py)
  - Master workbook    (Stocks_Buy_Strategy.xlsx) — needed by the master-sheet-update
    stage later in the run

Optional (noted, never blocks):
  - Fidelity pending transaction export — spending_summary.py already treats this as
    optional and simply skips the pending-holdings adjustment if absent.

CLI usage: python preflight_check.py [downloads_dir]
  Prints a JSON report to stdout and exits 1 if anything REQUIRED is missing, 0 if
  the run can proceed.
"""
import sys
import os
import json

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fidelity_file_classifier import find_latest


def main():
    downloads_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.expanduser("~/Downloads")

    required = []
    found = {}
    missing = []

    def check_simple(key, filename, why):
        path = os.path.join(downloads_dir, filename)
        exists = os.path.isfile(path)
        found[key] = path if exists else None
        if not exists:
            missing.append({"key": key, "expected": path, "why": why})
        return exists

    check_simple("amex", "activity.csv", "Amex spending export — needed to build spending_summary.xlsx")
    check_simple("barclays", "data.csv", "Barclays spending export — needed to build spending_summary.xlsx")
    check_simple("fidelity_account_summary", "AccountSummary.csv",
                  "Fidelity AccountSummary export — needed for accurate current holdings")
    check_simple("master_workbook", "Stocks_Buy_Strategy.xlsx",
                  "Master trading workbook — needed by the master-sheet-update stage")

    fidelity = find_latest(downloads_dir)
    found["fidelity_historic"] = fidelity["historic"]["path"] if fidelity["historic"] else None
    found["fidelity_pending"] = fidelity["pending"]["path"] if fidelity["pending"] else None
    if fidelity["historic"] is None:
        missing.append({
            "key": "fidelity_historic",
            "expected": "a TransactionHistory*.csv / transactions*.* export with real completion dates",
            "why": "Fidelity historic transaction export — needed to build spending_summary.xlsx",
        })
    if fidelity["pending"] is None:
        found["fidelity_pending_note"] = "No pending export found — optional, holdings will use settled positions only."

    report = {
        "ok": len(missing) == 0,
        "found": found,
        "missing": missing,
    }
    print(json.dumps(report, indent=2))
    sys.exit(0 if report["ok"] else 1)


if __name__ == "__main__":
    main()
