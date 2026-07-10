#!/usr/bin/env python3
"""Classify a Fidelity export CSV as 'historic' or 'pending' by CONTENT, not
filename — Fidelity reuses filenames across export types (e.g. both a historic and
a pending export can be called "TransactionHistory (4).csv"), so filename alone is
not a reliable signal. See CLAUDE.md / fidelity_import_state.json for the original
rule this codifies:
  - pending exports: every data row has Completion date == "Pending", Status is
    Priced/Ordered, and there's an extra "Expiry Date" column.
  - historic exports: real completion dates, Status is Completed/Cancelled, no
    "Expiry Date" column.

CLI usage:
  python fidelity_file_classifier.py <downloads_dir>
    Prints JSON: {"historic": {"path", "mtime"} | null, "pending": {...} | null}
    — the newest file of each type found in downloads_dir (any file matching
    TransactionHistory*.csv or transactions*.* is considered a candidate).
"""
import sys
import os
import csv
import io
import json
import re
from datetime import datetime

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

CANDIDATE_RE = re.compile(r"^(TransactionHistory|transactions).*\.csv$", re.IGNORECASE)


def classify(path):
    """Returns 'historic', 'pending', or None (not a recognizable Fidelity export)."""
    try:
        with open(path, encoding="utf-8-sig") as f:
            content = f.read()
    except (OSError, UnicodeDecodeError):
        return None

    lines = content.replace("\r", "").split("\n")
    header_idx = next((i for i, l in enumerate(lines) if l.startswith("Run Date")
                        or l.startswith("Order date") or l.startswith("Trade Date")), None)
    if header_idx is None:
        return None

    reader = csv.DictReader(io.StringIO("\n".join(lines[header_idx:])))
    fieldnames = reader.fieldnames or []
    has_expiry = any("Expiry Date" in f for f in fieldnames)

    rows = list(reader)
    data_rows = [r for r in rows if any((v or "").strip() for v in r.values())]
    if not data_rows:
        return None

    completion_field = next((f for f in fieldnames if "Completion date" in f), None)
    if completion_field and has_expiry:
        all_pending = all((r.get(completion_field, "") or "").strip().lower() == "pending" for r in data_rows)
        if all_pending:
            return "pending"
    return "historic"


def find_latest(downloads_dir):
    best = {"historic": None, "pending": None}
    for name in os.listdir(downloads_dir):
        if not CANDIDATE_RE.match(name) or name.startswith("Delete "):
            continue
        path = os.path.join(downloads_dir, name)
        if not os.path.isfile(path):
            continue
        kind = classify(path)
        if kind is None:
            continue
        mtime = os.path.getmtime(path)
        current = best[kind]
        if current is None or mtime > current["mtime"]:
            best[kind] = {"path": path, "mtime": mtime}
    return best


def main():
    downloads_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.expanduser("~/Downloads")
    result = find_latest(downloads_dir)
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
