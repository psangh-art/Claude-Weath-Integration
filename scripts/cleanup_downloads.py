#!/usr/bin/env python3
"""Final pipeline stage: flag redundant files in ~/Downloads for manual deletion.

Never deletes anything itself — only renames a candidate file by prepending
"Delete " to its name, in place, in the same folder. This is reversible (a plain
rename) and leaves the actual decision to remove the file with the user. Run
standalone with: python scripts/cleanup_downloads.py [--apply]
Without --apply, only prints what it would rename (dry run).

Rules (conservative — only flags files that are clearly superseded by something
newer that already exists):

1. Backup families: any file matching "<base>.bak-*" is grouped by <base>. Within
   each group, the single most-recently-modified backup is kept untouched; every
   older backup in that group is flagged. (A backup is only ever "redundant" once
   a newer backup of the same base file exists — never flags the sole/newest one.)
2. Fidelity export files: reads scripts/fidelity_import_state.json (see CLAUDE.md /
   [[fidelity-import-state]]) and flags any TransactionHistory*.csv / transactions*.*
   file strictly OLDER (by mtime) than the recorded source_mtime for its type
   (historic/pending). The currently-tracked file itself, and anything newer
   (a new export not yet ingested), are never flagged.
3. Numbered duplicate workbook saves: files matching "<stem>_<N>.<ext>" where a
   canonical "<stem>.<ext>" file exists and is newer than the numbered copy are
   flagged as pre-pipeline manual saves.
"""
import sys
import os
import re
import json
from datetime import datetime

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

from config import downloads_dir
DOWNLOADS = downloads_dir()
STATE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fidelity_import_state.json")

BAK_RE = re.compile(r"^(?P<base>.+)\.bak-.+$")
NUMBERED_RE = re.compile(r"^(?P<stem>.+)_(?P<n>\d+)(?P<ext>\.[A-Za-z0-9]+)$")
FIDELITY_RE = re.compile(r"^(TransactionHistory|transactions).*\.csv$", re.IGNORECASE)


def find_backup_candidates(entries):
    groups = {}
    for name, path, mtime in entries:
        m = BAK_RE.match(name)
        if m:
            groups.setdefault(m.group("base"), []).append((name, path, mtime))

    candidates = []
    for base, files in groups.items():
        if len(files) < 2:
            continue  # never flag the sole backup of a base file
        files.sort(key=lambda f: f[2], reverse=True)
        for name, path, mtime in files[1:]:  # all but the newest
            candidates.append((path, f"superseded by a newer backup of {base}"))
    return candidates


def find_fidelity_candidates(entries):
    if not os.path.exists(STATE_PATH):
        return []
    with open(STATE_PATH, "r", encoding="utf-8") as f:
        state = json.load(f)

    tracked_mtimes = {}
    for kind in ("historic", "pending"):
        info = state.get(kind)
        if info and info.get("source_mtime"):
            tracked_mtimes[kind] = datetime.fromisoformat(info["source_mtime"])
    tracked_files = {state[k]["source_file"] for k in ("historic", "pending") if state.get(k)}

    if not tracked_mtimes:
        return []
    newest_tracked = max(tracked_mtimes.values())

    candidates = []
    for name, path, mtime in entries:
        if not FIDELITY_RE.match(name):
            continue
        if name in tracked_files:
            continue  # currently-ingested file — never flag
        file_mtime = datetime.fromtimestamp(mtime)
        if file_mtime < newest_tracked:
            candidates.append((path, "superseded — older than the currently-ingested Fidelity export"))
    return candidates


def find_numbered_duplicate_candidates(entries):
    by_name = {name: (path, mtime) for name, path, mtime in entries}
    candidates = []
    for name, path, mtime in entries:
        m = NUMBERED_RE.match(name)
        if not m:
            continue
        canonical = f"{m.group('stem')}{m.group('ext')}"
        if canonical in by_name and by_name[canonical][1] > mtime:
            candidates.append((path, f"pre-pipeline manual save, superseded by current {canonical}"))
    return candidates


def main():
    apply = "--apply" in sys.argv[1:]

    entries = []
    for name in os.listdir(DOWNLOADS):
        path = os.path.join(DOWNLOADS, name)
        if name.startswith("Delete "):
            continue
        try:
            if not os.path.isfile(path):
                continue
            mtime = os.path.getmtime(path)
        except OSError:
            # File vanished between listing and stat (a browser .crdownload or an
            # OneDrive sync temp being moved). Skip it — this cosmetic step must
            # never abort the run over a transient file. See CLAUDE.md.
            continue
        entries.append((name, path, mtime))

    candidates = (
        find_backup_candidates(entries)
        + find_fidelity_candidates(entries)
        + find_numbered_duplicate_candidates(entries)
    )
    # de-dupe (a file could theoretically match more than one rule)
    seen = set()
    unique = []
    for path, reason in candidates:
        if path not in seen:
            seen.add(path)
            unique.append((path, reason))

    if not unique:
        print("No redundant files found in Downloads.")
        return

    print(f"{'Renaming' if apply else 'Would rename'} {len(unique)} file(s):")
    for path, reason in unique:
        name = os.path.basename(path)
        new_name = f"Delete {name}"
        new_path = os.path.join(os.path.dirname(path), new_name)
        print(f"  {name}\n    -> {new_name}   ({reason})")
        if apply:
            if os.path.exists(new_path):
                print(f"    SKIPPED — target already exists: {new_name}")
                continue
            try:
                os.rename(path, new_path)
            except OSError as e:
                # File locked (open in Excel / mid OneDrive sync) or otherwise
                # un-renamable. Warn and carry on — flagging a redundant file is
                # cosmetic and must never fail the run. See CLAUDE.md.
                print(f"    SKIPPED — could not rename ({e.__class__.__name__}: {e})")
                continue

    if not apply:
        print("\nDry run only — re-run with --apply to actually rename these files.")


if __name__ == "__main__":
    main()
