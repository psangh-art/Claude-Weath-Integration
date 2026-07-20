"""Recycle the Downloads files the pipeline has FLAGGED as redundant.

`cleanup_downloads.py` runs at the end of every pipeline run and marks files it
believes are redundant — superseded workbook backups, ingested Fidelity exports,
numbered duplicate saves — by RENAMING them with a `Delete ` prefix. It never
removes anything: the rename is reversible and the actual deletion is deliberately
left as a human decision.

This is that decision, made explicit. It is the only thing in the repo that
removes a file the user can see, so its blast radius is kept deliberately tiny:

  * It touches ONLY files whose name begins exactly with `Delete ` — the prefix
    the pipeline itself wrote. Nothing is matched by age, size or extension, so a
    file can never be caught by accident. To have something removed, rename it
    with that prefix; to spare one, take the prefix off.
  * It RECYCLES (config.recycle_to_bin, FOF_ALLOWUNDO). Nothing is hard-deleted,
    so a mistake is a trip to the Recycle Bin, not a restore from backup.
  * It lists everything and asks before acting. `--yes` skips the prompt for
    scripted use; `--dry-run` only ever lists.

Usage:
    python scripts/purge_flagged_downloads.py             # list, confirm, recycle
    python scripts/purge_flagged_downloads.py --dry-run
    python scripts/purge_flagged_downloads.py --yes
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from config import downloads_dir, recycle_to_bin  # noqa: E402

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

PREFIX = "Delete "


def flagged_files(folder):
    try:
        names = os.listdir(folder)
    except OSError as e:
        print(f"Could not read {folder}: {e}")
        return []
    out = []
    for n in sorted(names):
        if not n.startswith(PREFIX):
            continue
        p = os.path.join(folder, n)
        try:
            if os.path.isfile(p):
                out.append((p, os.path.getsize(p)))
        except OSError:
            # Vanished between listdir and stat (a sync temp) — skip it, exactly
            # as cleanup_downloads.py does. A transient must never fail the run.
            continue
    return out


def human(n):
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:,.0f} {unit}" if unit == "B" else f"{n/1:,.1f} {unit}"
        n /= 1024
    return f"{n:.1f} GB"


def main():
    ap = argparse.ArgumentParser(description="Recycle Downloads files flagged 'Delete '.")
    ap.add_argument("--dry-run", action="store_true", help="list only")
    ap.add_argument("--yes", "-y", action="store_true", help="skip the confirmation")
    args = ap.parse_args()

    folder = downloads_dir()
    print(f"Downloads: {folder}\n")

    files = flagged_files(folder)
    if not files:
        print("Nothing flagged 'Delete ' — nothing to do.")
        return 0

    total = sum(sz for _, sz in files)
    for p, sz in files:
        print(f"  {human(sz):>10}  {os.path.basename(p)}")
    print(f"\n{len(files)} file(s), {human(total)} total.")

    if args.dry_run:
        print("\nDry run — nothing was recycled.")
        return 0

    if not args.yes:
        print("\nThese go to the RECYCLE BIN (recoverable), not permanent deletion.")
        try:
            reply = input("Recycle them? [y/N] ").strip().lower()
        except EOFError:
            print("No console to confirm on — nothing recycled. Use --yes to skip the prompt.")
            return 1
        if reply not in ("y", "yes"):
            print("Cancelled — nothing was recycled.")
            return 0

    recycled = recycle_to_bin([p for p, _ in files])
    if not recycled:
        print("\nNothing was recycled — the files may be locked by another program "
              "(Excel or PowerPoint holding a workbook open is the usual cause).")
        return 1
    print(f"\nRecycled {len(recycled)} file(s), {human(total)} freed.")
    missed = len(files) - len(recycled)
    if missed:
        print(f"{missed} file(s) could not be recycled — likely locked or already gone.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
