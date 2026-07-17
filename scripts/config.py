#!/usr/bin/env python3
"""Single source of truth for machine-specific paths/IDs/ports (added
2026-07-12, external-review follow-up approved by the user): everything that
used to be hard-coded per-script now comes from scripts/config.json, so moving
machines means editing ONE file.

Usage:
    from config import CFG, downloads_dir, downloads_file, tv_layout_url
"""
import ctypes
import json
import os
import re

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

with open(os.path.join(_SCRIPT_DIR, 'config.json'), encoding='utf-8') as _f:
    CFG = json.load(_f)


def downloads_dir():
    return os.path.expanduser(CFG['downloadsDir'])


def downloads_file(key):
    """Absolute path of a Downloads artifact named by its config key,
    e.g. downloads_file('masterWorkbook')."""
    return os.path.join(downloads_dir(), CFG[key])


def tv_layout_url(chart_id):
    return CFG['tvLayoutUrlTemplate'].format(chart_id=chart_id)


# --- Recycle-bin delete (Windows SHFileOperation, FOF_ALLOWUNDO) -------------
# Repo policy (see consume_input_files.py): outputs/inputs are NEVER hard-deleted
# — they go to the Recycle Bin so the removal is always reversible.
_FO_DELETE = 0x0003
_FOF_SILENT = 0x0004
_FOF_NOCONFIRMATION = 0x0010
_FOF_ALLOWUNDO = 0x0040
_FOF_NOERRORUI = 0x0400


class _SHFILEOPSTRUCTW(ctypes.Structure):
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


def recycle_to_bin(paths):
    """Send existing `paths` to the Windows Recycle Bin in one operation.
    Returns the list actually recycled ([] if none existed / on failure)."""
    existing = [os.path.abspath(p) for p in paths if os.path.isfile(p)]
    if not existing:
        return []
    src = '\0'.join(existing) + '\0\0'
    op = _SHFILEOPSTRUCTW()
    op.wFunc = _FO_DELETE
    op.pFrom = src
    op.fFlags = _FOF_ALLOWUNDO | _FOF_NOCONFIRMATION | _FOF_SILENT | _FOF_NOERRORUI
    rc = ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op))
    return existing if rc == 0 else []


def purge_old_versions(target_path):
    """Before (re)building an output file, clear the browser/Office "` (N)`"
    duplicate copies that pile up next to it in the same directory
    (e.g. `Investment_Review_Deck (1).pptx`). The canonical `target_path`
    itself is left for the caller's own overwrite, so a build that later fails
    never removes the last good copy. Recycles rather than hard-deletes.
    Returns the list recycled."""
    d = os.path.dirname(os.path.abspath(target_path))
    stem, ext = os.path.splitext(os.path.basename(target_path))
    rx = re.compile(r'^' + re.escape(stem) + r' \(\d+\)' + re.escape(ext) + r'$', re.IGNORECASE)
    try:
        names = os.listdir(d)
    except OSError:
        return []
    dupes = [os.path.join(d, n) for n in names if rx.match(n)]
    return recycle_to_bin(dupes)
