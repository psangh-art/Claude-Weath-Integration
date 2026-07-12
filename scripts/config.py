#!/usr/bin/env python3
"""Single source of truth for machine-specific paths/IDs/ports (added
2026-07-12, external-review follow-up approved by the user): everything that
used to be hard-coded per-script now comes from scripts/config.json, so moving
machines means editing ONE file.

Usage:
    from config import CFG, downloads_dir, downloads_file, tv_layout_url
"""
import json
import os

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
