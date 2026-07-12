#!/usr/bin/env python3
"""Per-run history store (added 2026-07-12, external-review follow-up approved
by the user): every pipeline run appends its manifests to data/history.db
(SQLite) instead of the run simply overwriting yesterday's tmp JSONs. This is
the gateway to day-over-day comparison, indicator change detection, and every
"when did X first happen" question — the tmp manifests stay the pipeline's
working files; this is the archive.

Commands:
  python history_store.py record    append the current tmp manifests as a run
                                    (called automatically by run_full_pipeline.js)
  python history_store.py summary   list recorded runs
  python history_store.py diff      price/alert changes between the last two runs

The DB lives at <repo>/data/history.db (gitignored — per-machine data).
"""
import json
import os
import sqlite3
import sys
from datetime import datetime

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(os.path.dirname(SCRIPT_DIR), 'data', 'history.db')

SCHEMA = """
CREATE TABLE IF NOT EXISTS runs (
  run_id INTEGER PRIMARY KEY AUTOINCREMENT,
  recorded_at TEXT NOT NULL,
  charts INTEGER, alerts INTEGER, indicators INTEGER, channel_reads INTEGER
);
CREATE TABLE IF NOT EXISTS chart_snapshots (
  run_id INTEGER NOT NULL REFERENCES runs(run_id),
  ticker TEXT, chart_id TEXT, layout TEXT,
  price REAL, price_checked_at TEXT
);
CREATE TABLE IF NOT EXISTS alert_snapshots (
  run_id INTEGER NOT NULL REFERENCES runs(run_id),
  alert_id INTEGER, symbol TEXT, condition_type TEXT,
  target_price REAL, active INTEGER, expiration TEXT
);
CREATE TABLE IF NOT EXISTS indicator_snapshots (
  run_id INTEGER NOT NULL REFERENCES runs(run_id),
  ticker TEXT, layout TEXT, indicator TEXT, field TEXT, value TEXT
);
CREATE TABLE IF NOT EXISTS channel_snapshots (
  run_id INTEGER NOT NULL REFERENCES runs(run_id),
  ticker TEXT, kind TEXT, lower REAL, upper REAL, reason TEXT
);
CREATE TABLE IF NOT EXISTS master_updates (
  run_id INTEGER NOT NULL REFERENCES runs(run_id),
  ticker TEXT, action TEXT, alert_low REAL, alert_high REAL
);
CREATE INDEX IF NOT EXISTS idx_chart_ticker ON chart_snapshots(ticker, run_id);
CREATE INDEX IF NOT EXISTS idx_alert_symbol ON alert_snapshots(symbol, run_id);
CREATE INDEX IF NOT EXISTS idx_ind_ticker ON indicator_snapshots(ticker, run_id);
"""


def load_tmp(name):
    path = os.path.join(SCRIPT_DIR, name)
    if not os.path.exists(path):
        return None
    with open(path, encoding='utf-8') as f:
        return json.load(f)


def connect():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    con = sqlite3.connect(DB_PATH)
    con.executescript(SCHEMA)
    return con


def num(v):
    return v if isinstance(v, (int, float)) else None


def record():
    charts = load_tmp('layout_manifest_tmp.json') or []
    alerts = load_tmp('alerts_manifest_tmp.json') or []
    indicators = load_tmp('indicator_manifest_tmp.json') or []
    channels = load_tmp('channel_results_tmp.json') or []
    master = load_tmp('master_update_result_tmp.json') or {}

    if not charts and not alerts:
        print('No manifests found — nothing to record.')
        return

    con = connect()
    cur = con.cursor()
    cur.execute('INSERT INTO runs (recorded_at, charts, alerts, indicators, channel_reads) '
                'VALUES (?, ?, ?, ?, ?)',
                (datetime.now().isoformat(timespec='seconds'),
                 len(charts), len(alerts), len(indicators), len(channels)))
    run_id = cur.lastrowid

    cur.executemany(
        'INSERT INTO chart_snapshots VALUES (?, ?, ?, ?, ?, ?)',
        [(run_id, c.get('ticker'), c.get('chartId'), c.get('name'),
          num(c.get('price')), c.get('priceCheckedAt')) for c in charts])
    cur.executemany(
        'INSERT INTO alert_snapshots VALUES (?, ?, ?, ?, ?, ?, ?)',
        [(run_id, a.get('alertId'), a.get('symbol'), a.get('conditionType'),
          num(a.get('targetPrice')), 1 if a.get('active') else 0, a.get('expiration'))
         for a in alerts])
    cur.executemany(
        'INSERT INTO indicator_snapshots VALUES (?, ?, ?, ?, ?, ?)',
        [(run_id, i.get('ticker'), i.get('layoutName'), i.get('indicator'),
          i.get('field'), str(i.get('value'))) for i in indicators])
    cur.executemany(
        'INSERT INTO channel_snapshots VALUES (?, ?, ?, ?, ?, ?)',
        [(run_id, c.get('ticker'), c.get('kind'), num(c.get('lower')),
          num(c.get('upper')), c.get('reason')) for c in channels])

    rows = []
    for a in master.get('applied', []):
        rows.append((run_id, a.get('ticker'), 'applied', num(a.get('alert_low')), num(a.get('alert_high'))))
    for r in master.get('rejected', []):
        rows.append((run_id, r.get('ticker'), 'rejected', None, None))
    for b in master.get('below_alert_rows', []):
        rows.append((run_id, b.get('ticker'), 'below_alert', num(b.get('alert_low')), num(b.get('alert_high'))))
    cur.executemany('INSERT INTO master_updates VALUES (?, ?, ?, ?, ?)', rows)

    con.commit()
    print(f'Recorded run {run_id}: {len(charts)} charts, {len(alerts)} alerts, '
          f'{len(indicators)} indicator readings, {len(channels)} channel reads -> {DB_PATH}')


def summary():
    con = connect()
    rows = con.execute('SELECT run_id, recorded_at, charts, alerts, indicators, channel_reads '
                       'FROM runs ORDER BY run_id DESC LIMIT 20').fetchall()
    if not rows:
        print('No runs recorded yet.')
        return
    print(f'{"run":>4}  {"recorded_at":19}  {"charts":>6}  {"alerts":>6}  {"indic.":>6}  {"chans":>5}')
    for r in rows:
        print(f'{r[0]:>4}  {r[1]:19}  {r[2]:>6}  {r[3]:>6}  {r[4]:>6}  {r[5]:>5}')


def diff():
    con = connect()
    ids = [r[0] for r in con.execute('SELECT run_id FROM runs ORDER BY run_id DESC LIMIT 2')]
    if len(ids) < 2:
        print('Need at least two recorded runs to diff.')
        return
    new, old = ids
    print(f'Comparing run {old} -> run {new}\n')

    price_rows = con.execute("""
        SELECT n.ticker, o.price, n.price,
               ROUND((n.price - o.price) / o.price * 100, 2) AS pct
        FROM chart_snapshots n JOIN chart_snapshots o
          ON o.ticker = n.ticker AND o.run_id = ? AND n.run_id = ?
        WHERE o.price IS NOT NULL AND n.price IS NOT NULL AND o.price != 0
              AND n.price != o.price
        ORDER BY ABS(pct) DESC LIMIT 25""", (old, new)).fetchall()
    print('Price moves (top 25 by magnitude):' if price_rows else 'No price changes.')
    for t, po, pn, pct in price_rows:
        print(f'  {t:8} {po:>12,.2f} -> {pn:>12,.2f}  ({pct:+.2f}%)')

    for label, query in (
        ('New alerts', 'SELECT symbol, target_price FROM alert_snapshots WHERE run_id=? '
                       'AND alert_id NOT IN (SELECT alert_id FROM alert_snapshots WHERE run_id=?)'),
        ('Removed alerts', 'SELECT symbol, target_price FROM alert_snapshots WHERE run_id=? '
                           'AND alert_id NOT IN (SELECT alert_id FROM alert_snapshots WHERE run_id=?)'),
    ):
        a, b = (new, old) if label == 'New alerts' else (old, new)
        rows = con.execute(query, (a, b)).fetchall()
        if rows:
            print(f'\n{label} ({len(rows)}):')
            for s, tp in rows[:15]:
                print(f'  {s}  @ {tp}')


def main():
    cmd = sys.argv[1] if len(sys.argv) > 1 else 'record'
    {'record': record, 'summary': summary, 'diff': diff}.get(cmd, record)()


if __name__ == '__main__':
    main()
