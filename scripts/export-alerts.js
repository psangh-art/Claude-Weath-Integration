#!/usr/bin/env node
// Export every alert on the account (price crossings, channel/shape alerts) to a
// CSV file — this is the same data behind both the white dotted line drawn on a
// chart at the alert price and the Alerts panel in TradingView's UI. Unlike the
// layout/indicator exports, this doesn't need to switch layouts or focus panes —
// the pricealerts API returns every alert on the account in one call.
import os from 'os';
import path from 'path';
import { writeFileSync, mkdirSync, appendFileSync } from 'fs';
import { fileURLToPath } from 'url';
import * as health from '../src/core/health.js';
import * as alerts from '../src/core/alerts.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const OUT_PATH = path.join(os.homedir(), 'Downloads', 'tradingview_alerts.csv');

const LOGS_DIR = path.join(__dirname, '..', 'logs');
mkdirSync(LOGS_DIR, { recursive: true });
const RUN_STAMP = new Date().toISOString().replace(/[:.]/g, '-');
const LOG_PATH = path.join(LOGS_DIR, `alerts_${RUN_STAMP}.log`);
const LOG_LATEST_PATH = path.join(LOGS_DIR, 'latest-alerts.log');
const SUMMARY_PATH = path.join(LOGS_DIR, `alerts_${RUN_STAMP}.summary.json`);
const SUMMARY_LATEST_PATH = path.join(LOGS_DIR, 'latest-alerts-summary.json');

function log(...args) {
  const line = args.map(a => (a instanceof Error ? a.stack : String(a))).join(' ');
  console.log(line);
  appendFileSync(LOG_PATH, line + '\n');
  appendFileSync(LOG_LATEST_PATH, line + '\n');
}

function csvEscape(value) {
  const s = String(value ?? '');
  return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}

// A simple price-crossing alert's target is condition.series[1].value (series[0]
// is always the price barset being watched). Channel/shape-based alerts don't have
// a single target price, so this is left blank for those rather than guessed at.
function targetPrice(condition) {
  const valueEntry = (condition?.series || []).find(s => s.type === 'value');
  return valueEntry ? valueEntry.value : null;
}

const startedAt = new Date();

async function main() {
  log('Checking TradingView CDP connection on port 9222...');
  try {
    await health.healthCheck();
  } catch {
    log('\nCould not connect to TradingView.');
    log('Start it with CDP enabled first, e.g. run scripts\\launch_tv_debug.bat, then re-run this.');
    writeSummary({ ok: false, reason: 'CDP connection failed', rows: [] });
    process.exitCode = 1;
    return;
  }

  log('Fetching alerts...');
  const result = await alerts.list();
  if (result.error) {
    log(`\nFailed to fetch alerts: ${result.error}`);
    writeSummary({ ok: false, reason: result.error, rows: [] });
    process.exitCode = 1;
    return;
  }
  log(`Found ${result.alert_count} alerts (${result.alerts.filter(a => a.active).length} active).\n`);

  const rows = result.alerts.map(a => ({
    alertId: a.alert_id,
    symbol: a.symbol,
    message: a.message,
    conditionType: a.condition?.type || null,
    targetPrice: targetPrice(a.condition),
    resolution: a.resolution,
    active: a.active,
    created: a.created,
    lastFired: a.last_fired,
    expiration: a.expiration,
  }));

  const header = ['Alert ID', 'Symbol', 'Message', 'Condition', 'Target Price', 'Resolution', 'Active', 'Created', 'Last Fired', 'Expiration'];
  const lines = [header.join(',')];
  for (const r of rows) {
    lines.push([r.alertId, r.symbol, r.message, r.conditionType, r.targetPrice, r.resolution, r.active, r.created, r.lastFired, r.expiration].map(csvEscape).join(','));
  }
  writeFileSync(OUT_PATH, lines.join('\n') + '\n');

  log(`Done. ${rows.length} alerts saved to ${OUT_PATH}`);
  writeSummary({ ok: true, reason: null, rows });
}

function writeSummary({ ok, reason, rows }) {
  const finishedAt = new Date();
  const summary = {
    timestamp: startedAt.toISOString(),
    durationSeconds: Math.round((finishedAt - startedAt) / 1000),
    ok,
    reason,
    outputPath: OUT_PATH,
    logPath: LOG_PATH,
    totalAlerts: rows.length,
    activeAlerts: rows.filter(r => r.active).length,
    rows,
  };
  writeFileSync(SUMMARY_PATH, JSON.stringify(summary, null, 2));
  writeFileSync(SUMMARY_LATEST_PATH, JSON.stringify(summary, null, 2));
  log(`Summary written to ${SUMMARY_PATH}`);
}

main()
  .catch(err => {
    log('\nExport failed:', err);
    writeSummary({ ok: false, reason: err.message, rows: [] });
    process.exitCode = 1;
  })
  .finally(() => {
    // src/connection.js keeps an open CDP WebSocket that would otherwise hold
    // the event loop (and this process) open forever — force exit once main()
    // has genuinely finished.
    process.exit(process.exitCode || 0);
  });
