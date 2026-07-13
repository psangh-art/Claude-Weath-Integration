#!/usr/bin/env node
// Export the CURRENT numeric value of every indicator that exposes one via the
// Data Window (RSI, MACD, moving averages, custom plot()-based indicators, etc.),
// for every chart in every saved layout, into a CSV file. This is a snapshot, not
// a history — one reading per indicator, taken at run time.
//
// Marker/label-style indicators (e.g. "Dividend yield %") don't populate the Data
// Window and are simply absent from the output — that's expected, not a failure.
//
// "Reset chart view" is NOT used here (unlike export-layouts-excel.js) since this
// only reads numeric values, not pixel geometry, so a stale zoom/pan doesn't affect
// correctness.
import path from 'path';
import { writeFileSync, mkdirSync, appendFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { downloadsDir } from './config.js';
import * as health from '../src/core/health.js';
import * as ui from '../src/core/ui.js';
import * as pane from '../src/core/pane.js';
import * as data from '../src/core/data.js';
import { evaluateAsync } from '../src/connection.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const OUT_PATH = path.join(downloadsDir(), 'tradingview_indicator_values.csv');

const LOGS_DIR = path.join(__dirname, '..', 'logs');
mkdirSync(LOGS_DIR, { recursive: true });
const RUN_STAMP = new Date().toISOString().replace(/[:.]/g, '-');
const LOG_PATH = path.join(LOGS_DIR, `indicator_values_${RUN_STAMP}.log`);
const LOG_LATEST_PATH = path.join(LOGS_DIR, 'latest-indicator-values.log');
const SUMMARY_PATH = path.join(LOGS_DIR, `indicator_values_${RUN_STAMP}.summary.json`);
const SUMMARY_LATEST_PATH = path.join(LOGS_DIR, 'latest-indicator-values-summary.json');

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

  // Same hard guard as export-layouts-excel.js: auto-save must be OFF before any
  // layout switching/clicking, or interactions could silently persist back into
  // the user's saved layouts (see CLAUDE.md 2026-07-13). Throws if it can't be
  // disabled.
  log('Checking TradingView layout auto-save is off...');
  const autosave = await ui.ensureAutosaveDisabled();
  if (!autosave.found) {
    log(`  WARNING: could not locate the auto-save setting (${autosave.error}) — verify it is off manually (save-menu dropdown). Proceeding: this run makes no view changes.`);
  } else if (autosave.wasEnabled) {
    log('  auto-save WAS ENABLED — disabled it for this and future sessions.');
  } else {
    log('  auto-save already off.');
  }

  log('Fetching saved layouts...');
  const chartListJson = await evaluateAsync(`
    JSON.stringify((window.TradingViewApi._loadChartService._state.value().chartList || []).map(c => ({id: c.id, url: c.url, name: c.name})))
  `);
  const allLayouts = JSON.parse(chartListJson || '[]');
  const PLACEHOLDER_NAMES = /^test$/i;
  const layouts = allLayouts.filter(l => !PLACEHOLDER_NAMES.test((l.name || '').trim()));
  if (layouts.length === 0) {
    log('No saved layouts found.');
    writeSummary({ ok: false, reason: 'No saved layouts found', rows: [] });
    process.exitCode = 1;
    return;
  }
  log(`Found ${layouts.length} layouts.\n`);

  const rows = [];

  for (let i = 0; i < layouts.length; i++) {
    const layout = layouts[i];
    log(`[${i + 1}/${layouts.length}] ${layout.name} — switching...`);

    try {
      await ui.layoutSwitch({ name: layout.name });
    } catch (err) {
      log(`  FAILED to switch: ${err.message}`);
      continue;
    }

    const paneData = await pane.listWithRects();
    const panes = paneData.panes || [];

    let chartsWithValues = 0;
    for (let pi = 0; pi < panes.length; pi++) {
      const p = panes[pi];
      try {
        await pane.focus({ index: p.index });
      } catch (err) {
        log(`  pane ${pi}: FAILED to focus: ${err.message}`);
        continue;
      }
      const result = await data.getStudyValues();
      if (result.study_count === 0) continue;
      chartsWithValues++;
      for (const study of result.studies) {
        for (const [key, value] of Object.entries(study.values)) {
          rows.push({
            layoutId: layout.id,
            chartId: layout.url,
            layoutName: layout.name,
            ticker: p.ticker || p.symbol || null,
            company: p.description || null,
            indicator: study.name,
            field: key,
            value,
          });
        }
      }
    }
    log(`  ${chartsWithValues}/${panes.length} chart(s) had reportable indicator values`);
  }

  const header = ['Layout ID', 'Chart ID', 'Layout Name', 'Symbol', 'Company', 'Indicator', 'Field', 'Value'];
  const lines = [header.join(',')];
  for (const r of rows) {
    lines.push([r.layoutId, r.chartId, r.layoutName, r.ticker, r.company, r.indicator, r.field, r.value].map(csvEscape).join(','));
  }
  writeFileSync(OUT_PATH, lines.join('\n') + '\n');

  log(`\nDone. ${rows.length} indicator readings across ${layouts.length} layouts saved to ${OUT_PATH}`);
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
    totalReadings: rows.length,
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
