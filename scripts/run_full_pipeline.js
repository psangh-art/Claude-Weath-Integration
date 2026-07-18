#!/usr/bin/env node
// Full periodic pipeline: capture every TradingView layout -> build
// tradingview_layouts.xlsx (Charts/Indicators/Alerts, via export-layouts-excel.js) ->
// OCR channel-boundary detection per ticker -> apply Alert Low/High into
// Stocks_Buy_Strategy.xlsx -> update the Feedback_for_Claude_Code.md coverage
// tracker. Run this (instead of export-layouts-excel.js directly) for the whole
// thing end to end.
//
// The OCR step needs the Tesseract OCR binary installed (one-time, interactive —
// see channel_detect.py's docstring). If it's missing, this still completes the
// chart export/screenshot half and tells you exactly what's missing, rather than
// failing the whole run.
import path from 'path';
import { writeFileSync, readFileSync, existsSync } from 'fs';
import { fileURLToPath } from 'url';
import { spawnSync } from 'child_process';

import { downloadsFile } from './config.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const MASTER_SHEET_PATH = downloadsFile('masterWorkbook');
const FEEDBACK_PATH = downloadsFile('feedbackMd');
const CHARTS_MANIFEST = path.join(__dirname, 'layout_manifest_tmp.json');
const CHANNEL_INPUT = path.join(__dirname, 'channel_input_tmp.json');
const CHANNEL_RESULTS = path.join(__dirname, 'channel_results_tmp.json');

function run(cmd, args) {
  return spawnSync(cmd, args, { stdio: ['ignore', 'pipe', 'inherit'], encoding: 'utf-8' });
}

const CDP_PORT = 9222;

function cdpUp() {
  const r = spawnSync('curl', ['-s', '-m', '3', `http://localhost:${CDP_PORT}/json/version`], { encoding: 'utf-8' });
  return r.status === 0 && /Protocol-Version/.test(r.stdout || '');
}

// Guarantee TradingView is reachable over CDP before capture. If the debug port
// is already up we do NOTHING — an already-debug-ready TradingView is never
// force-killed, so unsaved chart edits are safe. Only when the port is down do we
// relaunch via launch_tv_debug.bat (which taskkills + reopens TV with the port).
// TradingView layouts live server-side, so a relaunch loads them fresh and does
// not modify the user's charts. This makes "start TV with the debug port" a
// non-issue on every run (user request 2026-07-14).
function ensureCdp() {
  if (cdpUp()) {
    console.log(`TradingView CDP already up on ${CDP_PORT}.`);
    return true;
  }
  console.log(`TradingView CDP not responding on ${CDP_PORT} — relaunching TradingView with the debug port...`);
  console.log('(Layouts are saved server-side; this reloads them and does not modify your charts.)');
  const res = spawnSync('cmd.exe', ['/c', path.join(__dirname, 'launch_tv_debug.bat'), String(CDP_PORT)],
    { stdio: 'inherit', timeout: 90000 });
  if (!cdpUp()) {
    console.error(`Could not bring up TradingView CDP on ${CDP_PORT}` +
      (res.error ? ` (${res.error.message})` : '') +
      '. Start it with scripts\\launch_tv_debug.bat and re-run.');
    return false;
  }
  return true;
}

function runDeck() {
  // Build the per-investment PowerPoint review deck (charts + live price +
  // master-sheet holdings/alerts + OCR read + TradingView alerts, one slide per
  // chart, flagging missing charts/alerts). Reads whatever manifests exist, so
  // it produces a useful "what's missing" deck even after a partial run. See
  // build_review_deck.py. Never blocks the run — a missing dep just logs here.
  console.log('\n=== Step 4/6: building PowerPoint review deck ===\n');
  const deckResult = spawnSync('python', [path.join(__dirname, 'build_review_deck.py')], { stdio: 'inherit' });
  if (deckResult.status !== 0) {
    console.error('Review-deck build could not run (see python output above) — nothing else in this run is affected.');
  }
}

function runVerify() {
  console.log('\n=== Step 5/6: verifying this run ===\n');
  const verifyResult = spawnSync('python', [path.join(__dirname, 'verify_pipeline.py'), '--live-alert-check'], { stdio: 'inherit' });
  if (verifyResult.status !== 0) {
    console.error('Verification report could not run (see python output above) — the export/update steps above may still be fine.');
  }
}

function runCleanup() {
  // Never deletes anything — only renames clearly-superseded files in Downloads
  // with a "Delete " prefix so they're easy to spot and remove by hand. See
  // cleanup_downloads.py's docstring for the exact rules. Runs last, after
  // verification, so a file is never flagged before this run has actually
  // confirmed producing its replacement.
  console.log('\n=== Step 6/6: flagging redundant Downloads files ===\n');
  const cleanupResult = spawnSync('python', [path.join(__dirname, 'cleanup_downloads.py'), '--apply'], { stdio: 'inherit' });
  if (cleanupResult.status !== 0) {
    console.error('Downloads cleanup could not run (see python output above) — nothing else in this run is affected.');
  }

  // Ensure only ONE version of each output file remains (user request 2026-07-17):
  // recycle the "X (N).ext" duplicate copies of every canonical output. Distinct
  // from cleanup_downloads above (rename-only) — this one removes (to Recycle Bin).
  const purgeResult = spawnSync('python', [path.join(__dirname, 'purge_output_duplicates.py')], { stdio: 'inherit' });
  if (purgeResult.status !== 0) {
    console.error('Output-duplicate purge could not run (see python output above) — nothing else in this run is affected.');
  }
}

function main() {
  if (!ensureCdp()) {
    process.exitCode = 1;
    return;
  }
  console.log('=== Step 1/6: capturing charts from TradingView ===\n');
  const exportResult = spawnSync('node', [path.join(__dirname, 'export-layouts-excel.js')], { stdio: 'inherit' });
  if (exportResult.status !== 0) {
    console.error('\nChart export failed — stopping before the Google Finance / master-sheet steps.');
    process.exitCode = 1;
    return;
  }

  if (!existsSync(MASTER_SHEET_PATH)) {
    console.warn(`\nNo master sheet found at ${MASTER_SHEET_PATH} — skipping the Alert Low/High update step.`);
    console.warn('Charts export is done; drop Stocks_Buy_Strategy.xlsx in Downloads and re-run to also update it.');
    return;
  }

  console.log('\n=== Step 2/6: OCR channel-boundary detection ===\n');
  const charts = JSON.parse(readFileSync(CHARTS_MANIFEST, 'utf-8'));
  const seen = new Set();
  const channelInput = [];
  for (const row of charts) {
    if (!row.ticker || !row.screenshot || seen.has(row.ticker)) continue;
    seen.add(row.ticker);
    channelInput.push({ ticker: row.ticker, screenshot: row.screenshot, known_price: row.price ?? null });
  }
  writeFileSync(CHANNEL_INPUT, JSON.stringify(channelInput, null, 2));
  console.log(`Running channel detection on ${channelInput.length} distinct tickers...`);

  const detectResult = run('python', [path.join(__dirname, 'channel_detect.py'), '--batch', CHANNEL_INPUT]);
  if (detectResult.status !== 0) {
    console.error('\nChannel detection could not run (see message below) — skipping the master-sheet update step.');
    console.error(detectResult.stderr || '(no error detail returned)');
    console.error('\nChart export is still done and up to date; re-run this once Tesseract is installed to also update Stocks_Buy_Strategy.xlsx.');
    return;
  }
  writeFileSync(CHANNEL_RESULTS, detectResult.stdout);

  console.log('\n=== Step 3/6: applying results into Stocks_Buy_Strategy.xlsx ===\n');
  const updateResult = spawnSync('python', [
    path.join(__dirname, 'update_master_sheet.py'),
    MASTER_SHEET_PATH,
    CHARTS_MANIFEST,
    CHANNEL_RESULTS,
    MASTER_SHEET_PATH,
    FEEDBACK_PATH,
  ], { stdio: 'inherit' });

  if (updateResult.status !== 0) {
    console.error('\nMaster-sheet update failed (see python output above).');
    process.exitCode = 1;
    return;
  }

  // Sub-step of step 3 (not a numbered "=== Step N/M ===" marker — the Production
  // Centre parses those): re-section and refresh the 'Stocks of Interest' section
  // tables. They were hand-maintained and drifted badly — audited 2026-07-15, 17 of
  // 25 rows sat in the wrong section against levels last touched on 07-06, with BEZ
  // and AUTO still listed 'at lower boundary' while nine stocks actually at their buy
  // point sat below. Runs AFTER update_master_sheet.py, since it reads the Alert Low /
  // Alert High that step has just written.
  console.log('\nRe-sectioning the Stocks of Interest tables...');
  const soiResult = spawnSync('python', [
    path.join(__dirname, 'refresh_soi_sections.py'), MASTER_SHEET_PATH, '--apply',
  ], { stdio: 'inherit' });
  if (soiResult.status !== 0) {
    console.error('Stocks of Interest re-sectioning failed (the Investments tab itself is already updated and intact).');
    process.exitCode = 1;
    return;
  }

  // Sub-step of step 3 (deliberately NOT a numbered "=== Step N/M ===" marker —
  // the Production Centre parses those): mirror spending_summary.xlsx's tabs
  // into the master workbook so the Finance Google Sheet import carries them.
  // A missing spending_summary.xlsx is a skip inside the script, not a failure.
  console.log('\nIntegrating spending-summary tabs into Stocks_Buy_Strategy.xlsx...');
  const integrateResult = spawnSync('python', [
    path.join(__dirname, 'integrate_spending_tabs.py'),
  ], { stdio: 'inherit' });
  if (integrateResult.status !== 0) {
    console.error('Spending-tab integration failed (master workbook itself is already updated and intact).');
    process.exitCode = 1;
    return;
  }

  console.log('\nDone. tradingview_layouts.xlsx, Stocks_Buy_Strategy.xlsx, and Feedback_for_Claude_Code.md are all up to date in Downloads.');
}

function recordHistory() {
  // Quiet step (deliberately NOT a numbered "=== Step N/M ===" stage — the
  // Production Centre parses those markers, and this needs no UI slot):
  // append this run's manifests to data/history.db so day-over-day
  // comparison and indicator change detection have something to work on.
  const result = spawnSync('python', [path.join(__dirname, 'history_store.py'), 'record'], { encoding: 'utf-8' });
  const line = (result.stdout || result.stderr || '').trim().split('\n').pop();
  if (line) console.log(`History: ${line}`);
}

function refreshDashboard() {
  // Wire the Investment Dashboard into the pipeline as one flow (user request
  // 2026-07-18): regenerate its data from the just-updated workbook + history.db,
  // then notify any open dashboard so it reflects this run live. Quiet sub-step
  // (no numbered "=== Step N/M ===" marker). Reads history.db, so runs AFTER
  // recordHistory(); spending_summary already feeds the workbook earlier in the run.
  const gen = spawnSync('python', [path.join(__dirname, 'dashboard_data.py')], { encoding: 'utf-8' });
  if (gen.status !== 0) {
    console.error('Dashboard: data refresh failed —', (gen.stderr || '').slice(-200));
    return;
  }
  const summary = (gen.stdout || '').trim().split('\n').find(l => l.includes('Portfolio value'));
  console.log(`Dashboard: data refreshed${summary ? ' (' + summary.trim() + ')' : ''}`);
  // Best-effort: tell a running dashboard (:4600) to reload and SSE-notify open
  // tabs. Silent no-op if the dashboard isn't running — the data is already on disk.
  spawnSync('curl', ['-s', '-m', '2', 'http://localhost:4600/pipeline-updated'], { encoding: 'utf-8' });
}

try {
  main();
} finally {
  // Always produce the review deck and verification report, even if an earlier
  // step stopped partway through — both report honestly on whatever manifests do
  // or don't exist rather than requiring a fully clean run (the deck's whole job
  // is to show what's missing).
  recordHistory();
  refreshDashboard();
  runDeck();
  runVerify();
  runCleanup();
}
