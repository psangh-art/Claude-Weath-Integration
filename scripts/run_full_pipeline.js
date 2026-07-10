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
import os from 'os';
import path from 'path';
import { writeFileSync, readFileSync, existsSync } from 'fs';
import { fileURLToPath } from 'url';
import { spawnSync } from 'child_process';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DOWNLOADS = path.join(os.homedir(), 'Downloads');
const MASTER_SHEET_PATH = path.join(DOWNLOADS, 'Stocks_Buy_Strategy.xlsx');
const FEEDBACK_PATH = path.join(DOWNLOADS, 'Feedback_for_Claude_Code.md');
const CHARTS_MANIFEST = path.join(__dirname, 'layout_manifest_tmp.json');
const CHANNEL_INPUT = path.join(__dirname, 'channel_input_tmp.json');
const CHANNEL_RESULTS = path.join(__dirname, 'channel_results_tmp.json');

function run(cmd, args) {
  return spawnSync(cmd, args, { stdio: ['ignore', 'pipe', 'inherit'], encoding: 'utf-8' });
}

function runVerify() {
  console.log('\n=== Step 4/4: verifying this run ===\n');
  const verifyResult = spawnSync('python', [path.join(__dirname, 'verify_pipeline.py'), '--live-alert-check'], { stdio: 'inherit' });
  if (verifyResult.status !== 0) {
    console.error('Verification report could not run (see python output above) — the export/update steps above may still be fine.');
  }
}

function main() {
  console.log('=== Step 1/4: capturing charts from TradingView ===\n');
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

  console.log('\n=== Step 2/4: OCR channel-boundary detection ===\n');
  const charts = JSON.parse(readFileSync(CHARTS_MANIFEST, 'utf-8'));
  const seen = new Set();
  const channelInput = [];
  for (const row of charts) {
    if (!row.ticker || !row.screenshot || seen.has(row.ticker)) continue;
    seen.add(row.ticker);
    channelInput.push({ ticker: row.ticker, screenshot: row.screenshot });
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

  console.log('\n=== Step 3/4: applying results into Stocks_Buy_Strategy.xlsx ===\n');
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

  console.log('\nDone. tradingview_layouts.xlsx, Stocks_Buy_Strategy.xlsx, and Feedback_for_Claude_Code.md are all up to date in Downloads.');
}

try {
  main();
} finally {
  // Always produce a verification report, even if an earlier step stopped
  // partway through — verify_pipeline.py reports honestly on whatever manifests
  // do or don't exist rather than requiring a fully clean run.
  runVerify();
}
