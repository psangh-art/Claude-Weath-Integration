#!/usr/bin/env node
// Export one row per individual chart (not per layout) — Layout ID + Chart ID +
// Layout Name + Symbol + Company + a cropped high-res screenshot of just that one
// chart — into an Excel workbook. Double-click scripts/export_layouts.bat to run
// this from the desktop.
//
// Per-layout capture strategy: take ONE full-window screenshot at 2x device scale
// (sharp enough to read axis labels even in a 6-pane grid — plain 1x captures
// compress each pane's price axis past reliable legibility), then crop out each
// pane's own region using its DOM bounding rect so every chart gets its own image
// instead of being squeezed into one grid screenshot. Cropping is delegated to
// crop_panes.py (PIL) since Node has no image lib here — same delegation pattern
// as the openpyxl workbook build below.
//
// Workbook assembly is delegated to build_layout_excel.py (openpyxl) — that embedding
// method is what's actually been verified to produce images Excel/openpyxl can both
// read back correctly. exceljs's own image anchoring (oneCellAnchor) was tried and
// silently produced pictures openpyxl couldn't parse at all (0 images read back) with
// no way to confirm real Excel handled it either, so don't switch back to it without
// re-verifying against real Excel first.
import os from 'os';
import path from 'path';
import { writeFileSync, mkdirSync, appendFileSync, existsSync, readdirSync } from 'fs';
import { fileURLToPath } from 'url';
import { spawnSync } from 'child_process';
import * as health from '../src/core/health.js';
import * as ui from '../src/core/ui.js';
import * as pane from '../src/core/pane.js';
import * as capture from '../src/core/capture.js';
import * as data from '../src/core/data.js';
import * as alerts from '../src/core/alerts.js';
import { evaluateAsync } from '../src/connection.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const OUT_PATH = path.join(os.homedir(), 'Downloads', 'tradingview_layouts.xlsx');
const MANIFEST_PATH = path.join(__dirname, 'layout_manifest_tmp.json');
const CROP_MANIFEST_PATH = path.join(__dirname, 'crop_manifest_tmp.json');
const INDICATOR_MANIFEST_PATH = path.join(__dirname, 'indicator_manifest_tmp.json');
const ALERTS_MANIFEST_PATH = path.join(__dirname, 'alerts_manifest_tmp.json');
const CAPTURE_SCALE = 2;
const MAX_RETRIES = 3;
const RETRY_BACKOFF_MS = 4000;

async function sleep(ms) {
  return new Promise(r => setTimeout(r, ms));
}

// Reuse a previous run's screenshots for this layout tag if every retry has failed —
// better than leaving an unattended periodic run with a permanently blank row. Crop
// filenames are deterministic per pane index, so an exact prefix match is enough.
function fallbackToStaleScreenshots(layout, tag) {
  const screenshotsDir = path.join(__dirname, '..', 'screenshots');
  const dirEntries = existsSync(screenshotsDir) ? readdirSync(screenshotsDir) : [];
  const charts = [];
  for (let pi = 1; pi <= 20; pi++) {
    const match = dirEntries.find(f => f.startsWith(`layout_${tag}_pane_${String(pi).padStart(2, '0')}_`));
    if (!match) break;
    charts.push({
      ticker: null,
      description: null,
      screenshot: path.join(screenshotsDir, match),
      error: 'stale screenshot reused after repeated capture failures',
    });
  }
  return charts;
}

// Every run is logged to its own timestamped file plus a fixed "latest" pointer
// (both the raw console output and a structured summary), so a run's results can
// be read back afterward without needing to have watched the terminal live.
const LOGS_DIR = path.join(__dirname, '..', 'logs');
mkdirSync(LOGS_DIR, { recursive: true });
const RUN_STAMP = new Date().toISOString().replace(/[:.]/g, '-');
const LOG_PATH = path.join(LOGS_DIR, `export_${RUN_STAMP}.log`);
const LOG_LATEST_PATH = path.join(LOGS_DIR, 'latest.log');
const SUMMARY_PATH = path.join(LOGS_DIR, `export_${RUN_STAMP}.summary.json`);
const SUMMARY_LATEST_PATH = path.join(LOGS_DIR, 'latest-summary.json');

function log(...args) {
  const line = args.map(a => (a instanceof Error ? a.stack : String(a))).join(' ');
  console.log(line);
  appendFileSync(LOG_PATH, line + '\n');
  appendFileSync(LOG_LATEST_PATH, line + '\n');
}

const startedAt = new Date();

async function main() {
  log('Checking TradingView CDP connection on port 9222...');
  try {
    await health.healthCheck();
  } catch {
    log('\nCould not connect to TradingView.');
    log('Start it with CDP enabled first, e.g. run scripts\\launch_tv_debug.bat, then re-run this.');
    writeSummary({ startedAt, ok: false, reason: 'CDP connection failed', layoutsSummary: [] });
    process.exitCode = 1;
    return;
  }

  log('Fetching saved layouts...');
  const chartListJson = await evaluateAsync(`
    JSON.stringify((window.TradingViewApi._loadChartService._state.value().chartList || []).map(c => ({id: c.id, url: c.url, name: c.name})))
  `);
  const allLayouts = JSON.parse(chartListJson || '[]');
  // Skip throwaway/placeholder layouts (e.g. a scratch "Test" layout) so they don't
  // end up as real rows in the workbook.
  const PLACEHOLDER_NAMES = /^test$/i;
  const layouts = allLayouts.filter(l => !PLACEHOLDER_NAMES.test((l.name || '').trim()));
  const skipped = allLayouts.length - layouts.length;
  if (skipped > 0) {
    log(`Skipping ${skipped} placeholder layout(s) (e.g. "Test").`);
  }
  if (layouts.length === 0) {
    log('No saved layouts found.');
    writeSummary({ startedAt, ok: false, reason: 'No saved layouts found', layoutsSummary: [] });
    process.exitCode = 1;
    return;
  }
  log(`Found ${layouts.length} layouts.\n`);

  const manifest = [];
  const layoutsSummary = [];
  const indicatorRows = [];

  // One attempt at capturing a single layout end to end. Throws on any transient
  // failure (switch, crop) so the caller can retry; the "no measurable panes" case
  // is a valid terminal outcome (full-layout screenshot), not a failure.
  async function captureLayoutOnce(layout, tag) {
    await ui.layoutSwitch({ name: layout.name });

    const paneData = await pane.waitForPanesToLoad();
    const panes = (paneData.panes || []).filter(p => p.rect && p.rect.width > 0 && p.rect.height > 0);
    const indicators = [];

    if (panes.length === 0) {
      log(`  no measurable panes — falling back to full-layout screenshot`);
      const shot = await capture.captureScreenshot({ region: 'full', filename: `layout_${tag}_raw`, scale: CAPTURE_SCALE });
      return { status: 'full_layout_fallback', charts: [{ ticker: null, description: null, screenshot: shot.file_path, error: null }], indicators };
    }

    // Reset zoom/pan to a consistent fitted view before capturing — a chart can be
    // left scrolled/zoomed from whenever it was last interacted with. This is a
    // view-only, unsaved change (see resetView()'s comment) so it never touches the
    // saved layout itself. "Reset chart view" (Alt+R) only affects the currently
    // FOCUSED pane, not the whole grid — confirmed by two independent runs producing
    // byte-identical broken output for the same non-focused panes regardless of wait
    // time, ruling out a load-timing race. So each pane must be focused individually
    // before resetting it.
    const lastPriceByPaneIndex = {};
    for (let pi = 0; pi < panes.length; pi++) {
      const p = panes[pi];
      await pane.focus({ index: p.index });
      await ui.resetView();
      await ui.waitForResetToSettle();

      // Piggyback current indicator values onto this same per-pane focus pass
      // (rather than a separate full layout-switching loop) — the pane is
      // already focused here for the reset-view fix, so reading Data Window
      // values now is free.
      const studyResult = await data.getStudyValues();
      for (const study of studyResult.studies || []) {
        for (const [field, value] of Object.entries(study.values)) {
          indicators.push({
            layoutId: layout.id,
            chartId: layout.url,
            layoutName: layout.name,
            ticker: p.ticker || p.symbol || null,
            company: p.description || null,
            indicator: study.name,
            field,
            value,
          });
        }
      }

      // Also read the pane's own last bar (close price) while it's focused —
      // this is the live "current price" used to decide single-trendline
      // direction (above/below price -> Alert High/Low) and to stamp how
      // fresh each chart's read is. It does NOT touch Stocks_Buy_Strategy.xlsx's
      // own GOOGLEFINANCE-formula price column (see CLAUDE.md) — this is a
      // separate, pipeline-owned value.
      try {
        lastPriceByPaneIndex[p.index] = await data.getLastPrice();
      } catch {
        lastPriceByPaneIndex[p.index] = null;
      }
    }

    const rawFilename = `layout_${tag}_raw`;
    const shot = await capture.captureScreenshot({ region: 'full', filename: rawFilename, scale: CAPTURE_SCALE });

    const crops = panes.map((p, pi) => ({
      x: p.rect.x * CAPTURE_SCALE,
      y: p.rect.y * CAPTURE_SCALE,
      width: p.rect.width * CAPTURE_SCALE,
      height: p.rect.height * CAPTURE_SCALE,
      filename: `layout_${tag}_pane_${String(pi + 1).padStart(2, '0')}_${(p.ticker || p.symbol || 'chart').replace(/[^a-zA-Z0-9]+/g, '_')}.png`,
    }));

    writeFileSync(CROP_MANIFEST_PATH, JSON.stringify(crops, null, 2));
    const cropDir = path.dirname(shot.file_path);
    const cropResult = spawnSync('python', [path.join(__dirname, 'crop_panes.py'), shot.file_path, CROP_MANIFEST_PATH, cropDir], { stdio: 'inherit' });
    if (cropResult.status !== 0) throw new Error('pane crop failed');

    const priceCheckedAt = new Date().toISOString();
    const layoutCharts = panes.map((p, pi) => {
      const lastPrice = lastPriceByPaneIndex[p.index] || null;
      return {
        ticker: p.ticker || p.symbol || null,
        description: p.description || null,
        screenshot: path.join(cropDir, crops[pi].filename),
        error: null,
        price: lastPrice ? lastPrice.close : null,
        priceBarTime: lastPrice ? lastPrice.time : null,
        priceCheckedAt,
      };
    });
    log(`  captured ${panes.length} chart(s)`);
    return { status: 'ok', charts: layoutCharts, indicators };
  }

  for (let i = 0; i < layouts.length; i++) {
    const layout = layouts[i];
    const tag = String(i + 1).padStart(2, '0');
    log(`[${i + 1}/${layouts.length}] ${layout.name} — switching...`);

    let result = null;
    let lastError = null;
    for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
      try {
        result = await captureLayoutOnce(layout, tag);
        break;
      } catch (err) {
        lastError = err;
        log(`  attempt ${attempt}/${MAX_RETRIES} failed: ${err.message}`);
        if (attempt < MAX_RETRIES) {
          await sleep(RETRY_BACKOFF_MS);
          try {
            await health.healthCheck();
          } catch {
            log('  still cannot reach TradingView — will retry anyway');
          }
        }
      }
    }

    if (!result) {
      const staleCharts = fallbackToStaleScreenshots(layout, tag);
      if (staleCharts.length > 0) {
        log(`  all ${MAX_RETRIES} attempts failed — reusing ${staleCharts.length} screenshot(s) from a previous run`);
        result = { status: 'stale_fallback', charts: staleCharts, indicators: [] };
      } else {
        log(`  FAILED after ${MAX_RETRIES} attempts, no previous screenshot to fall back on: ${lastError?.message}`);
        result = { status: 'failed', charts: [{ ticker: null, description: null, screenshot: null, error: lastError?.message || 'unknown error' }], indicators: [] };
      }
    }

    for (const chartEntry of result.charts) {
      manifest.push({ id: layout.id, chartId: layout.url, name: layout.name, ...chartEntry });
    }
    indicatorRows.push(...result.indicators);
    layoutsSummary.push({ layoutId: layout.id, chartId: layout.url, name: layout.name, status: result.status, error: result.charts[0]?.error || null, charts: result.charts });
  }

  writeFileSync(MANIFEST_PATH, JSON.stringify(manifest, null, 2));

  const failedCount = manifest.filter(m => !m.screenshot).length;
  if (failedCount > 0) {
    log(`\n${failedCount}/${manifest.length} rows failed to capture (see rows above) — fix connectivity and re-run to fill them in.`);
  }
  log(`Collected ${indicatorRows.length} indicator readings.`);

  log('\nFetching alerts...');
  const alertsResult = await alerts.list();
  const alertRows = (alertsResult.alerts || []).map(a => ({
    alertId: a.alert_id,
    symbol: a.symbol,
    message: a.message,
    conditionType: a.condition?.type || null,
    targetPrice: (a.condition?.series || []).find(s => s.type === 'value')?.value ?? null,
    resolution: a.resolution,
    active: a.active,
    created: a.created,
    lastFired: a.last_fired,
    expiration: a.expiration,
  }));
  log(`Fetched ${alertRows.length} alerts (${alertRows.filter(a => a.active).length} active).`);

  writeFileSync(INDICATOR_MANIFEST_PATH, JSON.stringify(indicatorRows, null, 2));
  writeFileSync(ALERTS_MANIFEST_PATH, JSON.stringify(alertRows, null, 2));

  log('\nBuilding Excel workbook (Charts, Indicators, Alerts sheets)...');
  const py = spawnSync('python', [path.join(__dirname, 'build_layout_excel.py'), MANIFEST_PATH, OUT_PATH, INDICATOR_MANIFEST_PATH, ALERTS_MANIFEST_PATH], { stdio: 'inherit' });
  if (py.status !== 0) {
    log('Failed to build the Excel file (see python output above).');
    writeSummary({ startedAt, ok: false, reason: 'Excel build failed', layoutsSummary, manifest, failedCount, indicatorRows, alertRows });
    process.exitCode = 1;
    return;
  }
  log(`\nDone. ${manifest.length - failedCount}/${manifest.length} charts, ${indicatorRows.length} indicator readings, ${alertRows.length} alerts captured successfully.`);
  writeSummary({ startedAt, ok: true, reason: null, layoutsSummary, manifest, failedCount, indicatorRows, alertRows });
}

function writeSummary({ startedAt, ok, reason, layoutsSummary, manifest = [], failedCount = 0, indicatorRows = [], alertRows = [] }) {
  const finishedAt = new Date();
  const summary = {
    timestamp: startedAt.toISOString(),
    durationSeconds: Math.round((finishedAt - startedAt) / 1000),
    ok,
    reason,
    outputPath: OUT_PATH,
    logPath: LOG_PATH,
    totalLayouts: layoutsSummary.length,
    totalCharts: manifest.length,
    failedCharts: failedCount,
    totalIndicatorReadings: indicatorRows.length,
    totalAlerts: alertRows.length,
    activeAlerts: alertRows.filter(a => a.active).length,
    layouts: layoutsSummary,
    indicators: indicatorRows,
    alerts: alertRows,
  };
  writeFileSync(SUMMARY_PATH, JSON.stringify(summary, null, 2));
  writeFileSync(SUMMARY_LATEST_PATH, JSON.stringify(summary, null, 2));
  log(`Summary written to ${SUMMARY_PATH}`);
}

main()
  .catch(err => {
    log('\nExport failed:', err);
    writeSummary({ startedAt, ok: false, reason: err.message, layoutsSummary: [] });
    process.exitCode = 1;
  })
  .finally(() => {
    // The CDP client (src/connection.js) keeps an open WebSocket that would
    // otherwise hold the event loop open forever, so a parent process chaining
    // this script via spawnSync (e.g. run_full_pipeline.js) would hang
    // indefinitely waiting for this process to exit. Force it closed once
    // main() has genuinely finished.
    process.exit(process.exitCode || 0);
  });
