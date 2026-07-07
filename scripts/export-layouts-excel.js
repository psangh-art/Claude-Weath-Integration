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
import { writeFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { spawnSync } from 'child_process';
import * as health from '../src/core/health.js';
import * as ui from '../src/core/ui.js';
import * as pane from '../src/core/pane.js';
import * as capture from '../src/core/capture.js';
import { evaluateAsync } from '../src/connection.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const OUT_PATH = path.join(os.homedir(), 'Downloads', 'tradingview_layouts.xlsx');
const MANIFEST_PATH = path.join(__dirname, 'layout_manifest_tmp.json');
const CROP_MANIFEST_PATH = path.join(__dirname, 'crop_manifest_tmp.json');
const CAPTURE_SCALE = 2;

async function main() {
  console.log('Checking TradingView CDP connection on port 9222...');
  try {
    await health.healthCheck();
  } catch {
    console.error('\nCould not connect to TradingView.');
    console.error('Start it with CDP enabled first, e.g. run scripts\\launch_tv_debug.bat, then re-run this.');
    process.exitCode = 1;
    return;
  }

  console.log('Fetching saved layouts...');
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
    console.log(`Skipping ${skipped} placeholder layout(s) (e.g. "Test").`);
  }
  if (layouts.length === 0) {
    console.error('No saved layouts found.');
    process.exitCode = 1;
    return;
  }
  console.log(`Found ${layouts.length} layouts.\n`);

  const manifest = [];

  for (let i = 0; i < layouts.length; i++) {
    const layout = layouts[i];
    const tag = String(i + 1).padStart(2, '0');
    console.log(`[${i + 1}/${layouts.length}] ${layout.name} — switching...`);

    try {
      await ui.layoutSwitch({ name: layout.name });
    } catch (err) {
      console.error(`  FAILED to switch: ${err.message}`);
      manifest.push({ id: layout.id, chartId: layout.url, name: layout.name, ticker: null, description: null, screenshot: null, error: err.message });
      continue;
    }

    // Reset zoom/pan to a consistent fitted view before capturing — a chart can be
    // left scrolled/zoomed from whenever it was last interacted with. This is a
    // view-only, unsaved change (see resetView()'s comment) so it never touches the
    // saved layout itself.
    await ui.resetView();
    await new Promise(r => setTimeout(r, 800));

    const rawFilename = `layout_${tag}_raw`;
    const shot = await capture.captureScreenshot({ region: 'full', filename: rawFilename, scale: CAPTURE_SCALE });

    const paneData = await pane.listWithRects();
    const panes = (paneData.panes || []).filter(p => p.rect && p.rect.width > 0 && p.rect.height > 0);

    if (panes.length === 0) {
      console.warn(`  no measurable panes — falling back to full-layout screenshot`);
      manifest.push({ id: layout.id, chartId: layout.url, name: layout.name, ticker: null, description: null, screenshot: shot.file_path, error: null });
      continue;
    }

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

    if (cropResult.status !== 0) {
      console.error(`  FAILED to crop panes — falling back to full-layout screenshot for this row`);
      manifest.push({ id: layout.id, chartId: layout.url, name: layout.name, ticker: null, description: null, screenshot: shot.file_path, error: 'pane crop failed' });
      continue;
    }

    for (let pi = 0; pi < panes.length; pi++) {
      const p = panes[pi];
      manifest.push({
        id: layout.id,
        chartId: layout.url,
        name: layout.name,
        ticker: p.ticker || p.symbol || null,
        description: p.description || null,
        screenshot: path.join(cropDir, crops[pi].filename),
        error: null,
      });
    }
    console.log(`  captured ${panes.length} chart(s)`);
  }

  writeFileSync(MANIFEST_PATH, JSON.stringify(manifest, null, 2));

  const failedCount = manifest.filter(m => !m.screenshot).length;
  if (failedCount > 0) {
    console.warn(`\n${failedCount}/${manifest.length} rows failed to capture (see rows above) — fix connectivity and re-run to fill them in.`);
  }

  console.log('\nBuilding Excel workbook...');
  const py = spawnSync('python', [path.join(__dirname, 'build_layout_excel.py'), MANIFEST_PATH, OUT_PATH], { stdio: 'inherit' });
  if (py.status !== 0) {
    console.error('Failed to build the Excel file (see python output above).');
    process.exitCode = 1;
    return;
  }
  console.log(`\nDone. ${manifest.length - failedCount}/${manifest.length} charts captured successfully.`);
}

main().catch(err => {
  console.error('\nExport failed:', err);
  process.exitCode = 1;
});
