#!/usr/bin/env node
// Local web app for the unified pipeline: one Execute button, live per-stage
// status, and a per-input data-age report. The bank/broker exports (Fidelity,
// Barclays, Amex) are OPTIONAL — without them the spending-summary stage is
// skipped and the TradingView stages still run; each feedstock box shows how
// old its data is and turns red past 6 weeks (see preflight_check.py). Only a
// missing master workbook halts the run.
//
// Reuses run_full_pipeline.js as-is for the chart-capture/OCR/master-update/
// review-deck/verify/cleanup stages (spawned as a child, its own "=== Step N/M:
// ... ===" markers are parsed to report those six as separate stages here)
// rather than re-implementing that logic — single source of truth stays there.
import http from 'http';
import fs from 'fs';
import path from 'path';
import os from 'os';
import { spawn, spawnSync } from 'child_process';
import { fileURLToPath } from 'url';

import { downloadsDir, downloadsFile, pythonExe, financeSheetUrl, onedriveProductsDir, productWebLink, CFG } from './config.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DOWNLOADS = downloadsDir();
const PORT = CFG.appPort;

const REPO_ROOT = path.join(__dirname, '..');
const APP_DIR = path.join(__dirname, 'pipeline_app');
const GALLERY_HTML = path.join(APP_DIR, 'review_deck.html');
const DECK_SUMMARY = path.join(APP_DIR, 'review_deck_summary.json');
const DECK_PPTX = downloadsFile('reviewDeckPptx');
const SPENDING_XLSX = downloadsFile('spendingSummaryXlsx');
const ARCH_PPTX = downloadsFile('architecturePptx');
const TIMINGS_PATH = path.join(REPO_ROOT, 'data', 'stage_timings.json');
// The Finance Google Sheet the pipeline syncs into (see CLAUDE.md).
const FINANCE_SHEET_URL = financeSheetUrl();
// Folder inside the OneDrive-synced tree that mirrors built products so
// PowerPoint/Excel Online can open them — ~/Downloads itself isn't synced.
const ONEDRIVE_PRODUCTS_DIR = onedriveProductsDir();

const CONTENT_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
};

function serveFile(res, filePath, { download } = {}) {
  if (!fs.existsSync(filePath)) {
    res.writeHead(404, { 'Content-Type': 'text/plain' });
    res.end('Not found');
    return;
  }
  const type = CONTENT_TYPES[path.extname(filePath).toLowerCase()] || 'application/octet-stream';
  const headers = { 'Content-Type': type };
  if (download) headers['Content-Disposition'] = `attachment; filename="${path.basename(filePath)}"`;
  res.writeHead(200, headers);
  fs.createReadStream(filePath).pipe(res);
}

// Only ever serve image files that live inside the repo (screenshots/) or the
// user's Downloads — never an arbitrary path from the query string.
function isAllowedAsset(abs) {
  const resolved = path.resolve(abs);
  const roots = [path.resolve(REPO_ROOT), path.resolve(DOWNLOADS)];
  return roots.some((root) => resolved === root || resolved.startsWith(root + path.sep))
    && /\.(png|jpe?g)$/i.test(resolved);
}

// Mirrors whichever of the three built products currently exist in Downloads
// into the OneDrive-synced products folder, overwriting in place (not
// delete+recreate) so OneDrive keeps the same item ID — that's what lets a
// one-time-pasted share link in config.json's productWebLinks stay valid
// across every future run. Downloads itself isn't OneDrive-synced, so this is
// the only way Office web apps can open these files. Called at server
// startup and after every run (success or failure — copy whatever exists).
function syncOneDriveProducts() {
  try {
    fs.mkdirSync(ONEDRIVE_PRODUCTS_DIR, { recursive: true });
  } catch (err) {
    console.error(`Could not create OneDrive products folder: ${err.message}`);
    return;
  }
  for (const src of [DECK_PPTX, ARCH_PPTX, SPENDING_XLSX]) {
    if (!fs.existsSync(src)) continue;
    try {
      fs.copyFileSync(src, path.join(ONEDRIVE_PRODUCTS_DIR, path.basename(src)));
    } catch (err) {
      console.error(`Could not sync ${path.basename(src)} to OneDrive: ${err.message}`);
    }
  }
}

// Fallback when no productWebLinks entry has been pasted yet: OneDrive's own
// web search for the exact synced filename — one click away from opening it
// in PowerPoint/Excel Online, versus the direct link's zero clicks.
function productSearchUrl(filename) {
  return `https://onedrive.live.com/?qt=search&q=${encodeURIComponent(filename)}`;
}

// First configured interpreter that exists — the AppData one has pandas/openpyxl,
// which spending_summary.py needs. Candidates live in config.json.
const PYTHON = pythonExe();

const STAGES = [
  { id: 1, name: 'Pre-flight file check' },
  { id: 2, name: 'Fidelity spending-summary build' },
  { id: 3, name: 'TradingView chart capture' },
  { id: 4, name: 'OCR channel-boundary detection' },
  { id: 5, name: 'Master-sheet update (Investments)' },
  { id: 6, name: 'PowerPoint review deck' },
  { id: 7, name: 'Verification' },
  { id: 8, name: 'Downloads cleanup' },
];

let running = false;
const sseClients = new Set();

function broadcast(event) {
  const payload = `data: ${JSON.stringify(event)}\n\n`;
  for (const res of sseClients) res.write(payload);
}

// ── Stage timings: per-stage durations from previous runs feed the front
// end's %-complete bar (each stage weighted by how long it took last time).
// Kept as {stageId: [last few seconds]} in data/stage_timings.json.
function loadTimings() {
  try { return JSON.parse(fs.readFileSync(TIMINGS_PATH, 'utf-8')); } catch { return {}; }
}

function expectedDurations() {
  const timings = loadTimings();
  const expected = {};
  for (const s of STAGES) {
    const samples = (timings[s.id] || []).slice().sort((a, b) => a - b);
    if (samples.length) expected[s.id] = samples[Math.floor(samples.length / 2)]; // median
  }
  return expected;
}

let stageStartedAt = {};   // id -> ms, while a stage is running
let runDurations = {};     // id -> seconds, completed stages of the current run

function recordRunTimings() {
  const timings = loadTimings();
  for (const [id, secs] of Object.entries(runDurations)) {
    timings[id] = [...(timings[id] || []), secs].slice(-5);
  }
  try {
    fs.mkdirSync(path.dirname(TIMINGS_PATH), { recursive: true });
    fs.writeFileSync(TIMINGS_PATH, JSON.stringify(timings, null, 2));
  } catch (err) {
    console.error(`Could not save stage timings: ${err.message}`);
  }
}

function stageEvent(id, status, message) {
  if (status === 'running') {
    stageStartedAt[id] = Date.now();
  } else if ((status === 'success' || status === 'failed') && stageStartedAt[id]) {
    runDurations[id] = (Date.now() - stageStartedAt[id]) / 1000;
    delete stageStartedAt[id];
  }
  broadcast({ type: 'stage', id, name: STAGES.find((s) => s.id === id).name, status, message: message || '' });
}

function logEvent(id, line) {
  broadcast({ type: 'log', id, line });
}

function runPreflight() {
  stageEvent(1, 'running');
  const result = spawnSync(PYTHON, [path.join(__dirname, 'preflight_check.py'), DOWNLOADS], { encoding: 'utf-8' });
  let report;
  try {
    report = JSON.parse(result.stdout);
  } catch {
    stageEvent(1, 'failed', 'Could not parse pre-flight report — see server console.');
    console.error(result.stdout, result.stderr);
    return null;
  }
  if (!report.ok) {
    const lines = report.missing.map((m) => `Missing: ${m.expected}\n  (${m.why})`).join('\n');
    stageEvent(1, 'failed', `${lines}\n\nSupply the file(s) in Downloads and click Execute again.`);
    return null;
  }
  broadcast({ type: 'files', report });
  const notes = [];
  const absent = Object.entries(report.files || {})
    .filter(([k, f]) => !f.present && k !== 'master_workbook' && k !== 'fidelity_pending');
  if (absent.length) {
    notes.push(`${absent.length} bank/broker export(s) not present — spending-summary stage will be skipped.`);
  }
  const stale = Object.values(report.files || {}).filter((f) => f.stale);
  if (stale.length) {
    notes.push(`Data over ${report.stale_days} days old (fresh export needed): ${stale.map((f) => f.label).join(', ')}.`);
  }
  stageEvent(1, 'success', notes.length ? notes.join('\n') : 'All input files present and fresh.');
  return report;
}

function runFidelityBuild(found) {
  stageEvent(2, 'running');
  const args = [
    path.join(__dirname, 'spending_summary.py'),
    found.amex,
    found.barclays,
    found.fidelity_historic,
    found.fidelity_account_summary,
    path.join(DOWNLOADS, 'spending_summary.xlsx'),
  ];
  if (found.fidelity_pending) args.push(found.fidelity_pending);

  const result = spawnSync(PYTHON, args, { encoding: 'utf-8' });
  const output = (result.stdout || '') + (result.stderr || '');
  for (const line of output.split('\n')) if (line.trim()) logEvent(2, line);

  if (result.status !== 0) {
    stageEvent(2, 'failed', 'spending_summary.py exited with an error — see log above.');
    return false;
  }
  const summaryLine = (result.stdout || '').split('\n').find((l) => l.includes('positions')) || 'Done.';
  stageEvent(2, 'success', summaryLine.trim());
  return true;
}

// Maps run_full_pipeline.js's own "=== Step N/M: <name> ===" console markers onto
// stages 3-8 here, so its existing chart-capture/OCR/master-update/deck/verify/cleanup
// logic doesn't need to be duplicated.
function runTradingViewPipeline() {
  return new Promise((resolve) => {
    let currentStage = null;
    const failedStages = new Set();
    const stageForStepName = (text) => {
      if (/capturing charts/i.test(text)) return 3;
      if (/OCR channel/i.test(text)) return 4;
      if (/applying results/i.test(text)) return 5;
      if (/review deck/i.test(text)) return 6;
      if (/verifying this run/i.test(text)) return 7;
      if (/flagging redundant/i.test(text)) return 8;
      return null;
    };
    // run_full_pipeline.js prints one of these when a step fails, then carries on
    // with its finally-block steps — so the exit code alone can't say WHICH stage
    // failed. Attribute the failure to the stage that printed its failure line.
    // (Coupled to run_full_pipeline.js's error strings, like the step markers.)
    const FAILURE_LINE = /^(Chart export failed|Channel detection could not run|Master-sheet update failed|Review-deck build could not run|Verification report could not run|Downloads cleanup could not run)/;

    const child = spawn('node', [path.join(__dirname, 'run_full_pipeline.js')], { cwd: path.join(__dirname, '..') });
    let buf = '';
    const onData = (chunk) => {
      buf += chunk.toString();
      let idx;
      while ((idx = buf.indexOf('\n')) !== -1) {
        const line = buf.slice(0, idx);
        buf = buf.slice(idx + 1);
        const marker = line.match(/=== Step \d+\/\d+: (.+?) ===/);
        if (marker) {
          if (currentStage && !failedStages.has(currentStage)) stageEvent(currentStage, 'success');
          currentStage = stageForStepName(marker[1]);
          if (currentStage) stageEvent(currentStage, 'running');
        } else if (line.trim() && currentStage) {
          if (FAILURE_LINE.test(line.trim()) && !failedStages.has(currentStage)) {
            failedStages.add(currentStage);
            stageEvent(currentStage, 'failed', line.trim());
          }
          logEvent(currentStage, line);
        }
      }
    };
    child.stdout.on('data', onData);
    child.stderr.on('data', onData);

    child.on('close', (code) => {
      if (currentStage && !failedStages.has(currentStage)) {
        // A non-zero exit whose failure was already pinned on an earlier stage
        // (e.g. chart capture) must not paint the last-running stage red too.
        const unattributed = code !== 0 && failedStages.size === 0;
        stageEvent(currentStage, unattributed ? 'failed' : 'success', unattributed ? `Exited with code ${code} — see log above.` : '');
      }
      // Any stage that never got a marker (e.g. chart export failed before
      // reaching later steps) is left as 'pending' by the caller and shown skipped.
      resolve(code === 0 && failedStages.size === 0);
    });
  });
}

async function executeRun() {
  if (running) return;
  running = true;
  stageStartedAt = {};
  runDurations = {};
  broadcast({ type: 'run-started', timings: expectedDurations() });

  for (const s of STAGES) stageEvent(s.id, 'pending');

  const report = runPreflight();
  if (!report) {
    for (const s of STAGES) if (s.id > 1) stageEvent(s.id, 'skipped', 'Skipped — pre-flight check failed.');
    syncOneDriveProducts();
    broadcast({ type: 'run-complete', ok: false });
    running = false;
    return;
  }

  // The bank/broker exports are optional (user policy 2026-07-13): without
  // them the spending build is skipped and the TradingView stages still run.
  // A failed spending build likewise doesn't block the chart pipeline.
  let spendingOk = true;
  if (!report.spending_ready) {
    stageEvent(2, 'skipped', 'Skipped — bank/broker exports not in Downloads. Data ages shown in the feedstock panel; supply fresh files to rebuild spending_summary.xlsx.');
  } else {
    spendingOk = runFidelityBuild(report.found);
  }

  const tvOk = await runTradingViewPipeline();
  if (tvOk && spendingOk) consumeInputFiles();
  recordRunTimings();
  syncOneDriveProducts();
  broadcast({ type: 'run-complete', ok: tvOk && spendingOk });
  running = false;
}

// After a FULLY successful run, the consumed bank/broker exports (and every
// other version of them in Downloads) are sent to the Recycle Bin — user
// policy 2026-07-12: "once files have been used they must be deleted". Failed
// runs keep their inputs so they can be re-run. The master workbook is never
// touched. See consume_input_files.py for the exact family matching.
function consumeInputFiles() {
  const result = spawnSync(PYTHON, [path.join(__dirname, 'consume_input_files.py'), DOWNLOADS, '--apply'], { encoding: 'utf-8' });
  const output = (result.stdout || '') + (result.stderr || '');
  for (const line of output.split('\n')) if (line.trim()) logEvent(8, line);
  if (result.status !== 0) {
    stageEvent(8, 'success', 'Cleanup done, but removing used input files failed — see log.');
  }
}

const server = http.createServer((req, res) => {
  if (req.method === 'GET' && req.url === '/') {
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
    res.end(fs.readFileSync(path.join(__dirname, 'pipeline_app', 'index.html')));
    return;
  }
  if (req.method === 'GET' && req.url === '/events') {
    res.writeHead(200, {
      'Content-Type': 'text/event-stream',
      'Cache-Control': 'no-cache',
      Connection: 'keep-alive',
    });
    res.write(`data: ${JSON.stringify({ type: 'hello', stages: STAGES, running, timings: expectedDurations() })}\n\n`);
    sseClients.add(res);
    req.on('close', () => sseClients.delete(res));
    return;
  }
  if (req.method === 'POST' && req.url === '/run') {
    if (running) {
      res.writeHead(409, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: 'A run is already in progress.' }));
      return;
    }
    executeRun();
    res.writeHead(202, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ started: true }));
    return;
  }

  // Feedstock status on load: which required Downloads files exist, without
  // running the pipeline. Mirrors preflight_check.py's report shape.
  if (req.method === 'GET' && req.url === '/files') {
    const result = spawnSync(PYTHON, [path.join(__dirname, 'preflight_check.py'), DOWNLOADS], { encoding: 'utf-8' });
    let report;
    try { report = JSON.parse(result.stdout); } catch { report = { ok: false, found: {}, missing: [], error: 'preflight parse failed' }; }
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify(report));
    return;
  }

  // Output bay: which products exist and where to open them. Office web apps
  // (PowerPoint/Excel Online) can only open a file that lives in a
  // OneDrive-synced folder (~/Downloads isn't), so each pptx/xlsx product
  // offers a one-time-pasted direct webUrl (config.json productWebLinks) or,
  // failing that, a searchUrl into OneDrive web search for the synced copy.
  if (req.method === 'GET' && req.url === '/products') {
    let deckSummary = null;
    try { deckSummary = JSON.parse(fs.readFileSync(DECK_SUMMARY, 'utf-8')); } catch { /* not built yet */ }
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({
      googleSheet: { url: FINANCE_SHEET_URL },
      deck: {
        exists: fs.existsSync(GALLERY_HTML),
        viewUrl: '/deck',
        summary: deckSummary,
        webUrl: productWebLink('reviewDeck'),
        searchUrl: productSearchUrl(path.basename(DECK_PPTX)),
      },
      spending: {
        exists: fs.existsSync(SPENDING_XLSX),
        webUrl: productWebLink('spendingSummary'),
        searchUrl: productSearchUrl(path.basename(SPENDING_XLSX)),
      },
      architecture: {
        exists: fs.existsSync(ARCH_PPTX),
        webUrl: productWebLink('architecturePptx'),
        searchUrl: productSearchUrl(path.basename(ARCH_PPTX)),
      },
    }));
    return;
  }

  if (req.method === 'GET' && req.url === '/deck') { serveFile(res, GALLERY_HTML); return; }
  if (req.method === 'GET' && req.url === '/deck.pptx') { serveFile(res, DECK_PPTX, { download: true }); return; }
  if (req.method === 'GET' && req.url === '/architecture.pptx') { serveFile(res, ARCH_PPTX, { download: true }); return; }
  if (req.method === 'GET' && req.url === '/download/spending') { serveFile(res, SPENDING_XLSX, { download: true }); return; }

  if (req.method === 'GET' && req.url.startsWith('/asset?')) {
    const p = new URL(req.url, `http://localhost:${PORT}`).searchParams.get('p');
    if (p && isAllowedAsset(p)) { serveFile(res, path.resolve(p)); return; }
    res.writeHead(403, { 'Content-Type': 'text/plain' });
    res.end('Forbidden');
    return;
  }

  res.writeHead(404);
  res.end('Not found');
});

syncOneDriveProducts();

server.listen(PORT, () => {
  console.log(`Pipeline app running at http://localhost:${PORT}`);
});
