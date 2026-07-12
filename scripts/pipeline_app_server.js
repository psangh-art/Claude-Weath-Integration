#!/usr/bin/env node
// Local web app for the unified pipeline: one Execute button, live per-stage
// status, and a hard stop (not a warning) if a required input file — e.g. a
// Fidelity export — is missing, so the user can supply it and re-run rather than
// the process failing confusingly deep into a multi-minute run.
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

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DOWNLOADS = path.join(os.homedir(), 'Downloads');
const PORT = 4590;

const REPO_ROOT = path.join(__dirname, '..');
const APP_DIR = path.join(__dirname, 'pipeline_app');
const GALLERY_HTML = path.join(APP_DIR, 'review_deck.html');
const DECK_SUMMARY = path.join(APP_DIR, 'review_deck_summary.json');
const DECK_PPTX = path.join(DOWNLOADS, 'Investment_Review_Deck.pptx');
const SPENDING_XLSX = path.join(DOWNLOADS, 'spending_summary.xlsx');
// The Finance Google Sheet the pipeline syncs into (see CLAUDE.md).
const FINANCE_SHEET_URL =
  'https://docs.google.com/spreadsheets/d/1UjAz_QUuh86_e6yq8QJf2veI8IpkRCyVfWaK6maqiyc/edit';

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

const PYTHON_CANDIDATES = [
  'C:\\Users\\Paul\\AppData\\Local\\Python\\bin\\python.exe', // has pandas/openpyxl — spending_summary.py needs these
  'python',
];
const PYTHON = PYTHON_CANDIDATES.find((p) => {
  if (p === 'python') return true;
  try { return fs.existsSync(p); } catch { return false; }
}) || 'python';

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

function stageEvent(id, status, message) {
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
    stageEvent(1, 'failed', `${report.missing.length} required file(s) missing:\n${lines}\n\nSupply the file(s) in Downloads and click Execute again.`);
    return null;
  }
  const note = report.found.fidelity_pending_note ? ` (${report.found.fidelity_pending_note})` : '';
  stageEvent(1, 'success', `All required files found.${note}`);
  return report.found;
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
    const stageForStepName = (text) => {
      if (/capturing charts/i.test(text)) return 3;
      if (/OCR channel/i.test(text)) return 4;
      if (/applying results/i.test(text)) return 5;
      if (/review deck/i.test(text)) return 6;
      if (/verifying this run/i.test(text)) return 7;
      if (/flagging redundant/i.test(text)) return 8;
      return null;
    };

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
          if (currentStage) stageEvent(currentStage, 'success');
          currentStage = stageForStepName(marker[1]);
          if (currentStage) stageEvent(currentStage, 'running');
        } else if (line.trim() && currentStage) {
          logEvent(currentStage, line);
        }
      }
    };
    child.stdout.on('data', onData);
    child.stderr.on('data', onData);

    child.on('close', (code) => {
      if (currentStage) stageEvent(currentStage, code === 0 ? 'success' : 'failed', code === 0 ? '' : `Exited with code ${code} — see log above.`);
      // Any stage that never got a marker (e.g. chart export failed before
      // reaching later steps) is left as 'pending' by the caller and shown skipped.
      resolve(code === 0);
    });
  });
}

async function executeRun() {
  if (running) return;
  running = true;
  broadcast({ type: 'run-started' });

  for (const s of STAGES) stageEvent(s.id, 'pending');

  const found = runPreflight();
  if (!found) {
    for (const s of STAGES) if (s.id > 1) stageEvent(s.id, 'skipped', 'Skipped — pre-flight check failed.');
    broadcast({ type: 'run-complete', ok: false });
    running = false;
    return;
  }

  const fidelityOk = runFidelityBuild(found);
  if (!fidelityOk) {
    for (const s of STAGES) if (s.id > 2) stageEvent(s.id, 'skipped', 'Skipped — Fidelity build failed.');
    broadcast({ type: 'run-complete', ok: false });
    running = false;
    return;
  }

  const tvOk = await runTradingViewPipeline();
  broadcast({ type: 'run-complete', ok: tvOk });
  running = false;
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
    res.write(`data: ${JSON.stringify({ type: 'hello', stages: STAGES, running })}\n\n`);
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

  // Output bay: which products exist and where to open them.
  if (req.method === 'GET' && req.url === '/products') {
    let deckSummary = null;
    try { deckSummary = JSON.parse(fs.readFileSync(DECK_SUMMARY, 'utf-8')); } catch { /* not built yet */ }
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({
      googleSheet: { url: FINANCE_SHEET_URL },
      deck: { exists: fs.existsSync(GALLERY_HTML), viewUrl: '/deck', pptxUrl: '/deck.pptx', summary: deckSummary },
      spending: { exists: fs.existsSync(SPENDING_XLSX), url: '/download/spending' },
    }));
    return;
  }

  if (req.method === 'GET' && req.url === '/deck') { serveFile(res, GALLERY_HTML); return; }
  if (req.method === 'GET' && req.url === '/deck.pptx') { serveFile(res, DECK_PPTX, { download: true }); return; }
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

server.listen(PORT, () => {
  console.log(`Pipeline app running at http://localhost:${PORT}`);
});
