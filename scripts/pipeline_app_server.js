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
  res.writeHead(404);
  res.end('Not found');
});

server.listen(PORT, () => {
  console.log(`Pipeline app running at http://localhost:${PORT}`);
});
