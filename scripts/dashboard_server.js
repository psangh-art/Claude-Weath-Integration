// Investment Dashboard — server (Phase 1).
// Plain Node http + SSE, no dependencies. Serves the self-contained front end in
// dashboard_app/, the generated JSON in dashboard_app/data/, regenerates that data
// via dashboard_data.py on startup and on demand, and pushes an SSE 'refresh' when
// it changes. Separate from pipeline_app_server.js (that RUNS the pipeline; this one
// only CONSUMES its output). Port 4600.
import http from 'http';
import fs from 'fs';
import path from 'path';
import { spawn } from 'child_process';
import { fileURLToPath } from 'url';
import os from 'os';
import { productWebLink, CFG } from './config.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const PORT = 4600;
const APP_DIR = path.join(__dirname, 'dashboard_app');
const DATA_DIR = path.join(APP_DIR, 'data');
const PYTHON = process.env.PYTHON || 'python';

const MIME = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8', '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml', '.png': 'image/png', '.ico': 'image/x-icon',
};

const sseClients = new Set();
function broadcast(event, data) {
  const msg = `event: ${event}\ndata: ${JSON.stringify(data)}\n\n`;
  for (const res of sseClients) res.write(msg);
}

let regenerating = false;
function regenerate(reason) {
  if (regenerating) return;
  regenerating = true;
  console.log(`[data] regenerating (${reason})...`);
  const py = spawn(PYTHON, [path.join(__dirname, 'dashboard_data.py')], { cwd: path.dirname(__dirname) });
  let err = '';
  py.stderr.on('data', (d) => { err += d; });
  py.on('close', (code) => {
    regenerating = false;
    if (code === 0) { console.log('[data] refreshed'); broadcast('refresh', { at: Date.now() }); }
    else { console.error('[data] FAILED:', err.slice(-400)); broadcast('error', { message: 'data refresh failed' }); }
  });
}

// ---- Intelligence: live index quotes (Yahoo Finance chart API) ------------
// The 'refresh just these widgets' data source. Fetched server-side (the browser
// can't call the Yahoo API cross-origin), cached briefly, forced fresh with
// ?refresh=1. Symbols are config-driven (config.json -> intelligenceIndices) so a
// symbol that doesn't resolve can be corrected without touching code. Yahoo's
// chart endpoint returns clean JSON (day change from meta + a close series for the
// sparkline) and, unlike Stooq's CSV, isn't behind a bot/JS challenge.
const INTEL_INDICES = Array.isArray(CFG.intelligenceIndices) ? CFG.intelligenceIndices : [];
const INTEL_TTL_MS = 5 * 60 * 1000;      // serve cache for 5 min unless forced
let intelCache = { at: 0, payload: null };

async function fetchJson(url, ms) {
  const ctrl = new AbortController();
  const t = setTimeout(() => ctrl.abort(), ms);
  try {
    const r = await fetch(url, { signal: ctrl.signal, headers: { 'User-Agent': 'Mozilla/5.0' } });
    if (!r.ok) throw new Error('HTTP ' + r.status);
    return await r.json();
  } finally {
    clearTimeout(t);
  }
}

async function fetchIndex(idx) {
  const url = 'https://query1.finance.yahoo.com/v8/finance/chart/'
            + encodeURIComponent(idx.symbol) + '?range=1mo&interval=1d';
  const base = { key: idx.key, label: idx.label, symbol: idx.symbol, currency: idx.currency || null };
  try {
    const j = await fetchJson(url, 8000);
    const result = j && j.chart && j.chart.result && j.chart.result[0];
    if (!result) return { ...base, error: 'no data (check symbol)' };
    const meta = result.meta || {};
    const closes = (((result.indicators || {}).quote || [])[0] || {}).close || [];
    const series = closes.filter(x => typeof x === 'number').slice(-30);
    const last = typeof meta.regularMarketPrice === 'number' ? meta.regularMarketPrice
               : (series.length ? series[series.length - 1] : null);
    const prev = typeof meta.chartPreviousClose === 'number' ? meta.chartPreviousClose
               : (typeof meta.previousClose === 'number' ? meta.previousClose : null);
    if (last == null) return { ...base, error: 'no data (check symbol)' };
    const change = (prev != null) ? last - prev : null;
    const change_pct = (prev) ? (change / prev) * 100 : null;
    const ts = (result.timestamp && result.timestamp.length)
             ? new Date(result.timestamp[result.timestamp.length - 1] * 1000).toISOString().slice(0, 10)
             : null;
    return {
      ...base,
      value: Math.round(last * 100) / 100,
      change: change == null ? null : Math.round(change * 100) / 100,
      change_pct: change_pct == null ? null : Math.round(change_pct * 100) / 100,
      as_of: ts,
      series: series.length >= 2 ? series : [],
    };
  } catch (e) {
    return { ...base, error: String(e.message || e) };
  }
}

async function getIntelligence(force) {
  const now = Date.now();
  if (!force && intelCache.payload && (now - intelCache.at) < INTEL_TTL_MS) return intelCache.payload;
  const results = await Promise.all(INTEL_INDICES.map(fetchIndex));
  const indices = {};
  for (const r of results) indices[r.key] = r;
  intelCache = { at: now, payload: { at: new Date().toISOString(), indices } };
  return intelCache.payload;
}

function serveFile(res, file) {
  fs.readFile(file, (e, buf) => {
    if (e) { res.writeHead(404); res.end('Not found'); return; }
    res.writeHead(200, { 'Content-Type': MIME[path.extname(file).toLowerCase()] || 'application/octet-stream' });
    res.end(buf);
  });
}

const server = http.createServer((req, res) => {
  const url = new URL(req.url, `http://localhost:${PORT}`);
  const pathname = decodeURIComponent(url.pathname);

  if (pathname === '/events') {
    res.writeHead(200, { 'Content-Type': 'text/event-stream', 'Cache-Control': 'no-cache', Connection: 'keep-alive' });
    res.write(`event: hello\ndata: ${JSON.stringify({ at: Date.now() })}\n\n`);
    sseClients.add(res);
    req.on('close', () => sseClients.delete(res));
    return;
  }
  if (pathname === '/refresh') { regenerate('manual'); res.writeHead(200, { 'Content-Type': 'application/json' }); res.end('{"ok":true}'); return; }

  // Live index quotes for the Intelligence screen. ?refresh=1 bypasses the cache.
  if (pathname === '/api/intelligence') {
    const force = url.searchParams.get('refresh') === '1';
    getIntelligence(force).then(payload => {
      res.writeHead(200, { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' });
      res.end(JSON.stringify(payload));
    }).catch(e => {
      res.writeHead(502, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: String(e.message || e), indices: {} }));
    });
    return;
  }

  // Serve the built decks (from Downloads) so the nav links work standalone.
  // Each deck ALSO has a productWebLinks key (scripts/config.json) — the same
  // one-time-pasted OneDrive share link pipeline_app_server.js's Output Bay
  // uses. When it's set, redirect straight there instead of streaming the raw
  // file: an Office share link opens directly in PowerPoint Online, which is
  // what "open in PowerPoint Online" actually requires (a real HTTPS URL Office
  // can serve — localhost can't be wrapped with the officeapps.live.com viewer,
  // since that service has to fetch the file itself). Falls back to serving the
  // raw .pptx (today's behaviour — downloads/opens locally) until the link is
  // pasted in, exactly like the Production Centre's Output Bay tiles already do.
  const DOWNLOADS = path.join(os.homedir(), 'Downloads');
  const DECKS = {
    '/decks/architecture.pptx': { file: 'Financial_Data_Pipeline_Architecture.pptx', linkKey: 'architecturePptx' },
    '/decks/alert-rules.pptx': { file: 'Alert_Rules_Model.pptx', linkKey: 'alertRules' },
    '/decks/review-deck.pptx': { file: 'Investment_Review_Deck.pptx', linkKey: 'reviewDeck' },
  };
  if (DECKS[pathname]) {
    const deck = DECKS[pathname];
    const webLink = productWebLink(deck.linkKey);
    if (webLink) { res.writeHead(302, { Location: webLink }); res.end(); return; }
    const f = path.join(DOWNLOADS, deck.file);
    fs.readFile(f, (e, buf) => {
      if (e) { res.writeHead(404); res.end('Deck not found — run the pipeline to build it.'); return; }
      res.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'Content-Disposition': `inline; filename="${deck.file}"`,
      });
      res.end(buf);
    });
    return;
  }

  // Static: data JSON, then app files; default to index.html. No path traversal.
  let rel = pathname === '/' ? 'index.html' : pathname.replace(/^\/+/, '');
  const target = path.normalize(path.join(APP_DIR, rel));
  if (!target.startsWith(APP_DIR)) { res.writeHead(403); res.end('Forbidden'); return; }
  serveFile(res, target);
});

regenerate('startup');
server.listen(PORT, () => console.log(`Investment Dashboard -> http://localhost:${PORT}`));
