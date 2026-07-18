// Investment Dashboard — server (Phase 1).
// Plain Node http + SSE, no dependencies. Serves the self-contained front end in
// dashboard_app/, the generated JSON in dashboard_app/data/, regenerates that data
// via dashboard_data.py on startup and on demand, and pushes an SSE 'refresh' when
// it changes. Separate from pipeline_app_server.js (that RUNS the pipeline; this one
// only CONSUMES its output). Port 4600.
import http from 'http';
import net from 'net';
import fs from 'fs';
import path from 'path';
import { spawn, spawnSync } from 'child_process';
import { fileURLToPath } from 'url';
import os from 'os';
import { productWebLink, CFG } from './config.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const PORT = 4600;
const APP_DIR = path.join(__dirname, 'dashboard_app');
const DATA_DIR = path.join(APP_DIR, 'data');
const PYTHON = process.env.PYTHON || 'python';
const REPO_ROOT = path.join(__dirname, '..');
const REVIEW_GALLERY = path.join(__dirname, 'pipeline_app', 'review_deck.html');

// In-app readable view of the architecture deck. No slide renderer is available,
// so a raw .pptx link only downloads the file — render_architecture_html.py lays
// the SAME .pptx out as HTML instead. Regenerated lazily when the .pptx is newer
// than the cached HTML, so the view always reflects the latest deck build.
const ARCH_PPTX = path.join(os.homedir(), 'Downloads', 'Financial_Data_Pipeline_Architecture.pptx');
const ARCH_HTML = path.join(APP_DIR, 'architecture.html');
const ALERTRULES_PPTX = path.join(os.homedir(), 'Downloads', 'Alert_Rules_Model.pptx');
const ALERTRULES_HTML = path.join(APP_DIR, 'alert_rules.html');
// Render a .pptx deck to its in-app HTML if the HTML is missing/stale. `renderArgs`
// are passed to render_architecture_html.py (the shared renderer): [pptx, html] for
// the architecture deck, ['alert-rules'] for the alert-rules deck (self-pathed).
function ensureDeckHtml(pptxPath, htmlPath, renderArgs) {
  let src;
  try { src = fs.statSync(pptxPath); } catch { return false; }  // no deck built yet
  let stale = true;
  try { stale = fs.statSync(htmlPath).mtimeMs < src.mtimeMs; } catch { stale = true; }
  if (stale) {
    const r = spawnSync(PYTHON, [path.join(__dirname, 'render_architecture_html.py'), ...renderArgs],
      { cwd: REPO_ROOT, env: { ...process.env, PYTHONUTF8: '1' } });
    if (r.status !== 0) { console.error('[deck] render failed:', String(r.stderr || '').slice(-300)); return false; }
  }
  return true;
}
const ensureArchitectureHtml = () => ensureDeckHtml(ARCH_PPTX, ARCH_HTML, [ARCH_PPTX, ARCH_HTML]);
const ensureAlertRulesHtml = () => ensureDeckHtml(ALERTRULES_PPTX, ALERTRULES_HTML, ['alert-rules']);

// Whitelist for the /asset image proxy the review-deck gallery uses: only images
// inside the repo or Downloads (same rule as pipeline_app_server.js).
function isAllowedAsset(abs) {
  const resolved = path.resolve(abs);
  const roots = [path.resolve(REPO_ROOT), path.resolve(path.join(os.homedir(), 'Downloads'))];
  return roots.some(root => resolved === root || resolved.startsWith(root + path.sep))
    && /\.(png|jpe?g)$/i.test(resolved);
}

const MIME = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8', '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml', '.png': 'image/png', '.ico': 'image/x-icon',
};

// ---- Pipeline app (Investment Production Centre, port 4590) ---------------
// The dashboard's "Pipeline" nav link opens the build screen that runs the
// spending/OCR/master-sheet pipeline. That's a SEPARATE server (pipeline_app_
// server.js). The launcher tries to start it in the background, but if that
// failed (or it crashed / was never started), clicking Pipeline hit a dead port
// and the browser said "site cannot be reached" (user report 2026-07-18). The
// nav link now points at THIS server's /pipeline route, which starts the app on
// demand if it isn't already listening, then redirects — so it always works.
const PIPELINE_PORT = (CFG && CFG.appPort) || 4590;

function portOpen(port, cb) {
  const sock = net.connect({ port, host: '127.0.0.1' });
  let settled = false;
  const finish = (up) => { if (!settled) { settled = true; sock.destroy(); cb(up); } };
  sock.setTimeout(1000);
  sock.once('connect', () => finish(true));
  sock.once('timeout', () => finish(false));
  sock.once('error', () => finish(false));
}

let pipelineStarting = false;
function ensurePipelineApp(cb) {
  portOpen(PIPELINE_PORT, (up) => {
    if (up) return cb(true);
    if (!pipelineStarting) {
      pipelineStarting = true;
      console.log('[pipeline] not running — starting the Pipeline app...');
      try {
        const child = spawn(process.execPath, [path.join(__dirname, 'pipeline_app_server.js')],
          { cwd: REPO_ROOT, detached: true, stdio: 'ignore' });
        child.unref();
      } catch (e) {
        console.error('[pipeline] spawn failed:', e.message);
        pipelineStarting = false;
        return cb(false);
      }
    }
    let tries = 0;                                   // poll up to ~8s for it to listen
    const iv = setInterval(() => {
      portOpen(PIPELINE_PORT, (u) => {
        if (u) { clearInterval(iv); pipelineStarting = false; cb(true); }
        else if (++tries >= 16) { clearInterval(iv); pipelineStarting = false; cb(false); }
      });
    }, 500);
  });
}

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

// ---- Watchlist LIVE prices (user request 2026-07-18) ----------------------
// The top Refresh button re-derives the watchlist from the last CAPTURED prices
// (history.db). This endpoint instead pulls LIVE market prices for the watchlist
// tickers straight from Yahoo, on demand (the Watchlist screen's ↻). LSE equities
// are quoted in pence (GBp) on Yahoo, matching the sheet; the default symbol rule
// is '<TICKER>.L' with '.'→'-' (BT.A → BT-A.L). config.json → watchlistYahooSymbols
// overrides it, and '' marks a ticker (e.g. a USD/oz commodity) as having no
// pence-denominated live source, so it's reported unsupported rather than guessed.
const WATCH_SYMBOL_OVERRIDE = (CFG.watchlistYahooSymbols && typeof CFG.watchlistYahooSymbols === 'object')
  ? CFG.watchlistYahooSymbols : {};
function yahooSymbolFor(ticker) {
  if (!ticker) return null;
  const t = String(ticker).trim().toUpperCase();
  if (Object.prototype.hasOwnProperty.call(WATCH_SYMBOL_OVERRIDE, t)) {
    return WATCH_SYMBOL_OVERRIDE[t] || null;   // '' => explicitly unsupported
  }
  return t.replace(/\./g, '-') + '.L';          // default: LSE equity in pence
}

async function fetchQuote(symbol) {
  const url = 'https://query1.finance.yahoo.com/v8/finance/chart/'
            + encodeURIComponent(symbol) + '?range=5d&interval=1d';
  const j = await fetchJson(url, 8000);
  const result = j && j.chart && j.chart.result && j.chart.result[0];
  if (!result) throw new Error('no data (check symbol)');
  const meta = result.meta || {};
  const last = typeof meta.regularMarketPrice === 'number' ? meta.regularMarketPrice : null;
  const prev = typeof meta.chartPreviousClose === 'number' ? meta.chartPreviousClose
             : (typeof meta.previousClose === 'number' ? meta.previousClose : null);
  if (last == null) throw new Error('no price');
  return { price: last, prev, currency: meta.currency || null };
}

async function getWatchlistLive() {
  let tickers = [];
  try {
    const wl = JSON.parse(fs.readFileSync(path.join(DATA_DIR, 'watchlist.json'), 'utf-8'));
    const rows = Array.isArray(wl) ? wl : (wl.rows || []);
    tickers = [...new Set(rows.map(r => r.ticker).filter(Boolean).map(t => String(t).trim().toUpperCase()))];
  } catch (e) {
    return { at: new Date().toISOString(), prices: {}, error: 'watchlist not built yet' };
  }
  const prices = {};
  await Promise.all(tickers.map(async (t) => {
    const sym = yahooSymbolFor(t);
    if (!sym) { prices[t] = { error: 'no live source' }; return; }
    try {
      const q = await fetchQuote(sym);
      const change = (q.prev != null) ? q.price - q.prev : null;
      prices[t] = {
        price: Math.round(q.price * 100) / 100,
        change: change == null ? null : Math.round(change * 100) / 100,
        change_pct: (q.prev) ? Math.round((change / q.prev * 100) * 100) / 100 : null,
        currency: q.currency, symbol: sym,
      };
    } catch (e) {
      prices[t] = { error: String(e.message || e), symbol: sym };
    }
  }));
  return { at: new Date().toISOString(), prices };
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

  // The pipeline finished and already regenerated the data files (run_full_pipeline
  // -> dashboard_data.py). Just tell open dashboards to reload via SSE — no need to
  // regenerate again here. This is what ties a pipeline run into the live dashboard.
  if (pathname === '/pipeline-updated') {
    broadcast('refresh', { at: Date.now(), source: 'pipeline' });
    res.writeHead(200, { 'Content-Type': 'application/json' }); res.end('{"ok":true}');
    return;
  }

  // Pipeline status for the Overview status button: last pipeline run (latest
  // history.db capture) + the age of each input file (from preflight_check.py).
  if (pathname === '/api/pipeline-status') {
    let files = [], lastRun = null;
    try {
      const pf = spawnSync(PYTHON, [path.join(__dirname, 'preflight_check.py')],
        { encoding: 'utf-8', cwd: REPO_ROOT, env: { ...process.env, PYTHONUTF8: '1' } });
      const r = JSON.parse(pf.stdout);
      files = Object.values(r.files || {}).map(f => ({
        label: f.label, present: f.present, as_of: f.as_of, age_days: f.age_days, stale: f.stale }));
    } catch (e) { /* leave files empty */ }
    try {
      const hr = spawnSync(PYTHON, ['-c',
        "import sqlite3,os;d=os.path.join('data','history.db');"
        + "print(sqlite3.connect(d).execute('SELECT MAX(price_checked_at) FROM chart_snapshots').fetchone()[0] if os.path.exists(d) else '')"],
        { encoding: 'utf-8', cwd: REPO_ROOT });
      lastRun = (hr.stdout || '').trim() || null;
    } catch (e) { /* leave null */ }
    res.writeHead(200, { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' });
    res.end(JSON.stringify({ last_run: lastRun, files }));
    return;
  }

  // Open the Pipeline app (formerly "Investment Production Centre"), starting it first.
  if (pathname === '/pipeline') {
    ensurePipelineApp((up) => {
      if (up) { res.writeHead(302, { Location: `http://localhost:${PIPELINE_PORT}/` }); res.end(); return; }
      res.writeHead(503, { 'Content-Type': 'text/html; charset=utf-8' });
      res.end('<body style="font-family:system-ui;padding:40px;max-width:640px">'
        + '<h2>Couldn’t start the Pipeline app</h2>'
        + `<p>The pipeline build screen (port ${PIPELINE_PORT}) did not come up. `
        + 'Try launching it directly with <b>Run Pipeline App.bat</b>, then click Pipeline again.</p></body>');
    });
    return;
  }

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

  // Live market prices for the Watchlist tickers (Watchlist screen's ↻ button).
  if (pathname === '/api/watchlist-live') {
    getWatchlistLive().then(payload => {
      res.writeHead(200, { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' });
      res.end(JSON.stringify(payload));
    }).catch(e => {
      res.writeHead(502, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: String(e.message || e), prices: {} }));
    });
    return;
  }

  // Review-deck IN-APP GALLERY (user request 2026-07-17): serve the image
  // gallery build_review_deck.py emits (scripts/pipeline_app/review_deck.html),
  // rewriting its relative "asset?p=" image srcs to the absolute "/asset?p="
  // route below so they resolve no matter the page path. (Architecture and
  // alert-rules stay raw .pptx downloads — no slide renderer is available.)
  if (pathname === '/decks/review-deck') {
    fs.readFile(REVIEW_GALLERY, 'utf-8', (e, html) => {
      if (e) { res.writeHead(404); res.end('Review-deck gallery not built yet — run the pipeline (or build_review_deck.py).'); return; }
      const fixed = html.replace(/(["'])asset\?p=/g, '$1/asset?p=');
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      res.end(fixed);
    });
    return;
  }
  // Architecture deck IN-APP READABLE VIEW (user request 2026-07-18): a raw .pptx
  // link only downloaded the file, so render the deck as HTML and serve it here.
  if (pathname === '/decks/architecture') {
    if (!ensureArchitectureHtml()) {
      res.writeHead(404); res.end('Architecture deck not built yet — run the pipeline (or the deck scripts).'); return;
    }
    fs.readFile(ARCH_HTML, 'utf-8', (e, html) => {
      if (e) { res.writeHead(404); res.end('Architecture view not available.'); return; }
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      res.end(html);
    });
    return;
  }
  // Alert Rules deck IN-APP READABLE VIEW (user request 2026-07-18): same as the
  // architecture view — render the .pptx to HTML rather than download it.
  if (pathname === '/decks/alert-rules') {
    if (!ensureAlertRulesHtml()) {
      res.writeHead(404); res.end('Alert Rules deck not built yet — run build_rules_deck.py.'); return;
    }
    fs.readFile(ALERTRULES_HTML, 'utf-8', (e, html) => {
      if (e) { res.writeHead(404); res.end('Alert Rules view not available.'); return; }
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      res.end(html);
    });
    return;
  }
  // Image proxy for the gallery — whitelisted to the repo + Downloads only.
  if (pathname === '/asset') {
    const p = url.searchParams.get('p');
    if (p && isAllowedAsset(p)) { serveFile(res, path.resolve(p)); return; }
    res.writeHead(403); res.end('Forbidden'); return;
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
