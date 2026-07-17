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
import { productWebLink } from './config.js';

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
