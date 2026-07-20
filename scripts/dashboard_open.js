// Opens the Investment Dashboard in a DEDICATED Chrome profile, closing any
// dashboard tab left over from a previous launch (user decision 2026-07-20).
//
// Why a dedicated profile. The front end's singleInstanceGuard() broadcasts a
// claim and older tabs step down, but a browser refuses window.close() on a tab
// the USER opened — so a stale tab can only be covered with an "inactive"
// overlay, never actually removed. Closing it for real needs CDP, and Chrome
// only exposes CDP when it is started with --remote-debugging-port, which it
// will NOT do for an already-running normal profile. Hence a separate profile
// under data/: the user's everyday Chrome (history, sessions, extensions,
// logins) is untouched, and this one exists purely to host the dashboard.
//
// Trade-off the user accepted: the dashboard opens in its own Chrome window
// signed into nothing, so it is not sharing the everyday profile's session.
//
// Run standalone: node scripts/dashboard_open.js [url]
import { spawn } from 'child_process';
import fs from 'fs';
import http from 'http';
import net from 'net';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const DEBUG_PORT = 9333;                       // NOT 9222 — that is TradingView Desktop's
const PROFILE_DIR = path.join(__dirname, '..', 'data', 'dashboard-chrome-profile');

const CHROME_CANDIDATES = [
  'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
  'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
  path.join(process.env.LOCALAPPDATA || '', 'Google\\Chrome\\Application\\chrome.exe'),
];

function chromeExe() {
  for (const p of CHROME_CANDIDATES) {
    try { if (p && fs.existsSync(p)) return p; } catch { /* keep looking */ }
  }
  return null;
}

// ── CDP helpers ────────────────────────────────────────────────────────────
// The HTTP endpoints (/json/list, /json/close, /json/new) are enough here —
// no WebSocket needed, so no dependency.
function cdp(pathname, { method = 'GET', timeout = 1500 } = {}) {
  return new Promise((resolve) => {
    const req = http.request(
      { host: '127.0.0.1', port: DEBUG_PORT, path: pathname, method, timeout },
      (res) => {
        let body = '';
        res.on('data', (c) => { body += c; });
        res.on('end', () => {
          try { resolve(JSON.parse(body)); } catch { resolve(body); }
        });
      });
    req.on('error', () => resolve(null));
    req.on('timeout', () => { req.destroy(); resolve(null); });
    req.end();
  });
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function waitForCdp(ms = 10000) {
  const deadline = Date.now() + ms;
  while (Date.now() < deadline) {
    const v = await cdp('/json/version');
    if (v && typeof v === 'object') return true;
    await sleep(250);
  }
  return false;
}

// Every page target already showing the dashboard. Matched on ORIGIN, so the
// dashboard's own sub-routes (/decks/architecture, /pipeline) count too and a
// stale one doesn't survive the sweep.
async function dashboardTabs(url) {
  const origin = new URL(url).origin;
  const list = await cdp('/json/list');
  if (!Array.isArray(list)) return [];
  return list.filter((t) => t.type === 'page' && typeof t.url === 'string'
                            && t.url.startsWith(origin));
}

async function closeStaleTabs(url) {
  const tabs = await dashboardTabs(url);
  for (const t of tabs) await cdp(`/json/close/${t.id}`);
  return tabs.length;
}

// The launcher starts this alongside the server, so the server may not be
// listening yet. Opening first would land the tab on a connection error.
function waitForServer(url, ms = 15000) {
  const { hostname, port } = new URL(url);
  const deadline = Date.now() + ms;
  return new Promise((resolve) => {
    const attempt = () => {
      const sock = net.connect({ host: hostname, port: Number(port) || 80 });
      const done = (ok) => {
        sock.destroy();
        if (ok) return resolve(true);
        if (Date.now() > deadline) return resolve(false);
        setTimeout(attempt, 250);
      };
      sock.setTimeout(1000);
      sock.on('connect', () => done(true));
      sock.on('error', () => done(false));
      sock.on('timeout', () => done(false));
    };
    attempt();
  });
}

// ── main ───────────────────────────────────────────────────────────────────
export async function openDashboard(url) {
  const exe = chromeExe();
  await waitForServer(url);

  // Already running with the debug port: close what's there, open one fresh tab.
  if (await waitForCdp(400)) {
    const closed = await closeStaleTabs(url);
    await cdp(`/json/new?${encodeURIComponent(url)}`, { method: 'PUT' });
    return { launched: false, closed };
  }

  if (!exe) {
    // No Chrome on this machine — fall back to the OS default browser rather
    // than failing the launch. Stale tabs then keep the overlay behaviour.
    spawn('cmd', ['/c', 'start', '', url], { detached: true, stdio: 'ignore' }).unref();
    return { launched: true, closed: 0, fallback: 'default-browser' };
  }

  fs.mkdirSync(PROFILE_DIR, { recursive: true });
  spawn(exe, [
    `--remote-debugging-port=${DEBUG_PORT}`,
    `--user-data-dir=${PROFILE_DIR}`,
    '--no-first-run',
    '--no-default-browser-check',
    `--app=${url}`,                 // chrome-app window: no omnibox, reads as an app
  ], { detached: true, stdio: 'ignore' }).unref();

  await waitForCdp(10000);
  return { launched: true, closed: 0 };
}

const invokedDirectly = process.argv[1]
  && path.resolve(process.argv[1]) === path.resolve(fileURLToPath(import.meta.url));

if (invokedDirectly) {
  const url = process.argv[2] || 'http://localhost:4600';
  openDashboard(url).then((r) => {
    if (r.fallback) console.log(`Opened ${url} in the default browser (Chrome not found).`);
    else if (r.launched) console.log(`Launched dashboard Chrome profile -> ${url}`);
    else console.log(`Reused dashboard Chrome profile -> ${url} (closed ${r.closed} stale tab(s))`);
    process.exit(0);
  });
}
