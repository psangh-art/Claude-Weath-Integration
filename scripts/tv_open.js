// Drive the running TradingView Desktop app to a specific saved layout/symbol.
//
// Used by the Investment Dashboard's Activity widget: clicking a ticker under
// "Charts to mark up" should put that chart in front of you ready to draw on
// (user request 2026-07-19), not open a read-only tab in the browser.
//
// This is the only dashboard code that talks to TradingView. Two project rules it
// has to honour, both learned the hard way (see CLAUDE.md):
//   * ensureAutosaveDisabled() FIRST. TradingView's layout auto-save silently
//     persists whatever view state it finds back into the saved layout — that once
//     overwrote every hand-zoomed chart and had to be redrawn by hand.
//   * NO view reset/refit of any kind. The saved layout IS the view to show.
// Navigation is a plain window.location.assign to the layout's chart URL, the same
// approach layoutSwitch() uses, because loadChartFromServer(id) silently no-ops on
// a numeric layout id.
import { spawn, spawnSync } from 'child_process';
import path from 'path';
import { fileURLToPath } from 'url';
import { evaluate, getClient } from '../src/connection.js';
import { ensureAutosaveDisabled } from '../src/core/ui.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const CDP_HOST = 'localhost';
const CDP_PORT = 9222;
const LAUNCH_BAT = path.join(__dirname, 'launch_tv_debug.bat');
const FOREGROUND_PS1 = path.join(__dirname, 'tv_foreground.ps1');

/**
 * Raise the TradingView Desktop window to the foreground (user request 2026-07-20).
 * Best-effort: navigation has already succeeded, so a failure here (window gone,
 * PowerShell blocked) must not fail the open — just don't present the window.
 */
function bringToForeground() {
  try {
    spawnSync('powershell',
      ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', FOREGROUND_PS1],
      { timeout: 10000, stdio: 'ignore', windowsHide: true });
  } catch { /* presentation is best-effort */ }
}

/** Is TradingView Desktop reachable over CDP right now? */
export async function tvAvailable() {
  try {
    const res = await fetch(`http://${CDP_HOST}:${CDP_PORT}/json/version`,
      { signal: AbortSignal.timeout(1500) });
    return res.ok;
  } catch {
    return false;
  }
}

/**
 * Bring TradingView Desktop up WITH the debugging port, on demand. Reuses the
 * pipeline's proven launcher (launch_tv_debug.bat): it taskkills any running
 * TradingView and reopens it with --remote-debugging-port=9222. Only called when
 * CDP is already DOWN, so an already-debug-ready TradingView is never disturbed;
 * and because layouts live server-side, a relaunch loads them fresh without
 * touching the user's charts. Polls until CDP answers or the timeout is hit.
 */
async function launchTradingView(timeoutMs = 30000) {
  try {
    const child = spawn('cmd.exe', ['/c', LAUNCH_BAT, String(CDP_PORT)],
      { detached: true, stdio: 'ignore', windowsHide: true });
    child.unref();
  } catch {
    return false;
  }
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    await new Promise(r => setTimeout(r, 1200));
    if (await tvAvailable()) {
      // CDP answers, but the chart page may still be mounting — let it settle
      // before ensureAutosaveDisabled reads the chart service.
      await new Promise(r => setTimeout(r, 3000));
      return true;
    }
  }
  return false;
}

const NOT_RUNNING =
  'Could not start TradingView Desktop with the debugging port. Check it is '
  + 'installed, or launch it manually with scripts\\launch_tv_debug.bat, then retry.';

/**
 * Bring `chartId`'s layout up in TradingView Desktop, on `symbol` if given.
 * Returns {ok, url, ...} — never throws for the "app isn't running" case, which is
 * expected and needs a readable message rather than a stack trace.
 */
export async function openChart({ chartId, symbol, layout }) {
  if (!chartId && !symbol) return { ok: false, error: 'Nothing to open: no chartId or symbol.' };
  // If TradingView Desktop isn't on the debug port, start it on demand rather than
  // just erroring — selecting a layout should bring the app up ready to draw on.
  if (!(await tvAvailable()) && !(await launchTradingView())) {
    return { ok: false, error: NOT_RUNNING, needsTradingView: true };
  }

  try {
    await getClient();
    // Must run before any navigation — see the note at the top of this file.
    await ensureAutosaveDisabled();

    const url = chartId
      ? `https://www.tradingview.com/chart/${chartId}/`
        + (symbol ? `?symbol=${encodeURIComponent(symbol)}` : '')
      : `https://www.tradingview.com/chart/?symbol=${encodeURIComponent(symbol)}`;

    await evaluate(`window.location.assign(${JSON.stringify(url)})`);

    // The "unsaved changes" dialog blocks the navigation if it appears.
    await new Promise(r => setTimeout(r, 500));
    const dismissed = await evaluate(`
      (function() {
        var btns = document.querySelectorAll('button');
        for (var i = 0; i < btns.length; i++) {
          if (/open anyway|don't save|discard/i.test(btns[i].textContent.trim())) {
            btns[i].click();
            return true;
          }
        }
        return false;
      })()
    `);
    if (dismissed) {
      await new Promise(r => setTimeout(r, 800));
      await evaluate(`window.location.assign(${JSON.stringify(url)})`);
    }

    // Present the window — otherwise the layout loads behind the browser/dashboard
    // and the user has to alt-tab to the chart they just asked to draw on.
    bringToForeground();

    return { ok: true, url, layout: layout || null, symbol: symbol || null };
  } catch (e) {
    return { ok: false, error: String(e && e.message || e) };
  }
}
