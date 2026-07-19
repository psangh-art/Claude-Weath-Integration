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
import { evaluate, getClient } from '../src/connection.js';
import { ensureAutosaveDisabled } from '../src/core/ui.js';

const CDP_HOST = 'localhost';
const CDP_PORT = 9222;

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

const NOT_RUNNING =
  'TradingView Desktop is not reachable on the debugging port. Start it with '
  + '--remote-debugging-port=9222 (the pipeline launcher does this) and try again.';

/**
 * Bring `chartId`'s layout up in TradingView Desktop, on `symbol` if given.
 * Returns {ok, url, ...} — never throws for the "app isn't running" case, which is
 * expected and needs a readable message rather than a stack trace.
 */
export async function openChart({ chartId, symbol, layout }) {
  if (!chartId && !symbol) return { ok: false, error: 'Nothing to open: no chartId or symbol.' };
  if (!(await tvAvailable())) return { ok: false, error: NOT_RUNNING, needsTradingView: true };

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

    return { ok: true, url, layout: layout || null, symbol: symbol || null };
  } catch (e) {
    return { ok: false, error: String(e && e.message || e) };
  }
}
