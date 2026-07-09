import { evaluate } from '../connection.js';

const CWC = 'window.TradingViewApi._chartWidgetCollection';

/**
 * List all panes in the current layout with their symbol metadata and DOM bounding
 * rect (CSS px, relative to viewport) — used to crop individual chart images out of
 * one full-layout screenshot instead of maximizing/re-capturing each pane separately.
 */
export async function listWithRects() {
  const panes = await evaluate(`
    (function() {
      var cwc = ${CWC};
      var all = cwc.getAll();
      var out = [];
      for (var i = 0; i < all.length; i++) {
        try {
          var c = all[i];
          var model = c.model ? c.model() : null;
          var mainSeries = model ? model.mainSeries() : null;
          var info = mainSeries && mainSeries.symbolInfo ? mainSeries.symbolInfo() : null;
          var rect = c._mainDiv ? c._mainDiv.getBoundingClientRect() : null;
          out.push({
            index: i,
            symbol: mainSeries ? mainSeries.symbol() : null,
            ticker: info ? info.name : null,
            description: info ? info.description : null,
            rect: rect ? { x: rect.x, y: rect.y, width: rect.width, height: rect.height } : null,
          });
        } catch (e) { out.push({ index: i, error: e.message }); }
      }
      return out;
    })()
  `);
  return { success: true, panes };
}

/**
 * Poll listWithRects() until two consecutive reads agree on every pane's symbol
 * (index, symbol, ticker) instead of trusting a single read right after
 * layoutSwitch(). layoutSwitch()'s own poll only confirms the NAVIGATION landed
 * (layoutId matches) — it says nothing about whether every individual pane's
 * chart widget has finished initializing its own mainSeries(). On layouts with
 * more panes (8-pane grids especially), some panes can still be mid-init at that
 * point and report a stale/shared symbol from whatever was previously loaded —
 * this was confirmed independently (chart_import_findings.md, 2026-07-09): a
 * genuinely 8-distinct-chart layout produced only 3 distinct ticker labels across
 * 3 consecutive export runs, with the affected panes' images verified correct via
 * file hash (so it's a metadata-timing bug, not a rendering one). Bounded so a
 * layout that never stabilizes doesn't hang the export — falls back to the last
 * read on timeout.
 */
export async function waitForPanesToLoad({ maxAttempts = 12, intervalMs = 400 } = {}) {
  let prev = null;
  let result = null;
  for (let attempt = 0; attempt < maxAttempts; attempt++) {
    result = await listWithRects();
    const panes = result.panes || [];
    const fingerprint = JSON.stringify(panes.map(p => [p.index, p.symbol, p.ticker]));
    const allHaveSymbol = panes.length > 0 && panes.every(p => p.symbol);
    if (allHaveSymbol && fingerprint === prev) {
      return result;
    }
    prev = fingerprint;
    await new Promise(r => setTimeout(r, intervalMs));
  }
  return result;
}

/**
 * Focus a specific pane by index (clicks its main div, same as a user click) —
 * "Reset chart view" (Alt+R) only resets the currently focused/active pane, not
 * every pane in a grid layout, so each pane must be focused individually before
 * resetting it.
 */
export async function focus({ index }) {
  const idx = Number(index);
  const result = await evaluate(`
    (function() {
      var cwc = ${CWC};
      var all = cwc.getAll();
      if (${idx} >= all.length) return { error: 'Pane index ' + ${idx} + ' out of range (have ' + all.length + ' panes)' };
      var chart = all[${idx}];
      if (chart._mainDiv) chart._mainDiv.click();
      return { focused: ${idx}, total: all.length };
    })()
  `);
  if (result?.error) throw new Error(result.error);
  return { success: true, focused_index: result.focused, total_panes: result.total };
}
