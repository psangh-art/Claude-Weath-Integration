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
