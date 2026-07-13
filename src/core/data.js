import { evaluate } from '../connection.js';

/**
 * Read current indicator values from the Data Window for the currently FOCUSED
 * pane only (focus-scoped — see pane.focus()). Marker/label-style indicators (e.g. "Dividend yield %")
 * don't populate the Data Window at all and are simply omitted, not an error.
 */
/**
 * Read the latest bar (time/OHLCV) for the currently FOCUSED pane's main series —
 * same focus-scoping as getStudyValues(). This is the live "current price" used to
 * decide single-trendline direction (above/below price -> Alert High/Low) and to
 * stamp how fresh a chart's read is; it does NOT touch Stocks_Buy_Strategy.xlsx's
 * own GOOGLEFINANCE-formula price column, which only calculates in real Google
 * Sheets, not this pipeline.
 */
export async function getLastPrice() {
  const data = await evaluate(`
    (function() {
      var chart = window.TradingViewApi._activeChartWidgetWV.value()._chartWidget;
      var bars = chart.model().mainSeries().bars();
      if (!bars || typeof bars.lastIndex !== 'function') return null;
      var last = bars.valueAt(bars.lastIndex());
      if (!last) return null;
      return { time: last[0], open: last[1], high: last[2], low: last[3], close: last[4], volume: last[5] || 0 };
    })()
  `);
  return data;
}

export async function getStudyValues() {
  const data = await evaluate(`
    (function() {
      var chart = window.TradingViewApi._activeChartWidgetWV.value()._chartWidget;
      var model = chart.model();
      var sources = model.model().dataSources();
      var results = [];
      for (var si = 0; si < sources.length; si++) {
        var s = sources[si];
        if (!s.metaInfo) continue;
        try {
          var meta = s.metaInfo();
          var name = meta.description || meta.shortDescription || '';
          if (!name) continue;
          var values = {};
          try {
            var dwv = s.dataWindowView();
            if (dwv) {
              var items = dwv.items();
              if (items) {
                for (var i = 0; i < items.length; i++) {
                  var item = items[i];
                  if (item._value && item._value !== '∅' && item._title) values[item._title] = item._value;
                }
              }
            }
          } catch(e) {}
          if (Object.keys(values).length > 0) results.push({ name: name, values: values });
        } catch(e) {}
      }
      return results;
    })()
  `);
  return { success: true, study_count: (data || []).length, studies: data || [] };
}
