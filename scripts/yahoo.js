// Yahoo Finance chart API — the one place this repo talks to it (added 2026-07-19).
//
// Two dashboard features use it and each had grown its own copy of the URL build
// and the last/previous-close extraction: the Intelligence screen's index widgets
// and the Watchlist's "Live prices" button. Same endpoint, same quirks, so they now
// share one reader.
//
// Why Yahoo at all: Stooq's CSV endpoint returns a bot/JS-challenge page (verified
// dead 2026-07-17), and the browser can't call Yahoo cross-origin — so these fetches
// have to happen server-side.

const CHART_URL = 'https://query1.finance.yahoo.com/v8/finance/chart/';

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

/** Round to 2dp, passing null through — every figure here is money or a percent. */
export const r2 = (n) => (n == null || !isFinite(n) ? null : Math.round(n * 100) / 100);

/**
 * One symbol's chart result, normalised.
 * Returns {price, prev, change, change_pct, currency, as_of, series} and THROWS on
 * a symbol Yahoo can't resolve — callers decide whether that's an error row or a
 * thrown failure.
 */
export async function fetchQuote(symbol, { range = '5d', interval = '1d', timeoutMs = 8000 } = {}) {
  const url = `${CHART_URL}${encodeURIComponent(symbol)}?range=${range}&interval=${interval}`;
  const j = await fetchJson(url, timeoutMs);
  const result = j && j.chart && j.chart.result && j.chart.result[0];
  if (!result) throw new Error('no data (check symbol)');

  const meta = result.meta || {};
  const closes = (((result.indicators || {}).quote || [])[0] || {}).close || [];
  const series = closes.filter(x => typeof x === 'number').slice(-30);
  // meta.regularMarketPrice is the live print; the close series is the fallback for
  // a symbol (some indices) that omits it outside market hours.
  const price = typeof meta.regularMarketPrice === 'number' ? meta.regularMarketPrice
              : (series.length ? series[series.length - 1] : null);
  const prev = typeof meta.chartPreviousClose === 'number' ? meta.chartPreviousClose
             : (typeof meta.previousClose === 'number' ? meta.previousClose : null);
  if (price == null) throw new Error('no price');

  const change = (prev != null) ? price - prev : null;
  return {
    price,
    prev,
    change,
    change_pct: prev ? (change / prev) * 100 : null,
    currency: meta.currency || null,
    as_of: (result.timestamp && result.timestamp.length)
      ? new Date(result.timestamp[result.timestamp.length - 1] * 1000).toISOString().slice(0, 10)
      : null,
    series,
  };
}
