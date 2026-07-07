import { evaluate, evaluateAsync, getClient } from '../connection.js';

/**
 * Dispatch TradingView's "Reset chart view" shortcut (Alt+R) so the capture always
 * starts from a consistent, fitted view instead of whatever pan/zoom state the chart
 * happened to be left in. This is a view-only change — it's never saved (no save
 * call follows it), and layoutSwitch's own navigation discards any unsaved state when
 * moving to the next layout, so it never persists back to the saved layout.
 */
export async function resetView() {
  const client = await getClient();
  await client.Input.dispatchKeyEvent({ type: 'keyDown', modifiers: 1, key: 'r', code: 'KeyR', windowsVirtualKeyCode: 82 });
  await client.Input.dispatchKeyEvent({ type: 'keyUp', key: 'r', code: 'KeyR', windowsVirtualKeyCode: 82 });
  return { success: true, action: 'reset_view' };
}

export async function layoutList() {
  const layouts = await evaluateAsync(`
    new Promise(function(resolve) {
      try {
        window.TradingViewApi.getSavedCharts(function(charts) {
          if (!charts || !Array.isArray(charts)) { resolve({layouts: [], source: 'internal_api', error: 'getSavedCharts returned no data'}); return; }
          var result = charts.map(function(c) { return { id: c.id || c.chartId || null, name: c.name || c.title || 'Untitled', symbol: c.symbol || null, resolution: c.resolution || null, modified: c.timestamp || c.modified || null }; });
          resolve({layouts: result, source: 'internal_api'});
        });
        setTimeout(function() { resolve({layouts: [], source: 'internal_api', error: 'getSavedCharts timed out'}); }, 5000);
      } catch(e) { resolve({layouts: [], source: 'internal_api', error: e.message}); }
    })
  `);
  return { success: true, layout_count: layouts?.layouts?.length || 0, source: layouts?.source, layouts: layouts?.layouts || [], error: layouts?.error };
}

export async function layoutSwitch({ name }) {
  const escaped = JSON.stringify(name);
  const match = await evaluateAsync(`
    new Promise(function(resolve) {
      try {
        var target = ${escaped};
        var charts = (window.TradingViewApi._loadChartService._state.value().chartList) || [];
        var found = null;
        if (/^\\d+$/.test(target)) {
          for (var i = 0; i < charts.length; i++) { if (String(charts[i].id) === target) { found = charts[i]; break; } }
        }
        if (!found) {
          for (var i = 0; i < charts.length; i++) { var cname = charts[i].name || ''; if (cname === target || cname.toLowerCase() === target.toLowerCase()) { found = charts[i]; break; } }
        }
        if (!found) {
          for (var j = 0; j < charts.length; j++) { var cn = (charts[j].name || '').toLowerCase(); if (cn.indexOf(target.toLowerCase()) !== -1) { found = charts[j]; break; } }
        }
        if (!found) { resolve({success: false, error: 'Layout "' + target + '" not found.'}); return; }
        if (!found.url) { resolve({success: false, error: 'Layout "' + (found.name || target) + '" has no chart url to navigate to.'}); return; }
        resolve({success: true, id: found.id, url: found.url, name: found.name});
      } catch(e) { resolve({success: false, error: e.message}); }
    })
  `);
  if (!match?.success) throw new Error(match?.error || 'Unknown error switching layout');

  // loadChartFromServer(id) silently no-ops when given the numeric layout id instead of
  // the alphanumeric chart url slug, so navigate directly instead of relying on it.
  const targetUrl = `https://www.tradingview.com/chart/${match.url}/`;
  await evaluate(`window.location.assign(${JSON.stringify(targetUrl)})`);

  // Handle "unsaved changes" confirmation dialog, which can block the navigation
  await new Promise(r => setTimeout(r, 500));
  const dismissed = await evaluate(`
    (function() {
      var btns = document.querySelectorAll('button');
      for (var i = 0; i < btns.length; i++) {
        var text = btns[i].textContent.trim();
        if (/open anyway|don't save|discard/i.test(text)) {
          btns[i].click();
          return true;
        }
      }
      return false;
    })()
  `);
  if (dismissed) {
    await new Promise(r => setTimeout(r, 1000));
    await evaluate(`window.location.assign(${JSON.stringify(targetUrl)})`);
  }

  // Poll for the navigation to actually land instead of trusting a single fixed wait
  let verified = null;
  for (let attempt = 0; attempt < 6; attempt++) {
    await new Promise(r => setTimeout(r, 2500));
    verified = await evaluate(`
      (function() {
        try { return window.TradingViewApi._saveChartService.layoutId(); } catch(e) { return null; }
      })()
    `);
    if (verified === match.url) break;
  }

  if (verified !== match.url) {
    throw new Error(`Navigated to layout "${match.name || name}" but verification failed (expected chart id "${match.url}", got "${verified}"). The chart may still be loading — retry layout_switch.`);
  }

  return {
    success: true,
    layout: match.name || name,
    layout_id: match.id,
    chart_id: match.url,
    source: 'window_location_assign',
    action: 'switched',
    unsaved_dialog_dismissed: dismissed,
  };
}
