import { evaluate, evaluateAsync } from '../connection.js';
import { spawnSync } from 'child_process';

// NOTE: a resetView() helper (Alt+R "Reset chart view" per pane before capture)
// used to live here. Removed 2026-07-13 by user decision: TradingView's reset
// snaps to its default zoom, not the user's saved wide channel view, so it zoomed
// captures in past the drawn trendlines. Capture the saved layout as-is — the
// user sizes charts in TradingView. Do not reintroduce a pre-capture view reset.

/**
 * Force TradingView's layout auto-save OFF before any capture run (user policy
 * 2026-07-13): with auto-save ON, TradingView silently persisted capture-time view
 * changes (the old Alt+R reset) back into the user's saved layouts — no "unsaved
 * changes" dialog ever appeared to stop it, and every chart's saved zoom was
 * overwritten. The capture flow no longer changes views at all, but this stays as
 * a hard guard: a run must never be able to write back to saved layouts.
 *
 * API located by live CDP probe (2026-07-13): _saveChartService.autoSaveEnabled()
 * returns a WatchedValue; setAutoSaveEnabled(false) is the toggle the save-menu
 * checkbox uses. Throws if auto-save is on and the toggle fails to stick. If the
 * API can't be found at all (TradingView update), returns found:false so the
 * caller can warn loudly — the run itself makes no view changes, so proceeding is
 * safe, but the guard being blind is worth surfacing every run.
 */
export async function ensureAutosaveDisabled() {
  const result = await evaluate(`
    (function() {
      try {
        var s = window.TradingViewApi._saveChartService;
        if (!s || typeof s.autoSaveEnabled !== 'function' || typeof s.setAutoSaveEnabled !== 'function') {
          return { found: false, error: 'autoSaveEnabled/setAutoSaveEnabled not present on _saveChartService' };
        }
        var read = function() {
          var wv = s.autoSaveEnabled();
          return (wv && typeof wv.value === 'function') ? !!wv.value() : !!wv;
        };
        var before = read();
        if (before) s.setAutoSaveEnabled(false);
        return { found: true, wasEnabled: before, nowEnabled: read() };
      } catch (e) { return { found: false, error: e.message }; }
    })()
  `);

  if (result?.found && result.nowEnabled) {
    throw new Error('TradingView auto-save is ON and could not be disabled — aborting so the run cannot write back to saved layouts. Turn it off manually (save-menu dropdown) and re-run.');
  }
  return {
    success: true,
    found: !!result?.found,
    wasEnabled: result?.wasEnabled ?? null,
    error: result?.error || null,
  };
}

/**
 * Ensure the TradingView window is maximized before any chart capture (user policy
 * 2026-07-13): pane size — and therefore candle/axis legibility AND the visible
 * date range (TradingView keeps bar spacing, so a narrower pane shows less history)
 * — depends directly on window size. The user sizes his charts with the window
 * maximized, so captures must run maximized too or they won't match what he saved.
 *
 * TradingView Desktop is Electron and does NOT implement CDP's
 * Browser.getWindowForTarget/setWindowBounds (probed live 2026-07-13), so the
 * check is done via CDP (innerWidth vs screen.availWidth) and the fix natively:
 * ShowWindowAsync(hwnd, SW_MAXIMIZE) on TradingView.exe's main window via
 * PowerShell — same no-focus-needed user32 approach as drive_open_dialog.ps1.
 * Throws if the window still isn't maximized after the attempt: capturing at the
 * wrong size silently changes what's in frame, which is worse than not running.
 */
export async function ensureWindowMaximized() {
  const readDims = () => evaluate(
    `({ w: window.innerWidth, h: window.innerHeight, aw: screen.availWidth, ah: screen.availHeight })`
  );
  // Width within 8px of the available screen width is "maximized" — height is left
  // slack (title bar / OS chrome sit outside innerHeight and vary by machine).
  const isMaximized = d => d && d.w >= d.aw - 8;

  const before = await readDims();
  if (isMaximized(before)) {
    return { success: true, wasMaximized: true, width: before.w, height: before.h };
  }

  const ps = spawnSync('powershell', ['-NoProfile', '-Command', `
    Add-Type -Namespace Native -Name Win -MemberDefinition '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
    $hit = $false
    Get-Process TradingView -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 } | ForEach-Object {
      [Native.Win]::ShowWindowAsync($_.MainWindowHandle, 3) | Out-Null
      $hit = $true
    }
    if (-not $hit) { exit 1 }
  `], { timeout: 15000 });
  if (ps.status !== 0) {
    throw new Error(`TradingView window is not maximized (${before.w}x${before.h}, screen ${before.aw}x${before.ah}) and no TradingView.exe window could be found to maximize. Maximize it manually and re-run.`);
  }

  await new Promise(r => setTimeout(r, 1500));
  const after = await readDims();
  if (!isMaximized(after)) {
    throw new Error(`TradingView window could not be maximized (still ${after.w}x${after.h}, screen ${after.aw}x${after.ah}). Captures at the wrong size change what's in frame — maximize the window manually and re-run.`);
  }
  return { success: true, wasMaximized: false, width: after.w, height: after.h };
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
