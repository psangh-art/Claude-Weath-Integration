# Claude-Weath-Integration — Notes for Claude Code

This file is read automatically by Claude Code at the start of a session in this repo.
It tracks feedback from reviewing pipeline output (in claude.ai) so fixes aren't lost
between runs. Update this file (don't just fix and forget) whenever a review surfaces
something worth remembering.

## Confirmed working (do not regress)

- **One image per chart, not one image per layout.** `pane.js` + `crop_panes.py` crop
  individual panes out of a full-layout screenshot; `build_layout_excel.py` produces one
  row per chart with its own image. This fixed a real problem: multi-pane grid
  screenshots made per-symbol price-axis text too small to read reliably. Do not go back
  to shipping one shared image per multi-symbol layout.
- **Device scale factor on capture.** `capture.js` uses
  `Emulation.setDeviceMetricsOverride` with a `scale > 1` before capturing, specifically
  to sharpen axis-label text without changing CSS layout. Keep this — 1x captures were
  the original resolution complaint.
- **`layoutSwitch()` verification loop.** Navigates via direct URL (not
  `loadChartFromServer(id)`, which silently no-ops on numeric ids), dismisses the
  "unsaved changes" dialog, and polls up to 6x/2.5s to confirm the layout actually
  loaded before returning success. This is what fixed the earlier "SWITCH FAILED"
  failures (CDP connection surviving navigation but the chart not actually having
  changed). Don't replace this with a single fixed `sleep()`.

## Open items / things to verify on the next export run

- [ ] Confirm cropped single-chart images are legible at 100% zoom for price-axis
      numbers specifically (not just ticker labels) — last review couldn't fully verify
      this due to a rendering issue on the reviewer's end, not a confirmed pass.
- [ ] Spot-check that `crop_panes.py`'s pixel coordinates (from `listWithRects()` in
      `pane.js`) are being scaled correctly when `capture.js`'s `deviceScaleFactor` > 1 —
      DOM rects are in CSS px, but the screenshot is now higher pixel density, so crop
      boxes need to be multiplied by `scale` or crops will be misaligned/offset.
- [ ] Any layout with a "Test" entry (e.g. `layout_04_Test.png` seen in an earlier
      export) should probably be filtered out of the manifest before it reaches
      `build_layout_excel.py` — check if `export-layouts-excel.js` excludes non-real
      layouts already.

## How to log new feedback

Append a dated entry under "Open items" or move confirmed fixes up to "Confirmed
working" — whichever a reviewer's notes indicate.
