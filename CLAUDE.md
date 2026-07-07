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

## Resolved (2026-07-07)

- **Legibility confirmed.** Cropped chart images were extracted from a real built
  workbook and viewed directly — price-axis numbers (not just ticker labels) are
  crisp and fully readable at native resolution. Confirmed pass, not just a claim.
- **Pixel-scaling was already correct, not a bug.** `export-layouts-excel.js` scales
  each pane's CSS-px rect by `CAPTURE_SCALE` before building the crop list:
  `x: p.rect.x * CAPTURE_SCALE, y: p.rect.y * CAPTURE_SCALE, ...` (same for
  width/height). `crop_panes.py` receives already-scaled pixel coordinates, so crops
  land correctly even with `scale > 1`. No fix needed — verified by reading the code.
- **"Test" layout now filtered.** `export-layouts-excel.js` drops any layout whose
  name matches `/^test$/i` (case-insensitive, exact match) before it reaches the
  manifest/workbook. Confirmed with the user this should be excluded going forward.

## Open items / things to verify on the next export run

(none currently — add new items here as reviews surface them)

## How to log new feedback

Append a dated entry under "Open items" or move confirmed fixes up to "Confirmed
working" — whichever a reviewer's notes indicate.
