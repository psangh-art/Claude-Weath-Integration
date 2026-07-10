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

## Critical fix (2026-07-10): scripts never actually exited

`export-layouts-excel.js`, `export-indicator-values.js`, and `export-alerts.js` each
called `main().catch(...)` with no explicit exit. `src/connection.js` keeps a
module-level CDP WebSocket client open, which holds Node's event loop alive forever —
so none of these scripts ever actually terminated on their own after finishing; the
terminal just sat there after the final "Done." line. Running any of them standalone
via the `.bat` files masked this (the user just closes the window), but it's fatal for
`run_full_pipeline.js`, which chains steps with `spawnSync` — that call blocks until
the child process *fully exits*, so the whole pipeline hung indefinitely after chart
capture finished, never reaching the OCR/master-sheet steps. Fixed by forcing
`process.exit(process.exitCode || 0)` in a `.finally()` after `main()` settles, in all
three scripts. **Any new standalone script that imports `src/connection.js` needs the
same explicit exit** — don't assume Node will terminate on its own once `main()`
resolves.

## Full pipeline (added 2026-07-09): Google Finance export + Stocks Buy Strategy update

`npm run pipeline` (or `Run Full Pipeline.bat` at the repo root) runs the whole chain
end to end, beyond just the chart/indicator/alert export:

- **Charts sheet** now also has a **Google Finance Ticker** and **Google Finance
  Formula** column per chart (`ticker_normalize.py` — handles the trailing-dot
  artifacts, `BT.A`-style class suffixes, and commodity name mapping documented in
  `Claude_Code_Handoff_Instructions.md` section 5-7).
- **`channel_detect.py`** OCRs each chart's price axis and detects the two
  channel-boundary lines by colour (port of the algorithm in the handoff doc section
  4) — cross-validates and rejects rather than guessing. Requires the **Tesseract OCR
  binary** on PATH (not just the `pytesseract` pip package); this is a one-time,
  *interactive* install (https://github.com/UB-Mannheim/tesseract/wiki) that cannot be
  scripted from an unattended run because the installer needs a UAC prompt. Until it's
  installed, the pipeline still completes the chart export and tells you exactly
  what's missing rather than failing outright.
- **`update_master_sheet.py`** writes `Alert Low` (= lower boundary × 1.05),
  `Alert High`, and `Chart = Yes` (colour rule applied atomically — value and
  font/fill change together, never separately) into `~/Downloads/Stocks_Buy_Strategy.xlsx`,
  then updates the Coverage Tracker table and inserts a new dated session entry in
  `~/Downloads/Feedback_for_Claude_Code.md` (right after the tracker, keeping it
  pinned at the top of the file — a real bug from an earlier attempt at this insertion
  put the new entry above the tracker instead, now fixed and covered by a manual test
  before trusting it against the real file). Rows with `Alert Low Source = "Manual"`
  are never touched; a re-read within 3% of the existing Alert Low is treated as noise
  and left alone.
- `export-layouts-excel.js` itself now retries each layout's switch+capture up to 3x
  with a health-check/backoff between attempts if TradingView's CDP connection drops
  transiently, and falls back to reusing a same-run-window screenshot from a previous
  capture rather than leaving a row permanently blank if every retry fails.
- **`verify_pipeline.py`** runs automatically as the last step (`python
  scripts/verify_pipeline.py --live-alert-check` to re-run standalone without
  recapturing) and writes `logs/verify_<timestamp>.md` / `logs/latest-verify.md`: a
  per-layout/per-chart capture status, how many tickers got a real upper/lower channel
  read vs rejected (and why), an alert count cross-checked against TradingView live,
  and confirmation that `tradingview_layouts.xlsx`'s sheet row counts actually match
  what was captured. Runs even if an earlier step stopped partway through — it reports
  honestly on whatever manifests do or don't exist rather than requiring a fully clean
  run.

**`Claude_Code_Handoff_Instructions.md`** (kept in the user's Downloads, not this repo)
is the full behavioural spec this all implements — schema, colour rules, ticker
normalization, the exact channel-reading algorithm, and the "silence over guessing"
philosophy (an unresolved row is fine; a wrong number in a live trading sheet is not).
Where it disagrees with this file on chart-interpretation or spreadsheet-update
behavior, it wins — update this file to match rather than the reverse.

**Repo consolidation (2026-07-09):** this repo is now the *only* one for chart
extraction / TradingView-to-Excel work — `tradingview-mcp`'s duplicate copy of this
pipeline was reverted back to its prior state, since it's a separate 68-tool MCP
server repo unrelated to this specific workflow.

## Open items / things to verify on the next export run

(none currently — add new items here as reviews surface them)

## Resolved (2026-07-10)

- **`verify_pipeline.py` crashed printing its own report to console.** Windows'
  default console codepage (cp1252) can't encode the ✅/⚠️/⚪ emoji used in the
  report text, so the final `print(report_text)` raised `UnicodeEncodeError` even
  though the report had already been written correctly to
  `logs/latest-verify.md`. Fixed by catching the encode error and falling back to
  writing UTF-8 bytes directly to `sys.stdout.buffer`. The file write path was
  never affected — only the console echo. Confirmed fixed by re-running
  `python scripts/verify_pipeline.py --live-alert-check` standalone after the fix.

## Resolved (2026-07-10, part 2): 'Stocks Buy Strategy' tab renamed to 'Investments'; Alert Low highlighting

- **Sheet tab rename.** The worksheet tab in `Stocks_Buy_Strategy.xlsx` previously named
  `Stocks Buy Strategy` is now named **`Investments`** — tab only, per explicit user
  confirmation; the workbook **filename** stays `Stocks_Buy_Strategy.xlsx`.
  `SHEET_NAME` in `update_master_sheet.py` was updated to `'Investments'` to match.
  **If you add any new code that opens this workbook and looks up a sheet by name,
  use `'Investments'`, not `'Stocks Buy Strategy'`.**
- **Renaming a sheet in openpyxl does NOT rewrite formulas elsewhere in the workbook
  that reference it by quoted name** (unlike renaming a tab via the Excel UI, which
  does fix those up). ~36 live formula cells — VLOOKUPs in `Stocks of Interest`
  column F and `History` column K — referenced `'Stocks Buy Strategy'!$C:$I` by
  literal string and would have silently become `#REF!` errors on next open if left
  alone. Caught by scanning all formula cells across every sheet for the old quoted
  reference before considering the rename safe, and rewriting each one to
  `'Investments'!$C:$I`. **Any future sheet rename in this workbook must repeat this
  scan-and-rewrite step** — don't assume `ws.title = ...` is sufficient.
  `scripts/fix_relx_history_2026-07-10.py` (a historical one-off, already executed)
  still contains the old hardcoded `'Stocks Buy Strategy'!` string in its source —
  left as-is since it's a dated record of a completed run, not something meant to be
  re-run; do not reuse it as a template without updating that reference first.
- **Alert Low green highlight.** The `Alert Low` cells (column E) for the four rows
  under "🟢 AT LOWER BOUNDARY — within 5% of alert low (Highest priority)" in
  `Stocks of Interest` (rows 5–8: Glencore, Beazley, Auto Trader, Rio Tinto) are now
  filled green (`FFC6EFCE` fill / `FF276221` bold font — the same "good/highest
  priority" green already used for `Chart = Yes` in `update_master_sheet.py` and for
  the Proximity column in this same table). This section is manually maintained (not
  rebuilt by any script), so the highlight persists until the section's row range or
  membership changes by hand — it isn't reapplied automatically by the pipeline.

## How to log new feedback

Append a dated entry under "Open items" or move confirmed fixes up to "Confirmed
working" — whichever a reviewer's notes indicate.
