# Claude-Weath-Integration — Notes for Claude Code

This file is read automatically by Claude Code at the start of a session in this repo.
It tracks feedback from reviewing pipeline output (in claude.ai) so fixes aren't lost
between runs. Update this file (don't just fix and forget) whenever a review surfaces
something worth remembering.

## Confirmed working (do not regress)

- **NO pre-capture view reset — capture each layout exactly as saved (user decision
  2026-07-13).** The capture loop used to press Alt+R ("Reset chart view") on every
  pane before screenshotting, on the theory that it normalised leftover pan/zoom.
  In reality TradingView's reset snaps to its DEFAULT bar spacing/right-offset, not
  the user's saved wide channel view — it zoomed charts in past their drawn
  trendlines and made whole layouts unreadable (the 2026-07-13 analyst run flagged
  PRU/LGEN/BEZ/SDLF in "FT100 Insurance" as over-zoomed with months of blank future
  space; that was Alt+R's doing, NOT the saved layout's zoom state as first
  suspected). The user sizes each chart in TradingView; `layoutSwitch()` loads the
  saved state fresh from the server, so it's already the view to capture.
  `resetView()`/`waitForResetToSettle()` were deleted from `src/core/ui.js` — do not
  reintroduce any view reset/refit before capture. Per-pane `pane.focus()` clicks
  remain: the Data Window / last-price reads are focus-scoped and clicking doesn't
  change the view.
- **Auto-save guard before every capture run (user policy 2026-07-13).** TradingView's
  layout auto-save silently persisted the old Alt+R-reset views back into the user's
  saved layouts (no "unsaved changes" dialog fires when auto-save is on) — the user
  had to re-zoom every chart by hand. `ui.ensureAutosaveDisabled()` (in
  `src/core/ui.js`) now runs first in both `export-layouts-excel.js` and
  `export-indicator-values.js`: it reads `_saveChartService.autoSaveEnabled()` (a
  WatchedValue; API confirmed by live CDP probe) and calls `setAutoSaveEnabled(false)`
  if on — **aborting the run if the toggle doesn't stick**, and warning loudly if a
  TradingView update ever hides the API. Any new script that switches layouts or
  clicks panes must call this guard first.
- **Window-maximize guard before capture (user policy 2026-07-13).**
  `ui.ensureWindowMaximized()` runs before the capture loop in
  `export-layouts-excel.js`: pane size drives both legibility and the visible date
  range (TradingView keeps bar spacing, so a narrower window shows less history),
  and the user sizes his charts with the window maximized. TradingView Desktop
  (Electron) does NOT implement CDP `Browser.setWindowBounds` (probed live), so the
  check reads `innerWidth` vs `screen.availWidth` over CDP and the fix is native:
  PowerShell `ShowWindowAsync(hwnd, SW_MAXIMIZE)` on TradingView.exe — the same
  no-focus user32 approach as `drive_open_dialog.ps1`. The run ABORTS if the window
  still isn't maximized after the attempt (wrong-size captures silently change
  what's in frame).

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
- **`cleanup_downloads.py`** (added 2026-07-10) runs automatically as the final step
  (`python scripts/cleanup_downloads.py --apply`; omit `--apply` for a dry run) and
  flags redundant files in `~/Downloads` by renaming them with a `Delete ` prefix —
  it **never deletes anything itself**, so the rename is always reversible and the
  actual removal stays a manual, human decision. Flags: (1) old workbook backups
  (`<name>.bak-*`) once a newer backup of the same base file exists — never flags the
  sole/newest backup; (2) Fidelity `TransactionHistory*.csv`/`transactions*.*` files
  older than whatever `fidelity_import_state.json` currently has recorded as
  ingested for that type; (3) numbered duplicate workbook saves (`<stem>_<N>.<ext>`)
  once a newer canonical `<stem>.<ext>` exists. Any file already prefixed `Delete `
  is skipped on the next run, so re-running is always safe/idempotent.

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

## Pipeline App — "Investment Production Centre" (added 2026-07-10, rebuilt 2026-07-12)

`Run Pipeline App.bat` starts `scripts/pipeline_app_server.js` (plain Node `http`,
no new dependency) and opens `http://localhost:4590` — the **Investment Production
Centre**: a control-room "production line" screen. Input feedstock (the required
files, found/loaded) flows through a PowerPoint-style filmstrip of the pipeline
stages (live status, spotlighted active slide, deck-progress rail, per-stage log
drawer), out to an **output bay** linking the built products. Front end is
`scripts/pipeline_app/index.html` (self-contained, theme-aware). **The
`app-developer` agent (`.claude/agents/app-developer.md`) owns this front end** —
route it any UI/presentation change; it's scoped away from pipeline logic.

Server routes beyond `/` + SSE `/events` + `POST /run`: `/files` (preflight status
on load), `/products` (built-product availability + links), `/deck` (in-Chrome
gallery of the review deck), `/deck.pptx` (download / open in PowerPoint),
`/download/spending`, and `/asset?p=` (image proxy for the gallery, **whitelisted to
the repo + Downloads only** — never serve an arbitrary path). `build_review_deck.py`
emits the gallery (`pipeline_app/review_deck.html`, gitignored) + a summary JSON
that feeds the output bay. Products linked: Finance Google Sheet, the review deck
(view + download), `spending_summary.xlsx`, and the architecture deck
(`Financial_Data_Pipeline_Architecture.pptx`, `/architecture.pptx`, added
2026-07-13). Per-stage durations of each run are kept in
`data/stage_timings.json` (last 5 per stage, gitignored); the SSE `hello` /
`run-started` events carry the medians as `timings` so the front end can show a
%-complete bar weighted by how long each stage took in previous runs.

Eight stages, in order:

1. **Pre-flight file check** (`preflight_check.py`) — reports on the input files
   in `~/Downloads` and their DATA AGE. **Rule change 2026-07-13 (user policy):
   the bank/broker exports — Amex (`activity.csv`), Barclays (`data.csv`),
   Fidelity `AccountSummary.csv` and the Fidelity historic export (classified by
   content, `fidelity_file_classifier.py`) — are OPTIONAL.** The pipeline runs
   without them (stage 2 is skipped, TradingView stages still run); only a
   missing master workbook `Stocks_Buy_Strategy.xlsx` halts the run. Each input
   reports an as-of date — file mtime while it sits in Downloads, or the
   ingestion date recorded in `data/ingestion_state.json` (written by
   `consume_input_files.py` just before recycling, gitignored) once consumed —
   and is flagged **stale (red in the app) past 42 days / 6 weeks**, meaning a
   fresh export is needed. A pending Fidelity export never blocks and is never
   flagged.
2. **Fidelity spending-summary build** — runs `spending_summary.py` against the
   files pre-flight found, writing `spending_summary.xlsx`. Skipped (not failed)
   when the exports aren't present; a failure here no longer blocks stages 3-8.
3–8. **TradingView chart capture → OCR → master-sheet update → PowerPoint review
   deck → verification → Downloads cleanup** — these are `run_full_pipeline.js`'s
   six steps (the deck build was added 2026-07-12 as step 4/6, in the always-run
   finally block so a partial run still gets a "what's missing" deck), run
   unmodified as a single child process; the app parses that script's own
   `=== Step N/M: ... ===` console markers (regex is version-agnostic, `\d+/\d+`)
   to report them as separate stages. **If you rename or reorder the step log
   lines in `run_full_pipeline.js`, update the `STAGES` array + `stageForStepName()`
   regexes in `pipeline_app_server.js` to match**, or the app will silently stop
   attributing log lines to the right stage. The per-step **failure strings**
   ("Chart export failed…", "Downloads cleanup could not run…", etc.) are parsed
   too (`FAILURE_LINE` in the server, added 2026-07-13): run_full_pipeline's
   finally block keeps running after an early failure, so the child's exit code
   alone lands on the *last* stage — a real chart-capture failure was once shown
   as a stage-8 cleanup failure this way. Failure is now pinned on the stage that
   printed its failure line; renaming those strings needs the same regex update.

Each stage reports success/failure with the tail of what it actually did (e.g. a
position-adjustment count, an applied/rejected/unmatched tally) — not just a
checkmark. Only one run at a time (`POST /run` returns 409 if one is already in
progress).

## Google Sheets sync (added 2026-07-11): Finance spreadsheet refresh

After a pipeline run, `Stocks_Buy_Strategy.xlsx` gets synced into the user's
**"Finance"** Google Sheet (id `1UjAz_QUuh86_e6yq8QJf2veI8IpkRCyVfWaK6maqiyc`).
The xlsx is always the source of truth — the Google-side data tabs are wiped and
re-imported wholesale each refresh, so never hand-edit them. Workflow (user's
design, confirmed 2026-07-11):

1. The Finance sheet keeps a permanent **`ClaudeCode`** tab that is NEVER deleted
   (Google Sheets requires ≥1 sheet to exist; the tab doubles as an issues/run log —
   e.g. the 2026-07-11 GOOGLEFINANCE commodity test results live there).
2. Delete all data tabs (keep only `ClaudeCode`). This is the standing authorized
   workflow — it prevents Google's "(1)" name-suffixing when the fresh tabs import.
3. File → Import → **Upload** tab → click Browse → drive the native picker with
   **`scripts/drive_open_dialog.ps1`** (see below) → Import location:
   **"Insert new sheet(s)"** → tick **"Import theme"** (explicit user preference) →
   Import data.

**`scripts/drive_open_dialog.ps1`** types a path into the already-open native
Windows file dialog and clicks its Open button, entirely via
`SendMessage(WM_SETTEXT/BM_CLICK)` to the dialog's child controls — no window
focus needed. Do NOT "simplify" it back to `AppActivate`+`SendKeys` (keystrokes
landed in the wrong window — once typing the path straight into a terminal) or
`SetForegroundWindow` (blocked by Windows foreground-lock for background
processes). Both failure modes were hit for real before landing on this approach.

**GOOGLEFINANCE cannot price commodities (verified 2026-07-11).** `TVC:GOLD`-style
formulas return `#N/A`, and so do `CURRENCY:XAUUSD`/`XAGUSD`/`XPTUSD`/`XPDUSD` —
Google has dropped metals support — while equities (`GOOG`) and FX
(`CURRENCY:GBPUSD`) still work. So `ticker_normalize.py`'s
`RELIABLE_GOOGLEFINANCE_COMMODITIES` claim is falsified. **Implemented 2026-07-12**:
`RELIABLE_GOOGLEFINANCE_COMMODITIES` is now empty (commodities get no formula) and
`update_master_sheet.py` writes each commodity's TradingView-captured price into
Investments' Current Price as a plain VALUE, stamped via 'Chart Last Checked'.
Only commodities with a TradingView chart get a price — **Brent (UKOIL), Palladium
and Copper have no chart in the layouts**, so their cells keep whatever they had
(Brent/Palladium show `#N/A` from old dead formulas; Copper's `CPER`-ETF formula
still works). Adding TV charts for those three would auto-fix them on the next run.

**Sync fully automated end-to-end (verified 2026-07-12):** tab deletion (right-click
→ Delete → OK, driven by accessibility refs from the `find` tool — screenshot pixel
coordinates were unreliable because the viewport/screenshot scale differ), then the
Import dialog (its contents are in an iframe invisible to find/read_page, so
screenshot coordinates ARE needed there), then `drive_open_dialog.ps1` for the native
picker. A dated run entry gets appended to the `ClaudeCode` tab each sync.

## Per-run verification of master-sheet writes (added 2026-07-12)

`update_master_sheet.py`'s result JSON now records this run's `commodity_prices`
and `below_alert_rows`, and `verify_pipeline.py` cross-checks BOTH against the
saved workbook (value actually present in Investments col I; below-alert row count
in the 'Stocks of Interest' rows 5-38 reserved block; no resurrected 'Below Alert
Low' sheet) and fails the overall verdict on mismatch. This is deliberately the
last gate before the workbook is imported into the Finance Google Sheet.

## Stocks of Interest formatting + TradingView links (2026-07-12)

- The below-alert block (rows 1-40) is styled to **match the section tables** below
  it: Arial, thin borders, navy title band, `FF2E5077` subtitle/header bands, a red
  `FFC00000` "BELOW ALERT LOW" section band, pale-red data rows, Stock-name-then-
  Ticker identity columns. `refresh_block()` reproduces this every run.
- **Every stock row links to its TradingView layout** via
  `=HYPERLINK("…/chart/<chartId>/","📊 Layout")` (same pattern as Investments' col
  AJ). Below-alert block: column J, reproduced by the pipeline (chartId threaded
  through `update_master_sheet`'s matches). Section tables (hand-maintained): column
  Q (P is 'Last Updated'), added by `add_tv_links_soi_2026-07-12.py`, gated on a
  real-ticker regex so the FTSE-review-calendar / dividend-cover reference tables
  further down the sheet are left untouched. `ticker→chartId` is matched through
  `ticker_normalize`. Rows whose ticker has no captured chart get no link.

## Review deck (added 2026-07-12)

`python scripts/build_review_deck.py [out.pptx]` (default
`~/Downloads/Investment_Review_Deck.pptx`) builds a PowerPoint-Online-compatible
deck from the latest run: summary page (missing charts / no-alert tickers / held
investments with no chart, each flagged red when non-empty), a section slide per
layout, one slide per chart with the cropped image + live price + master-sheet
holdings/alerts + OCR channel read + TradingView alerts, and an appendix of
master rows with no chart. Requires `python-pptx` (installed for the
`AppData\Local\Python` interpreter).

## Input-file consumption + preserved tabs (2026-07-12)

- **Used input files are deleted after a fully successful app run** (user policy
  2026-07-12): `consume_input_files.py` sends the consumed bank/broker exports —
  and every other version of them in Downloads (` (N)` duplicates, `Delete `
  copies) — to the **Recycle Bin** (never hard-deleted; that's the deliberate
  safety floor). Families: activity.csv (Amex), data.csv (Barclays),
  AccountSummary.csv, TransactionHistory*/transactions*.csv (Fidelity). The
  master workbook is NEVER touched. Wired into `pipeline_app_server.js` after
  all stages succeed; failed runs keep their inputs for the re-run.
  `cleanup_downloads.py` itself is still rename-only.
- **`spending_summary.py` preserves manual tabs across rebuilds**: it recreates
  the workbook from scratch each run, which used to silently drop
  hand-maintained tabs — 'Payslip Summary' and 'Retirement Income Plan' were
  lost this way (restored 2026-07-12 from the user's 2026-07-02 Drive copy,
  data as of that date). `preserve_manual_sheets()` now carries over any sheet
  in the existing file that the run didn't generate.
- **Stats charts sit in a 3-across grid from the top**
  (`arrange_stats_charts_2026-07-12.py`, one-off — re-run it if charts are
  added/moved). 'Chart Last Checked' (Investments col AK) is styled to match
  the other headers, enforced idempotently by `get_or_create_last_checked_col`.
- Agents: **`test-analyst`** (end-to-end data-quality audits across every
  pipeline hand-off, read-only) joins `app-developer` and `excel-formatter`.
  Later additions: **`product-owner`** (owns BACKLOG.md, 2026-07-12),
  **`investment-analyst`** (stock analysis, buy prices, daily brief drafts,
  2026-07-12), and **`data-developer`** (data ingestion + transforms — CSV
  loaders, fidelity_file_classifier, ticker_normalize, pivots, master-sheet
  derivations, 2026-07-13). The architecture deck's agents slide is refreshed
  on request by re-running `add_agents_slide_2026-07-12.py` (re-runnable —
  replaces the slide).

## Open items / things to verify on the next export run

- Brent (UKOIL), Palladium and Copper have no TradingView chart, so no captured
  price and no working formula (Copper's CPER formula still works) — user to add
  TV charts if live pricing for them is wanted.
- WPP's Alert Low (1121.92) is ~4x its live price (274.6, gap −75%) — looks like a
  stale/misread level; flagged to the user 2026-07-12, needs a manual look.
- Orphaned `sync_*` temp Google Sheets in the user's Drive (left by a failed
  2026-07-10 sync attempt) still need manual deletion by the user — no Drive
  delete tool available.

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
