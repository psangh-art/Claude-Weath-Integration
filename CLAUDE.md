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
- **Channel boundaries are read at TODAY'S DATE, not the frame edge (user decision
  2026-07-13).** `channel_detect.py` fits each drawn boundary's straight line across
  up to 17 sample x-positions and evaluates the fit at the rightmost candle column
  (= today; found via tight candle-colour mask so the dotted last-price line can't
  match). The original handoff-doc section-4 scan took the first clean sample
  right-to-left, which landed in the blank future space and returned
  projected-forward boundaries — overstating Alert Low/High on ascending channels
  (e.g. SDLF read 836/981 vs 774/919 at today). This refinement supersedes the
  handoff doc's scan order on this specific point, per explicit user instruction —
  if the handoff doc is revised, align it to this. Verified against SDLF/PRU/AEP
  images; BEZ's genuine breakout still correctly rejects.
  - **`find_today_x` fix (2026-07-14): full-width scan + last-price-chip exclusion.**
    `find_today_x` used to cap its candle scan at a hardcoded `0.85w` "to exclude
    the price axis". But the right-offset varies per chart — some run candles to
    ~0.95w, others leave a wide blank future band so the last candle is at ~0.65w —
    so the cap silently read the channel at frac 0.849 regardless: understating
    ASCENDING channels that reach further right (CCH lower rail 3679 vs true 3810,
    +3.8%) and reading a projected-forward rail where there's blank space. The cap
    also over-STATED descending channels (CCR read 94/135 five months back-dated vs
    the true 66/108 at today, −30% — visually confirmed a steep descending channel).
    Now `find_today_x` scans the FULL width and excludes the last-price LABEL CHIP
    (teal up / red down, a fixed ~119px block that matches the candle mask and sits
    flush to the far edge, frac ~0.99): if the rightmost candle-coloured run reaches
    frac ≥ 0.975 it's the chip — walk left across that solid block to its left edge
    (the plot/axis boundary) and take today as the rightmost real candle strictly
    left of it. The boundary SAMPLING still stays in the clean ≤0.85w region and
    EXTRAPOLATES the straight-line fit to today_x (sampling into the candle-occluded
    near-today columns mispairs rail-vs-midline clusters and poisons the fit — tried
    and reverted). Validated: patched fit matches direct pixel reads at today_x to
    ±2 on CCH/CCR/ALW/FCIT/PSON/RKT; clean A/B (committed vs patched on the same
    images) shows only magnitude shifts in the expected direction, zero detections
    gained or lost. (AZN's yellow-drawn channel still yields a stray-blue false
    single — pre-existing, unrelated to this fix, worth a separate look.)
  - **Price-bracket guard now extrapolates one tick past the read labels
    (2026-07-16).** `fit_price_axis`'s bracket check used to reject any read where the
    known price fell outside `[lo_lbl×0.9, hi_lbl×1.1]` of the OCR'd axis labels. The
    lowest/highest label is often OCCLUDED by the last-price chip, the crosshair or a
    drawn marker and never OCRs — WPP's "200" label sits behind the price chip +
    dashed line + arrow, so the axis OCR'd as [400-1200], 278.5 fell below 360, and
    the whole chart was rejected, leaving WPP stuck at a stale Alert Low of 1121.92
    (from when it traded ~£11) and wrongly parked in BELOW ALERT LOW. The axis is
    linear, so the fit off 400..1200 is perfect and the missed label is exactly one
    tick out: the guard now allows the price up to ONE median tick-spacing beyond the
    read labels (`margin = max(0.1×label, tick)`). Measured on the 22 charts the old
    guard rejected: only **3 change** — WPP now reads its yellow rails (202.65/437.34,
    matches the chart), SMWH reads a single high, OCDO's axis is trusted (still no line
    near price). The other **19 stay rejected** and correctly so — their prices are
    genuinely off the visible frame (NXT 14765 vs a 3000-10000 axis, KGF 284 vs
    500-700, IMB 2731 vs 1200-2200): stale/wrong-zoom charts with no drawn line near
    today's price, which still need a redraw in TradingView, not a guard change. Do
    NOT widen the one-tick margin further without re-measuring — 1.5+ ticks starts
    admitting the off-frame reads.
  - **Yellow hand-drawn TREND LINES now feed alerts (user rule 2026-07-14).** Some
    charts have no blue TradingView channel — the user marks support/resistance with
    straight YELLOW trend lines instead (AZN is all-yellow: alert-low line ~10,960,
    alert-high line ~15,700, no blue). `channel_detect.py` now detects them:
    `trend_yellow_mask` (R≥220,G≥205,B≤95 — line core ~#FDEA3B; excludes the pale
    app icon and amber event chips), `_extract_straight_lines` (RANSAC + a COVERAGE
    check so a wavy yellow indicator/EMA is rejected — real trend lines are dead
    straight, measured <2px residual), and `read_yellow_trendlines` (prices at
    today_x, must reach near today, land in-frame, within 0.3–3× price).
    **Governing selection rule (replaces blue-only):** on each side of today's price,
    use the line CLOSEST to price among {blue parallel rails, yellow trend lines} —
    Alert Low = nearest support below, Alert High = nearest resistance above; a
    yellow line beyond the blue rail is out-competed automatically. Every candidate
    is split by which side of price it sits on (NOT by blue lower/upper role — a
    price that has broken just outside the channel would otherwise give a degenerate
    alert_high < alert_low; that's how the SILVER 58.38/58.36 bug appeared and was
    fixed). Blue-only charts are byte-identical to before (verified). The axis OCR
    was refactored into shared `fit_price_axis` (blue + yellow price the same axis).
    Two behaviour consequences flagged to the user: (1) ~6 charts where price sits
    just outside the blue channel now emit a single-sided alert instead of a full
    parallel; (2) where the nearest yellow support is <5% below price, `alert_low ×
    1.05` (in `update_master_sheet.py`) lands at/above current price, so those
    stocks (CNA, BLND, LAND, SILVER, HIK, UU., CRDA…) flag as at/below Alert Low —
    inherent to the nearest-line rule, revisit the ×1.05 buffer if unwanted.

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

## Alert-rules model — signed off by the user (2026-07-15)

The user reviewed `Alert_Rules_Model.pptx` (built by `scripts/build_rules_deck.py`
from a `channel_detect --batch` run) slide by slide and **agreed the governing rule
and all six named patterns**: IN-CHANNEL, BREAKOUT ABOVE, BREAKDOWN BELOW, TREND
LINES ONLY, ON THE LINE, NO READ. Treat these as settled; re-open only on explicit
instruction. Two amendments came out of that review:

- **A wide band is not evidence of a misread.** `build_rules_deck.py` used to flag
  any read whose band was >60% of the share price (or Alert High >1.6x price, etc.)
  as "is this really one channel?" — the theory being that two unrelated drawings had
  been paired. The user confirmed the markup on **every** chart it caught: PAF, IBST,
  BKG (trend lines only), ENT, LSEG, BREE, TW. (parallel only) and AUTO ("correctly
  marked up" — a blue channel with yellow lines over it, where Alert Low came from a
  rail and Alert High from a yellow line). The heuristic was **retired**, not just
  narrowed. Don't reintroduce a span/plausibility guard on the strength of a wide
  band alone — the user's charts genuinely span that much and he wants them published.
- **Below a parallel channel, the band still governs** — this AMENDS the BREAKDOWN
  BELOW pattern. The side-of-price split leaves a broken-down chart with no support
  at all (both rails above price), shipping a bare Alert High while a stale Alert Low
  sits above it in the sheet — that's how RIO/CPG/STJ ended up inverted. The user
  reads a parallel channel as the trading band itself: **price dropping out of the
  bottom IS the buy signal**, not a reason to re-cast the bottom rail as resistance.
  So `channel_detect.py`'s `below_parallel` branch keeps the rails in their drawn
  roles: Alert Low = bottom rail, Alert High = top rail. Deliberately narrow, and the
  narrowness is the user's own scoping — **parallel-ONLY** charts (yellow lines were
  drawn nearer to price on purpose and still win the nearest-line rule: CTEC, ADM);
  a **real channel** of two distinct rails, not a lone blue line (AV., MTRO, MKS);
  and **a breakout ABOVE is untouched** — the top rail still flips to support, which
  the user confirmed explicitly ("only for price below the channel"). This does mean
  Alert Low lands above the current price on these rows, so they flag as below Alert
  Low — that is the intent, not a bug.

**The ×1.05 buffer is skipped once price has REACHED *or PASSED* the support line**
(user decision 2026-07-15). `buffered_alert_low()` in `update_master_sheet.py` used to
skip the buffer only on `on_alert` (price within 0.25% of the line). Once the band rule
above started reading a full pair off a broken-down channel, price sat BELOW the rail
rather than on it — so `on_alert` was false, the buffer applied, and RIO would have been
written 7,407.07 against a rail at 7,054.35 and a live price of 6,927: a level 6.9%
above a price already in hand, and not the "Alert Low = bottom of the parallel" the user
asked for. The guard now takes `price` (from `row.get('price')` at both call sites) and
returns the bare level when `price <= lower`. The CLAMP guard is unaffected, and rows
where price sits above the support line still buffer exactly as before.

### SUPERSEDED 2026-07-16 — ×1.05 buffer RETIRED + parallel channel = the band

The user reviewed the live below-alert list and corrected **GLEN, NG and ADM**, which
together redefined two things above. Both are now settled; the paragraphs above are
kept only as history.

- **The ×1.05 buffer is GONE. Alert Low IS the drawn support line.** GLEN's rail read
  518.37 but shipped 544.29 (= ×1.05); the user wants ~515, i.e. the rail. Same for
  NG (1278→1183) and ADM (3687→3510). `buffered_alert_low()` in
  `update_master_sheet.py` now returns `lower` unchanged (the CLAMP-under-Alert-High
  guard stays; `ALERT_LOW_BUFFER`/`on_alert`/`price` are dead for buffering). This
  drops **every** Alert Low ~5% to its line — a re-run re-applied 93 of them and the
  below-alert list fell 23→6, because most rows were only "below" the *buffered* level.
  Do not reintroduce a buffer without an explicit new user decision.
- **A blue PARALLEL CHANNEL is the trading band and OVERRIDES the nearest-line rule**
  (supersedes the 2026-07-14 "nearest line each side" governing rule *for parallel
  channels*). In `channel_detect.py`'s `process_one`, when two distinct blue rails are
  present and price has **not** broken out above the top rail:
  - **Alert High = TOP rail, always** — even when price has broken DOWN through the
    channel. The old `below_parallel` branch cast the *bottom* rail as resistance and
    shipped that as the high (ADM read 3729.57; the user wants the top rail 4258.30).
  - **Alert Low = BOTTOM rail while price is INSIDE** the channel. A yellow trend line
    sitting between the bottom rail and price no longer raises it (NG: the 1217 yellow
    is ignored, the 1183 rail wins). This reverses the old "yellow drawn nearer to
    price still wins" note for CTEC/ADM above.
  - **If price has broken BELOW the bottom rail, Alert Low = the nearest yellow trend
    line beneath price** (the next support down, ADM → 3510.45), or the broken bottom
    rail if none is drawn.
  - **Two things are deliberately left alone:** a **breakout ABOVE** the top rail
    (top rail flips to support — nearest-line default), and **`on_alert`** rows where
    price is sitting ON a drawn line within `ON_ALERT_TOL` (DGE 1537, CRDA 2879, HSBA
    1476 keep the reached line — the band rule is guarded by `not on_alert` so it can't
    drag Alert Low down to the far channel bottom). The band rule changed 24 charts on
    the 2026-07-16 batch (11 raised Alert High to the top rail, 10 lowered Alert Low to
    the bottom rail/trend beneath, 3 both); the user reviewed the full diff and approved.

## Investments 'Marked Up' column at B (2026-07-16)

The user asked for a column after Investments!A ('Chart' = does a chart exist)
confirming whether HE has **marked up** the chart (drawn channel/trend lines) —
`Yes`/`No`, auto-derived, at **column B**. `update_master_sheet.marked_up_flag()` is
the single source of the rule (a detected pattern OR an axis-read failure ⇒ drawings
present ⇒ Yes; 'no channel or trend line found near price', a macro/reference symbol,
or no capture ⇒ No) and `update_master_sheet` writes it every run.

**Inserting that column shifted EVERY Investments data column one to the right**, so
this is the sheet-rename-VLOOKUP trap (see 2026-07-10 note) at column scale. Two
things had to move together and MUST stay in sync on any future structural change:

- **~5,600 in-sheet formulas + the cross-sheet VLOOKUPs** in History and 'Stocks of
  Interest' that point at Investments. openpyxl does NOT adjust formula references on
  insert, and no LibreOffice/Excel-COM engine is available here, so
  `insert_marked_up_col_2026-07-16.py` (a one-off, already run) shifts them with the
  formula Tokenizer — touching only real cell refs that resolve to Investments, never
  string literals (`googlefinance("INDEXFTSE:UKX")`, HYPERLINK urls), `'Base Data'!`
  refs, or a sheet's own refs. It also moved the merged title A1:W1→A1:X1, the
  conditional-format range A4:T278→A4:U278, 20 column widths, and freeze D3→E3.
- **The hardcoded column constants in the CODE.** All are now +1 vs the old layout:
  `update_master_sheet.py` (`COL_MARKED_UP=2`, SHARE_NAME 3, TICKER 4, HOLDINGS 5,
  TARGET 7, CURRENT_PRICE 10, ALERT_LOW 13, ALERT_LOW_SOURCE 14, ALERT_HIGH 16,
  CLAUDE_NOTES 32), `refresh_soi_sections.py` (`INV_TICKER/LOW/HIGH = 4/13/16` and its
  `VLOOKUP('Investments'!$D:$J,7`), `build_review_deck.py`, and `verify_pipeline.py`
  (ticker col 4, price col 10). **Add or move an Investments column again and every
  one of these needs the same coordinated bump** — a wrong constant writes Alert Low
  into the wrong column of a live trading sheet.

## 'Stocks of Interest' section tables are pipeline-maintained (2026-07-15)

The section tables BELOW the auto-built below-alert block (rows 41-81: at lower
boundary / near / watchlist / breakouts) used to be hand-maintained, and drifted:
audited 2026-07-15, **17 of 25 rows sat in the wrong section** and 21 of 25 carried an
Alert Low that no longer matched the pipeline, all stamped 'Last Updated 2026-07-06'.
The drift ran the DANGEROUS way — BEZ (+15.5%) and AUTO (+11.4%) were still listed as
'AT LOWER BOUNDARY — highest priority' while nine stocks genuinely at their buy point
(DGE, AZN, GSK, GOLD, HSBA, CRDA, PLAT, PRU, ALW) sat in lower-priority sections, and
GSK was listed twice. `scripts/refresh_soi_sections.py` now rebuilds them, wired into
`run_full_pipeline.js` as a sub-step of step 3 AFTER `update_master_sheet.py` (it reads
the Alert Low/High that step writes). Idempotent — a second run is a byte-for-byte
no-op. Ownership, which must be preserved by any change:

- **Pipeline-owned** (rewritten each run): C Pattern, E Alert Low, G Alert High,
  P Last Updated, section placement, and within-section order (nearest its alert low
  first).
- **Regenerated each run**: D/F/H/I/J/Q are row-relative formulas and MUST be rewritten
  rather than copied — a row that moves section otherwise keeps pointing at its old
  row's cells. Same class of bug as the sheet-rename VLOOKUP breakage above.
- **Hand-written, carried per ticker and never invented**: A Stock, K Chart Note,
  L Analyst Rating, M Holdings, N Target Value, O Notes.
- **MEMBERSHIP IS NOT AUTOMATIC.** These tables are a curated watchlist: the script
  re-sections and refreshes the stocks already listed and never adds or drops one. Add
  a stock by hand and the next run places and maintains it. (The block above it IS a
  full auto-generated list — different thing, same sheet.)

Two traps found while building it, both worth keeping in mind:
- **Unmerge before clearing.** The section bands are merged ranges and a `MergedCell`'s
  `.value` is read-only — clearing first raises AttributeError.
- **'No lines drawn' ≠ 'not read this run'.** `channel_detect` sets
  `pattern='no lines read'` for BOTH an undrawn chart and a failed axis read, so only
  `reason` separates them (every axis failure names the axis; an undrawn chart says
  'no channel or trend line found near price'). A row whose axis read failed still HAS
  the user's drawings and keeps its Alert Low from an earlier run — labelling it 'No
  lines drawn' would tell him his trendlines had vanished, and would hide that the
  level is inherited. 7 of 353 charts read cleanly at 09:22 but not at 16:00 (AUTO,
  NWG, TW., SMT, BA, ENT, III), six of them withheld by the price-bracket guard.

**The below-alert block's SECTION LAYOUT is coupled to `verify_pipeline.py`'s row
count.** The block is no longer one section: 'On Alert' sits above 'Below Alert Low'
(2026-07-15), so a second section band and column-header row now fall inside the
rows 5-38 data range. Verify counted "column A is not empty" over that range and so
counted the band and header as data — 25 against the 23 actually written, failing a
run whose block was correct. It now counts a ticker in column B plus a numeric price
in **column F** (was column C until 2026-07-16 — see the schema-widening note below;
column C now holds the Pattern text, so the price moved to F). **Add another section
to the block, or move the price column again, and this count needs revisiting** — it
is the last gate before the Google Sheet import, so a false failure here is as costly
as a missed one.

**The below-alert block now shares the section tables' 17-column schema (user
request 2026-07-16).** The auto-built block (`add_below_alert_sheet.py`) used to
carry only 10 columns (Stock, Ticker, Current, Alert Low, Gap %, Alert High,
Holdings, Target, Checked At, TradingView); the user wanted "Trading below alert /
Below alert low" to read the same as the 'Near Lower Boundary' section table below
it — "it's missing chart note, div yield etc". `HEADER` and the row builder now
mirror `refresh_soi_sections.py`'s columns exactly: Stock, Ticker, **Pattern,
Proximity, Alert Low, Current, Alert High, Upside %, P/E, Div Yield, Chart Note,
Analyst Rating**, Holdings, Target, **Notes, Last Updated**, TradingView. Consequences
baked in on purpose: (1) **Current sits in column F as a literal value, NOT a
VLOOKUP formula** like the section tables use — the pipeline already has the captured
price and verify's block-count gate must see a real number there (a formula string
reads as non-numeric under openpyxl and would fail the gate). (2) **Chart Note (K),
Analyst Rating (L) and Notes (O) are left BLANK** — they're hand-curated per ticker
in the section tables only, and the block is a full auto-generated list, so inventing
them here isn't wanted (user decision 2026-07-16). (3) Pattern comes from
channel_detect via a new `'detection'` field carried on each match in
`update_master_sheet.py`, labelled with `refresh_soi_sections.pattern_label` — so a
row whose axis read failed this run shows 'Not read this run — level inherited', not
the misleading 'No lines drawn'. `LAST_COL` is 17; the bands/headers merge A..Q.

**WEDGE — a seventh pattern (user request 2026-07-15).** Two yellow trend lines
converging in the near future. It **does not change Alert Low / Alert High** — the
user was explicit that the trend-line rules still stand, so a wedge is a
classification of the chart's shape plus one signal: **price breaking ABOVE the wedge
is a potential buy**. `detect_wedge()` in `channel_detect.py` needs the fitted SLOPES,
so `_read_lines_at_today` now returns `(records, meta)` carrying each line's pixel-space
fit; `read_yellow_trendlines_geom` exposes it and `process_one` derives the prices from
that one call (a separate wedge read would have added a third `fit_price_axis` OCR pass
per chart — that OCR is why a batch run takes ~4 minutes).

Both gates are load-bearing, and the numbers came from measuring the real charts —
**don't loosen either without re-measuring**:
- `WEDGE_HORIZON_FRAC = 0.15` — the apex must be at most 0.15 pane-widths past today
  (~3 months; the user picked this from 0.15/0.25/0.40 options). "Two lines that meet
  somewhere to the right" is not a pattern: *any* two non-parallel lines do, and the
  ungated test fired on **31 of 353 charts**. At 0.15w it fires on 6 (AML, BAG, CNA,
  DGE, EZJ, HOC); 0.40w doubles that to 14.
- `WEDGE_MIN_GAP_FRAC = 0.03` — the lines must still be 3% of price apart at today.
  This only rejects one thick line double-detected as several (PALLADIUM's three
  "lines" within 2% of each other, apex 0.001w ahead; also PLATINUM/LAND/ADM). Varying
  it 3→8% changes nothing else, so the horizon is the real lever.

The pattern is named ahead of the blue classification, because the wedge is made by the
yellow lines and the user reads it as the chart's shape — CNA and DGE sit inside a blue
channel *and* carry a wedge. Alert levels are untouched either way; CNA is the case to
check a change against, since its wedge lines (141.88/168.49) are NOT its alert lines
(177.73/227.42).

Known deck defect, not yet fixed: **CTEC is a bad example on the BREAKDOWN BELOW
slide** — it's a mixed chart whose alert levels both came from yellow lines
(166.73/236.22), so the rail-flip the slide describes never drove it. Pick a
parallel-only example if that slide survives the amendment above.

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
gallery of the review deck), and `/asset?p=` (image proxy for the gallery,
**whitelisted to the repo + Downloads only** — never serve an arbitrary path).
`/deck.pptx`, `/architecture.pptx` and `/download/spending` still work as direct
URLs but the UI no longer offers downloads (user request 2026-07-13): the Output
Bay links open products in their web apps instead — Finance Google Sheet, review
deck (in-app gallery view + "Edit in PowerPoint Online"), `spending_summary.xlsx`
("Open in Excel Online"), architecture deck ("Open in PowerPoint Online").
Because Office Online can only open files that live in OneDrive (Downloads is
NOT synced), `syncOneDriveProducts()` copies the three product files to
`C:\Users\Paul\OneDrive\Investment Production\` (config `onedriveProductsDir`)
at startup and after every run — `copyFileSync` overwrite keeps OneDrive item
IDs (and share links) stable. Direct one-click open needs each file's OneDrive
link pasted ONCE into config.json → `productWebLinks`; until then the buttons
fall back to an honest "Find in OneDrive" search link. `build_review_deck.py`
emits the gallery (`pipeline_app/review_deck.html`, gitignored) + a summary JSON
that feeds the output bay. Per-stage durations of each run are kept in
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

## Spending tabs mirrored into the master workbook (2026-07-13)

`scripts/integrate_spending_tabs.py` copies EVERY sheet of
`spending_summary.xlsx` into `Stocks_Buy_Strategy.xlsx` as same-named tabs
(user request 2026-07-13), so the master — and the Finance Google Sheet it's
imported into — carries Wealth Summary / Targets / Payslip Summary /
Retirement Income Plan alongside the investment tabs. Runs as a sub-step of
run_full_pipeline's step 3 (deliberately not a numbered step marker — the
Production Centre parses those). spending_summary.xlsx is the source of truth
for these tabs (replaced in place each run, same tab position); the master's
own six tabs are in a PROTECTED_MASTER_TABS set and can never be overwritten
by a colliding sheet name. A missing spending_summary.xlsx skips cleanly (the
spending build is optional). The cross-workbook sheet-copy logic lives in
`scripts/xlsx_sheet_copy.py`, shared with spending_summary.py's
preserve_manual_sheets() — one copy of that code, don't re-inline it.

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
Only commodities with a TradingView chart get a price. The user added charts for
Brent, Palladium and Copper on 2026-07-13; `ticker_normalize.py` now maps the TV
chart symbol `BRENT` to master ticker `UKOIL` and strips TradingView's
continuous-futures `1!`/`2!` suffix (`COPPER1!` → `COPP`), so all three populate
on the next run. Their old dead `TVC:`/`CPER` formulas in col J are superseded by
the written value in col I.

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

- Brent/Palladium/Copper charts were added by the user 2026-07-13 and the symbol
  mapping fixed (`BRENT`→UKOIL, futures `1!` suffix stripped) — verify on the next
  run that all three get a captured price written to Investments col I. Their dead
  col-J formulas (`TVC:UKOIL`, `TVC:PALLADIUM`) will still show `#N/A` — consider
  clearing them once the value-writes are confirmed.
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
