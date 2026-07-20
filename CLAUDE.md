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
    gained or lost. (AZN's stray-blue false single — flagged here originally — was
    RESOLVED by the yellow-trend-line work: on the current capture AZN reads
    `blue_lines: []` and three legit yellow lines, see the AZN note under the
    all-yellow section below. This parenthetical is kept only as history.)
  - **`find_today_x` now excludes the OHLC legend band + reaches through
    fill-dimmed candles (2026-07-17).** Two related today_x errors, both found by
    reviewing the annotated review deck against the drawn charts:
    (1) **Teal/red OHLC legend text** at the very top ("O.. H.. L.. C.. +66..") is
    drawn in the up-candle teal / down-candle red and matches `candle_mask`. On a
    chart with blank future space the legend extends further RIGHT than the last real
    candle, so today_x landed in the legend and the ascending rail was read out in the
    future projection — **ADM read a 4345 top rail vs the true ~3935 at today** (user:
    "alert high should be the top of the parallel"). Fix: zero the top `min(100, h/2)`
    px before counting candle columns. (2) **Candles drawn INSIDE a channel fill** are
    dimmed by the navy overlay (down #F23645→~#9A293E, up #089981→~#086764) and fall
    below the tight mask, so on **FRAS** today_x stopped at 0.73w and understated the
    top rail (827→843). Fix: a `find_today_x`-only dim-candle mask (`b<=110` keeps it
    off the magenta event chips), unioned with the tight mask. THREE guards make the
    dim extension safe — do not remove: the chip's LEFT EDGE is found on the TIGHT mask
    (the dim mask fills the axis-panel gap and would swallow the whole series — cost
    SILVER/NATGAS their read); the extension is a CONTIGUOUS walk right from the last
    bright candle that stops at the first real gap (so genuine blank future is never
    crossed and the chip's own antialiasing bleed can't drag today onto the chip edge —
    that put NXT +16% and PAGE onto a stale reading); gap tolerance `max(15, .012w)`.
    Isolated A/B over the 342-ticker batch: legend fix **14 changed, 0 lost** (all the
    expected direction); dim fix **8 changed, 0 gained/lost** (FRAS 843, 7 small
    toward-today shifts). Verified on-chart: ADM/FRAS sit on the rail at the last bar,
    SILVER/NATGAS/NXT/EZJ/PAGE unregressed.
  - **Axis bracket guard retries psm 11 when the read doesn't bracket the price
    (2026-07-17).** `fit_price_axis` rejected a chart when its clean labels didn't
    bracket the known price — but the labels NEAREST the price are the ones most often
    missed: on a chart zoomed wide to show an old spike, today's ticks sit in the busy
    candle region where the default OCR block-segments over them (**AO World** read only
    [150-450] and was dropped, so its two yellow trend lines never fed alerts; user:
    "AO World should use the 2 trend lines"). Fix: when the default clean labels don't
    bracket the price AND psm 11 hasn't run, retry with sparse-text psm 11 and accept it
    ONLY if the recovered set has ≥3 labels AND actually brackets the price. A genuinely
    off-frame price still finds no bracketing labels and stays rejected. Batch A/B: **24
    recovered (None→read), 0 lost, 0 existing reads altered** — incl. AO 87.8/119.64,
    and KGF/NXT/IMB (the tickers the 2026-07-16 one-tick note said to leave rejected,
    now correct because the user redrew them as on-frame monthly charts — verified the
    recovered rails sit on the drawings).
  - **Review-deck level tags shrunk to "Low"/"High" + value, no background box
    (2026-07-17, user request).** `annotate_chart_levels` in `build_review_deck.py`
    draws small coloured text with a thin dark stroke instead of a filled label chip,
    so the level markers stay legible without covering the chart — the user reads many
    charts fast to confirm the detected levels.
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
  - **Axis OCR now falls back to sparse-text mode (psm 11) when the default reads too
    few labels (2026-07-16).** `fit_price_axis` OCR'd the axis crop with Tesseract's
    DEFAULT page segmentation (psm 3), which treats the tall gutter as one block and on
    some charts finds only ONE tick label — HTWS/Helios Towers read just "120" of eight
    clearly-visible labels (240/220/.../100), failed the `>= 3 clean labels` gate, and
    was silently dropped, leaving a STALE Alert High of 251.2 with no low. psm 11
    ("sparse text") is built for scattered labels like a price axis and recovers them
    (HTWS: 6 of 8). The fix tries the default FIRST — so every chart that already reads
    stays byte-identical — and only re-runs with psm 11 when the default yields < 3
    clean labels, taking the sparse read only if it has MORE clean labels. Measured
    against the committed code over the 353-chart batch: **17 charts changed, ALL
    None -> now-read recoveries, ZERO reads lost, ZERO reads altered** (HTWS, HSBA,
    SHEL, SVT, ANTO, BT.A, VOD, SBRY, ULVR, SMT, BATS, WIZZ, WEIR, OSB, UTG, NATGAS,
    COPPER1!). The axis price-bracket guard (2c) still applies to every recovered read,
    so a recovered-but-off-frame axis is still rejected — the fallback only makes a
    genuinely-readable axis readable, it does not relax the trust check.
  - **Yellow trend-line minimum span lowered 0.12 -> 0.10 (`YELLOW_MIN_SPAN_FRAC`,
    2026-07-16).** A genuine drawn resistance line can be SHORT when it only marks a
    recent leg: HTWS/Helios Towers' descending WEDGE line spans 0.117w (646 collinear
    px, 97% coverage) and was dropped at 0.12, leaving HTWS with only its Alert Low and
    no wedge classification. Lowered so it's admitted; a line this collinear is never
    noise. Isolated against the batch (psm-11 held constant, span 0.12 vs 0.10):
    **exactly 2 charts change** — HTWS (single_low -> parallel + WEDGE, gaining the
    descending rail 198.7) and COPPER1! (single_low -> parallel, gaining a real short
    yellow resistance line at 1407.52, verified on the chart). Do NOT drop below 0.10
    without re-measuring — the RANSAC subsample (cap 4000) already under-represents a
    minority line, so shorter floors start admitting fragments. (Note: raising the
    subsample cap alone does NOT recover these lines — the span floor is the only gate
    that was rejecting them; verified at caps 8000/20000.)
  - **Yellow hand-drawn TREND LINES now feed alerts (user rule 2026-07-14).** Some
    charts have no blue TradingView channel — the user marks support/resistance with
    straight YELLOW trend lines instead (AZN is all-yellow: an ascending channel
    — lower rail ~10,958, upper rail ~15,707 — PLUS a flat horizontal support line
    at ~12,218, no blue). `channel_detect.py` now detects them:
    `trend_yellow_mask` (R≥220,G≥205,B≤95 — line core ~#FDEA3B; excludes the pale
    app icon and amber event chips), `_extract_straight_lines` (RANSAC + a COVERAGE
    check so a wavy yellow indicator/EMA is rejected — real trend lines are dead
    straight, measured <2px residual), and `read_yellow_trendlines` (prices at
    today_x, must reach near today, land in-frame, within 0.3–3× price).
    **Governing selection rule (replaces blue-only):** on each side of today's price,
    use the line CLOSEST to price among {blue parallel rails, yellow trend lines} —
    Alert Low = nearest support below, Alert High = nearest resistance above; a
    yellow line beyond the blue rail is out-competed automatically. **AZN confirms
    this rule (user decision 2026-07-17):** its Alert Low reads the horizontal
    support line at ~12,218 (nearest support below the 12,594 price), NOT the
    ascending lower rail at ~10,958 — the user drew the flat line as the buy level
    and price sitting just above it is the intended signal; keep the nearest-support
    read, do not special-case AZN back to the rail. Every candidate
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
  `oneoff/insert_marked_up_col_2026-07-16.py` (a one-off, already run) shifts them with the
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

**A NEW COLUMN MUST FOLLOW THE EXISTING FORMAT of the column beside it — never leave
cells on openpyxl's Calibri/Arial-10, no-fill, centred defaults (user rule
2026-07-16).** 'Marked Up' shipped with its VALUES set but NO styling, so its data
cells were visibly out of step with the green/grey Yes/No 'Chart' flag right next to
it (Arial 10 vs 9, no fill vs the green `FFC6EFCE`/grey `FFF2F2F2` pair, centred vs
left/top). Fixed by `set_marked_up_flag()` in `update_master_sheet.py` — which styles
column B atomically (value + Arial-9 font + green/grey fill + thin border + left/top)
to MATCH column A 'Chart', the same way `set_chart_flag()` does — plus a one-off,
`oneoff/fix_marked_up_format_2026-07-16.py` (already run, backs the workbook up first), that
applied that style to all 355 existing rows and fixed the B2 header alignment. **Any
future added/inserted column: mirror the adjacent column's font, fill, border and
alignment in the writing code AND back-fill existing rows — matching the neighbour is
the rule, openpyxl defaults are never acceptable in this workbook.**

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
  - **Transient-failure hardening (2026-07-16).** This step once failed a whole
    Production Centre run with "error code 1 on Downloads cleanup" and then wouldn't
    reproduce — because both its filesystem calls could raise and abort the (purely
    cosmetic) step: `os.path.getmtime` on a file that vanished between `os.listdir`
    and stat (a browser `.crdownload` / OneDrive sync temp), and `os.rename` on a
    file locked by Excel or mid-OneDrive-sync. Both are now wrapped in `try/except
    OSError` — a vanished file is skipped, a locked rename prints `SKIPPED — could
    not rename …` and continues. The step must never fail the run over a transient
    file; keep it that way.

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
Bay links open products in their web apps instead. **Output Bay tiles (user
request 2026-07-16): Finance Google Sheet, review deck (in-app gallery "View"
only — user confirmed the in-app view is enough, no PowerPoint-Online edit link),
architecture deck ("View in PowerPoint Online").** The standalone `spending_summary.xlsx` tile was REMOVED — its tabs are
now mirrored into `Stocks_Buy_Strategy.xlsx` / the Finance Google Sheet by
`integrate_spending_tabs.py`, so the Finance workbook tile covers it (the spending
BUILD stage stays — it's the source those tabs are copied from; only the
duplicate product tile went). Because Office Online can only open files that live
in OneDrive (Downloads is NOT synced), `syncOneDriveProducts()` copies the product
files to `C:\Users\Paul\OneDrive\Investment Production\` (config
`onedriveProductsDir`) at startup and after every run — `copyFileSync` overwrite
keeps OneDrive item IDs (and share links) stable. Each pptx tile's "View/Edit in
PowerPoint Online" needs that file's OneDrive share link pasted ONCE into
config.json → `productWebLinks`; **the old "Find in OneDrive" search-link fallback
was removed (user request 2026-07-16) — the user wants a straight view of the end
result, not a lookup**, so a tile with no pasted link shows "Not available yet"
until the link is added. `build_review_deck.py`
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

**THE TAB MUST BE VISIBLE, and the browser sync CANNOT be fully automated
(2026-07-20).** Claude-in-Chrome drives tabs in the BACKGROUND —
`document.visibilityState` reads `"hidden"` — and **Chrome will not open a native
file picker for a hidden tab**. The Browse click silently does nothing,
`drive_open_dialog.ps1` correctly reports "no visible Open dialog" (there is no
dialog to drive), and because the data tabs are deleted FIRST, **the user's live
Finance Sheet sits EMPTY** until someone finishes by hand — which is exactly what
happened on 2026-07-20. Foregrounding the window with user32, `window.focus()` and
creating a fresh MCP tab all failed to change `visibilityState`; only the user
clicking the tab did. **Check `visibilityState` BEFORE deleting any tab.** Note also
the Import dialog now defaults Import location to **"Create new spreadsheet"** (not
Replace) — it must be changed to **Insert new sheet(s)**, and only then does the
**Import theme** checkbox appear.

### Sheets API sync — service-account auth is set up (2026-07-20)

The real fix for the above is to stop driving a browser: write each tab in place
with the Sheets API (`values.batchUpdate`) — no deletion, no picker, no empty
window, and re-runnable unattended as a sub-step of `run_full_pipeline.js`. The
CREDENTIAL half is done and verified; the sync itself is still to build.

- **A SERVICE ACCOUNT, not OAuth** (user decision 2026-07-20). An OAuth desktop app
  left in "Testing" mode has its refresh token expired by Google every 7 days, which
  would break the sync roughly weekly. A service account has no consent screen, no
  browser step and no expiry.
- Project `finance-sheet-sync-503009`, account
  `finance-sync@finance-sheet-sync-503009.iam.gserviceaccount.com`, Sheets API only
  (the Drive API is NOT needed to edit an existing sheet). **The step that is always
  missed is sharing the Finance sheet with the service-account email as Editor** —
  the key is valid and every call still 403s until you do.
- **The key lives at `C:\Users\Paul\.secrets\finance-sheets-sync.json` — OUTSIDE the
  repo**, because the repo is inside OneDrive and a private key must not sync to the
  cloud. Overridable via `config.json → sheetsServiceAccountKey`. `.gitignore` carries
  `*service-account*.json` / `*-sheets-sync.json` so an in-tree copy can't be
  committed by accident.
- `scripts/check_sheets_auth.py` verifies all three independently-breakable things
  (key parses and is a service-account key; Sheets API enabled; sheet actually
  shared) and names which one failed. Read-only — safe against the live sheet.

**NORTON INTERCEPTS HTTPS, and it breaks Python (not Node or Chrome).** Norton's
"Web/Mail Shield" re-signs every TLS certificate: connect to
`oauth2.googleapis.com` and the certificate presented is issued by `CN=Norton
Web/Mail Shield Root, OU=generated by Norton Antivirus for SSL/TLS scanning`. Norton
installs that root into the **Windows** store, so Chrome and Node are unaffected —
but Python verifies against certifi and dies with `CERTIFICATE_VERIFY_FAILED:
unable to get local issuer certificate`. **Pointing at certifi does NOT help**; an
explicit `verify=certifi.where()` fails identically. (Compounding it, this
interpreter ships with no trust store at all: `cafile=None` and its OpenSSL default
`C:\Program Files\Common Files\SSL/cert.pem` doesn't exist.) `scripts/ssl_certs.py`
`ensure_ca_bundle()` injects **`truststore`** so Python verifies against the OS
store, with a certifi env-var fallback; call it before the first HTTPS request in any
script that talks to an external service. **Never `verify=False`** — that trades a
local trust-config problem for a permanent silent downgrade on a path carrying
financial credentials, and would hide a genuine MITM.

The architecture deck carries all of this on a dedicated slide, rebuilt by the
re-runnable `scripts/oneoff/add_sheets_sync_slide_2026-07-20.py`.

### `sync_finance_sheet.py` — the API sync itself (built 2026-07-20)

`python scripts/sync_finance_sheet.py [--dry-run] [--tabs A,B]`. Writes each tab
IN PLACE in ~20s. **Verified against the pre-sync snapshot at 0 differing cells on
all 9 tabs**, twice (so it's idempotent) — the sheet had just been populated by a
manual xlsx import, which made "byte-identical to the import" the acceptance test.

Four things are load-bearing; changing any of them silently corrupts a live
financial sheet:

- **TWO WRITE PASSES, because one `valueInputOption` cannot serve the whole grid.**
  RAW for everything except formulas and dates; USER_ENTERED for those only.
  USER_ENTERED **coerces**: it turned `6.7%` into `0.067` rendered as `6.70%`, and
  would turn a free-text Chart Note reading `1-2` into a date. There are ~6,100
  plain strings exposed to that. RAW stores them verbatim; formulas and dates
  genuinely must be parsed (a formula written RAW lands as literal text).
- **Which cells get parsed is decided by the openpyxl CELL TYPE, never by sniffing
  the rendered string.** The first version tested "does this text look like an ISO
  date", which converted Investments' 'Chart Last Checked' column — 347 cells that
  are plain TEXT in the workbook — into real dates. Exactly the coercion the
  two-pass design exists to prevent. Only the source type can tell them apart.
- **NUMBER FORMATS MUST BE RE-PUSHED AFTER the USER_ENTERED pass, which clears
  them.** Measured: after the first full sync every formula cell came back
  `numberFormat: None`, so `-0.4%` rendered as `-0.003685720404` while RAW-written
  literals kept their `#,##0.00`. `sheets_number_format()` maps the workbook's
  format codes across (Excel and Sheets share the pattern syntax for everything
  this workbook uses). Taking them from the WORKBOOK rather than reading them back
  off the sheet means the sync can also repair a sheet whose formats are already
  lost — which is how the live sheet was fixed.
- **`ClaudeCode` and `Stats` are never written.** ClaudeCode is the run log (and
  Google needs ≥1 sheet); **Stats holds CHARTS and no data** (1×1 in the xlsx) —
  the API cannot rebuild charts from openpyxl, so clearing it would destroy them.

Fills, borders, merges and column widths are NOT pushed and don't need to be:
`clear()` wipes values only. So the model is **import the xlsx by hand ONCE to
establish the appearance, then let the API keep the numbers current**. A
STRUCTURAL change (inserting a column, moving a section) misaligns that surviving
formatting and needs a fresh manual import — the known cost of the approach.

The run-log line is appended at the true next free row, computed explicitly.
`append_row(table_range='A1')` treats the first contiguous block as "the table"
and INSERTS after it, which put five test entries at the TOP of ClaudeCode and
pushed the history down (repaired from the pre-sync snapshot).

**Take a snapshot before any bulk sheet operation.** `get_all_values(
value_render_option='FORMULA')` over every tab dumps to JSON in seconds, and it
is what made both mistakes above cheaply reversible.

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
  Q (P is 'Last Updated'), added by `oneoff/add_tv_links_soi_2026-07-12.py`, gated on a
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

**Alert Low/High drawn ONTO each chart image (user request 2026-07-16).** So the
user can eyeball the detected levels across many charts fast, `annotate_chart_levels()`
draws a horizontal line at each level's pixel row — **Alert Low in green, Alert High
in orange** — with a labelled tag, straight onto the chart before it goes into BOTH
the .pptx and the in-app gallery (`write_gallery`). The pixel row comes from
channel_detect's axis fit (`price = a*y + b`), which `process_one` now returns as
`axis_a`/`axis_b` (+ `pane_h`/`pane_w`) in its result — surfaced from
`_read_lines_at_today`'s meta so it costs no extra OCR pass. A level that maps off
the visible frame is skipped; a chart with no axis read (nothing to place a level
against) is shown un-annotated. Annotated copies go to
`scripts/pipeline_app/_annotated_charts/` (gitignored) — the .pptx embeds the bytes
at `add_picture` time and the gallery serves them through the `/asset` proxy (that
dir is inside the repo, which the proxy whitelists), so the files are throwaway.
Green/orange were chosen to stay distinct from the chart's own yellow trend lines
and blue channel. This is a VERIFICATION aid — it draws whatever channel_detect
read this run (`lower`/`upper`), so a wrong level shows up wrong, which is the point.

## Input-file consumption + preserved tabs (2026-07-12)

- **Used input files are ARCHIVED after a fully successful app run — the LAST
  copy of each is kept** (user policy 2026-07-12, amended 2026-07-18, re-confirmed
  2026-07-19). `consume_input_files.py` MOVES the newest export of each type into
  **`~/Downloads/old_pipeline/`** under a canonical name, so exactly ONE version of
  each is retained there: `activity.csv` (Amex), `data.csv` (Barclays),
  `AccountSummary.csv`, `TransactionHistory.csv` (historic) and
  `TransactionHistory_pending.csv` (pending) — the last two split by
  `fidelity_file_classifier` content, not filename. A prior copy already in
  old_pipeline is replaced, so history never accumulates. **Every OTHER version left
  in Downloads** (the browser's ` (N)` duplicates, `Delete ` copies) goes to the
  Recycle Bin. Nothing is ever hard-deleted — that's the deliberate safety floor;
  the recycle call is `config.recycle_to_bin` (this file used to carry a second,
  identical SHFileOperation implementation — removed 2026-07-19). The master
  workbook is NEVER touched. Wired into `pipeline_app_server.js` after all stages
  succeed; failed runs keep their inputs for the re-run. `cleanup_downloads.py`
  itself is still rename-only.
  - **`purge_flagged_downloads.py` + `Clean Up Downloads.bat` on the Desktop
    (added 2026-07-20) are the human decision cleanup_downloads.py defers.** They
    recycle ONLY files whose name begins exactly with the `Delete ` prefix the
    pipeline itself wrote — nothing is matched on age, size or extension, so a file
    can never be caught by accident; rename a file with that prefix to have it
    removed, take the prefix off to spare it. Recycle Bin, never a hard delete, and
    it lists everything and asks first (`--yes` skips the prompt, `--dry-run` only
    lists). With no console to confirm on it recycles NOTHING rather than assuming
    yes. This is the only thing in the repo that removes a user-visible file, which
    is why the blast radius is kept this narrow — do not widen it to pattern
    matching.
  - Note: the Amex family matches **`activity.csv` only**. An `activity.xlsx` sitting
    in Downloads is neither read by `spending_summary.py` nor archived — export Amex
    as CSV.
- **`spending_summary.py` preserves manual tabs across rebuilds**: it recreates
  the workbook from scratch each run, which used to silently drop
  hand-maintained tabs — 'Payslip Summary' and 'Retirement Income Plan' were
  lost this way (restored 2026-07-12 from the user's 2026-07-02 Drive copy,
  data as of that date). `preserve_manual_sheets()` now carries over any sheet
  in the existing file that the run didn't generate.
- **Stats charts sit in a 3-across grid from the top**
  (`oneoff/arrange_stats_charts_2026-07-12.py`, one-off — re-run it if charts are
  added/moved). 'Chart Last Checked' (Investments col AK) is styled to match
  the other headers, enforced idempotently by `get_or_create_last_checked_col`.
- Agents: **`test-analyst`** (end-to-end data-quality audits across every
  pipeline hand-off, read-only) joins `app-developer` and `excel-formatter`.
  Later additions: **`product-owner`** (owns BACKLOG.md, 2026-07-12),
  **`investment-analyst`** (stock analysis, buy prices, daily brief drafts,
  2026-07-12), and **`data-developer`** (data ingestion + transforms — CSV
  loaders, fidelity_file_classifier, ticker_normalize, pivots, master-sheet
  derivations, 2026-07-13), and **`validation`** (audits the review deck's
  detected patterns/alert levels against the signed-off pattern rules,
  fixes every instance of a shared root cause with an A/B measurement, and
  flags genuinely new patterns for the user, 2026-07-17). The architecture
  deck's agents slide is refreshed on request by re-running
  `oneoff/add_agents_slide_2026-07-12.py` (re-runnable — replaces the slide).

## Investment Dashboard + output hygiene (2026-07-17)

A batch of dashboard/deck changes, all committed:

- **Only ONE version of each output file per run.** `config.py` now has
  `recycle_to_bin()` + `purge_old_versions(target)` (Windows SHFileOperation,
  `FOF_ALLOWUNDO` — Recycle Bin, never hard-delete). `purge_old_versions` clears
  the browser/Office `X (N).ext` duplicates next to a canonical output, leaving the
  canonical for the caller's own overwrite. Wired into `build_review_deck.py` and
  `build_rules_deck.py` before `.save()`, and a new `purge_output_duplicates.py`
  runs as the last action of `run_full_pipeline.js`'s cleanup step over every output
  (master, layouts, feedback md, the 3 decks, spending). `cleanup_downloads.py`
  stays deliberately rename-only. **Any new output-writer should call
  `purge_old_versions` too.**
- **Intelligence screen (dashboard) uses the Yahoo Finance chart API, NOT Stooq.**
  Stooq's CSV endpoint returns a bot/JS-challenge page — verified dead. The six
  index symbols live in `config.json → intelligenceIndices` (Yahoo `^FTSE`, `^GDAXI`,
  `^STOXX50E`, `^GSPC`, `^IXIC`, `^DJI`). `dashboard_server.js` `/api/intelligence`
  fetches server-side (browser can't call Yahoo cross-origin), 5-min cache,
  `?refresh=1` forces. Each widget's ↻ pulls live; value + day change (from
  `meta.chartPreviousClose`) + a 30-pt close-series sparkline.
- **Watchlist `Chg (1d)` column** comes from `dashboard_data._price_changes()` —
  day-over-day: latest history.db run vs the most recent run on an EARLIER calendar
  day (same-day re-runs don't count).
- **Watchlist is strictly the 'AT LOWER BOUNDARY' band** (already noted); the top
  **Refresh** regenerates from the latest CAPTURED prices (history.db) — it does NOT
  fetch live market prices; only the Intelligence ↻ does.
- **Review-deck link → in-app gallery.** `dashboard_server.js` serves the existing
  `pipeline_app/review_deck.html` at `/decks/review-deck` (rewriting relative
  `asset?p=` to the absolute `/asset?p=` route) with an `/asset` image proxy
  whitelisted to repo + Downloads. Architecture/alert-rules stay raw `.pptx`
  downloads — **no slide renderer (LibreOffice) is available**, so there's no image
  gallery for pptx-only diagram decks.
- **Architecture deck now reflects the Investment Dashboard.**
  `oneoff/add_investment_dashboard_slide_2026-07-17.py` inserts a dedicated readable slide
  (after the flow diagram) with click-through links to `localhost:4600`; the Agents
  slide (`oneoff/add_agents_slide_2026-07-12.py`) now lists all 8 agents with a
  count-adaptive row layout. AZN's Alert Low = the 12,218 horizontal support (the
  old stray-blue note was stale — resolved).

## Validation audit (2026-07-17) — Alert High no longer goes stale independently

The `validation` agent's first run found a live write-path bug: `is_noise_refresh()`
(update_master_sheet.py) gated the whole parallel-channel row on the **Alert Low**
delta alone. The two rails of a channel drift at different rates (different slopes),
so an Alert Low within the 3% noise band would skip BOTH cell writes and let a
drifted **Alert High** go stale indefinitely. **Fixed:** `is_noise_refresh` now
takes `alert_high`; a parallel row is 'noise' only when BOTH rails are within
threshold (and if the sheet has no comparable High yet, that's a change, not noise).
Single-sided reads (`single_low`) still check only the Low — unchanged. Clean A/B
(old vs new code, same workbook copy, same `channel_results_tmp.json`): **7 rows
moved noise→applied — ADM (High 4258→3933), BARC (491→578), TSCO (598→526), FCIT
(346→363), QQ, SHEL, ULVR — zero regressions, zero other applied rows changed.** The
live sheet's stale Highs correct themselves on the next pipeline run.

FIXED from that audit (all A/B-verified, applied to the live sheet + committed):
- **AV. (Aviva) — pale-cyan channel now detected.** `channel_blue_mask` was widened to
  also match the pale-cyan channel-rail colour (~#B2EBF2, R150-200/G210-245/B220-250)
  alongside the saturated blue. AV.'s ascending channel is drawn in that pale cyan, which
  the saturated-only mask couldn't see — so the only "blue" it found was a manually-drawn
  "Sell at 693" #2962FF horizontal ray, shipped as Alert High 693.54. Now reads the real
  rails: Alert High **869.16** (top rail), and PNN gains a real second rail too
  (single_high → parallel, Alert Low 456.83). Full-batch A/B (committed vs widened, same
  342 images, independently re-run 2026-07-17): **only AV. + PNN change structurally, 0
  rails lost, ~7 tickers shift <1 price unit (sub-0.02%, RANSAC reseed noise), 0 new false
  lines.** Do NOT tighten the pale-cyan band without re-measuring. (This also resolves the
  AZN-class "stray-blue" concern for good.)
- **Single-blue-boundary label (MKS/TPK/MTRO/HSBA/BBY/III…).** When only ONE blue rail
  survives the plausibility gates (its partner filtered as implausible on a wide/old
  channel), `process_one` no longer calls it a breakout/breakdown — it can't tell top from
  bottom — and labels it 'single blue boundary above/below price (nearest-line default)'.
  Alert LEVELS are unchanged (computed earlier); only the pattern text. `refresh_soi_sections
  .PATTERN_LABELS` gained the matching 'Single blue boundary' entry so the SOI Pattern
  column doesn't go blank.

Still OPEN from that audit (data/workflow, NOT a code fix — need a redraw, not a patch):
- **IAG** carries a ~9x-stale Alert Low (49.19 vs price ~444) — the captured layout has
  no markup so the read is honestly rejected and the OLD level is inherited (working as
  designed). Needs the user to re-mark IAG's Monthly view (the yellow COVID-low trendline)
  or clear the level by hand. **PFD** similar (axis OCR [300-800] doesn't bracket price
  191; unusual chart with a huge historic range) — a dedicated look, not a quick patch.

## Dashboard: in-app architecture view + live Watchlist prices (2026-07-18)

Three user requests, all committed:

- **Architecture deck now opens IN THE APP as a readable HTML view, not a raw
  .pptx download.** The nav "Architecture" link pointed at `/decks/architecture.pptx`,
  which (no OneDrive share link pasted, no LibreOffice renderer available) only ever
  streamed the raw file for the browser to download — the user reported it "just not
  working". `scripts/render_architecture_html.py` reads the SAME
  `Financial_Data_Pipeline_Architecture.pptx` the deck scripts produce and lays every
  slide out as absolutely-positioned HTML boxes from each shape's EMU geometry, fill,
  border and per-run font styling, so it stays in sync with the deck automatically
  (no hand-maintained second copy). Fully responsive with pure CSS: each slide is a
  `container-type:inline-size` stage at the slide's aspect ratio, children positioned
  in %, fonts in `cqw`; lines/connectors drawn in one SVG per slide whose viewBox is
  the slide's EMU coords. `dashboard_server.js` serves it at `/decks/architecture`
  (`ensureArchitectureHtml()` regenerates lazily when the .pptx is newer than the
  cached `dashboard_app/architecture.html`, which is gitignored). Front-end nav link
  updated to `/decks/architecture`. The old `/decks/architecture.pptx` route still
  works as a raw-download fallback. A few boxes clip a long single-line label
  (e.g. "TradingView"→"TradingVie…") because the box height is fixed and text is
  `overflow:hidden` — acceptable, the diagram is faithful and readable; don't switch
  to `overflow:visible` on 131 boxes (they'd overlap).
- **Watchlist now has a LIVE-price refresh, separate from the top Refresh button.**
  The top **Refresh** re-derives the watchlist from the last CAPTURED prices
  (history.db); the new **"Live prices"** ↻ button on the Watchlist screen pulls LIVE
  market prices from Yahoo on demand. Server route `/api/watchlist-live`
  (`getWatchlistLive()` in `dashboard_server.js`) reads the current watchlist tickers
  from `watchlist.json`, maps each to a Yahoo symbol (default rule = LSE equity
  `<TICKER>.L` with `.`→`-`, quoted in pence to match the sheet), and fetches
  `regularMarketPrice` + `chartPreviousClose` via the same Yahoo chart API the
  Intelligence screen uses. `config.json → watchlistYahooSymbols` overrides the rule;
  `''` marks a ticker as having NO pence-denominated live source (commodities are
  USD/oz — PALL/PLAT/GOLD/SILVER/COPP/NATGAS/UKOIL are set to `''` and report
  "no live source" rather than being guessed). The front end overlays the live
  price/change onto the rows, recomputes proximity, and stamps "Live · HH:MM · N no
  live source". Verified end-to-end in-browser: 5 LSE equities updated with real
  day-changes, Palladium correctly unsupported.
- **Architecture deck old-version deletion.** Item #3 of the request. The newest
  architecture editor (`oneoff/add_investment_dashboard_slide_2026-07-17.py`) already called
  `purge_old_versions`, and the end-of-pipeline `purge_output_duplicates.py` sweeps the
  architecture deck too. Added the same `purge_old_versions(deck_path)` call to
  `oneoff/add_agents_slide_2026-07-12.py` (which CLAUDE.md says is re-run on request) for
  consistency. Alert_Rules_Model (`build_rules_deck.py`) and Investment_Review_Deck
  (`build_review_deck.py`) already purged — confirmed, no change needed.

- **Widget WIDTH vs SIZE, and medium widths halved.** In the dashboard grid,
  `size` (S/M/L) is the widget's HEIGHT (grid-row span 1/2/3) and `span` is its WIDTH
  (columns out of 12) — they are independent. User asked to "reduce width of medium by
  50%": the four `size:'M'` widgets had their `span` halved — Portfolio Value Over Time
  8→4, and Alert Status / Targets / Relevant News 4→2 each. (If a widget's width ever
  needs changing, edit its `span`, not its `size`.)

- **"Pipeline" nav link now self-heals (was "site cannot be reached").** The link opens
  the Investment Production Centre — a SEPARATE server (`pipeline_app_server.js`, port
  `CFG.appPort`=4590). `Run Investment Dashboard.bat` tries to start it in the
  background, but if that failed / it crashed / node was killed, clicking Pipeline hit a
  dead port. The link now points at the dashboard's own **`/pipeline`** route, which
  `ensurePipelineApp()` uses to check port 4590 (raw TCP via `net.connect`), spawn
  `pipeline_app_server.js` detached if it's down, poll up to ~8s for it to listen, then
  302-redirect. Verified: with 4590 down, hitting /pipeline starts it and redirects to a
  live HTTP 200. The old direct `http://localhost:4590` link is gone from the front end.

- **Single-instance guard: opening a new dashboard supersedes old tabs.** User wanted
  old dashboard tabs closed when a new one opens. `singleInstanceGuard()` (front end,
  runs first in `init`) uses a `BroadcastChannel('investment-dashboard')`: the newest tab
  (largest `ts`) broadcasts a claim, every older tab calls `window.close()` and — since
  **browsers refuse to close a tab the USER opened** (only script-opened windows can be
  closed) — falls back to a full-screen "this tab is inactive, reopened elsewhere" cover
  and closes its SSE (`_sse`) so it stops polling. So user-opened stale tabs go quiet and
  clearly inactive rather than literally vanishing; a script-opened tab does close. Do
  NOT expect `window.close()` alone to remove a normal tab — the cover is the reliable
  part. Verified in-browser: opening a 2nd tab dropped the cover on the 1st.
  - **Superseded as the primary mechanism 2026-07-20 — the launcher now opens a
    DEDICATED Chrome profile and closes stale tabs for real over CDP** (user accepted
    the trade-off). Chrome only exposes CDP when started with
    `--remote-debugging-port`, and it will not do that for an already-running normal
    profile, so `scripts/dashboard_open.js` runs its own profile under
    `data/dashboard-chrome-profile/` (gitignored — a whole browser profile) on port
    **9333** (NOT 9222, that is TradingView Desktop's) in an `--app` window. Second
    launch: find the profile already up, `/json/close` every page target on the
    dashboard's ORIGIN (sweeping `/decks/architecture`, `/pipeline` too), open one
    fresh tab. `Run Investment Dashboard.bat` calls it; it waits for the server's port
    first so the tab can't land on a connection error, and falls back to the default
    browser when Chrome isn't installed. **The BroadcastChannel cover stays** — it is
    the only thing that helps a tab the user opened by hand in their everyday Chrome.
    The cost Paul accepted: the dashboard runs in a profile signed into nothing.

- **A SERVER MUST BIND ITS PORT BEFORE DOING ANY STARTUP WORK THAT WRITES TO DISK
  (2026-07-20).** `Run Investment Dashboard.bat` started `dashboard_server.js`
  unconditionally, so relaunching it while one was already up died with a raw
  `EADDRINUSE` stack trace on 4600 — and because `regenerate('startup')` ran BEFORE
  `server.listen()`, the instance that was about to die had already rewritten the
  dashboard JSON, racing the live server that owns those files. Now the bat
  port-checks 4600 exactly as it already did 4590 (curl, reuse the running one and
  close the window), and the server binds FIRST, regenerating only from inside the
  `listen` callback, with an `error` handler that reports EADDRINUSE in plain English
  and exits 0. Both paths tested: cold start binds → regenerates → serves 200; a
  second launch reuses the live server and leaves its data alone. **Any new
  long-running server here (the Production Centre included) should follow the same
  order — bind, then do startup work** — otherwise a duplicate launch corrupts the
  running instance's state on its way out.

## Dashboard: income-fund positions, real dividend dates, Fidelity-accounts total (2026-07-19)

- **Total Portfolio Value and the Portfolio Value Over Time series read Wealth Summary
  ROW 33, 'Fidelity accounts'** (user decision 2026-07-19), not holdings + income funds
  and not the account block summed by hand — `WS_FIDELITY_TOTAL_ROW = 33` is the sheet's
  own headline Fidelity line, so the dashboard, the sheet and the Finance Google Sheet
  cannot drift apart. £3.76M for Jul 2026. Row 33 spans the account block (rows 5-13)
  PLUS the two Fidelity SIPPs, which live up in the pension blocks (Paul row 15, Susan
  row 25) — `WS_ACCOUNT_ROWS` includes them so the Accounts widget ties back to row 33
  (1,938,320 + 1,823,862 = 3,762,182). Those SIPP labels carry no holder name
  ('  SIPP Savings - Fidelity (2000001606)'), hence the `WS_SIPP_ROWS` row->holder map.
- **Accounts widget (Overview, medium)** — `overview.accounts` lists each Fidelity
  account row for the **last FULL month** (the current month is still accruing, so
  comparing it to the previous month understates the increase) with value and the
  month-on-month increase; the widget rewrites its own title to "Accounts — <month>".
  Account/wrapper/number are parsed out of the sheet's `'  Investment ISA (Paul)
  (SANX002282)'` label format. NOTE: most rows show a £0 increase because the Wealth
  Summary carries the same figure across months for any account that hasn't been
  re-imported — that's the sheet, not the widget.
- **FAMILY ACCOUNTS ONLY** (user decision 2026-07-19). The Fidelity export covers
  accounts administered for wider relatives (Dorothy Wall, Olive Elizabeth Sangha,
  Freda Hibbert) as well as the family's. `FAMILY_HOLDERS = {Paul, Susan, Liam,
  Jayne}` filters `_fidelity_positions()` at source, so every downstream total is the
  family's: 13 -> 12 investments (Vodafone was Olive's), 44 -> 22 income positions,
  and the income-fund P&L flipped −£29,985 -> +£6,178 because the losses sat in the
  relatives' accounts. The Wealth Summary accounts block and the History sheet were
  already family-only, so those needed no change.
- **Portfolio → Income Funds is now per fund PER ACCOUNT** (`portfolio.income_positions`,
  44 rows) with Account, Wrapper, P&L, Yield, Last Income, Ex-Div and Paid, plus a
  **sticky total row that totals the FILTERED rows**. The sheet-level `income_funds`
  roll-up is untouched — Overview's metrics, Relevant News and the Historic dividends
  table all still read it, and rebuilding those off per-account rows would have changed
  every headline number.
  - **Dividend dates come from `TransactionHistory*.csv`, the only source of them.**
    `_fund_income_events()` reads the 'Income Received' rows: Fidelity's **Order date IS
    the ex-dividend date and Completion date the payment date**, keyed by (fund name,
    account number). The export only covers ~30 days, so this is "the latest payment",
    not a history — a holding with no row shows no dates rather than an inferred one
    (23 of 44 have dates; the Acc units legitimately never pay cash).
  - **Yield is derived from that payment** (`amount x 12 / holdings`, these funds
    distribute monthly), NOT matched to the Income Funds sheet by name. A token-overlap
    matcher was tried and REVERTED: the generic words ('high', 'yield', 'bond', 'inc',
    'global') outnumber the distinguishing ones, so it paired AXA and IFSL Marlborough
    with other managers' funds. Don't reintroduce name matching for these.
- **Relevant News rows DESCRIBE the event (user request 2026-07-19).** A date + amount
  alone doesn't say what happens. Each row now carries a `description`: equities get
  "Goes ex-dividend in N days — Xp per share (Y% yield). Hold through <date> to qualify"
  (or "Last went ex-dividend <date> (N days ago)… next date not yet published" when
  Base Data's ex-div date is in the past — most of them are); income funds describe the
  REAL last distribution from the transaction export ("Monthly income of £9,268 paid
  into cash on 30 Jun 2026 (ex-div 24 Jun), across 6 holdings"), and Acc units say the
  income is reinvested rather than paid out. Equity rows are grouped BY TICKER — the
  same stock in three accounts is one dividend event (Aviva was listed three times).
- **Chart Statistics widget (Overview, medium)** — `overview.chart_stats`, built from
  the same `channel_results_tmp.json` the Activity items read: charts captured, % marked
  up, how many yielded alert levels, and a stacked bar + legend of the pattern mix.
  'Marked up' mirrors `update_master_sheet.marked_up_flag` (a detected pattern OR an axis
  failure means drawings are present), so the two never disagree. Latest run: 340 charts,
  56% marked up, 181 with levels.
- **Cash Available reads the BROKER export, not the Wealth Summary (2026-07-19).** It
  showed £79 — the Wealth Summary's three standalone Cash Account rows — while the real
  figure is **£384,106**, because most of the cash sits INSIDE each ISA/SIPP.
  `_fidelity_cash()` sums the 'Cash available' column (index 12) on each ACCOUNT row of
  AccountSummary's detail section, family holders only (the broker's own £384,647 total
  includes the relatives' £541). The metric carries the per-account breakdown as `rows`
  (tooltip on the figure) and falls back to the Wealth Summary rows if the export is
  missing.
- **Row unit is 44px, not 104px (2026-07-19).** The user asked for Small +50% height.
  Heights are grid-row spans, so the unit shrank and every span scaled: S 3 (164px),
  M 4 (224px), L 6 (344px), XL 16 (944px) — M/L/XL are byte-identical to before, only
  S changed. `height(n) = 44n + 16(n-1)`; keep that formula in mind before changing
  `--row` again.
- **The Design screen RESIZES widgets — stop editing spans in code (user request
  2026-07-19).** Every row now has a Height (S/M/L/XL) and a Width (in twelfths:
  1 / 1.5 / 2 / 3 / 4 / 6 / 12) dropdown alongside its order number, plus ▲/▼ move buttons; **everything saves
  the moment it changes — there is no Apply button** (user request 2026-07-19). Overrides live in `localStorage['dashboard:sizes:v1']` as `{id:{size,span}}` and
  `renderGrid` reads them through `sizeOf()`/`spanOf()`. **Only values that DIFFER from
  the code default are stored**, so a later default change still reaches any widget the
  user never resized — and 'Reset all to default' clears both the order and the sizes.
  A size/width request from the user is now a question of whether the DEFAULT should
  change, not the only way to change it.
- **The grid is 24 columns, not 12 (2026-07-19).** The user asked for Smalls at 1.5
  columns and Mediums at 3 — grid spans are integers, so the track count doubled and
  every span with it. Widths in 24ths: Small 3, Medium 6, Large 12, XL 24. Talk about
  widths in the user's twelfths (1.5 / 3 / 6 / 12) but write them doubled in `span`.
- **Overview rows are banded by size (user request 2026-07-19):** row 1 = the six
  Small metric cards, rows 2-3 = the Mediums (Portfolio Value Over Time + Accounts at
  span 6; Alert Status / Targets / Relevant News at span 4), row 4 = Activity. A run of
  half-width Smalls leaves the row part-full, so `renderGrid()` appends an inert
  `.grid-spacer` to consume the leftover columns. **Forcing `grid-column-start` on the
  first Medium does NOT work** — CSS grid will still place a later auto item back into
  the gap (Accounts jumped up beside the Smalls), and because `.widget` uses
  `grid-column: span var(--span)`, setting only the start collapses the item to one
  column. The spacer is the reliable mechanism.
- **Layout (user requests, same day):** Overview small-widget order is Total Portfolio
  Value, Gain, Trading Profit, Monthly Dividend, Accumulative, Cash — with Portfolio
  Value Over Time + Alert Status on the row beneath; Activity narrowed to `span:4` to
  match Portfolio Value Over Time. **Saved layouts are keyed `dashboard:layout:v2:` —
  bump that version whenever the DEFAULT order changes**, because `getOrder()` lets a
  saved layout win for widgets it already lists, so an existing browser would otherwise
  keep the old arrangement for ever. Long identity cells are capped (`td.cell-name`
  260px, 190px in the sold table) with the full name on hover.

## Dashboard Portfolio holdings come from the BROKER export, not the workbook (2026-07-19)

The Portfolio screen listed 5 positions off the Investments sheet's `Holdings (£)`
column and was missing real holdings (Centrica, AstraZeneca, Anglo American, Glencore,
Vodafone) — that column goes stale and only carries a total, so it can't supply
Account, Wrapper or book cost either. `dashboard_data._fidelity_positions()` now reads
`~/Downloads/AccountSummary.csv`'s **'View all account details'** section (one row per
holding-per-account) as the authoritative list of what is held. The section ABOVE it is
an aggregate across accounts — reading both double-counts, so the parser only starts at
that marker line. Per row: Account = account holder's first name (matches the History
sheet's 'Paul'/'Susan' convention), Wrapper = Product, Value £ (col 10), Book cost
(col 13), **Buy Price = book cost ÷ quantity × 100** (pence, to match the pence-quoted
current price), **P&L if sold today = the broker's own Gain/loss £ (col 14)** — not
derived, so it can't drift from the statement. Each row is joined back to the
Investments sheet BY TICKER for Alert Low/High, Type and the TradingView link
(`AS_TICKER_ALIASES` maps broker tickers to sheet tickers: `AV.`→`AV`, the platinum
ETC `SPLT`→`PLAT`, `SPDM`→`PALL`; `AS_FUND_AS_INVESTMENT` pulls WS Guinness in by name
since it has no ticker). Funds with no ticker are skipped — income funds keep coming
from the Income Funds sheet, which is the only source of their yield/income. **If the
export is absent the old workbook-Holdings path still runs**, so the dashboard never
goes blank. Known cosmetic mismatch: PLAT/PALL rows show the broker's ETC buy price
against the METAL's chart price (the alert levels are on the metal) — the £ P&L is
still the broker's, which is the number that matters.

Front-end, same batch (user requests): small Overview widgets halved in width
(`span:2`→`span:1`) with `.metric` given `overflow:hidden` so a graphic can't spill out
of the fixed 104px card — the graphic budget in an S card is only ~22px, so the
portfolio sparkline is 52×18 and Cash Available's ring dropped to 18px with its % moved
into the sub line (a centre label is illegible at that size). Holdings and Sold Positions
tables went `L`→`XL` (grid-row span 8) to fill the page.

## Dashboard wired into the pipeline as one flow (2026-07-18)

The Investment Dashboard used to be a separate downstream app: `dashboard_data.py`
regenerated its JSON only on the dashboard server's own startup / manual Refresh, so
a pipeline run did NOT update it. User asked for a single logic flow tied into the
pipeline. Now `run_full_pipeline.js`'s `finally` block runs `refreshDashboard()`
right after `recordHistory()` (dashboard reads `history.db`, so it must run after the
record step): it runs `python dashboard_data.py` to regenerate the JSON from the
just-updated workbook + history.db, then best-effort `curl`s the dashboard's new
**`/pipeline-updated`** route (`dashboard_server.js`), which broadcasts an SSE
`refresh` so any open dashboard reloads live. It's a QUIET sub-step (no numbered
`=== Step N/M ===` marker — same as `recordHistory`/`integrate_spending_tabs`), so
the Production Centre stage count is unchanged. `spending_summary.py` already runs
in-pipeline (Production Centre stage 2 → mirrored into the master by
`integrate_spending_tabs.py`), so the full chain is now: preflight → spending build →
TradingView capture → OCR → master sheet (+SOI +spending tabs) → history.db →
**dashboard data + notify** → review deck → verify → cleanup. `/pipeline-updated`
only broadcasts (no re-generation — the pipeline already wrote the files); the dashboard
isn't required to be running (data lands on disk regardless).

## spending_summary month anchors are DERIVED, never hardcoded (2026-07-19)

`spending_summary.py` was hand-calibrated with literal `2026-05` / `2026-06` /
`2026-07` periods scattered through it, and they went wrong the moment the calendar
moved past them (bug #20, 2026-07-18, was one symptom). On the 18 Jul 2026 export the
user reported the **Income table missing May**, and the cause was exactly this: May
fell in the GAP between the pinned Jan–Apr history tables and the 60-day transaction
window, so it was classified as an *actual of zero* and never estimated, while **July
was likewise treated as a complete actual and shipped its part-month £192**.

The fix is a single `MonthAnchors` object (`resolve_anchors()`), built from the
AccountSummary `Export date` header, that every month boundary now reads from. **The
governing distinction, which any future edit must respect:**

- **DATA pinned to the month it was measured** keeps its literal month for ever — the
  Jan–Apr `load_spend_history()`/`load_income_history()` tables, `load_history()`'s
  wealth series, and the May-2026 pension/house/car estimates (now anchored on the
  named `ESTIMATES_AS_OF` constant). These project forward from their own as-of date,
  so the calendar advancing never invalidates them.
- **ANCHORS describing where the report sits** are always derived: `year`, `months`,
  `data_month` (the snapshot month — the anchor for holdings, prices, account values
  and the Targets "what's it worth now" reads), `partial_month` + `partial_scale` (the
  snapshot lands mid-month, so `18/31` is computed, not typed), `hold_from`,
  `hist_cutoff` and `wealth_cutoff` (both = one month past the end of the pinned
  tables, so extending a table moves the boundary by itself).

Verified by re-resolving against synthetic export dates (18 Jul, 31 Jul, 3 Sep 2026,
5 Jan 2027): the data month, partial flag, hold-from and actual/estimated split all
track the export, and a 2027 export raises a loud WARNING that the pinned 2026 tables
no longer contribute (they need a fresh full-year export — the anchors can't invent
that data, and silently estimating the whole year would be worse).

Three things worth keeping in mind:

- **The actual/estimated split is computed PER PIVOT**, from that pivot's own month
  coverage. The spending sources and the Fidelity export cover different months; when
  the bank exports are absent entirely, a shared split made their months read as
  actual zeroes — which blanked the salary row for June and dragged the median down
  for every other month. `anchors.split_months(tx_months)` is called separately for
  the spend pivot and the income pivots.
- **A dividend already banked in the partial month is never scaled.** The partial-month
  rule scales a raw value up by `1/partial_scale`, which is right for a continuous flow
  (spend, monthly fund income) and wrong for a discrete payment — it would invent cash.
  The stock branches keep a banked payment as-is instead. Same reason `Salary` keeps a
  real payment but falls back to the median when the payslip hasn't landed yet (this
  subsumes the old hardcoded "June Salary estimate" special case).
- **`EQUITY_DIVIDENDS_INLINE` is now module-level and `EQUITY_DIVIDENDS_ANNUAL` sums
  it.** The Targets table used to carry a duplicated `413 + 168 + 178 + 315 + 78`
  literal that had to be re-added by hand whenever a payment changed. The annual income
  figures are likewise now row sums of the already-estimated pivots rather than
  month-literal arithmetic — that alone corrected Income per month £15,418 → £22,359
  (the old calculation under-counted the gap months and used a `salary_may * 8` hack).

Measured against a pre-change baseline on the same inputs: the Wealth Summary's gap
months fill in (May and July across every income/spend row), Paul's SIPP reads its
July snapshot value in the July column instead of stamping it onto May, SANQ000468's
back-ramp shifts to end at the real snapshot month, and the Targets SIPP/ISA rows come
out byte-identical (they now read the July column, which holds what the May column
used to).

**Applied to the live workbook 2026-07-20** — the fix had only ever been run to a temp
file, so `Stocks_Buy_Strategy.xlsx` still carried the reported bug (Salary/SIPP blank
for May, July stuck on its part-month £192). Rebuilt `spending_summary.xlsx` →
`integrate_spending_tabs.py` → `dashboard_data.py`. Note the Amex/Barclays exports were
consumed on 2026-07-12 and are gone, so the spend side is estimated throughout; June
Salary moved 6535 → the 6379 median, which is the honest estimate for a month with no
bank export behind it. **The Finance Google Sheet has NOT been re-synced** — that is
still outstanding, along with the 'Strategic' rename.

- **The gap months between the pinned history and the snapshot are INTERPOLATED, not
  projected (2026-07-20).** Rebuilding surfaced a second defect the anchors work exposed
  rather than caused: months after the history tables end but BEFORE the snapshot month
  were grown forward on fund income — a branch written when the snapshot was the very
  next month and the gap was empty. With a July snapshot the gap opened and the
  projection overshot: Paul's SIPP read £1.498M for June against a July snapshot of
  £1.432M, so the row climbed above the snapshot and fell back onto it, and rows 18/26/
  33/34 inherited the spike. `sheet_assets.py` now interpolates across the gap so the
  series is monotone between the two months actually MEASURED and lands exactly on the
  snapshot; months after the snapshot still grow on income. Measured A/B on the same
  inputs: **6 cells move, all June**. May is NOT affected — 2026-05 is a pinned value in
  `load_history()`'s 'Paul Pension' table (1,490,236), real data, not a projection.
  Consequence worth knowing: the dashboard's Gain-vs-last-month now picks June and reads
  **−£27,266** (half of the real May→July fall) where it used to read +£48,856 off a
  flat carry-forward. The direction is now right; the magnitude is an interpolation.

## Payslips screen + pension allowance (2026-07-19)

A **Payslips** screen was added to the Investment Dashboard, and payslip loading was
deliberately kept OUT of the pipeline (user decision 2026-07-19: "we don't need to
include payslip uploads in the pipeline — we'll load it via this new payslips screen").

- **Data**: `dashboard_data.build_payslips()` reads the workbook's hand-maintained
  `Payslip Summary` tab (band rows and TOTAL rows are skipped — they're formulas and
  read as None locally, the same constraint as everything else in that file) and emits
  `payslips.json`: every row, per-tax-year totals, and the allowance block below.
- **Upload**: the screen's button POSTs the PDF as a raw body to
  `/api/payslips/upload` (no multipart parser — `dashboard_server.js` has no deps),
  which runs `scripts/payslip_ingest.py` and, on success, regenerates the dashboard
  data so the row appears at once. `?dry=1` exercises the whole path without writing.
  `payslip_ingest.py` keeps extraction and writing separate ON PURPOSE: a payslip whose
  layout it doesn't recognise raises `PayslipParseError` and the UI shows why, rather
  than writing a guessed number into a live financial sheet. Field labels live in
  `FIELD_LABELS` so a new payroll layout is a one-line addition; `--dump` prints the
  PDF text to find the right labels. A new row copies the styling of the row above it
  (the "match the neighbour" rule), backs the workbook up, and the source PDF is
  archived to `data/payslips/`. **An image-only/scanned PDF is rejected with a clear
  message** — there is no OCR in this path.
- **Pension annual allowance panel, at the top of the screen** (user request). Shows
  what's left this tax year and **which month contributions must stop**.
  `dashboard_data.pension_allowance()` consumes carry-forward HMRC's way — the year's
  own allowance first, then unused allowance from the previous three years OLDEST
  FIRST, walking the years in order so each year's carry-forward reflects what later
  years already ate. Current read: £60,000 + £13,263 carried from 2023/24 = £73,263
  available, £12,789 used, **£60,474 left, stop after Feb 2027**.
  - **THIS IS ARITHMETIC ON PAUL'S OWN PAYSLIPS, NOT TAX ADVICE**, and the three things
    it cannot know are listed on the card rather than buried: the **taper** (adjusted
    income over £260k, which covers all income and can't be read off a payslip), the
    **MPAA**, and **contributions paid outside payroll** (a personal payment into the
    SIPP is an input to the same allowance). Any one of them changes the answer.
  - A pay month with no payslip uploaded is still counted at the current rate and
    flagged `est` — dropping it would understate the year's input and overstate what's
    left. That's why June appears in the projection.
- The pipeline top-bar buttons (`Run <date>`, `Refresh`) are **hidden on Payslips**
  via `NO_PIPELINE_ACTIONS` — nothing on that screen comes from a pipeline run.
  Note they are hidden with `style.display`, not the `hidden` attribute: `.btn` sets
  `display:inline-flex`, which beats the UA `[hidden]{display:none}` rule.

## Dashboard batch (2026-07-19, same session)

- **'Long Term' investment type renamed to 'Strategic'** (user request). One-off
  `oneoff/rename_type_strategic_2026-07-19.py` updated the Investments `Type` cell and any
  data-validation dropdown offering the old wording; `dashboard_data` classifies on
  `'strategic'` but still accepts `'long term'` so an un-migrated sheet reads the same,
  and the Targets breakdown label is now 'Strategic'.
- **Gain vs Last Month now measures the last month with FRESH data, and names it in
  the widget title** ("Gain — May 2026"). Two conditions, both needed: the month must
  be COMPLETE (the current one is still accruing — July read +£193) *and* it must have
  MOVED. The Wealth Summary carries the previous figure forward for any account not
  re-imported, so June sat at May's exact total and a June-vs-May reading was **£0** —
  just as misleading. The real answer is **May vs Apr, +£48,856**, and the caveat says
  how many months since carry the same figure forward.
- **Total Income** small metric on Overview, right of Accumulative Fund Income:
  monthly dividend + accumulative fund income (£26.1k/mo, £313.1k/yr) — the two cards
  to its left added up, so the split stays visible.
- **Accounts widget has a name filter** (holder, wrapper or account number; the account
  number is only in the row tooltip, so filtering on it is the quickest way to find
  one). **The total follows the filter** — a total ignoring it would sit under a short
  list and read as their sum.
- **Widget height scale gained a step.** There was nothing between L (6 rows, 344px)
  and XL (16 rows, 944px), so "make Activity twice as tall" had no token to land on.
  **XL is now 12 rows and the old 16-row XL became XXL**, which keeps the tokens
  ascending and leaves Holdings / Sold Positions / the Payslips table at their existing
  height. Activity went L → XL, exactly the requested doubling. `LAYOUT_KEY` bumped to
  `v4` for the Total Income card.

## 'Charts to mark up' opens the layout in TradingView Desktop (2026-07-19)

The list moved from the **Activity** widget to **Chart Statistics** (user request) —
that widget is about what the detection run found, and its '149 not marked up' figure
is exactly what the list enumerates. Chart Statistics went `M`→`L` and its body
scrolls; the status line sits ABOVE the chip list because the list scrolls inside the
widget and a message underneath it was off-screen exactly when it mattered.

**Clicking a ticker now drives the running TradingView Desktop to that chart's SAVED
LAYOUT** rather than opening a browser tab — the point is to land on the layout ready
to DRAW on, which a browser tab can't do. `scripts/tv_open.js` is the only dashboard
code that talks to TradingView; `POST /api/open-chart {chart_id, ticker, layout}`
drives it and `GET /api/tradingview-status` reports reachability.

- **It obeys the two standing TradingView rules** (see the top of this file):
  `ensureAutosaveDisabled()` runs FIRST — auto-save silently persists whatever view
  state it finds back into the saved layout, which once overwrote every hand-zoomed
  chart — and there is **NO view reset/refit of any kind**; the saved layout is the
  view to show. Navigation is a plain `window.location.assign` to the layout's chart
  URL plus the unsaved-changes dismissal, the same approach `layoutSwitch()` uses,
  because `loadChartFromServer(id)` silently no-ops on a numeric layout id.
- **ticker → layout comes from `layout_manifest_tmp.json`** (`{id, chartId, name,
  ticker}` per captured pane) — the only place that mapping exists. `dashboard_data`
  attaches `chart_id` / `layout` / `layout_id` to each `charts_to_markup` entry; all
  108 currently resolve. A ticker with no captured layout keeps the browser link.
- **TradingView Desktop must be running with `--remote-debugging-port=9222`.** It is
  NOT by default — the app can be open (it was, PID 4044) with the port closed. The
  route returns a 503 and a plain-English message saying exactly that rather than
  hanging or throwing. **The happy path has not been exercised end to end yet**: doing
  so needs TradingView relaunched with the debug port, which would drop the user's
  running session, so it was left for them to confirm.

## Codebase banner + consolidation pass (2026-07-19)

- **Overview carries a live "Codebase" banner** (user request): lines of code, file
  count, last-updated date and a per-language split, above the widget grid on the
  Overview screen only. Both figures are COMPUTED, never typed — `scripts/codebase_stats.js`
  (served at `/api/codebase`, 60s cache) walks the tree counting hand-written source
  and takes "last updated" from the last **git commit**, not a file mtime (a mtime
  changes when a generated file is rewritten or OneDrive touches something, which
  would claim an update that never happened). Three things are excluded on purpose,
  or the number measures the wrong thing: dependencies (`node_modules`), GENERATED
  artefacts in the tree (`architecture.html` / `alert_rules.html` / `review_deck.html`
  are rendered from the .pptx decks — 1.9 MB of machine output — plus
  `package-lock.json`), and run scratch (`*_tmp.json`, `logs/`, `screenshots/`,
  `data/`, `__pycache__`). Docs (`.md`) are out too: CLAUDE.md alone is thousands of
  lines of prose. Reads ~19,300 lines / 73 files today.
  - The banner is hidden off-Overview with the `hidden` attribute plus an explicit
    **`.codebanner[hidden]{display:none}`** rule — `display:flex` beats the UA
    `[hidden]` rule, the same trap the top-bar `.btn` hit (it showed on every screen
    until that rule was added).

- **The trading-profit YEAR is derived, not a literal.** `dashboard_data.py` pinned
  `2026` in four places and the front end hardcoded it in two widget titles, so on
  1 Jan the figure would have silently gone to £0 under a heading still saying 2026.
  Now `profit_year` = the calendar year of the most recent completed sale (not
  today's date — an early-January dashboard should still report the year that has
  trades in it), carried in the JSON as `trading_profit.year` /
  `historic.summary.year`, and both widgets rewrite their own title from it. Keys
  renamed `trading_profit_2026`→`trading_profit`, `total_profit_2026`→`total_profit`.
  Values unchanged (£38,255.54 / 31 sells, verified before and after).

- **Duplication removed (same behaviour, one implementation each):**
  - `scripts/server_util.js` — `serveFile`, the content-type map and `isAllowedAsset`
    were duplicated across `pipeline_app_server.js` and `dashboard_server.js`, with a
    comment in one saying "same rule as the other". That's exactly how a **security**
    rule (the `/asset` proxy whitelist) gets tightened in one place and not the other.
    Both now import it; the surviving `serveFile` is the streaming one.
  - `scripts/yahoo.js` — the Intelligence widgets and the Watchlist "Live prices"
    button each had their own copy of the Yahoo chart URL and the last/previous-close
    extraction. One `fetchQuote(symbol, {range, interval})` + an `r2()` rounder now.
  - `config.recycle_to_bin` — `consume_input_files.py` carried a second, identical
    `SHFILEOPSTRUCTW` block.
  - `xlsx_sheet_copy.copy_cell_style` + `offset_formula` — the six-line style-copy was
    re-inlined in several writers, and `offset_formula` (re-anchor a row-relative
    formula when a row moves — the recurring "formula still points at its old row"
    bug) lived inside an archived one-off that `tests/` imported by module name.
    Both are now in the shared module; `xlsx_sheet_copy.py` is "shared openpyxl
    helpers", no longer just sheet copying.
  - Unused imports dropped from `build_rules_deck.py`, `fidelity_file_classifier.py`,
    `spending_summary.py`.

- **`scripts/oneoff/` — the 18 dated, already-executed migration scripts moved out of
  `scripts/`** (~1,800 lines), leaving the ~24 that actually run. Nothing in the
  pipeline imported them; the only cross-references were documentation. `scripts/oneoff/README.md`
  carries the rules — they are a RECORD, not a template (several target a workbook
  layout that no longer exists, e.g. `oneoff/fix_relx_history_2026-07-10.py` still names the
  pre-rename `'Stocks Buy Strategy'` sheet), and only the ones this file calls
  re-runnable should ever be re-run, now as `python scripts/oneoff/<name>.py`.

All verified after the pass: 27/27 tests pass, all 24 live Python modules import,
both servers start and every route answers (`/`, `/api/codebase`, `/api/intelligence`,
`/api/watchlist-live`, `/files`, `/products`), and `dashboard_data.py` reproduces
identical figures.

## spending_summary.py split into a `spending` package (2026-07-19)

At 4,020 lines it was the largest file in the repo. The reading/categorising/
pivoting half now lives in **`scripts/spending/`**; `spending_summary.py` (2,529
lines) keeps the Excel writer and the CLI:

| module | what's in it |
|---|---|
| `spending/constants.py` | CATEGORIES, ACCOUNT_OWNER/LABELS, FAMILY_ORDER, and the **pinned** figures (`ESTIMATES_AS_OF`, `EQUITY_DIVIDENDS_*`) — with the pinned-data-vs-derived-anchor comment that governs them |
| `spending/anchors.py` | `MonthAnchors` / `resolve_anchors` — every month boundary, derived from the export date |
| `spending/categorise.py` | `categorise_amex` / `categorise_barclays` — pure string rules, no I/O |
| `spending/loaders.py` | the Amex/Barclays/Fidelity CSV readers, the pinned Jan–Apr tables, `load_history` |
| `spending/holdings.py` | `build_holdings`, `apply_pending_holdings`, `build_acc_holdings` |
| `spending/pivots.py` | the monthly pivots, `estimate_future_months`, `build_summary_data` |

**The code was MOVED verbatim — no behaviour was touched.** Proof, and the method
to repeat for any future change here: the workbook build is fully deterministic
(two runs of the unchanged code produced a zero-line diff), so a cell-level dump of
every sheet — value, number format, font/fill/alignment/border, merges, column
widths, row heights, freeze panes, tab colour — is a valid regression gate.
**Both code paths came out at a 0-line diff**: a fresh build (1,836 dump lines) and
a rebuild over an existing workbook exercising `preserve_manual_sheets` (4,710 dump
lines across all four tabs). 27/27 tests pass.

Three **dead** functions were deleted in the same pass (zero references anywhere):
`hdr_style`, `data_cell`, `build_fidelity_fund_pivot`.

### write_excel split into phases (same day)

`write_excel` was the remaining 2,115-line function; it is now **six phases** under
`scripts/spending/sheet/`, and `spending_summary.py` is **467 lines** — the setup,
the context, six calls, and `main`:

| module | lines | phase |
|---|---|---|
| `sheet/style.py` | 146 | the palette — colours, `P`/`F`, the named fills/fonts, `fml`, `HIST_MAP`, and the two data-dependent factories `make_col_fill` / `make_val_font` |
| `sheet/sheet_assets.py` | 646 | summary table, Fidelity accounts, both pension sets, assets, growth/total rows, calculations |
| `sheet/sheet_spending.py` | 231 | Section 1 spend-by-category + reimbursements |
| `sheet/sheet_income.py` | 483 | Sections 2 and 3 (income by account, accumulative holdings) |
| `sheet/sheet_totals.py` | 202 | the cross-section totals that need row numbers from above |
| `sheet/sheet_targets.py` | 478 | Investment Risk Metrics + Targets tables |
| `sheet/sheet_finish.py` | 182 | widths, the split onto the Targets tab, note row, save |

**How it was done, because the method is the safety argument.** The sections share a
mutable row cursor and back-patch each other, so the split was driven by a
**flow-sensitive read/write analysis** rather than by eye: for each candidate
boundary, which names does the section read before binding (its inputs) and which
does it bind that a later section reads before rebinding (its outputs). Every phase
body was then **sliced verbatim as one contiguous region** of the original file
(comments and formatting byte-exact) and the unpack/repack lines around it were
GENERATED from that analysis. Interfaces came out at 5–11 in, 0–5 out.

**Do not hand-maintain those `ctx.x = x` lines carelessly** — they are the phase
contract. A new shared variable needs adding to both the producing phase's repack
and the consuming phase's unpack, or it silently reverts to a stale value.

Getting the analysis right mattered twice: a naive version counted nested-function
parameters and comprehension variables as shared state (inventing `m` and `s`
interfaces that don't exist), and counted leaked loop variables as real outputs.

**Two move-hazards this surfaced, both fixed** — worth expecting on any future move:
- `sheet_targets.py` computes the repo root from `__file__` to write
  `data/spending_dividends.json` (which `dashboard_data.py` reads for the Monthly
  Dividend figure). The file went from two levels below the root to four, so the
  literal `dirname(dirname(...))` had to change — a wrong path here fails SILENTLY.
- Moving a function moves its dependencies: `preserve_manual_sheets` and
  `_resolve_cell_num` went with the phases that use them and needed `os`, `sys`,
  `load_workbook` and `copy_cell_style` imports that the old module already had.
  **Re-run an unresolved-name scan after every move**, not once at the end.

Same gate as the module split, and it is the reason this was safe to do at all:
**both code paths at a 0-line diff** — fresh build (1,836 dump lines) and the
`preserve_manual_sheets` rebuild (4,710 lines, all four tabs) — plus 27/27 tests,
every module importing, and a run from a foreign cwd.

**Import note:** the package is imported as `from spending.constants import …`,
which works from any cwd because Python puts the *script's* directory on `sys.path`
— the pipeline spawns `spending_summary.py` by absolute path from the repo root, and
that was verified. Any new module that needs one of these values must import it from
`spending.*`, not re-declare it.

**Trap this created, now fixed:** moving the 18 one-offs into `scripts/oneoff/`
broke five of them, because `sys.path.insert(0, dirname(__file__))` used to resolve
to `scripts/` and now resolved to `oneoff/`. All five (including both scripts this
file calls re-runnable) now insert `dirname(__file__)/..`. **Any further script
moves must re-check that line.**

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
  `scripts/oneoff/fix_relx_history_2026-07-10.py` (a historical one-off, already executed)
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
