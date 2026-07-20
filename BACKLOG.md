# Backlog — Claude-Weath-Integration

Owned by the `product-owner` agent (`.claude/agents/product-owner.md`).
Single source of truth for pending improvements. Statuses: Proposed →
Approved → In progress → Done (with commit/PR ref). Items marked
**user decision** are blocked until the user decides.

## Now

| Item | Why it matters | Size | Status | Builder |
|---|---|---|---|---|
| Re-sync the Finance Google Sheet | The master workbook was rebuilt 2026-07-20 with the month-anchor + gap-interpolation fixes, and the 'Strategic' rename (2026-07-19) has never reached Google either. The Sheet is currently behind the xlsx on both | S | **User decision** — the sync deletes and re-imports every data tab; needs Paul present | main assistant |
| First investment-analyst run (analyst notes + daily brief draft) | Exercises the agent end to end; populates the deck's Analyst view | M | Proposed | investment-analyst |
| Verify the TradingView chart-open happy path | `POST /api/open-chart` (dashboard → TV Desktop over CDP 9222) has never run end to end. Needs TV Desktop relaunched with `--remote-debugging-port=9222`, and it navigates the live session — any in-progress drawing is lost | S | **User decision** — needs a moment Paul isn't mid-markup | user |

## Next

| Item | Why it matters | Size | Status | Builder |
|---|---|---|---|---|
| Redraw stale/off-frame charts in TradingView | The detection is honest about these — it rejects the read and the OLD level is inherited, which is why they look stale. **IAG** carries a ~9× stale Alert Low (49.19 vs price ~444) with no markup on the captured layout; **PFD**'s axis OCR [300-800] doesn't bracket price 191. Earlier batch: **BEZ, HIK, MNG, SDR** show price ABOVE the detected channel, most likely a leftover OLD drawing competing with the new one | S | **User decision** — redraw in TradingView, then re-run | user |
| Investigate remaining OCR axis-label failures | 2026-07-16/17 recovered most of these: the psm-11 sparse-text fallback (17 charts), the one-tick bracket extrapolation (WPP, SMWH) and the psm-11 bracket retry (24 more, incl. AO World). What's left is genuinely off-frame charts needing a redraw, plus true faint/small-axis failures (BAG, SMT, ULVR, GBPUSD) that need a capture-resolution or preprocessing fix | M | Proposed | data-developer |
| Schedule the daily brief (routine/cron) | Makes the analyst's daily report automatic instead of on-demand | S | Proposed — needs user OK on timing | main assistant |
| Sample payslip PDF to calibrate `FIELD_LABELS` | `payslip_ingest.py` refuses a layout it doesn't recognise rather than guessing a number into a live financial sheet — so the upload button on the Payslips screen can't be proven until one real payslip has been run through it | S | **User decision** — needs a PDF from Paul | user |
| Clear the dead commodity col-J formulas | Brent/Palladium/Copper now get a captured price written into Investments col I, but their old `TVC:`/`CPER` GOOGLEFINANCE formulas still sit in col J showing `#N/A` | S | Proposed — confirm the value-writes land on the next run first | data-developer |

## Later

| Item | Why it matters | Size | Status | Builder |
|---|---|---|---|---|
| "Investment OS" evolution (history-driven analytics, briefings) | Long-term direction from the 2026-07-12 external review | L | Blocked on SQLite history store; never a rebuild of the working pipeline | TBD |

## Done

| Item | Shipped |
|---|---|
| Dashboard opens in a dedicated Chrome profile — stale tabs now actually CLOSE over CDP, instead of only being covered by the BroadcastChannel guard (a browser refuses `window.close()` on a user-opened tab). Paul accepted the trade-off: the dashboard runs in a profile signed into nothing | 2026-07-20, commit ebbedde |
| Month-anchor + gap-interpolation fixes applied to the LIVE workbook — May/July gaps filled, July's part-month £192 replaced by a full-month estimate, June SIPP interpolated between two measured months instead of projected past the snapshot | 2026-07-20, commit 3070b2f |
| Pipeline app reskin finished — the Dashboard's sidebar shell, not just its palette. Rail scroll-spies the four sections; collapses to icons under ~900px | 2026-07-20, commit 36472e2 |
| Payslips screen + pension annual-allowance panel (carry-forward oldest-first, stop-after month), deliberately OUT of the pipeline | 2026-07-19 |
| YELLOW hand-drawn trend lines feed alerts (nearest-line rule each side of price), plus the WEDGE pattern | 2026-07-14 / 2026-07-15 |
| WPP stale Alert Low (1121.92) fixed — occluded "200" axis label meant its read was rejected and it stuck at an ~£11-era level, wrongly parked in BELOW ALERT LOW. Bracket guard now extrapolates one tick past the OCR'd labels; WPP reads 202.65/437.34 (×1.05 → 212.78) and drops out of below-alert. SMWH recovered too | 2026-07-16, commit 7adfa3f |
| Below-alert block widened to the section-table 17-column schema (Pattern, Proximity, Upside %, P/E, Div Yield, …) — user asked for "Trading below alert / Below alert low" to match 'Near Lower Boundary'. Curated Chart Note/Analyst Rating/Notes left blank; verify block-count gate moved C→F | 2026-07-16, commit 7adfa3f |
| OCR/channel hardening: exclude macro (yield/FX/index) symbols, robust Theil-Sen axis fit, price-bracket guard rejecting wrong-scale reads (WPP/NXT/CRDA/PLATINUM), nan-token crash fix | 2026-07-14, commit 700e9bf |
| CDP preflight: pipeline auto-relaunches TradingView with the debug port only when 9222 is down (never force-kills an already-ready TV) | 2026-07-14, commit 700e9bf |
| Re-check 6 price-below-channel rejections | 2026-07-14: RIO now reads (frame-edge fix); ADM/REL/CTEC confirmed genuinely price-below-channel (correct rejects); MGNS/ULVR are OCR-axis failures |
| Channel reads at today's date (line-fit + last-candle x), not the frame edge | 2026-07-13 |
| Capture-corruption fixes: Alt+R reset removed, autosave guard, window-maximize guard | 2026-07-13, commits c1cee17 + 4902c03 |
| PR #8 merged (below-alert gap + full 2026-07-12 batch + architecture-deck updates) | 2026-07-13 |
| data-developer agent (data ingestion + transforms) | 2026-07-13 |
| `.claude/worktrees/` gitignored | 2026-07-12, commit beed2b3 |
| `config.json` single source of truth (Python + Node loaders, 12 consumers switched) | 2026-07-12, commit beed2b3 |
| Per-run manifest snapshots into SQLite (`history_store.py` record/summary/diff, wired into the pipeline) | 2026-07-12 |
| pytest suite for pure logic — 21 tests passing (`tests/test_pure_logic.py`) | 2026-07-12 |
| investment-analyst agent + Analyst-view hook in the review deck | 2026-07-12 |
| (earlier completed items live in CLAUDE.md's Resolved sections) | — |
