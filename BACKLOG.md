# Backlog — Claude-Weath-Integration

Owned by the `product-owner` agent (`.claude/agents/product-owner.md`).
Single source of truth for pending improvements. Statuses: Proposed →
Approved → In progress → Done (with commit/PR ref). Items marked
**user decision** are blocked until the user decides.

## Now

| Item | Why it matters | Size | Status | Builder |
|---|---|---|---|---|
| First investment-analyst run (analyst notes + daily brief draft) | Exercises the new agent end to end; populates the deck's Analyst view | M | Proposed | investment-analyst |

## Next

| Item | Why it matters | Size | Status | Builder |
|---|---|---|---|---|
| Redraw broken-out channels in TradingView | 2026-07-14 re-run after redraws: **BLND & LAND now read** ✅. **BEZ, HIK, MNG, SDR still show price ABOVE the detected channel** (e.g. SDR price 588 vs detected [140–251], MNG 345 vs [215–299]) — looks like a leftover OLD channel drawing competing with the new one; needs checking in TradingView. GSK & COPPER1! fail on axis OCR, not the channel | S | **User decision** — remove stale drawings on BEZ/HIK/MNG/SDR, then re-run | user |
| Investigate remaining OCR axis-label failures | 2026-07-14: 4 of the old failures were macro/index/FX charts, now **excluded** by design; wrong-scale misreads are now **withheld** by the price-bracket guard. 2026-07-16: the bracket guard now extrapolates one tick past the OCR'd labels, recovering **WPP + SMWH** (occluded bottom label) — the other 19 are genuinely off-frame (NXT/KGF/IMB…) and need a redraw, not OCR work. Remaining true faint/small-axis failures (BAG, SMT, ULVR, GBPUSD) still need a capture-resolution or preprocessing fix | M | Proposed | data-developer |
| Detect YELLOW trendlines, not just channel-blue | Several charts draw their levels in yellow (APN's converging trends, LGEN's resistance, DCC's support) — invisible to channel_detect today, so those tickers can never get an alert read. APN's old "read" was actually the blue BUY button (false positive, now rejected + cleared from the sheet 2026-07-13) | M | Proposed — needs care: yellow is also used for non-channel annotations | data-developer |
| Schedule the daily brief (routine/cron) | Makes the analyst's daily report automatic instead of on-demand | S | Proposed — needs user OK on timing | main assistant |
| TV charts for Brent (UKOIL), Palladium, Copper | Would give live captured prices for the 3 chartless commodities (currently #N/A / stale) | S | **User decision** — charts must be added in TradingView by the user | user |

## Later

| Item | Why it matters | Size | Status | Builder |
|---|---|---|---|---|
| "Investment OS" evolution (history-driven analytics, briefings) | Long-term direction from the 2026-07-12 external review | L | Blocked on SQLite history store; never a rebuild of the working pipeline | TBD |

## Done

| Item | Shipped |
|---|---|
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
