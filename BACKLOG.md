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
| Redraw broken-out channels in TradingView | 2026-07-13 run: price has genuinely broken ABOVE the drawn channel for **BEZ (1288 vs top 1088), GSK, HIK, MNG, BLND, LAND, SDR, Copper (COPPER1!)** — the detector correctly refuses to write alerts from an invalidated pattern, so these tickers get no Alert Low/High until redrawn | S | **User decision** — redraw in TradingView, then re-run pipeline | user |
| Re-check 6 price-below-channel rejections after next run | ADM, REL, CTEC, MGNS, RIO, ULVR were rejected with price *below* the detected channel — likely artifacts of the old frame-edge read (fixed 2026-07-13, reads now at today's date); expect most to self-resolve on the next run, redraw only what still fails | S | Proposed — verify on next pipeline run | main assistant |
| Investigate 17 OCR axis-label failures | 17 charts failed price-axis OCR (9 "no readable labels", 8 "<3 clean labels") in the 2026-07-13 run — those charts can never get a channel read until diagnosed | M | Proposed | data-developer |
| Schedule the daily brief (routine/cron) | Makes the analyst's daily report automatic instead of on-demand | S | Proposed — needs user OK on timing | main assistant |
| TV charts for Brent (UKOIL), Palladium, Copper | Would give live captured prices for the 3 chartless commodities (currently #N/A / stale) | S | **User decision** — charts must be added in TradingView by the user | user |
| WPP stale Alert Low (1121.92 vs live 274.6) | Wrong level in a live trading sheet; flagged 2026-07-12 | S | **User decision** — needs manual review | user |

## Later

| Item | Why it matters | Size | Status | Builder |
|---|---|---|---|---|
| "Investment OS" evolution (history-driven analytics, briefings) | Long-term direction from the 2026-07-12 external review | L | Blocked on SQLite history store; never a rebuild of the working pipeline | TBD |

## Done

| Item | Shipped |
|---|---|
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
