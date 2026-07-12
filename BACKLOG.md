# Backlog — Claude-Weath-Integration

Owned by the `product-owner` agent (`.claude/agents/product-owner.md`).
Single source of truth for pending improvements. Statuses: Proposed →
Approved → In progress → Done (with commit/PR ref). Items marked
**user decision** are blocked until the user decides.

## Now

| Item | Why it matters | Size | Status | Builder |
|---|---|---|---|---|
| Merge PR #8 (below-alert gap + 2026-07-12 batch) | Ships committed work to main | S | Waiting on user merge | user |
| Add `.claude/worktrees/` to `.gitignore` | Prevents accidental commit of working copies; closes external-review finding | S | Proposed | main assistant |
| `config.json` for hard-coded paths/IDs/port | Downloads paths, Python path, Finance sheet ID, port 4590 are scattered across ~10 scripts; single config makes machine migration painless | M | Proposed | main assistant |

## Next

| Item | Why it matters | Size | Status | Builder |
|---|---|---|---|---|
| Per-run manifest snapshots into SQLite | Gateway to indicator change detection, day-over-day comparison, and every "Investment OS" analytics ambition | M | Proposed | main assistant |
| pytest suite for pure logic (ticker_normalize, offset_formula, gap calc, fidelity classifier) | Catches regressions offline without TradingView running; complements test-analyst's data audits | M | Proposed | test-analyst (specs) + main assistant |
| TV charts for Brent (UKOIL), Palladium, Copper | Would give live captured prices for the 3 chartless commodities (currently #N/A / stale) | S | **User decision** — charts must be added in TradingView by the user | user |
| WPP stale Alert Low (1121.92 vs live 274.6) | Wrong level in a live trading sheet; flagged 2026-07-12 | S | **User decision** — needs manual review | user |

## Later

| Item | Why it matters | Size | Status | Builder |
|---|---|---|---|---|
| "Investment OS" evolution (history-driven analytics, briefings) | Long-term direction from the 2026-07-12 external review | L | Blocked on SQLite history store; never a rebuild of the working pipeline | TBD |

## Done

| Item | Shipped |
|---|---|
| (seeded 2026-07-12 — completed items before this date live in CLAUDE.md's Resolved sections) | — |
