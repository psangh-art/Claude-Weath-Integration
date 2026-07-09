# Chart Import Findings

Running log of channel readings extracted from `Trading_Layouts.xlsx` exports and applied (or not) to `Stocks Buy Strategy.xlsx`. One entry per import session, newest first.

---

## 📋 Coverage Tracker — check this FIRST every session

This is the persistent list of tickers that **have appeared in a TradingView export** but still don't have usable Alert Low data in the master sheet. Unlike a full gap analysis (which includes ~250 stocks never expected to have charts yet), this list is scoped to only the ones actually worth chasing — if a ticker's chart has been exported, it belongs here until it's resolved.

**Process going forward:** every time a new `tradingview_layouts.xlsx` comes in, check it against this list first. Update the status column. Don't let a ticker quietly disappear from tracking just because a new export didn't happen to fix it.

| Ticker | Company | Status | Why | Last checked |
|---|---|---|---|---|
| MTRO | Metro Bank | 🔴 Blocked | Chart has never had a channel drawn on it — confirmed across 4 separate exports. Needs someone to draw the channel on TradingView before this can ever resolve. | 2026-07-09 |
| OSB | OSB Group | 🟡 Ambiguous | Chart has 5 detected line segments instead of a clean 2 (channel + extra trendline/MA overlay). Needs a cleaner single-indicator screenshot. | 2026-07-09 |
| AML | Aston Martin Lagonda | ⚪ Not yet attempted | Appeared in export, not yet read | 2026-07-09 |
| BAG | A.G. Barr | ⚪ Not yet attempted | Appeared in export, not yet read | 2026-07-09 |
| CCR | C&C Group | ⚪ Not yet attempted | Appeared in export, not yet read | 2026-07-09 |
| CMCX | CMC Markets | ⚪ Not yet attempted | Appeared in export, not yet read | 2026-07-09 |
| PFD | Premier Foods | ⚪ Not yet attempted | Long-term chart (2011-2027), channel not yet reliably read | 2026-07-09 |

---

## 2026-07-09 — Full batch import (104 tickers, 35 layout groups)

**Source:** `tradingview_layouts.xlsx` uploaded same day, built via VS Code from the GitHub pipeline scripts.
**Method:** Pixel-level colour detection of channel boundary lines against OCR'd axis labels, filtered for plausibility (rejecting near-zero-width detections and cross-checking against known real prices/data where available).

### 🔴 Bug still present — needs a fix in the export script

**FT100 Support Services 1 layout mislabeling.** This is the **third consecutive export** with the same bug: the layout tags AUTO, then five rows all labeled "SGE", then REL, then two more "SGE" — eight panes, but only three distinct labels. Verified the underlying image files ARE genuinely different charts (distinct file hashes), so this is a **label-writing bug in the export**, not a rendering issue. The correct company names for those five/two "SGE" slots are still unknown — please check what the export script is doing when it writes the `Symbol` column for this specific layout, since it's clearly not reading each pane's actual ticker correctly.

### ✅ Successes — 36 stocks updated with fresh Alert Low/High

ANTO, RIO, GLEN, EDV, FRES, MNG, ICG, LSEG, STJ, CWR, SHAW, TBCG, RTO, HSBA, CRDA, CCH, DGE, CHG, SNR, INF, RMV, WPP, BTRW, HWDN, MNDI, SSE, UTG, NXT, KGF, EZJ, ENT, WTB, SLVR, TW, UU, QQ

8 more (SMIN, SPX, CNA, SVT, BLND, LAND, SGRO, GAW) were re-read but came within 3% of already-current values — left untouched to avoid introducing noise from re-reading the same channel twice.

### ❌ Rejected — nothing written, for cause

| Ticker | Reason |
|---|---|
| MTRO | Confirmed (4th time now) — this chart has never had a channel drawn on it |
| WEIR | Detected channel (1,652-2,504) conflicts with verified SMA data (50/100/200/350-day averages all sit in the 2,415-2,861 range) - held back rather than overwrite with contradictory data |
| IHG | Detection returned 143-190p; real IHG trades around 7,000-10,000p. Physically implausible - clear false read |
| III, JD., MKS, NG., SDR | Detected channel width exceeded 150% of the lower boundary - almost certainly picked up unrelated lines on the chart, not the actual channel |
| GDX, TLW | Genuinely new tickers with no existing row, but neither got a trustworthy read (TLW: no channel detected at all; GDX: detected but only 4.5% width, likely noise). Not added - need a cleaner chart before these can go in |

### Data quality note (unrelated to channels)

Copper's Simple Moving Averages have shown all four periods (50/100/200/350-day) as **exactly identical** in two separate exports now (see indicator CSV findings, 2026-07-08 session, not repeated here). Worth checking whether this is the same root cause as the Support Services 1 mislabeling - possibly the same underlying "indicator slot" bug affecting both SMA and channel indicators.

### Spreadsheet formatting fix (unrelated to this import, found during QA)

18 rows had `Chart = Yes` but were still showing the gray "No" cell colour (a styling bug from earlier sessions where the value was updated but the cell format wasn't). One row (Brent Oil / UKOIL) had the reverse issue - `Chart = No` but green formatting. All 19 corrected; every `Yes` cell is now consistently green (`#C6EFCE` fill / `#276221` text) and every `No` is consistently gray.

---

## 2026-07-08 — Insurance layout refresh + Ceres Power

**Source:** `Trading_Layouts.xlsx` (Insurance/Financials refresh + ad hoc uploads)

### Applied
| Ticker | Channel | Lower | Upper | Alert Low | Confidence |
|---|---|---|---|---|---|
| CWR (Ceres Power) | Ascending | 465p | 736p | 488p | High - current price sits almost exactly on the detected lower boundary |

### Read but NOT applied - pending confirmation (later superseded by 2026-07-09 batch)
| Ticker | Proposed Alert Low | Status |
|---|---|---|
| BNZL (Bunzl) | 2,520p | Never confirmed by Paul; not carried forward |
| PRU, SDLF | ~1,065p / ~835p | Superseded - not re-read in the 07-09 batch |

### Detection failed or conflicting
BEZ (contradictory chart pattern vs existing note), ADM, HSX (no channel detected - different chart style)

### Corrections
- **RIO**: moved from "NEAR" to "AT BOUNDARY" tier - live sheet showed 1.6% Diff to Low, not the ~9-16% estimated from stale web data
- **SBRY**: Alert Low corrected 252 -> 273, matching "Lower 260p" in Claude Notes

---

*Template for future entries: separate Applied / Rejected / Pending sections so nothing gets silently written to the live trading sheet without a confidence check, and always flag export bugs (mislabeling, duplicate rows, flat/identical indicator values) even when they don't block the numeric read.*
