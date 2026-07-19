# One-off migration scripts (archive)

Dated, **already-executed** scripts that made a one-time change to
`Stocks_Buy_Strategy.xlsx`, `spending_summary.xlsx` or one of the PowerPoint decks —
inserting a column, restyling a table, adding a slide, renaming a value.

They were moved here on 2026-07-19 (from `scripts/`) purely for structure: eighteen
dated files sat alongside the ~24 that actually run, which made it hard to see what
the live pipeline consists of. **Nothing in the pipeline imports them at runtime** —
the only cross-references are documentation.

## Rules

- **They are a record of what was done, not a template.** Each one was written
  against the workbook layout of its own date. `fix_relx_history_2026-07-10.py`, for
  instance, still refers to the sheet by its pre-rename name `'Stocks Buy Strategy'`
  (now `Investments`). Re-running one blind against today's workbook can corrupt it.
- **Re-run only the ones CLAUDE.md says are re-runnable**, from the repo root, e.g.
  `python scripts/oneoff/add_agents_slide_2026-07-12.py` (replaces the Agents slide)
  and `add_investment_dashboard_slide_2026-07-17.py`. Both back up / purge old deck
  versions first.
- **Don't add new ones here casually.** A change that needs repeating belongs in the
  pipeline; a genuine one-off lands here after it has been run.
- Shared helpers do NOT live here. `offset_formula` and `copy_cell_style` moved to
  `scripts/xlsx_sheet_copy.py` so the tests and live writers depend on a maintained
  module rather than on an archived script.
