"""Every month boundary this report depends on, DERIVED from the export date.

See constants.py for the pinned-data-vs-derived-anchor distinction: anything that
describes *where the report sits* is computed here and must never be a literal.

Extracted from spending_summary.py on 2026-07-19 — that file had grown to 4,020
lines. Behaviour is unchanged: the code below is the original, moved verbatim.
"""
import os

import pandas as pd

from .loaders import load_history, load_income_history, load_spend_history


def _account_summary_export_date(account_summary_path):
    """The 'Export date' header Fidelity writes at the top of AccountSummary.csv."""
    if not account_summary_path or not os.path.exists(account_summary_path):
        return None
    try:
        with open(account_summary_path, encoding="utf-8-sig") as f:
            for _ in range(20):
                line = f.readline()
                if not line:
                    break
                parts = line.replace("\r", "").rstrip("\n").split(",")
                if len(parts) >= 2 and parts[0].strip().lower() == "export date":
                    return pd.to_datetime(parts[1].strip(), dayfirst=True)
    except (OSError, ValueError):
        return None
    return None


def _last_month_in(*series_dicts):
    """Latest month present in any of the pinned history tables."""
    months = [m for d in series_dicts for vals in d.values() for m in vals]
    return max(months) if months else None


class MonthAnchors:
    """Where this report currently sits in the calendar. All derived — see above.

    export_date     the broker snapshot's own date (AccountSummary 'Export date')
    year / months   the calendar year being reported, Jan–Dec
    data_month      month of the snapshot: the anchor for holdings, prices and
                    account values, and the last month with any live data
    partial_month   data_month when the snapshot lands mid-month, so its
                    transactions are only partly captured (None on a month end)
    partial_scale   fraction of that month the snapshot covers (day / days in month)
    hold_from       first month with no live data at all — projections past the
                    snapshot are unreliable, so account values hold flat from here
    hist_cutoff     first month the live transaction files own; before it, the
                    pinned Jan–Apr spend/income tables win
    wealth_cutoff   same boundary for load_history()'s wealth series
    """

    def __init__(self, export_date, hist_last, wealth_hist_last):
        self.export_date = export_date
        self.hist_last = hist_last
        self.year = export_date.year
        self.jan = pd.Period(f"{self.year}-01", "M")
        self.dec = pd.Period(f"{self.year}-12", "M")
        self.months = [self.jan + i for i in range(12)]
        self.data_month = export_date.to_period("M")
        self.partial_scale = min(1.0, export_date.day / export_date.days_in_month)
        self.partial_month = self.data_month if self.partial_scale < 1.0 else None
        # Last month whose data is complete enough to average over.
        self.last_actual = (self.data_month - 1) if self.partial_month else self.data_month
        self.hold_from = self.data_month + 1
        self.hist_cutoff = (hist_last + 1) if hist_last else self.jan
        self.wealth_cutoff = (wealth_hist_last + 1) if wealth_hist_last else self.jan

    def split_months(self, tx_months):
        """Partition the year into (actual, estimated) months.

        A month is an ACTUAL only if it is complete AND the transaction files
        actually cover it. Everything else is estimated — including any GAP month
        that falls between the end of the pinned history and the start of the
        60-day transaction window, which is the case the old literal anchors
        missed entirely.
        """
        tx = set(tx_months)
        actual = [m for m in self.months if m <= self.last_actual and m in tx]
        gap = [m for m in self.months if m <= self.last_actual and m not in tx]
        later = [m for m in self.months if m > self.last_actual]
        return actual, gap + later

    def describe(self):
        parts = [f"snapshot {self.export_date:%d %b %Y} → data month {self.data_month}"]
        if self.partial_month:
            parts.append(f"partial ({self.partial_scale:.0%} of the month)")
        parts.append(f"history hands over at {self.hist_cutoff}")
        return "; ".join(parts)

    def warnings(self):
        """Anything about this run that a human should look at, in plain words."""
        out = []
        if self.hist_last is not None and self.hist_last.year != self.year:
            out.append(
                f"The pinned spend/income history ends at {self.hist_last} but this report "
                f"covers {self.year}. Those tables no longer contribute, so the early months "
                f"rest entirely on estimates — refresh load_spend_history() and "
                f"load_income_history() from a full-year export."
            )
        return out


def resolve_anchors(account_summary_path=None):
    """Build the MonthAnchors for this run from whatever the inputs tell us.

    The AccountSummary export header is the only exact statement of the snapshot
    date; with no AccountSummary at all, today is the honest fallback (the run is
    being built from a transaction export that was just downloaded).
    """
    export_date = _account_summary_export_date(account_summary_path)
    if export_date is None:
        export_date = pd.Timestamp.today().normalize()
    return MonthAnchors(
        export_date,
        _last_month_in(load_spend_history(), load_income_history()),
        _last_month_in(load_history()),
    )
