"""Fixed reference data for the spending summary: categories, family account
ownership/labels, and the PINNED figures that are anchored to the month they were
measured (as opposed to the derived anchors in anchors.py — the comment below
explains the distinction, which is load-bearing).

Extracted from spending_summary.py on 2026-07-19 — that file had grown to 4,020
lines. Behaviour is unchanged: the code below is the original, moved verbatim.
"""
import pandas as pd


# ── Categories ────────────────────────────────────────────────────────────────
CATEGORIES = [
    "Broadband & Phone", "Car Expenses", "Cash Withdrawals",
    "Clothing & Shopping", "Eating & Drinking Out", "Electronics",
    "Food Shopping", "Health & Wellbeing", "Holidays",
    "Liam Sangha", "Media & Subscriptions", "Professional Services",
    "Seatfrog", "Sport & Leisure", "Susan's Car Lease", "Tax",
    "Travel & Hotels", "Travel & Transport", "University Fees", "Utilities",
]

# Account ownership — only family accounts included
ACCOUNT_OWNER = {
    "2000001606": "Paul",
    "SANX002282": "Paul",
    "SANQ000468": "Paul",
    "2000001604": "Susan",
    "SANX002617": "Susan",
    "SANX002936": "Jayne",
    "AS10303823": "Liam",
    "AG10131710": "Liam",
}

ACCOUNT_LABELS = {
    "2000001606": "SIPP Savings",
    "SANX002282": "Investment ISA",
    "SANQ000468": "Investment Account (Joint)",
    "2000001604": "SIPP Savings",
    "SANX002617": "Investment ISA",
    "SANX002936": "Junior ISA",
    "AS10303823": "Investment ISA",
    "AG10131710": "Investment Account",
}

# Display order of family members
FAMILY_ORDER = ["Paul", "Susan", "Jayne", "Liam"]


# ── Month anchors ─────────────────────────────────────────────────────────────
# The hardcoded months in this file are two different kinds of thing, and the
# distinction is the whole point of this section:
#
#   DATA pinned to the month it was measured — the Jan–Apr spend/income tables,
#   load_history()'s wealth series, the May-2026 pension/house/car estimates.
#   These keep their literal months for ever. They project forward from their own
#   as-of date, so the calendar advancing never invalidates them.
#
#   ANCHORS describing where the report currently sits — which year it covers,
#   which month the broker snapshot was taken in, which month is only partly
#   captured, and where hardcoded history hands over to live data. These were
#   written as literal May/June 2026 periods and silently went wrong the moment
#   the calendar moved past them: on the 18 Jul 2026 export, May fell in the gap
#   between the Jan–Apr history and the 60-day transaction window and was
#   classified as an ACTUAL of zero (never estimated), while July was likewise
#   treated as a complete actual and reported its part-month £192. Anchors are
#   derived from the data below and must never be literals again.
ESTIMATES_AS_OF = pd.Period("2026-05", "M")   # as-of month of the pinned estimates

# Confirmed equity dividends for the reporting year, shown inline under Paul's SIPP
# and summed for the Targets table's annual income. Pinned data — one copy, so the
# table and the total can never disagree (they were separate literals until
# 2026-07-19, and the total had to be re-added by hand whenever a payment changed).
EQUITY_DIVIDENDS_INLINE = [
    ("  RELX PLC",         {"2026-06": 413, "2026-09": 168}),
    ("  Sage Group PLC",   {"2026-06": 178}),
    ("  Auto Trader Group", {"2026-09": 315}),
    ("  Weir Group",       {"2026-11": 78}),
]
EQUITY_DIVIDENDS_ANNUAL = sum(v for _, divs in EQUITY_DIVIDENDS_INLINE for v in divs.values())
