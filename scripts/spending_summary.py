#!/usr/bin/env python3
"""
Spending Summary Generator
--------------------------
Usage:
    python spending_summary.py                                    # uses defaults
    python spending_summary.py amex.csv barclays.csv fidelity.csv
    python spending_summary.py amex.csv barclays.csv fidelity.csv output.xlsx

Reads Amex (activity.csv), Barclays (data.csv) and Fidelity
(TransactionHistory.csv) and produces a categorised monthly spending summary
plus a Fidelity income-by-account section as a formatted Excel file.
"""

import sys
import os
import io
import csv
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Windows' default console codepage (cp1252) can't encode characters like
# "→" used in console-only status prints below — reconfigure to UTF-8 so a
# cosmetic console print can't crash the script after the workbook is
# already written (same class of bug fixed in verify_pipeline.py).
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass


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


def apply_pending_holdings(holdings: dict, pending_path: str) -> dict:
    """
    Adds net units from pending (not-yet-settled) Buy/Sell orders on top of a
    holdings dict. AccountSummary and completed-transaction history only
    reflect *settled* trades, so a pending sell/buy that hasn't cleared yet
    wouldn't otherwise show up until the next export. A pending export has
    Completion date == 'Pending' on every row and Status values of
    Priced/Ordered rather than Completed — it's read here without a Status
    filter since none of its rows are ever 'Completed'.

    Returns a new dict; does not mutate the input.
    """
    if not pending_path or not os.path.exists(pending_path):
        return holdings

    FAMILY_ACCS = set(ACCOUNT_OWNER.keys())
    with open(pending_path, encoding="utf-8-sig") as f:
        content = f.read()
    lines = content.replace("\r", "").split("\n")
    try:
        start = next(i for i, l in enumerate(lines) if l.startswith("Order date"))
    except StopIteration:
        return holdings

    result = dict(holdings)
    reader = csv.DictReader(io.StringIO("\n".join(lines[start:])))
    for row in reader:
        t = row.get("Transaction type", "").strip()
        if t not in ("Buy", "Sell"):
            continue
        acc = row.get("Account Number", "").strip()
        if acc not in FAMILY_ACCS:
            continue
        inv = row.get("Investments", "").strip().strip('"')
        if not inv or inv == "Cash":
            continue
        try:
            qty = float(row.get("Quantity", "0") or 0)
        except (ValueError, TypeError):
            continue
        delta = qty if t == "Buy" else -qty
        result[(acc, inv)] = result.get((acc, inv), 0) + delta

    return {k: v for k, v in result.items() if v > 0.001}


def build_holdings(fidelity_path: str, account_summary_path: str = None, pending_path: str = None) -> dict:
    """
    Returns dict: (account, fund_name) -> net units held.

    If account_summary_path is provided (AccountSummary CSV export), uses that
    as the authoritative source — it reflects exact holdings at export date.
    Falls back to calculating net units from transaction history otherwise.

    If pending_path is provided, pending Buy/Sell orders are layered on top
    via apply_pending_holdings — see that function for why.
    """
    FAMILY_ACCS = set(ACCOUNT_OWNER.keys())
    holdings = None

    # ── Primary: AccountSummary CSV ───────────────────────────────────────────
    if account_summary_path and os.path.exists(account_summary_path):
        with open(account_summary_path, encoding="utf-8-sig") as f:
            content = f.read()
        lines = content.replace("\r", "").split("\n")

        # Find the last "Type,Holdings,Account number" header — that's the detail section
        header_idx = None
        for i, l in enumerate(lines):
            if l.startswith("Type,Holdings,Account number"):
                header_idx = i

        if header_idx is not None:
            reader = csv.DictReader(iter(lines[header_idx:]))
            holdings = {}
            for row in reader:
                if row.get("Type", "").strip() != "Asset":
                    continue
                acc = row.get("Account number", "").strip()
                if acc not in FAMILY_ACCS:
                    continue
                fund = row.get("Holdings", "").strip().strip('"')
                try:
                    qty = float(row.get("Quantity", "0") or 0)
                except (ValueError, TypeError):
                    continue
                if qty > 0:
                    holdings[(acc, fund)] = qty

    # ── Fallback: calculate from transaction history ───────────────────────────
    if holdings is None:
        with open(fidelity_path) as f:
            content = f.read()
        lines = content.replace("\r", "").split("\n")
        start = next(i for i, l in enumerate(lines) if l.startswith("Order date"))
        reader = csv.DictReader(io.StringIO("\n".join(lines[start:])))
        rows = [r for r in reader if r.get("Status", "").strip() == "Completed"]

        from collections import defaultdict
        raw_holdings = defaultdict(float)
        for r in rows:
            t = r["Transaction type"].strip()
            acc = r["Account Number"].strip()
            inv = r["Investments"].strip().strip('"')
            try:
                qty = float(r["Quantity"])
            except (ValueError, KeyError):
                continue
            if t == "Buy":
                raw_holdings[(acc, inv)] += qty
            elif t == "Sell":
                raw_holdings[(acc, inv)] -= qty
        holdings = {k: v for k, v in raw_holdings.items() if v > 0.001}

    return apply_pending_holdings(holdings, pending_path)


# ── AMEX categorisation ────────────────────────────────────────────────────────
def categorise_amex(desc: str) -> str | None:
    d = str(desc).upper()
    if "PAYMENT RECEIVED" in d:
        return None
    if "SEATFROG" in d:
        return "Seatfrog"

    holiday_kw = [
        "ADEJE", "TENERIFE", "CRISTIANOS", "ARMENIME", "HIPERDINO",
        "RIU PALACE", "RED LION 005", "WALI HAII", "LIMONELLA",
        "TEMPLE BAR 367", "HARRY S BAR 047", "ROSSO SUL MARE",
        "SUGAR AND SPICY", "WAXY OSHEAS", "RTE. Y BARES", "CANDELARIA",
        "MURPHYS", "COSTA CALETA", "COSTA ADEJE", "MANGO SIAM", "VIPS SIAM",
        "PIZZERIA ALL ITALIANA", "SIAM MALL", "STRADIVARIUS SIAM",
        "ROCA NEGRA", "LM 176", "LM 179", "LM 95", "SSP-TFS",
        "CAFETERIA EL DUQUE", "ORANGE CAFE", "SUPERMERCADO", "MAQUINAS VENDING",
        "TAXI ADEJE", "TAXI-ADEJE", "TAXI LM", "TAXI L.M.", "TAXI VILLA DE ADEJ",
        "HD EX SAN SEBASTIA", "CASA PLAYA", "CHIRINGUITO EL PUE",
        "MINIMARKET TENERIF", "HERTZ UK", "TUI INFLIGHT", "ADMIRAL TRAVEL",
        "NEWCASTLE AIR", "3CPAYMENT*PREMIER INN 4 LONDON",
        "3CPAYMENT*PREMIER INN 4 SOUTHAMP", "3CPAYMENT*NEWCASTLE AIR",
        "IBIS LONDON", "DIVERSE DINING LTD",
    ]
    if any(kw in d for kw in holiday_kw):
        return "Holidays"

    if "EXPLEO" in d or "UK EXCL" in d:
        return "Salary"
    if "LIAM SANGHA" in d:
        return "Liam Sangha"
    if "HMRC" in d:
        return "Tax"

    if any(kw in d for kw in ["OCTOPUS ENERGY", "NORTHUMBRIAN WATER", "TALKTALK", "TV LICENCE", "PARENTPAY"]):
        return "Utilities"
    if any(kw in d for kw in ["TALKMOBILE", "VODAFONE", "APPLE.COM/BILL"]):
        return "Broadband & Phone"
    if any(kw in d for kw in ["NETFLIX", "NORTON", "CLAUDE.AI", "TRADINGVIEW", "THESHIFT.SUBSTACK", "BCS MEMBER"]):
        return "Media & Subscriptions"
    if any(kw in d for kw in ["SAINSBURY'S PETROL", "SAINSBURYS PETROL", "DENTON BURN FILLING",
                                "KWIK FIT", "HALFORDS", "TT2 LTD", "TYNE TUNN", "NYX*EVCHARG"]):
        return "Car Expenses"
    if any(kw in d for kw in ["3CPAYMENT*PREMIER INN 4 NEWCASTLE", "3CPAYMENT*SLALEY HALL",
                                "MALMAISON", "CALEDONIAN HOTEL", "SLALEY HALL", "LONSDALE HOTEL",
                                "6928 - LONSDALE HOTEL", "RAMSIDE HALL", "BEAMISH HALL"]):
        return "Travel & Hotels"
    if any(kw in d for kw in ["LNER CAR PARK", "LNERCARPK", "LONDON NORTH EASTERN", "TFL TRAVEL",
                                "UBER TRIP", "TRAINLINE", "NEXUS TRAVEL", "SUMUP*TAXI",
                                "RINGGO", "PAYBYPHONE", "NEWCASTLE CITY COU", "W H SMITH"]):
        return "Travel & Transport"
    if any(kw in d for kw in ["CINEWORLD", "TYNESIDE BADMINTON", "HOLLYWOOD BOWL", "GO OUTDOORS",
                                "NEWCASTLE THEATRE", "TYNE THEATRE", "SUMUP*TYNE THEATRE",
                                "TICKETMASTER", "TM *TICKETMASTER", "HUSTLER POOL",
                                "UTILITA ARENA", "O2 CITY HALL"]):
        return "Sport & Leisure"
    if any(kw in d for kw in ["BOOTS", "DENPLAN", "WINDMILL ORTHODONT", "WELLNESS EMPOWER",
                                "JG *MACMILLAN", "WWW.DEMENTIAUK", "DACCS HAIR", "L BRITTON",
                                "SP ANCIENT BRAVE", "SP UPCIRCLE", "SP MODIBODI", "SP STEPPRS",
                                "SP ARCHIES", "SP NUTRITION GEEKS", "ROSEDENE", "SEPHORA", "MOLTON BROWN"]):
        return "Health & Wellbeing"
    if any(kw in d for kw in ["BCS MEMBERS", "BCS MEMBERSHIP", "APM", "HM PASSPORT", "STARTFIRST", "PARK ROAD SF"]):
        return "Professional Services"
    if any(kw in d for kw in ["CURRY'S", "CURRYS", "APPLE STORE", "PROCOOK"]):
        return "Electronics"

    food_kw = [
        "SAINSBURY'S SUPERMARKET", "SAINSBURY'S             ", "SAINSBURYS SUPERMARKET",
        "ASDA STORES", "ASDA ", "M&S STORE", "M&S GATESHEAD", "M&S METRO", "M&S NEWCASTLE",
        "MARKS & SPENCER", "MARKS AND SPENCER", "CO-OP WASHINGTON", "TESCO STORE",
        "SANSBRY", "WAITROSE 619", "ELM TREE FARM", "BROCKSBU", "BROOKSIDE NURSERY",
        "RINGTONS", "SWEET SYMPHONY", "SWEETSYMPHONY", "TEAL FARM", "92174 TEAL FARM",
        "DOJO*PITY ME NURSERY",
    ]
    if any(kw in d for kw in food_kw):
        return "Food Shopping"

    eat_kw = [
        "STARBUCKS", "MCDONALDS", "MCDONALD'S", "COSTA ", "FIVE GUYS", "TORTILLA",
        "ZAPATISTA", "BREWDOG", "BABUCHO", "CHAMPS SPORTS BAR", "WASHINGTON ARMS",
        "9012297 WASHINGTON ARMS", "DEEP NORTH", "DOJO*", "FAT FRANKS", "SIMLA",
        "KAMAL", "KYLIN ORIENTAL", "KING NEPTUNE", "JD WETHERSPOON", "WETHERSPOON",
        "DOMINOS", "HARRY'S BAR NEWCASTLE", "LATHAMS", "PRIDE OF SPITALFIELDS",
        "THE SUN WHARF", "LIBERTY BOUNDS", "TRAITORS GATE", "E1 FOODHALL",
        "A & S TAKEAWAY", "ARCHER AND SCOTTS", "STARVED", "SIR WILLIAM DE WES",
        "THREE MILE", "JOB BULLMAN", "BONEMI", "ZETTLE *LAKESIDE", "ZETTLE *OLIVIAS",
        "SHIFT4*EL TORERO", "0033 - GATESHEAD", "0511 - WASHINGTON", "1637 - WASHING",
        "1675 - METRO CENTRE", "5698 - DEARNE VALLEY", "6222 -RAVENSW", "ASTRONOMER",
    ]
    if any(kw in d for kw in eat_kw):
        return "Eating & Drinking Out"

    clothing_kw = [
        "AMAZON", "AMZN", "AMZ*", "NEXT ONLINE", "PRIMARK", "HOLLIST", "LEVI ", "LEVIS",
        "ARGOS", "IKEA", "FENWICK", "OLIVER BONAS", "CARD FACTORY", "HOTEL CHOCOLAT",
        "JELLYCAT", "RADLEY", "WATERSTONES", "SP POM POM", "SP SAINTSOFIA", "SP WUKA",
        "VIVA*FLYING TIGER", "MANGO", "PULL AND BEAR", "STRADIVARIUS", "BECKLE",
        "HOME BARGAINS", "B&Q", "THE RANGE", "LAKELAND", "MORPETH GARDEN",
        "HUSH HOMEWEAR", "WWW.JOHNLEWIS", "JOHN LEWIS", "WHITE STUFF", "SPORTSDIRECT",
        "JG *ISLA COOK", "NATIONAL MERCH", "3CPAYMENT*GEORGE WASHIN",
    ]
    if any(kw in d for kw in clothing_kw):
        return "Clothing & Shopping"

    return "Clothing & Shopping"


# ── Barclays categorisation ────────────────────────────────────────────────────
def categorise_barclays(row) -> str | None:
    memo   = str(row.get("Memo", "") or "").upper().strip()
    subcat = str(row.get("Subcategory", "") or "").strip()
    amount = float(row.get("Amount", 0) or 0)
    spend  = -amount

    if any(kw in memo for kw in ["AMERICAN EXPRESS", "FASL PRIM CLIENT B"]):
        return None

    if amount > 0:
        if "EXPLEO" in memo or "UK EXCL" in memo:
            return "Salary"
        if "FIDELITY FASL PYMT" in memo:
            return "Fidelity"
        return None

    # Susan's BMW car lease — £525.60/month (ends Jul 2027)
    if "BMW FINANCIAL SERV" in memo:
        return "Susan's Car Lease"

    if any(kw in memo for kw in ["NOTEMACHINE", "CARDTRONICS", "LLOYDS", "HALIFAX PLC", "NAT WEST BANK", "HSBC"]):
        return "Cash Withdrawals"
    if subcat == "Cash Withdrawal":
        return "Cash Withdrawals"

    if "WWW.FIDELITY.CO.UK" in memo:
        return "Fidelity"
    if "SEATFROG" in memo:
        return "Seatfrog"
    if "LIAM SANGHA" in memo and "COURSE FEES" in memo:
        return "University Fees"
    if "LIAM SANGHA" in memo:
        return "Liam Sangha"
    if "HMRC" in memo:
        return "Tax"

    if any(kw in memo for kw in ["OCTOPUS ENERGY", "OCTOPUS ENE", "NORTHUMBRIAN WATER",
                                   "TALKTALK", "TV LICENCE", "PARENTPAY", "D&G BOILER"]):
        return "Utilities"
    if any(kw in memo for kw in ["H3G", "TALKMOBILE", "VODAFONE LTD"]):
        return "Broadband & Phone"
    if any(kw in memo for kw in ["NETFLIX", "NORTON", "TRADINGVIEWV", "THESHIFT.SUBSTACK",
                                   "BCS MEMBERS", "BCS MEMBERSHIP"]):
        return "Media & Subscriptions"
    if any(kw in memo for kw in ["KWIK FIT", "HALFORDS", "DENTON BURN", "SAINSBURYS PETROL",
                                   "SAINSBURY'S PETROL", "NYX*EVCHARG", "NYX*GEORGEWASHINGT",
                                   "TT2 LTD", "TYNE TUNN", "MERCEDES BENZ", "MBUK",
                                   "BMW FINANCIAL SERV"]):
        return "Car Expenses"
    if any(kw in memo for kw in ["PREMIER INN", "IBIS", "MALMAISON", "CALEDONIAN HOTEL",
                                   "SLALEY HALL", "LONSDALE HOTEL", "RAMSIDE HALL",
                                   "BEAMISH HALL", "BAY TREE HILGAY"]):
        return "Travel & Hotels"
    if any(kw in memo for kw in ["TUI", "HERTZ", "SPAIN", "EUR", "ADEJE", "TENERIFE",
                                   "CRISTIAN", "AIRPORT"]):
        return "Holidays"
    if any(kw in memo for kw in ["LNER", "TFL TRAVEL", "UBER", "TRAINLINE", "NEXUS TRAVEL",
                                   "RINGGO PARKING", "PAYBYPHONE", "NEWCASTLE CITY COU",
                                   "NEWCASTLE CC PARKI", "PPOINT_*SLATYFORD", "POST OFFICE COUNTE"]):
        return "Travel & Transport"
    if any(kw in memo for kw in ["EVERYONE ACTIVE", "BADMINTON CLUB", "MR THEYAGA SHANDRA",
                                   "HOLLYWOOD BOWL", "CINEWORLD", "NEWCASTLE THEATRE",
                                   "O2 CITY HALL", "UTILITA ARENA", "HUSTLER POOL",
                                   "SUNDERLAND CC", "KESWICK BREWING", "BALLYS NEWCASTLE"]):
        return "Sport & Leisure"
    if any(kw in memo for kw in ["BOOTS UK", "DENPLAN", "WINDMILL ORTHODONT", "WELLNESS EMPOWER",
                                   "L BRITTON", "WWW.HAIRTRADE", "DACCS HAIR",
                                   "GLC LIMITED", "ROSEDENE", "SP WUKA", "SP SAINTSOFIA", "SP POM POM"]):
        return "Health & Wellbeing"
    if any(kw in memo for kw in ["BCS MEMBERS", "BCS MEMBERSHIP", "APM", "HM PASSPORT OFFICE"]):
        return "Professional Services"
    if any(kw in memo for kw in ["CURRY'S", "CURRYS", "APPLE STORE", "AO MANCHESTER", "GEBERIT"]):
        return "Electronics"

    food_kw = [
        "SAINSBURYS S/MKTS", "MARKS&SPENCER PLC", "ASDA", "CO-OP", "TESCO",
        "CLR*KNITSLEY", "ELM TREE FARM", "BRIDGESSTONES", "COSY NEWS", "CLAY'S GARDEN",
        "RINGTONS LTD", "INDIRAN MINIMART", "MORPETH GARDEN", "D AND G NEWS",
        "HOCKWOLD STORES", "HOME BARGAINS", "THE RANGE", "LAKELAND",
        "SNEHAREDDY", "SINGING CANARY", "BOOTS UK ECOMM",
    ]
    if any(kw in memo for kw in food_kw):
        return "Food Shopping"

    eat_kw = [
        "CROSS KEYS", "GREGGS", "KYLIN ORIENTAL", "CHAMPS SPORTS BAR", "THE BOTANIST",
        "TACO BELL", "BABUCHO", "COLMANS CATERERS", "PING ON CHOP SUEY", "BEAMISH MUSEUM",
        "FRIEZ AND BURGZ", "THE GREEN", "SUBWAY", "PERSIAN BITE", "CONCORD TANDOORI",
        "SHIP LANGSTONE", "ZETTLE_*OLIVIAS CO", "6888 - PARK VIEW", "BALTIC FLOUR MILLS",
        "BALTIX FLOUR MILLS", "ELVET ICES", "DOJO*", "NOVELLO'S", "NOVELLO ",
        "O'BRIENS", "O BRIENS", "THE BAKE HOUSE", "ALVINOS",
        "SCREAM SUNDERLAND", "THE BRIDGES SUNDER", "NEWCASTLE HALAL SU",
    ]
    if any(kw in memo for kw in eat_kw):
        return "Eating & Drinking Out"

    clothing_kw = [
        "NEXT ", "PRIMARK", "HOLLISTER", "LEVI", "ARGOS", "IKEA", "FENWICK",
        "OLIVER BONAS", "CARD FACTORY", "HOTEL CHOCOLAT", "JELLYCAT", "RADLEY",
        "W H SMITH", "WATERSTONES", "PULL AND BEAR", "PULLANDBEAR", "STRADIVARIUS",
        "HUSH HOMEWEAR", "WWW.JOHNLEWIS", "WHITE STUFF", "SPORTSDIRECT",
        "JG *ISLA COOK", "MARKS&SPENCER PLC", "AMAZON*", "CLARKS",
    ]
    if any(kw in memo for kw in clothing_kw):
        return "Clothing & Shopping"

    if subcat in ("Funds Transfer", "Standing Order") and spend > 0:
        return "Clothing & Shopping"
    if subcat == "Direct Debit" and spend > 0:
        return "Utilities"

    # Catch-all for any remaining outgoing transactions
    if spend > 0:
        return "Clothing & Shopping"

    return None


# ── Data loading ───────────────────────────────────────────────────────────────
def load_amex(path: str) -> pd.DataFrame:
    df = pd.read_csv(path)
    df["Date"] = pd.to_datetime(df["Date"], dayfirst=True)
    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0)
    df["month"] = df["Date"].dt.to_period("M")
    df["category"] = df["Description"].apply(categorise_amex)
    df = df[(df["Amount"] > 0) & df["category"].notna()].copy()
    df["spend"] = df["Amount"]
    return df[["month", "category", "spend"]]


def load_barclays(path: str) -> pd.DataFrame:
    df = pd.read_csv(path)
    df["Date"] = pd.to_datetime(df["Date"], dayfirst=True, errors="coerce")
    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0)
    df["Memo"] = df["Memo"].fillna("")
    df["month"] = df["Date"].dt.to_period("M")
    df["category"] = df.apply(categorise_barclays, axis=1)
    df = df[df["category"].notna()].copy()

    sal = df[df["category"] == "Salary"].copy()
    sal["spend"] = sal["Amount"]

    fid = df[df["category"] == "Fidelity"].copy()
    fid["spend"] = fid["Amount"].abs()

    exp = df[(df["category"] != "Salary") & (df["category"] != "Fidelity") & (df["Amount"] < 0)].copy()
    exp["spend"] = exp["Amount"].abs()

    return pd.concat([sal, fid, exp])[["month", "category", "spend"]]


def load_fidelity_income(path: str) -> pd.DataFrame:
    """
    Returns income rows from Fidelity by account number and month.
    Includes: Income Received, Cash Dividend (positive amounts = money in).
    Includes: Income Payment (negative Amount, but Quantity = cash paid out to holder).
    """
    with open(path) as f:
        content = f.read()

    lines = content.replace("\r", "").split("\n")
    start = next(i for i, l in enumerate(lines) if l.startswith("Order date"))
    reader = csv.DictReader(io.StringIO("\n".join(lines[start:])))
    rows = [r for r in reader if r.get("Status", "").strip() == "Completed"]

    income_types = {"Income Received", "Cash Dividend", "Income Payment"}
    records = []
    for r in rows:
        txn_type = r["Transaction type"].strip()
        if txn_type not in income_types:
            continue

        # Parse date — use Completion date if available, else Order date
        date_str = (r.get("Completion date") or r.get("Order date") or "").strip()
        if not date_str or date_str.lower() == "pending":
            date_str = r["Order date"].strip()
        try:
            dt = pd.to_datetime(date_str, dayfirst=True)
        except Exception:
            continue

        account_num = r["Account Number"].strip()
        # Amount: Income Received / Cash Dividend → positive (money in)
        # Income Payment → negative Amount but Quantity = cash amount received
        raw_amount = str(r["Amount"]).strip().replace("£", "").replace(",", "")
        try:
            amount = float(raw_amount)
        except ValueError:
            continue

        if txn_type == "Income Payment":
            # These are distributions paid out — Quantity holds the positive cash figure
            raw_qty = str(r["Quantity"]).strip().replace("£", "").replace(",", "")
            try:
                amount = float(raw_qty)
            except ValueError:
                continue
        elif amount <= 0:
            continue  # skip negative income received (reversals etc.)

        if dt.year != 2026:
            continue

        # Fund name: Source investment field, fallback to Investments, fallback to txn type
        fund = r.get("Source investment", "").strip()
        if not fund:
            fund = r.get("Investments", "").strip()
        if not fund:
            fund = "(Cash / Other)"

        records.append({
            "month": dt.to_period("M"),
            "account": account_num,
            "fund": fund,
            "amount": amount,
        })

    df = pd.DataFrame(records)
    if df.empty:
        return df

    return df


def load_spend_history():
    """
    Hardcoded Amex + Barclays spend by category, Jan–Apr 2026.
    Updated from full-year transaction export 30/05/2026.
    These values are read-only — never overwritten by new uploads.
    """
    P = lambda s: pd.Period(s, "M")
    return {
        "Broadband & Phone":    {P("2026-01"):218,  P("2026-02"):103,  P("2026-03"):103,  P("2026-04"):108},
        "Car Expenses":         {P("2026-01"):251,  P("2026-02"):979,  P("2026-03"):910,  P("2026-04"):947},
        "Cash Withdrawals":     {P("2026-01"):530,  P("2026-02"):2060, P("2026-03"):940,  P("2026-04"):400},
        "Clothing & Shopping":  {P("2026-01"):954,  P("2026-02"):938,  P("2026-03"):2156, P("2026-04"):642},
        "Eating & Drinking Out":{P("2026-01"):1464, P("2026-02"):1171, P("2026-03"):1785, P("2026-04"):1256},
        "Electronics":          {P("2026-01"):1601, P("2026-02"):280,  P("2026-03"):0,    P("2026-04"):499},
        "Food Shopping":        {P("2026-01"):985,  P("2026-02"):1512, P("2026-03"):1615, P("2026-04"):1062},
        "Health & Wellbeing":   {P("2026-01"):667,  P("2026-02"):731,  P("2026-03"):1348, P("2026-04"):464},
        "Holidays":             {P("2026-01"):1611, P("2026-02"):0,    P("2026-03"):39,   P("2026-04"):1623},
        "Liam Sangha":          {P("2026-01"):300,  P("2026-02"):300,  P("2026-03"):1079, P("2026-04"):300},
        "Media & Subscriptions":{P("2026-01"):342,  P("2026-02"):13,   P("2026-03"):183,  P("2026-04"):31},
        "Professional Services":{P("2026-01"):134,  P("2026-02"):0,    P("2026-03"):7,    P("2026-04"):215},
        "Seatfrog":             {P("2026-01"):33,   P("2026-02"):29,   P("2026-03"):73,   P("2026-04"):31},
        "Sport & Leisure":      {P("2026-01"):158,  P("2026-02"):52,   P("2026-03"):108,  P("2026-04"):224},
        "Susan's Car Lease":    {P("2026-01"):526,  P("2026-02"):526,  P("2026-03"):526,  P("2026-04"):526},
        "Tax":                  {P("2026-01"):1230, P("2026-02"):0,    P("2026-03"):0,    P("2026-04"):0},
        "Travel & Hotels":      {P("2026-01"):154,  P("2026-02"):97,   P("2026-03"):35,   P("2026-04"):0},
        "Travel & Transport":   {P("2026-01"):184,  P("2026-02"):100,  P("2026-03"):178,  P("2026-04"):567},
        "University Fees":      {P("2026-01"):0,    P("2026-02"):4768, P("2026-03"):0,    P("2026-04"):0},
        "Utilities":            {P("2026-01"):1548, P("2026-02"):1251, P("2026-03"):413,  P("2026-04"):840},
        "Salary":               {P("2026-01"):6364, P("2026-02"):6394, P("2026-03"):6271, P("2026-04"):6535},
        "Fidelity":             {P("2026-01"):401,  P("2026-02"):12195,P("2026-03"):30954,P("2026-04"):22500},
    }


def load_income_history():
    """
    Hardcoded Fidelity income by account, Jan–Apr 2026.
    Updated from full-year TransactionHistory export 30/05/2026.
    These values are read-only — never overwritten by new uploads.
    """
    P = lambda s: pd.Period(s, "M")
    return {
        "2000001606": {P("2026-01"):5940,  P("2026-02"):6474,  P("2026-03"):10308, P("2026-04"):7631},
        "SANX002282": {P("2026-01"):2263,  P("2026-02"):2662,  P("2026-03"):3503,  P("2026-04"):3088},
        "SANX002617": {P("2026-01"):1936,  P("2026-02"):2185,  P("2026-03"):2895,  P("2026-04"):2713},
        "2000001604": {P("2026-01"):1468,  P("2026-02"):1653,  P("2026-03"):2944,  P("2026-04"):1975},
        "AS10303823": {P("2026-01"):266,   P("2026-02"):301,   P("2026-03"):266,   P("2026-04"):298},
        "SANQ000468": {P("2026-01"):3,     P("2026-02"):287,   P("2026-03"):249,   P("2026-04"):0},
        "SANX002936": {P("2026-01"):62,    P("2026-02"):60,    P("2026-03"):60,    P("2026-04"):74},
    }


def build_spending_pivot(amex_df, bar_df):
    combined = pd.concat([amex_df, bar_df])
    pivot = combined.pivot_table(
        index="category", columns="month", values="spend",
        aggfunc="sum", fill_value=0
    )
    months = sorted(pivot.columns)
    pivot = pivot.reindex(columns=months, fill_value=0)
    pivot["Total"] = pivot[months].sum(axis=1)
    pivot = pivot.reindex(CATEGORIES + ["Salary", "Fidelity"], fill_value=0)

    # Inject hardcoded Jan–Apr history — never overwrite with live data
    HIST_CUTOFF = pd.Period("2026-06", "M")
    spend_hist = load_spend_history()
    for cat, hist_vals in spend_hist.items():
        for period, val in hist_vals.items():
            if period < HIST_CUTOFF:
                if period not in pivot.columns:
                    pivot[period] = 0
                pivot.loc[cat, period] = val
    # Re-sort columns
    all_cols = sorted([c for c in pivot.columns if c != "Total"])
    pivot = pivot.reindex(columns=all_cols + ["Total"], fill_value=0)
    pivot["Total"] = pivot[all_cols].sum(axis=1)
    months = all_cols
    return pivot, months


def build_fidelity_pivot(fid_df):
    if fid_df.empty:
        pivot = pd.DataFrame()
    else:
        fid_df = fid_df[fid_df["account"].isin(ACCOUNT_OWNER)].copy()
        pivot = fid_df.pivot_table(
            index="account", columns="month", values="amount",
            aggfunc="sum", fill_value=0
        )

    # Inject hardcoded Jan–Apr income history — never overwrite with live data
    HIST_CUTOFF = pd.Period("2026-06", "M")
    inc_hist = load_income_history()
    for acc, hist_vals in inc_hist.items():
        for period, val in hist_vals.items():
            if period < HIST_CUTOFF:
                if acc not in pivot.index:
                    pivot.loc[acc] = 0
                if period not in pivot.columns:
                    pivot[period] = 0
                pivot.loc[acc, period] = val

    if pivot.empty:
        return pd.DataFrame(), []

    months = sorted([c for c in pivot.columns if c != "Total"])
    pivot = pivot.reindex(columns=months, fill_value=0)
    pivot["Total"] = pivot[months].sum(axis=1)
    pivot = pivot.sort_values("Total", ascending=False)
    return pivot, months


def estimate_future_months(pivot, actual_months, future_months, may_scale=18/31,
                           skip_funds=None):
    """
    Estimate future month values.
    - May (partial, in future_months): scaled up from its own raw value.
    - Other future months: median of actuals.
    - University Fees: Feb actual + Sep estimate only.
    - Tax: no extrapolation.
    - Stocks: use confirmed dividend payment months from market data.
    """
    import numpy as np

    result = pivot.copy()
    MAY = pd.Period("2026-05", "M")

    # Confirmed 2026 dividend payment months for stocks
    # Sold stocks get empty list (no future payments)
    # Amounts estimated from actuals; for stocks with no 2026 actuals, use 0 (unknown amount)
    STOCK_DIV_MONTHS = {
        # Currently held
        "AVIVA, ORD GBP0.328947368 (AV.)":          [10],        # Oct (May already in actuals)
        "AUTOTRADER GROUP PLC, ORD GBP0.01 (AUTO)":  [9],         # Sep interim (~Aug ex-div)
        "RELX PLC, ORD GBP0.1444 (REL)":             [6, 9],      # Jun final + Sep interim
        "THE SAGE GROUP PLC, GBP0.01051948 (SGE)":   [6],         # Jun interim (~May ex-div)
        "WEIR GROUP, ORD GBP0.125 (WEIR)":           [6, 11],     # Jun + Nov
        "VODAFONE GROUP, ORD USD0.2095238 (VOD)":    [7],         # Jul final FY26
        # Sold — no future payments
        "BP, ORD USD0.25 (BP.)":                     [],
        "SHELL PLC, ORD EUR0.07 (SHEL)":             [],
        "LEGAL & GENERAL GROUP, ORD GBP0.025 (LGEN)": [],
        "ITV, ORD GBP0.10 (ITV)":                   [],
        "SAP SE, ORD NPV (SAP)":                     [],
        # Funds with confirmed quarterly payment schedule
        # M&G Global Floating Rate HY: quarterly, pay dates 31 Mar/30 Jun/30 Sep/31 Dec
        # (Mar already in actuals; remaining: Jun, Sep, Dec)
        "M&G Global Floating Rate High Yield Fund Sterling I-H Inc": [6, 9, 12],
    }

    skip_funds = skip_funds or set()

    for idx in result.index:
        # Skip funds that were injected with precise pence-based estimates
        if idx in skip_funds:
            continue
        if idx == "University Fees":
            feb_val = 0
            for m in actual_months:
                if m.month == 2:
                    feb_val = float(result.loc[idx, m])
            for m in future_months:
                if m == MAY:
                    raw = float(result.loc[idx, m]) if m in result.columns else 0
                    result.loc[idx, m] = round(raw / may_scale) if raw > 0 else 0
                else:
                    result.loc[idx, m] = round(feb_val) if m.month == 9 else 0
            result.loc[idx, "Total"] = round(sum(float(result.loc[idx, m]) for m in actual_months + future_months))
            continue

        if idx == "Tax":
            for m in future_months:
                result.loc[idx, m] = 0
            result.loc[idx, "Total"] = round(sum(float(result.loc[idx, m]) for m in actual_months))
            continue

        if idx == "Salary":
            # Salary is a fixed payment — never scale up for partial months
            # Use May actual as-is, project forward using median of actuals
            actuals_vals = [float(result.loc[idx, m]) for m in actual_months if m in result.columns]
            monthly_salary = round(np.median(actuals_vals)) if actuals_vals else 0
            for m in future_months:
                if m == MAY:
                    pass  # keep May raw value unchanged
                else:
                    result.loc[idx, m] = monthly_salary
            result.loc[idx, "Total"] = round(sum(float(result.loc[idx, m]) for m in actual_months + future_months))
            continue

        if idx == "Holidays":
            ANNUAL_BUDGET = 12000
            actual_total = sum(float(result.loc[idx, m]) for m in actual_months)
            remaining = max(0, ANNUAL_BUDGET - actual_total)
            per_month = round(remaining / len(future_months)) if future_months else 0
            for m in future_months:
                result.loc[idx, m] = per_month
            result.loc[idx, "Total"] = round(actual_total + remaining)
            continue

        # ── Stock / equity dividends ──────────────────────────────────────────
        # Schedule: { fund_name: [(payment_month, pence_per_share), ...] }
        # Pence-per-share from confirmed 2026 announcements (dividenddata.co.uk / company IR)
        # Per-account cash estimate = units_held * pence / 100
        # If no units data available, falls back to historical mean payment
        STOCK_DIV_SCHEDULE = {
            # Semi-annual. Final: pay May 14 (already in actuals). Interim: pay Oct 16 (ex Aug 28)
            "AVIVA, ORD GBP0.328947368 (AV.)":         [(10, 13.1)],
            # Semi-annual. Interim: pay Jan 26 (already in actuals). Final: ~Sep 26 (ex ~Aug 27)
            "AUTOTRADER GROUP PLC, ORD GBP0.01 (AUTO)": [(9,  7.1)],
            # Semi-annual. Final: pay Jun 18 (ex May 7). Interim: ~Sep 11 (ex ~Aug 7)
            "RELX PLC, ORD GBP0.1444 (REL)":            [(6, 48.0), (9, 19.5)],
            # Semi-annual. Final: pay Feb 10 (already in actuals). Interim: ~Jun 27 (ex ~May 29)
            "THE SAGE GROUP PLC, GBP0.01051948 (SGE)":  [(6,  7.5)],
            # Semi-annual. Final: pay May 29 (already in actuals). Interim: ~Nov 4 (ex ~Oct 2)
            "WEIR GROUP, ORD GBP0.125 (WEIR)":          [(11, 19.6)],
            # Semi-annual. Interim: pay Feb 5 (already in actuals). Final: pay Jul 31 (ex Jun 5)
            "VODAFONE GROUP, ORD USD0.2095238 (VOD)":   [(7,  1.9)],
            # Semi-annual. Final: pay Jun 4 (ex Apr 23 — already past). Interim: pay Sep 25 (ex Aug 20)
            "LEGAL & GENERAL GROUP, ORD GBP0.025 (LGEN)": [(6, 15.36), (9, 6.12)],
            # Sold — no future payments
            "BP, ORD USD0.25 (BP.)":                     [],
            "SHELL PLC, ORD EUR0.07 (SHEL)":             [],
            "ITV, ORD GBP0.10 (ITV)":                    [],
            "SAP SE, ORD NPV (SAP)":                     [],
            # Funds with confirmed quarterly schedule — pay months only, no pence-per-share
            # M&G GFRHYF: quarterly pay dates 31 Mar/30 Jun/30 Sep/31 Dec (ex-div ~1 month before)
            "M&G Global Floating Rate High Yield Fund Sterling I-H Inc": [(6, None), (9, None), (12, None)],
        }

        if idx in STOCK_DIV_SCHEDULE:
            schedule = STOCK_DIV_SCHEDULE[idx]
            # For each future payment, try to compute from units × pence, else use mean of actuals

            # Gather actual payments for fallback
            actual_vals = [float(result.loc[idx, m]) for m in actual_months if float(result.loc[idx, m]) > 0]
            if not actual_vals:
                actual_vals = [float(result.loc[idx, m]) for m in future_months if m in result.columns and float(result.loc[idx, m]) > 0]

            for m in future_months:
                if m == MAY:
                    raw = float(result.loc[idx, m]) if m in result.columns else 0
                    if raw > 0:
                        result.loc[idx, m] = round(raw / may_scale)
                        continue
                # Check if this month has a scheduled payment
                pay_entry = next(((mo, ppm) for mo, ppm in schedule if mo == m.month), None)
                if pay_entry:
                    pay_month, ppm = pay_entry
                    if actual_vals:
                        per_payment = round(float(np.mean(actual_vals)))
                    else:
                        per_payment = 0
                    result.loc[idx, m] = per_payment
                else:
                    result.loc[idx, m] = 0
            result.loc[idx, "Total"] = round(sum(float(result.loc[idx, m]) for m in actual_months + future_months))
            continue

        # Any unscheduled stock: zero future months
        idx_upper = str(idx).upper()
        is_stock = any(p in idx_upper for p in [", ORD ", "ORD GBP", "ORD USD", "ORD EUR", "ORD NPV"])
        if is_stock:
            for m in future_months:
                result.loc[idx, m] = 0
            result.loc[idx, "Total"] = round(sum(float(result.loc[idx, m]) for m in actual_months))
            continue

        # Default: median of actuals for regular monthly/recurring income
        actuals = [float(result.loc[idx, m]) if m in result.columns else 0 for m in actual_months]
        non_zero = [v for v in actuals if v > 0]
        median_est = float(np.median(non_zero)) if non_zero else 0

        for m in future_months:
            if m == MAY:
                raw = float(result.loc[idx, m]) if m in result.columns else 0
                result.loc[idx, m] = round(raw / may_scale) if raw > 0 else round(median_est)
            else:
                result.loc[idx, m] = round(median_est)

    result["Total"] = result[actual_months + future_months].sum(axis=1)
    return result


def build_fidelity_fund_pivot(fid_df):
    pass


def build_account_fund_pivot(fid_df):
    """
    Returns OrderedDict: person -> { account -> DataFrame(fund x month) }
    Filtered to family accounts only, grouped by person in FAMILY_ORDER,
    accounts within each person sorted by total desc, funds alphabetically.
    """
    if fid_df.empty:
        return {}, []

    months = sorted(fid_df["month"].unique())

    # Filter to family accounts only
    fid_df = fid_df[fid_df["account"].isin(ACCOUNT_OWNER)].copy()

    result = {}
    for person in FAMILY_ORDER:
        person_accounts = [acc for acc, owner in ACCOUNT_OWNER.items() if owner == person]
        # Sort by total desc
        acc_totals = (
            fid_df[fid_df["account"].isin(person_accounts)]
            .groupby("account")["amount"].sum()
            .reindex(person_accounts, fill_value=0)
            .sort_values(ascending=False)
        )
        person_data = {}
        for acc in acc_totals.index:
            sub = fid_df[fid_df["account"] == acc]
            if sub.empty:
                continue
            pivot = sub.pivot_table(
                index="fund", columns="month", values="amount",
                aggfunc="sum", fill_value=0
            )
            pivot = pivot.reindex(columns=months, fill_value=0)
            pivot["Total"] = pivot[months].sum(axis=1)
            pivot = pivot.sort_index()
            person_data[acc] = pivot
        if person_data:
            result[person] = person_data

    return result, months


# ── Excel styles helper ────────────────────────────────────────────────────────
NUM_FMT = '#,##0;(#,##0);"-"'

def hdr_style(ws, row, col, value, fill, font):
    c = ws.cell(row=row, column=col, value=value)
    c.font = font
    c.fill = fill
    c.alignment = Alignment(horizontal="left" if col == 1 else "right", vertical="center")
    return c

def data_cell(ws, row, col, value, font, fill=None, num_fmt=NUM_FMT, halign="right"):
    c = ws.cell(row=row, column=col, value=value)
    c.font = font
    c.number_format = num_fmt
    c.alignment = Alignment(horizontal=halign, vertical="center")
    c.border = Border(bottom=Side(style="thin", color="D9E1F2"))
    if fill:
        c.fill = fill
    return c


def _read_account_summary_rows(account_summary_path):
    """Yield all rows from AccountSummary CSV as dicts."""
    if not account_summary_path or not os.path.exists(account_summary_path):
        return
    with open(account_summary_path, encoding="utf-8-sig") as f:
        content = f.read()
    lines = content.replace("\r", "").split("\n")
    try:
        hi = max(i for i, l in enumerate(lines) if l.startswith("Type,Holdings,Account number"))
        import io as _io_helper
        for row in csv.DictReader(_io_helper.StringIO("\n".join(lines[hi:]))):
            yield row
    except (ValueError, StopIteration):
        return


def build_acc_holdings(account_summary_path, fidelity_path=None, inc_income_df=None):
    """
    Returns dict: { acc: { fund: {'value', 'units', 'price_may26', 'monthly_values'} } }
    Only accumulation (Acc/ACC) funds in family accounts.
    monthly_values: { pd.Period -> value } based on cumulative units × price.
    """
    FAMILY_ACCS = set(ACCOUNT_OWNER.keys())
    result = {}
    if not account_summary_path or not os.path.exists(account_summary_path):
        return result

    with open(account_summary_path, encoding="utf-8-sig") as f:
        content = f.read()
    lines = content.replace("\r", "").split("\n")
    header_idx = max(i for i, l in enumerate(lines) if l.startswith("Type,Holdings,Account number"))
    reader = csv.DictReader(io.StringIO("\n".join(lines[header_idx:])))
    for row in reader:
        if row.get("Type", "").strip() != "Asset": continue
        acc = row.get("Account number", "").strip()
        if acc not in FAMILY_ACCS: continue
        fund = row.get("Holdings", "").strip().replace('"', '')
        if " Acc" not in fund and " ACC" not in fund: continue
        try:
            val = float(row.get("Value (£)", "0").replace(",", ""))
            qty = float(row.get("Quantity", "0").replace(",", ""))
        except (ValueError, TypeError):
            continue
        if acc not in result: result[acc] = {}
        price_may26 = val / qty if qty > 0 else 0
        result[acc][fund] = {"value": val, "units": qty, "price_may26": price_may26, "monthly_values": {}}

    if not fidelity_path or not os.path.exists(fidelity_path):
        return result

    # ── Build price history and unit history from transaction CSV ──────────────
    with open(fidelity_path) as f:
        fid_lines = f.read().replace("\r", "").split("\n")
    fid_start = next(i for i, l in enumerate(fid_lines) if l.startswith("Order date"))
    fid_reader = csv.DictReader(io.StringIO("\n".join(fid_lines[fid_start:])))
    fid_rows = [r for r in fid_reader if r.get("Status", "").strip() == "Completed"]

    # Price observations: (fund, period) -> price
    from collections import defaultdict
    price_obs = defaultdict(dict)   # fund -> {period: price}
    unit_changes = defaultdict(lambda: defaultdict(float))  # (acc, fund) -> {period: net_units}

    for r in fid_rows:
        t = r["Transaction type"].strip()
        if t not in ("Buy", "Sell"): continue
        acc = r["Account Number"].strip()
        inv = r["Investments"].strip().replace('"', '')
        if " Acc" not in inv and " ACC" not in inv: continue
        ds = (r.get("Completion date") or r["Order date"]).strip()
        if not ds or ds.lower() == "pending": ds = r["Order date"].strip()
        try:
            dt = pd.to_datetime(ds, dayfirst=True)
            period = dt.to_period("M")
            price = float(r["Price per unit"]) if r.get("Price per unit") else 0
            qty = float(r["Quantity"]) if r.get("Quantity") else 0
        except (ValueError, TypeError):
            continue
        if price > 0:
            price_obs[inv][period] = price
        if acc in FAMILY_ACCS:
            if t == "Buy":
                unit_changes[(acc, inv)][period] += qty
            else:
                unit_changes[(acc, inv)][period] -= qty

    # ── Build monthly value for each (acc, fund) ────────────────────────────────
    MAY_2026 = pd.Period("2026-05", "M")
    JAN_2026 = pd.Period("2026-01", "M")
    DEC_2026 = pd.Period("2026-12", "M")
    all_periods = [JAN_2026 + i for i in range(12)]  # Jan–Dec 2026

    # Annual yield overrides for funds where transaction history is sparse
    # For Acc funds, this represents total return (income reinvested into NAV)
    # Sources: Fidelity key statistics pages
    FUND_ANNUAL_YIELD = {
        "Aegon High Yield Bond B Acc":                    0.0732,  # 7.32% distribution yield
        "Schroder High Yield Opportunities Fund Z Acc":   0.0769,  # 7.69% distribution yield
        "WS Guinness Global Energy Fund I Acc":           0.0235,  # 2.35% historic yield
    }

    def interpolate_price(fund, period, price_may26):
        """Get price for a given period using observations + interpolation/extrapolation."""
        obs = dict(price_obs.get(fund, {}))
        obs[MAY_2026] = price_may26  # anchor at May 26
        periods_sorted = sorted(obs.keys())
        if not periods_sorted:
            return price_may26
        if period in obs:
            return obs[period]
        before = [p for p in periods_sorted if p <= period]
        after  = [p for p in periods_sorted if p >= period]
        if before and after:
            p0, p1 = before[-1], after[0]
            if p0 == p1: return obs[p0]
            w = (period - p0).n / max((p1 - p0).n, 1)
            return obs[p0] + w * (obs[p1] - obs[p0])
        elif before:
            p0 = before[-1]
            # Always use known annual yield for forward extrapolation — never use
            # short-term price trends which can be negative
            annual = FUND_ANNUAL_YIELD.get(fund, 0.05)
            monthly_growth = max(0.0, (1 + annual) ** (1/12) - 1)
            months_fwd = (period - p0).n
            return obs[p0] * ((1 + monthly_growth) ** months_fwd)
        else:
            # Only future observations — extrapolate backward using annual yield
            p0 = after[0]
            annual = FUND_ANNUAL_YIELD.get(fund, 0.05)
            monthly_growth = max(0.0, (1 + annual) ** (1/12) - 1)
            months_back = (p0 - period).n
            return obs[p0] / ((1 + monthly_growth) ** months_back)

    for acc, funds in result.items():
        for fund, data in funds.items():
            price_may26 = data["price_may26"]
            units_may26 = data["units"]

            # Work backwards/forwards from May 2026 to get units each month
            monthly_vals = {}
            monthly_units = {}   # units held at START of each month

            # Go forward from May 2026
            running = units_may26
            for p in sorted(all_periods):
                if p >= MAY_2026:
                    delta = unit_changes.get((acc, fund), {}).get(p, 0)
                    if p > MAY_2026:
                        running += delta
                    price = interpolate_price(fund, p, price_may26)
                    monthly_vals[p] = round(running * price)
                    monthly_units[p] = running
            # Go backward from May 2026
            running = units_may26
            # Find earliest buy date for this fund/account from unit_changes
            fund_unit_changes = unit_changes.get((acc, fund), {})
            # If there are no buy transactions, the fund was held before our data window
            # Use Jan 2026 as the start (fund was already held)
            buys = [p for p, d in fund_unit_changes.items() if d > 0]
            earliest_buy = min(buys) if buys else JAN_2026

            for p in sorted([pp for pp in all_periods if pp < MAY_2026], reverse=True):
                delta = fund_unit_changes.get(p, 0)
                running -= delta  # undo this month's purchase
                if running <= 0 and p < earliest_buy:
                    # Fund genuinely not held before this point
                    monthly_vals[p] = 0
                    monthly_units[p] = 0
                else:
                    price = interpolate_price(fund, p, price_may26)
                    monthly_vals[p] = round(running * price)
                    monthly_units[p] = running

            # Appreciation = Inc dividend ppu × Acc units held at start of month
            # ppu derived from: total Inc income paid / total Inc units held that month
            # For months with no Inc data, fall back to interpolated price change
            price_appreciation = {}

            # Build Inc ppu lookup: fund_base -> month -> ppu (pence per unit)
            inc_ppu = {}  # { (fund_base, period) -> ppu }
            INC_EQUIV = {
                "Aegon High Yield Bond B Acc":                  "Aegon High Yield Bond B Inc",
                "Schroder High Yield Opportunities Fund Z Acc": "Schroder High Yield Opportunities Fund Z Inc",
            }
            inc_fund_name = INC_EQUIV.get(fund)
            if inc_fund_name and inc_income_df is not None and not inc_income_df.empty:
                inc_df = inc_income_df[inc_income_df["fund"] == inc_fund_name]
                # Total Inc units across all accounts (from AccountSummary)
                total_inc_units = sum(
                    float(str(row.get("Quantity","0")).replace(",",""))
                    for row in _read_account_summary_rows(account_summary_path)
                    if row.get("Holdings","").strip() == inc_fund_name
                    and row.get("Type","").strip() == "Asset"
                )
                if total_inc_units > 0:
                    monthly_inc = inc_df.groupby("month")["amount"].sum()
                    for m_period, total_inc_amt in monthly_inc.items():
                        ppu = total_inc_amt / total_inc_units  # in £
                        inc_ppu[(fund, m_period)] = ppu

            sorted_periods = sorted(monthly_vals.keys())
            for i, p in enumerate(sorted_periods):
                if i == 0:
                    price_appreciation[p] = 0
                else:
                    p_prev = sorted_periods[i - 1]
                    units_start = monthly_units.get(p_prev, 0)
                    # Use Inc ppu if available for this period
                    ppu = inc_ppu.get((fund, p))
                    if ppu is not None and units_start > 0:
                        price_appreciation[p] = round(units_start * ppu)
                    else:
                        # Fallback: interpolated price change × units — floor at 0
                        price_now  = interpolate_price(fund, p, price_may26)
                        price_prev = interpolate_price(fund, p_prev, price_may26)
                        price_appreciation[p] = max(0, round(units_start * (price_now - price_prev)))

            data["monthly_values"] = monthly_vals
            data["price_appreciation"] = price_appreciation

    return result


def load_history(path: str = None) -> dict:
    """
    Returns hardcoded historical wealth data (Jan 2024 – May 2026).
    Previous_Data_-_Sheet.csv is no longer required as an input file.
    This data is read-only and is never overwritten by live CSV uploads.
    """
    P = lambda s: pd.Period(s, "M")
    return {
        "All Fidelity accounts": {
            P("2024-01"):2376960, P("2024-02"):2466948, P("2024-03"):2532566, P("2024-04"):2467355,
            P("2024-05"):2494411, P("2024-06"):2557748, P("2024-07"):2556915, P("2024-08"):2592210,
            P("2024-09"):2602223, P("2024-10"):2604253, P("2024-11"):2986977, P("2024-12"):2967004,
            P("2025-01"):3093088, P("2025-02"):3013405, P("2025-03"):2998913, P("2025-04"):2920244,
            P("2025-05"):3073417, P("2025-06"):3150166, P("2025-07"):3279978, P("2025-08"):3302847,
            P("2025-09"):3334836, P("2025-10"):3481679, P("2025-11"):3511850, P("2025-12"):3575468,
            P("2026-01"):3622768, P("2026-02"):3641768, P("2026-03"):3665439, P("2026-04"):3646654,
            P("2026-05"):3670357,
        },        "PS Arriva pension (Def Bene)": {
            P("2024-01"):58951,  P("2024-02"):59245,  P("2024-03"):59542,  P("2024-04"):59839,
            P("2024-05"):60139,  P("2024-06"):60439,  P("2024-07"):60742,  P("2024-08"):61045,
            P("2024-09"):61350,  P("2024-10"):61657,  P("2024-11"):61965,  P("2024-12"):62275,
            P("2025-01"):62587,  P("2025-02"):62900,  P("2025-03"):63214,  P("2025-04"):63530,
            P("2025-05"):63848,  P("2025-06"):64167,  P("2025-07"):64488,  P("2025-08"):64810,
            P("2025-09"):65134,  P("2025-10"):65460,  P("2025-11"):65787,  P("2025-12"):66116,
            P("2026-01"):66447,  P("2026-02"):66779,  P("2026-03"):67113,  P("2026-04"):67449,
            P("2026-05"):67786,
        },
        "Paul Explo Scottish Widows": {
            P("2024-01"):913,    P("2024-02"):8049,   P("2024-03"):15184,  P("2024-04"):21013,
            P("2024-05"):22084,  P("2024-06"):28724,  P("2024-07"):34754,  P("2024-08"):40784,
            P("2024-09"):47000,  P("2024-10"):53030,  P("2024-11"):60350,  P("2024-12"):73400,
            P("2025-01"):79430,  P("2025-02"):13256,  P("2025-03"):19286,  P("2025-04"):24000,
            P("2025-05"):30030,  P("2025-06"):36060,  P("2025-07"):39600,  P("2025-08"):47150,
            P("2025-09"):53505,  P("2025-10"):1500,   P("2025-11"):7530,   P("2025-12"):20965,
            P("2026-01"):7000,   P("2026-02"):13030,  P("2026-03"):19060,  P("2026-04"):25090,
            P("2026-05"):12000,
        },
        "Paul Pension": {
            P("2024-01"):1119525, P("2024-02"):1126955, P("2024-03"):1209103, P("2024-04"):1221890,
            P("2024-05"):1237856, P("2024-06"):1244045, P("2024-07"):1238096, P("2024-08"):1244286,
            P("2024-09"):1248966, P("2024-10"):1255211, P("2024-11"):1261487, P("2024-12"):1284913,
            P("2025-01"):1325372, P("2025-02"):1331999, P("2025-03"):1338659, P("2025-04"):1345352,
            P("2025-05"):1300435, P("2025-06"):1304984, P("2025-07"):1311509, P("2025-08"):1343960,
            P("2025-09"):1349580, P("2025-10"):1356328, P("2025-11"):1363110, P("2025-12"):1369925,
            P("2026-01"):1376775, P("2026-02"):1383659, P("2026-03"):1389272, P("2026-04"):1385000,
            P("2026-05"):1490236,
        },
        "Susan Fidelity Pension": {
            P("2025-09"):363535,  P("2025-10"):365353,  P("2025-11"):367179,  P("2025-12"):385900,
            P("2026-01"):387830,  P("2026-02"):389769,  P("2026-03"):391717,  P("2026-04"):393676,
            P("2026-05"):395644,
        },
        "Susan Pension": {
            P("2024-01"):1023600, P("2024-02"):1035344, P("2024-03"):1047279, P("2024-04"):1057937,
            P("2024-05"):1068646, P("2024-06"):1080505, P("2024-07"):1092422, P("2024-08"):1104395,
            P("2024-09"):1110104, P("2024-10"):1167356, P("2024-11"):1176452, P("2024-12"):1182455,
            P("2025-01"):1194387, P("2025-02"):1200359, P("2025-03"):1206361, P("2025-04"):1212392,
            P("2025-05"):1218454, P("2025-06"):1224547, P("2025-07"):1230669, P("2025-08"):1236823,
            P("2025-09"):1291390, P("2025-10"):1297847, P("2025-11"):1304336, P("2025-12"):1345864,
            P("2026-01"):1352593, P("2026-02"):1359356, P("2026-03"):1366153, P("2026-04"):1384465,
            P("2026-05"):1391387,
        },
        "Capita (RMSPS) - to 2018 (65 yrs)": {
            P("2024-01"):440596,  P("2024-02"):442799,  P("2024-03"):445013,  P("2024-04"):447238,
            P("2024-05"):449474,  P("2024-06"):451722,  P("2024-07"):453980,  P("2024-08"):456250,
            P("2024-09"):459672,  P("2024-10"):465637,  P("2024-11"):467187,  P("2024-12"):468743,
            P("2025-01"):470304,  P("2025-02"):471870,  P("2025-03"):473441,  P("2025-04"):475018,
            P("2025-05"):476600,  P("2025-06"):478187,  P("2025-07"):479779,  P("2025-08"):481377,
            P("2025-09"):482980,  P("2025-10"):478216,  P("2025-11"):479808,  P("2025-12"):481406,
            P("2026-01"):483009,  P("2026-02"):484618,  P("2026-03"):486231,  P("2026-04"):487850,
            P("2026-05"):489475,
        },
        "RMPP 2012 - 2023 (65 years)": {
            P("2024-01"):194206,  P("2024-02"):195177,  P("2024-03"):196152,  P("2024-04"):197133,
            P("2024-05"):198119,  P("2024-06"):199109,  P("2024-07"):200105,  P("2024-08"):201106,
            P("2024-09"):202111,  P("2024-10"):198700,  P("2024-11"):200190,  P("2024-12"):201692,
            P("2025-01"):203204,  P("2025-02"):204728,  P("2025-03"):206264,  P("2025-04"):207811,
            P("2025-05"):209369,  P("2025-06"):210940,  P("2025-07"):212522,  P("2025-08"):214116,
            P("2025-09"):215722,  P("2025-10"):217339,  P("2025-11"):218969,  P("2025-12"):220612,
            P("2026-01"):222266,  P("2026-02"):223933,  P("2026-03"):225613,  P("2026-04"):227305,
            P("2026-05"):229010,
        },
        "Collective Pension (2023 +)": {
            P("2024-01"):24779,   P("2024-02"):25318,   P("2024-03"):25856,   P("2024-04"):26395,
            P("2024-05"):26934,   P("2024-06"):27472,   P("2024-07"):28011,   P("2024-08"):28550,
            P("2024-09"):29088,   P("2024-10"):29627,   P("2024-11"):30166,   P("2024-12"):30704,
            P("2025-01"):31243,   P("2025-02"):31782,   P("2025-03"):32321,   P("2025-04"):32859,
            P("2025-05"):33468,   P("2025-06"):34076,   P("2025-07"):34685,   P("2025-08"):35293,
            P("2025-09"):35917,   P("2025-10"):36644,   P("2025-11"):37652,   P("2025-12"):38667,
            P("2026-01"):39690,   P("2026-02"):40721,   P("2026-03"):41759,   P("2026-04"):42805,
            P("2026-05"):43859,
        },
        "Cash Balance": {
            P("2024-01"):88013,   P("2024-02"):88453,   P("2024-03"):88895,   P("2024-04"):89340,
            P("2024-05"):89786,   P("2024-06"):90235,   P("2024-07"):90686,   P("2024-08"):91140,
            P("2024-09"):91596,   P("2024-10"):138005,  P("2024-11"):138005,  P("2024-12"):138695,
            P("2025-01"):139389,  P("2025-02"):140085,  P("2025-03"):140786,  P("2025-04"):141490,
            P("2025-05"):142197,  P("2025-06"):142908,  P("2025-07"):143623,  P("2025-08"):144341,
            P("2025-09"):145063,  P("2025-10"):154555,  P("2025-11"):155328,  P("2025-12"):156104,
            P("2026-01"):156885,  P("2026-02"):157669,  P("2026-03"):158458,  P("2026-04"):159250,
            P("2026-05"):160046,
        },
        "AVC Bonus Plan (Scottish widows)": {
            P("2024-01"):1466,    P("2024-02"):1486,    P("2024-03"):1505,    P("2024-04"):1525,
            P("2024-05"):1546,    P("2024-06"):1566,    P("2024-07"):1586,    P("2024-08"):1606,
            P("2024-09"):1627,    P("2024-10"):1648,    P("2024-11"):1668,    P("2024-12"):1689,
            P("2025-03"):14000,   P("2025-04"):14070,   P("2025-05"):14140,   P("2025-06"):14211,
            P("2025-07"):14282,   P("2025-08"):14354,   P("2025-09"):48174,   P("2025-10"):53174,
            P("2025-11"):58174,   P("2025-12"):63174,   P("2026-01"):68174,   P("2026-02"):73174,
            P("2026-03"):78174,   P("2026-04"):83174,   P("2026-05"):88174,
        },
        "MF": {
            P("2024-01"):180000,  P("2024-02"):180000,  P("2024-03"):180000,  P("2024-04"):180000,
            P("2024-05"):180000,  P("2024-06"):180000,  P("2024-07"):180000,  P("2024-08"):180000,
            P("2024-09"):180000,  P("2024-10"):180000,  P("2024-11"):180000,  P("2024-12"):187000,
            P("2025-01"):187000,  P("2025-02"):187000,  P("2025-03"):187000,  P("2025-04"):187000,
            P("2025-05"):187000,  P("2025-06"):187000,  P("2025-07"):187000,  P("2025-08"):190000,
            P("2025-09"):190000,  P("2025-10"):190000,  P("2025-11"):190000,  P("2025-12"):190000,
            P("2026-01"):190000,  P("2026-02"):190000,  P("2026-03"):190000,  P("2026-04"):190000,
            P("2026-05"):190000,
        },
        "House": {
            P("2024-01"):422694,  P("2024-02"):422906,  P("2024-03"):423117,  P("2024-04"):423329,
            P("2024-05"):423540,  P("2024-06"):423752,  P("2024-07"):423964,  P("2024-08"):424176,
            P("2024-09"):424388,  P("2024-10"):424600,  P("2024-11"):424813,  P("2024-12"):450000,
            P("2025-01"):450000,  P("2025-02"):450225,  P("2025-03"):450450,  P("2025-04"):450675,
            P("2025-05"):450901,  P("2025-06"):451126,  P("2025-07"):451352,  P("2025-08"):451577,
            P("2025-09"):451803,  P("2025-10"):452029,  P("2025-11"):452255,  P("2025-12"):481000,
            P("2026-01"):481241,  P("2026-02"):481481,  P("2026-03"):481722,  P("2026-04"):481963,
            P("2026-05"):482204,
        },
        "Liam ISA": {
            P("2024-01"):9344,    P("2024-02"):9344,    P("2024-03"):9344,    P("2024-04"):9344,
            P("2024-05"):10500,   P("2024-06"):10500,   P("2024-07"):30100,   P("2024-08"):32000,
            P("2024-09"):33000,   P("2024-10"):33000,   P("2024-11"):34000,   P("2024-12"):34000,
            P("2025-01"):34500,   P("2025-02"):34500,   P("2025-03"):34500,   P("2025-04"):34500,
            P("2025-05"):41000,   P("2025-06"):41000,   P("2025-07"):43500,   P("2025-08"):43500,
            P("2025-09"):47500,   P("2025-10"):47500,   P("2025-11"):50000,   P("2025-12"):50000,
            P("2026-01"):50000,   P("2026-02"):50000,   P("2026-03"):50000,   P("2026-04"):55617,
            P("2026-05"):55617,
        },
        "Jaynes ISA": {
            P("2024-01"):7558,    P("2024-02"):7558,    P("2024-03"):7558,    P("2024-04"):7558,
            P("2024-05"):7558,    P("2024-06"):7558,    P("2024-07"):7558,    P("2024-08"):9104,
            P("2024-09"):9104,    P("2024-10"):9104,    P("2024-11"):9104,    P("2024-12"):9104,
            P("2025-01"):9104,    P("2025-02"):9104,    P("2025-03"):9104,    P("2025-04"):9104,
            P("2025-05"):9104,    P("2025-06"):9104,    P("2025-07"):9104,    P("2025-08"):9104,
            P("2025-09"):9104,    P("2025-10"):9104,    P("2025-11"):9104,    P("2025-12"):9104,
            P("2026-01"):9104,    P("2026-02"):9104,    P("2026-03"):9104,    P("2026-04"):9104,
            P("2026-05"):9104,
        },
        "Cars Paul": {
            P("2024-01"):48182,   P("2024-02"):47379,   P("2024-03"):46589,   P("2024-04"):45813,
            P("2024-05"):45049,   P("2024-06"):44298,   P("2024-07"):47500,   P("2024-08"):46708,
            P("2024-09"):47000,   P("2024-10"):46217,   P("2024-11"):45446,   P("2024-12"):43000,
            P("2025-01"):42283,   P("2025-02"):41579,   P("2025-03"):44000,   P("2025-04"):43267,
            P("2025-05"):44000,   P("2025-06"):43267,   P("2025-07"):42546,   P("2025-08"):41836,
            P("2025-09"):41139,   P("2025-10"):40454,   P("2025-11"):39779,   P("2025-12"):39116,
            P("2026-01"):45000,   P("2026-02"):44250,   P("2026-03"):43513,   P("2026-04"):42787,
            P("2026-05"):42074,
        },
        "Cash": {
            P("2024-01"):12000,   P("2024-02"):11000,   P("2024-03"):7000,    P("2024-04"):3000,
            P("2024-05"):5000,    P("2024-06"):34000,   P("2024-07"):5000,    P("2024-08"):5000,
            P("2024-09"):5000,    P("2024-10"):5000,    P("2024-11"):5000,    P("2024-12"):5000,
            P("2025-01"):5000,    P("2025-02"):5000,    P("2025-03"):5000,    P("2025-04"):5000,
            P("2025-05"):44000,   P("2025-06"):10000,   P("2025-07"):5000,    P("2025-08"):5000,
            P("2025-09"):5000,    P("2025-10"):5000,    P("2025-11"):15000,   P("2025-12"):6000,
            P("2026-01"):6000,    P("2026-02"):5500,    P("2026-03"):5500,    P("2026-04"):13000,
            P("2026-05"):13000,
        },
    }


def build_summary_data(account_summary_path, all_months):
    """Build summary data. History is embedded in load_history() — no CSV needed."""
    FIDELITY_ACCS = {
        'AW10032966','SANX002282','2000001606','SANQ000468',
        'SANX002936','AW10261123','SANX002617','2000001604',
        'AW10580794','AS10303823','AG10131710'
    }

    # ── Fidelity total from AccountSummary ────────────────────────────────────
    fidelity_total = 0
    fidelity_by_acc = {}
    if account_summary_path and os.path.exists(account_summary_path):
        with open(account_summary_path, encoding="utf-8-sig") as f:
            content = f.read()
        lines = content.replace("\r", "").split("\n")
        header_idx = max(i for i, l in enumerate(lines) if l.startswith("Type,Holdings,Account number"))
        reader = csv.DictReader(io.StringIO("\n".join(lines[header_idx:])))
        for row in reader:
            if row.get("Type", "").strip() != "Account": continue
            acc = row.get("Account number", "").strip()
            if acc not in FIDELITY_ACCS: continue
            try:
                v = float(row.get("Value (£)", "0").replace(",", ""))
            except (ValueError, TypeError):
                v = 0
            fidelity_by_acc[acc] = v
            fidelity_total += v

    # ── Find May 2026 period index ────────────────────────────────────────────
    MAY_2026 = pd.Period("2026-05", "M")
    may_idx = next((i for i, m in enumerate(all_months) if m == MAY_2026), 0)

    def project_monthly(may_value, annual_pct=None, monthly_add=None):
        """Project values forward/backward from May 2026."""
        vals = {}
        for i, m in enumerate(all_months):
            offset = i - may_idx  # months from May 2026
            if annual_pct is not None:
                # Compound monthly: (1 + annual_pct) ^ (offset/12)
                vals[m] = round(may_value * ((1 + annual_pct) ** (offset / 12)))
            elif monthly_add is not None:
                vals[m] = round(max(0, may_value + offset * monthly_add))
        return vals

    # ── PS Arriva (Defined Benefit) ────────────────────────────────────────────
    arriva = project_monthly(67786, annual_pct=0.05)

    # ── Expleo Scottish Widows ────────────────────────────────────────────────
    expleo_sw = project_monthly(12000, monthly_add=6000)

    # ── Susan's pensions (5%/year growth from May 2026 estimates) ─────────────
    susan_pensions = {
        "Capita (RMSPS) – to 2018 (65 yrs)":    project_monthly(489475, annual_pct=0.05),
        "RMPP 2012–2023 (65 years)":             project_monthly(229010, annual_pct=0.05),
        "Collective Pension (2023+)":            project_monthly(43859,  annual_pct=0.05),
        "Cash Balance":                          project_monthly(160046, annual_pct=0.05),
        "AVC Bonus Plan (Scottish Widows)":      project_monthly(88174,  annual_pct=0.05),
    }

    # Susan's Fidelity SIPP (2000001604)
    # Jan–Apr: use 'Susan Fidelity Pension' history row
    # May+: start from AccountSummary value, updated with income growth in write_excel
    susan_fidelity_sipp = fidelity_by_acc.get("2000001604", 0)
    susan_fidelity_sipp_vals = {m: susan_fidelity_sipp for m in all_months}

    # ── Historical data (hardcoded — no file needed) ──────────────────────────
    history = load_history()

    # Override pre-cutoff months with history values for Susan Fidelity Pension
    susan_fid_hist = history.get("Susan Fidelity Pension", {})
    HISTORY_CUTOFF_PERIOD = pd.Period("2026-07", "M")
    for m in all_months:
        if m < HISTORY_CUTOFF_PERIOD and m in susan_fid_hist:
            susan_fidelity_sipp_vals[m] = susan_fid_hist[m]

    # ── Liam Fidelity (AS10303823 + AG10131710) ───────────────────────────────
    liam_fid_total = sum(fidelity_by_acc.get(a, 0) for a in ["AS10303823", "AG10131710"])

    # ── Jayne Fidelity (SANX002936) ───────────────────────────────────────────
    jayne_fid_total = fidelity_by_acc.get("SANX002936", 0)

    # ── Barclays balance by month ──────────────────────────────────────────────
    # Jan-May: established values (manual/previous tracking)
    # June: actual Barclays balance £6,916.87 (confirmed)
    # Jul+: hold at June balance pending further data
    barclays_by_month = {
        pd.Period("2026-01", "M"): 6000,
        pd.Period("2026-02", "M"): 6000,
        pd.Period("2026-03", "M"): 5500,
        pd.Period("2026-04", "M"): 5500,
        pd.Period("2026-05", "M"): 13000,
    }
    BARCLAYS_JUNE_BALANCE = 6916.87
    JUNE_2026_BAR = pd.Period("2026-06", "M")
    barclays_by_month[JUNE_2026_BAR] = round(BARCLAYS_JUNE_BALANCE)
    for m in all_months:
        if m not in barclays_by_month:
            barclays_by_month[m] = round(BARCLAYS_JUNE_BALANCE)

    return {
        "fidelity_total": fidelity_total,
        "fidelity_by_acc": fidelity_by_acc,
        "arriva": arriva,
        "expleo_sw": expleo_sw,
        "susan_pensions": susan_pensions,
        "susan_fidelity_sipp": susan_fidelity_sipp_vals,
        "barclays_by_month": barclays_by_month,
        "liam_fid_total": liam_fid_total,
        "jayne_fid_total": jayne_fid_total,
        "may_idx": may_idx,
        "_account_summary_path": account_summary_path,
    }


# ── Write Excel ────────────────────────────────────────────────────────────────
def write_excel(spend_pivot, actual_months, future_months, fid_pivot,
                acc_fund_map, holdings, summary_data, acc_holdings, output_path,
                reimbursements=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "Wealth Summary"

    # Single comprehensive sheet — no separate tabs

    def fml(cell, formula):
        """Write an Excel formula to a cell, ensuring it is stored as a formula not text."""
        cell.data_type = "f"
        cell.value = formula
        cell.number_format = '#,##0;(#,##0);"-"'

    all_months = actual_months + future_months
    spend_months = all_months
    fid_months   = all_months

    spend_month_labels = [m.strftime("%b %Y") for m in all_months]
    fid_month_labels   = [m.strftime("%b %Y") for m in all_months]
    actual_set = set(actual_months)  # months with real CSV data (Jan–Apr)

    # ── History boundary ───────────────────────────────────────────────────────
    # Previous Data owns everything BEFORE this month.
    # Live CSV files own this month and everything after.
    # This boundary never moves when new files are uploaded.
    HISTORY_CUTOFF = pd.Period("2026-07", "M")   # AccountSummary export date: 29 May 2026

    def is_history_month(m):
        """True if this month belongs to Previous Data (read-only)."""
        return m < HISTORY_CUTOFF

    def live_and_hist_safe(live_dict, hist_key_str):
        """Merge values with strict boundary:
        - Months BEFORE HISTORY_CUTOFF: Previous Data only (never overwritten by live)
        - Months FROM HISTORY_CUTOFF: live CSV files only
        """
        merged = {}
        hist_series = history.get(HIST_MAP.get(hist_key_str, ""), {})
        for m in all_months:
            if is_history_month(m):
                # History owns this month — ignore live value
                if m in hist_series:
                    merged[m] = hist_series[m]
                # If no history value exists, leave blank (don't backfill from live)
            else:
                # Live owns this month
                v = live_dict.get(m)
                if v:
                    merged[m] = v
        return merged

    def projection_with_hist_override(proj_dict, hist_key_str):
        """For formula-projected series (Expleo SW, Susan pensions, Cars):
        - Use hardcoded history value if it exists for a month (actuals take priority)
        - Otherwise use the live projection (covers all months incl. June+)
        """
        merged = {}
        hist_series = history.get(HIST_MAP.get(hist_key_str, ""), {})
        for m in all_months:
            if m in hist_series:
                merged[m] = hist_series[m]
            elif m in proj_dict:
                merged[m] = proj_dict[m]
        return merged


    C_NAVY   = "1F4E79"
    C_TEAL   = "1A5276"
    C_ALT    = "F2F7FB"
    C_SAL    = "E2EFDA"
    C_TOT    = "D9E1F2"
    C_FOOT   = "BDD7EE"
    C_FID_H  = "145A32"
    C_FID_A  = "EAF5EA"
    C_FID_T  = "A9DFBF"
    C_EST_H  = "4A4A6A"   # muted navy for estimated month headers
    C_EST_V  = "F7F7FC"   # very light lavender for estimated cells
    C_EST_A  = "EDEDF5"   # alt row estimated

    P = lambda c: PatternFill("solid", start_color=c, end_color=c)
    HDR_FILL  = P(C_NAVY)
    EST_HDR_FILL = P(C_EST_H)
    ALT_FILL  = P(C_ALT)
    EST_FILL  = P(C_EST_V)
    EST_ALT   = P(C_EST_A)
    SAL_FILL  = P(C_SAL)
    TOT_FILL  = P(C_TOT)
    FOOT_FILL = P(C_FOOT)
    FID_FILL  = P(C_FID_H)
    FIDA_FILL = P(C_FID_A)
    FIDT_FILL = P(C_FID_T)

    def F(bold=False, color="000000", size=10, name="Arial", italic=False):
        return Font(bold=bold, color=color, name=name, size=size, italic=italic)

    HDR_FONT  = F(bold=True,  color="FFFFFF")
    EST_FONT  = F(bold=True,  color="CCCCDD")   # dimmed for estimated headers
    BODY_FONT = F()
    EST_BODY  = F(color="555577")               # muted text for estimated values
    BOLD_FONT = F(bold=True)
    SAL_FONT  = F(bold=True,  color="375623")
    FID_HFONT = F(bold=True,  color="FFFFFF")
    FID_TFONT = F(bold=True,  color="145A32")

    def col_fill(m, is_alt, base_fill, base_alt):
        """Return appropriate fill for actual vs estimated column."""
        if m in actual_set:
            return base_alt if is_alt else None
        else:
            return EST_ALT if is_alt else EST_FILL

    def val_font(m, bold=False, sal=False):
        if sal: return SAL_FONT
        if m not in actual_set: return EST_BODY
        return BOLD_FONT if bold else BODY_FONT

    # ── Summary Table ──────────────────────────────────────────────────────────
    C_SUM_H  = "2C3E50"   # dark slate header
    C_SUM_A  = "F4F6F7"   # light alt row
    C_SUM_T  = "D5D8DC"   # total / subtotal
    C_SUM_S  = "1A5276"   # section subheader (dark blue)
    C_SUM_SA = "EAF2FF"   # section alt

    P_SUM_H  = P(C_SUM_H)
    P_SUM_A  = P(C_SUM_A)
    P_SUM_T  = P(C_SUM_T)
    P_SUM_S  = P(C_SUM_S)
    P_SUM_SA = P(C_SUM_SA)

    n_sum_cols = 3 + len(all_months)  # Label | blank | blank | month cols (live)
    # All tables: col1=Label, col2=Units/blank, col3=blank, col4+=months

    HIST_MAP = {
        "fidelity_all":      "All Fidelity accounts",
        "arriva":            "PS Arriva pension (Def Bene)",
        "expleo":            "Paul Explo Scottish Widows",
        "susan_fidelity":    "Susan Fidelity Pension",
        "capita":            "Capita (RMSPS) - to 2018 (65 yrs)",
        "rmpp":              "RMPP 2012 - 2023 (65 years)",
        "collective":        "Collective Pension (2023 +)",
        "cash_balance":      "Cash Balance",
        "avc":               "AVC Bonus Plan (Scottish widows)",
        "mf":                "MF",
        "house":             "House",
        "liam":              "Liam ISA",
        "jayne":             "Jaynes ISA",
        "cars":              "Cars Paul",
        "paul_pension":      "Paul Pension",
        "susan_pension":     "Susan Pension",
    }

    # ── Historic months (from Previous_Data CSV — used for calcs, NOT displayed) ─
    history = load_history()  # hardcoded — no file needed
    hist_periods_set = set()
    for series in history.values():
        hist_periods_set.update(series.keys())
    hist_months = sorted(p for p in hist_periods_set if p not in set(all_months))

    # Display only 2026 months — all tables share same columns
    sum_months = list(all_months)   # Jan–Dec 2026 only
    n_sum_cols = 3 + len(sum_months)

    # Helper: look up a value for a given month — history first for hist months,
    # live dict for live months. live_dict keys are pd.Period.
    def sum_val(live_dict, m):
        """Return value for month m: from live_dict if available, else history."""
        if m in live_dict and live_dict[m]:
            return live_dict[m]
        return None

    def write_total_row(r, label, monthly_dict, fill_col, font_col="042C53"):
        c = ws.cell(row=r, column=1, value=label)
        c.font = F(bold=True, color=font_col); c.fill = P(fill_col)
        c.alignment = Alignment(horizontal="left", vertical="center")
        ws.cell(row=r, column=2).fill = P(fill_col)
        ws.cell(row=r, column=3).fill = P(fill_col)
        for col_idx, m in enumerate(sum_months, 4):
            v = monthly_dict.get(m, 0)
            cell = ws.cell(row=r, column=col_idx, value=int(round(v)) if v else None)
            cell.number_format = '#,##0'; cell.font = F(bold=True, color=font_col)
            cell.fill = P(fill_col); cell.alignment = Alignment(horizontal="right", vertical="center")
        return r + 1

    def sum_row(r, label, values_dict, alt=False, bold=False, header_fill=None,
                label_indent="", font_color="000000", total_fill=None, hist_key=None):
        """Write one row: col1=Label, col2=blank, col3=blank, col4+=months.
        hist_key: if provided, uses history[hist_key] for historic months."""
        fill = header_fill or (P_SUM_A if alt else None)
        font = F(bold=bold, color=font_color)
        c = ws.cell(row=r, column=1, value=label_indent + label)
        c.font = font
        c.alignment = Alignment(horizontal="left", vertical="center")
        c.border = Border(bottom=Side(style="thin", color="D5D8DC"))
        if fill: c.fill = fill

        for col in (2, 3):
            bc = ws.cell(row=r, column=col)
            bc.border = Border(bottom=Side(style="thin", color="D5D8DC"))
            if fill: bc.fill = fill

        hist_row_name = HIST_MAP.get(hist_key, hist_key) if hist_key else None
        hist_series = history.get(hist_row_name, {}) if hist_row_name else {}

        for col_idx, m in enumerate(sum_months, 4):
            is_hist = is_history_month(m)
            if is_hist and hist_key and m in hist_series:
                # Previous Data owns this month and has a value — use it
                val = hist_series[m]
            elif is_hist and hist_key and m not in hist_series:
                # Previous Data owns this month but no entry — blank
                val = None
            else:
                # Either not a history month, or no hist_key — use values_dict
                val = values_dict.get(m, 0) if values_dict else 0
            is_est = (not is_hist) and (m not in actual_set)
            # Write value — allow zero (don't filter with truthiness)
            if val is not None:
                cell = ws.cell(row=r, column=col_idx, value=int(round(val)))
                cell.number_format = '#,##0'
            else:
                cell = ws.cell(row=r, column=col_idx, value=None)
                cell.number_format = '#,##0'
            if is_hist:
                cell.font = F(bold=bold, color="7F8C8D")
                cell.fill = P("E8EAEB") if alt else P("F2F3F4")
            elif is_est:
                cell.font = F(bold=bold, color="555577")
                if fill: cell.fill = fill
            else:
                cell.font = font
                if fill: cell.fill = fill
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.border = Border(bottom=Side(style="thin", color="D5D8DC"))

        return r + 1

    def sum_total_row(r, label, monthly_dict, fill_col, font_col="042C53"):
        """Write a total/subtotal row across all sum_months."""
        font = F(bold=True, color=font_col)
        p_fill = P(fill_col)
        c = ws.cell(row=r, column=1, value=label)
        c.font = font; c.fill = p_fill
        c.alignment = Alignment(horizontal="left", vertical="center")
        ws.cell(row=r, column=2).fill = p_fill
        ws.cell(row=r, column=3).fill = p_fill
        for col_idx, m in enumerate(sum_months, 4):
            v = monthly_dict.get(m, 0)
            cell = ws.cell(row=r, column=col_idx, value=int(round(v)) if v else None)
            cell.number_format = '#,##0'; cell.font = font; cell.fill = p_fill
            cell.alignment = Alignment(horizontal="right", vertical="center")
        return r + 1

    MAY_2026 = pd.Period("2026-05", "M")

    # Section title
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_sum_cols)
    t = ws.cell(row=1, column=1, value="Family Wealth Summary")
    t.font = F(bold=True, color="FFFFFF", size=12)
    t.fill = P_SUM_H
    t.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 24

    # Header row: Label | blank | blank | hist months (grey) | live months
    ws.cell(row=2, column=1, value="").fill = P_SUM_H
    ws.cell(row=2, column=2, value="").fill = P_SUM_H
    ws.cell(row=2, column=3, value="").fill = P_SUM_H
    for col_idx, m in enumerate(sum_months, 4):
        is_hist = m in hist_months
        is_est = (not is_hist) and (m not in actual_set)
        label = m.strftime("%b %Y")
        c = ws.cell(row=2, column=col_idx, value=label)
        if is_hist:
            c.font = F(bold=True, color="AAAAAA")
            c.fill = P("BFC9CA")
        elif is_est:
            c.font = F(bold=True, color="CCCCDD")
            c.fill = P("4A4A6A")
        else:
            c.font = F(bold=True, color="FFFFFF")
            c.fill = P_SUM_H
        c.alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[2].height = 18

    cur_row = 3

    # ── 1. Fidelity accounts ──────────────────────────────────────────────────
    # Sub-header
    ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=n_sum_cols)
    sh = ws.cell(row=cur_row, column=1, value="Fidelity accounts")
    sh.font = F(bold=True, color="FFFFFF", size=9)
    sh.fill = P_SUM_S
    sh.alignment = Alignment(horizontal="left", vertical="center")
    cur_row += 1

    # Mapping: summary row → history row name
    def hist(key):
        """Return history series for this key."""
        return history.get(HIST_MAP.get(key, ""), {})

    def live_and_hist(live_dict, hist_key_str):
        """Alias for live_and_hist_safe — history always wins for pre-cutoff months."""
        return live_and_hist_safe(live_dict, hist_key_str)

    # Each family account — exclude SIPP accounts (shown in pensions instead)
    fid_total_val = summary_data["fidelity_total"]
    fid_by_acc = summary_data["fidelity_by_acc"]
    SIPP_ACCS = {"2000001606", "2000001604"}  # moved to pensions sections
    FIDELITY_ACC_LABELS = {
        "AW10032966": "Cash Account (Paul)",
        "SANX002282": "Investment ISA (Paul)",
        "SANQ000468": "Investment Account (Joint)",   # removed "(Paul)"
        "SANX002936": "Junior ISA (Jayne)",
        "AW10261123": "Cash Account (Susan)",
        "SANX002617": "Investment ISA (Susan)",
        "AW10580794": "Cash Account (Liam)",
        "AS10303823": "Investment ISA (Liam)",
        "AG10131710": "Investment Account (Liam)",
    }

    fid_non_sipp_accs = [(acc, val) for acc, val in sorted(fid_by_acc.items(), key=lambda x: x[1], reverse=True)
                         if acc not in SIPP_ACCS]

    # Compute monthly growth for SANQ000468 from accumulated fund values
    sanq_monthly = {}
    for fund_h, fd_h in acc_holdings.get("SANQ000468", {}).items():
        mv = fd_h.get("monthly_values", {})
        for m, v in mv.items():
            sanq_monthly[m] = sanq_monthly.get(m, 0) + v
    # Add LGEN and any non-Acc holdings (flat) to get full account value
    sanq_non_acc_val = fid_by_acc.get("SANQ000468", 0) - sum(
        fd_h["value"] for fd_h in acc_holdings.get("SANQ000468", {}).values()
    )
    for m in all_months:
        sanq_monthly[m] = sanq_monthly.get(m, 0) + sanq_non_acc_val

    # Override with ACTUAL AccountSummary value for the current data month (June 2026)
    # The growth-projection model overshoots; June actual = £791,867.11
    JUNE_2026 = pd.Period("2026-06", "M")
    if JUNE_2026 in all_months:
        sanq_monthly[JUNE_2026] = fid_by_acc.get("SANQ000468", sanq_monthly.get(JUNE_2026, 0))

    fid_rows_start = cur_row  # track start row for Fidelity total formula

    for i, (acc, val) in enumerate(fid_non_sipp_accs):
        label = FIDELITY_ACC_LABELS.get(acc, acc)
        if acc == "SANQ000468" and sanq_monthly:
            monthly_vals = {m: sanq_monthly.get(m, val) for m in all_months}
        else:
            monthly_vals = {m: val for m in all_months}
        # Fidelity account rows: no per-account history, live values only
        cur_row = sum_row(cur_row, label, monthly_vals, alt=(i % 2 == 0),
                          label_indent="  ", font_color="042C53")
    fid_rows_end = cur_row

    # ── 2. Paul's pensions ────────────────────────────────────────────────────
    ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=n_sum_cols)
    sh = ws.cell(row=cur_row, column=1, value="Paul's pensions")
    sh.font = F(bold=True, color="FFFFFF", size=9); sh.fill = P_SUM_S
    sh.alignment = Alignment(horizontal="left", vertical="center")
    cur_row += 1

    paul_sipp_val = fid_by_acc.get("2000001606", 0)
    paul_sipp_vals_flat = {m: paul_sipp_val for m in all_months}
    cur_row = sum_row(cur_row, "SIPP Savings – Fidelity (2000001606)", paul_sipp_vals_flat,
                      alt=True, label_indent="  ", font_color="1A5276")

    arriva_merged = projection_with_hist_override(summary_data["arriva"], "arriva")
    cur_row = sum_row(cur_row, "PS Arriva (Defined Benefit)", arriva_merged,
                      alt=False, label_indent="  ", font_color="1A5276")  # pre-merged

    expleo_merged = projection_with_hist_override(summary_data["expleo_sw"], "expleo")
    cur_row = sum_row(cur_row, "Expleo Scottish Widows", expleo_merged,  # pre-merged
                      alt=True, label_indent="  ", font_color="1A5276")

    # Paul total placeholder — updated below once paul_sipp_growth_vals is computed
    paul_total_placeholder = {m: paul_sipp_val + summary_data["arriva"].get(m, 0) + summary_data["expleo_sw"].get(m, 0)
                              for m in all_months}
    cur_row = write_total_row(cur_row, "Total Paul's pensions", paul_total_placeholder, C_SUM_T)

    # ── 3. Susan's pensions ───────────────────────────────────────────────────
    ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=n_sum_cols)
    sh = ws.cell(row=cur_row, column=1, value="Susan's pensions")
    sh.font = F(bold=True, color="FFFFFF", size=9); sh.fill = P_SUM_S
    sh.alignment = Alignment(horizontal="left", vertical="center")
    cur_row += 1

    SUSAN_HIST_KEYS = {
        "Capita (RMSPS) – to 2018 (65 yrs)": "capita",
        "RMPP 2012–2023 (65 years)":          "rmpp",
        "Collective Pension (2023+)":         "collective",
        "Cash Balance":                        "cash_balance",
        "AVC Bonus Plan (Scottish Widows)":    "avc",
    }
    susan_total_by_month = {m: 0 for m in all_months}
    for i, (pen_label, pen_vals) in enumerate(summary_data["susan_pensions"].items()):
        hk = SUSAN_HIST_KEYS.get(pen_label)
        merged = projection_with_hist_override(pen_vals, hk) if hk else pen_vals
        cur_row = sum_row(cur_row, pen_label, merged, alt=(i % 2 == 0),
                          label_indent="  ", font_color="1A5276")  # hist_key omitted — pre-merged
        # Accumulate from merged (history-corrected) not raw pen_vals
        for m in all_months:
            susan_total_by_month[m] += merged.get(m, 0)

    sipp_vals = summary_data["susan_fidelity_sipp"]
    for m in all_months:
        if m not in sipp_vals:
            sipp_vals[m] = list(sipp_vals.values())[0] if sipp_vals else 0
    susan_fid_merged = live_and_hist(sipp_vals, "susan_fidelity")
    cur_row = sum_row(cur_row, "SIPP Savings – Fidelity (2000001604)", susan_fid_merged,
                      alt=True, label_indent="  ", font_color="1A5276",
                      hist_key="susan_fidelity")
    for m in all_months:
        susan_total_by_month[m] = susan_total_by_month.get(m, 0) + sipp_vals.get(m, 0)

    cur_row = write_total_row(cur_row, "Total Susan's pensions", susan_total_by_month, C_SUM_T)
    susan_excl_sipp = {m: susan_total_by_month.get(m, 0) - sipp_vals.get(m, 0) for m in all_months}

    # ── 4. Assets section ──────────────────────────────────────────────────────
    ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=n_sum_cols)
    sh = ws.cell(row=cur_row, column=1, value="Other Assets")
    sh.font = F(bold=True, color="FFFFFF", size=9); sh.fill = P_SUM_S
    sh.alignment = Alignment(horizontal="left", vertical="center")
    cur_row += 1

    mf_proj = {m: 190000 for m in all_months}
    mf_merged = projection_with_hist_override(mf_proj, "mf")
    mf_vals = {m: mf_merged.get(m, 190000) for m in all_months}
    cur_row = sum_row(cur_row, "MF", mf_merged, alt=True, label_indent="  ",
                      font_color="1A5276")  # pre-merged

    bar_data = summary_data.get("barclays_by_month", {})
    cur_row = sum_row(cur_row, "Cash (Barclays)", bar_data, alt=False, label_indent="  ",
                      font_color="1A5276")

    house_may = 450450
    house_vals = {}
    may_idx_h = next((i for i, m in enumerate(all_months) if m == MAY_2026), 0)
    for i, m in enumerate(all_months):
        offset = i - may_idx_h
        house_vals[m] = round(house_may * ((1.05) ** (offset / 12)))
    house_merged = projection_with_hist_override(house_vals, "house")
    cur_row = sum_row(cur_row, "House", house_merged, alt=True, label_indent="  ",
                      font_color="1A5276")  # pre-merged

    # Liam ISA (AS10303823) and Investment Account (AG10131710) already in Fidelity accounts — not duplicated here

    car_may = 42074  # actual May 2026 value from history (Cars Paul)
    car_vals = {}
    for i, m in enumerate(all_months):
        offset = i - may_idx_h
        car_vals[m] = round(car_may * ((1 - 0.05) ** (offset / 12)))
    car_merged = projection_with_hist_override(car_vals, "cars")
    cur_row = sum_row(cur_row, "Cars Paul (Mercedes AMG GTS 2016)", car_merged, alt=False,
                      label_indent="  ", font_color="1A5276")  # pre-merged

    assets_dicts = [mf_vals, bar_data, house_vals, car_vals]
    assets_by_month = {m: sum(d.get(m, 0) for d in assets_dicts) for m in all_months}
    cur_row = write_total_row(cur_row, "Total Other Assets", assets_by_month, C_SUM_T)

    # ── Compute Fidelity growth ───────────────────────────────────────────────
    fid_growth = {all_months[0]: round(fid_total_val)}
    for i in range(1, len(all_months)):
        m = all_months[i]
        monthly_income = fid_pivot[m].sum() if m in fid_pivot.columns else 0
        fid_growth[m] = round(fid_growth[all_months[i-1]] + monthly_income)

    # ── SIPP growth: history Jan–Apr, AccountSummary anchors May, income grows Jun+ ─
    # AccountSummary is the authoritative source for the export date (May 26 2026)
    paul_sipp_val_may = fid_by_acc.get("2000001606", 0)  # AccountSummary May value
    susan_sipp_val    = fid_by_acc.get("2000001604", 0)  # AccountSummary May value
    paul_hist_series  = hist("paul_pension")
    paul_sipp_growth_vals = {}
    susan_sipp_growth_vals = {}
    sipp_vals_from_summary = summary_data.get("susan_fidelity_sipp", {})

    ps_running = paul_hist_series.get(all_months[0], paul_sipp_val_may)
    ss_running = sipp_vals_from_summary.get(all_months[0], susan_sipp_val)

    for i, m in enumerate(all_months):
        if m == MAY_2026:
            # AccountSummary is ground truth for this month — override any projection
            ps_running = paul_sipp_val_may
            ss_running = susan_sipp_val
            paul_sipp_growth_vals[m] = paul_sipp_val_may
            susan_sipp_growth_vals[m] = susan_sipp_val
        elif is_history_month(m):
            # Pre-cutoff: use history if available
            if m in paul_hist_series:
                paul_sipp_growth_vals[m] = paul_hist_series[m]
                ps_running = paul_hist_series[m]
            else:
                paul_sipp_growth_vals[m] = ps_running
            ss_val = sipp_vals_from_summary.get(m, ss_running)
            susan_sipp_growth_vals[m] = ss_val
            ss_running = ss_val
        else:
            # Post-cutoff (Jun+): grow from previous month using income
            ps_inc = fid_pivot.loc["2000001606", m] if "2000001606" in fid_pivot.index and m in fid_pivot.columns else 0
            ss_inc = fid_pivot.loc["2000001604", m] if "2000001604" in fid_pivot.index and m in fid_pivot.columns else 0
            ps_running = round(ps_running + ps_inc)
            ss_running = round(ss_running + ss_inc)
            paul_sipp_growth_vals[m] = ps_running
            susan_sipp_growth_vals[m] = ss_running

    fid_non_sipp_total = sum(val for acc, val in fid_by_acc.items() if acc not in SIPP_ACCS)
    sanq_flat = fid_by_acc.get("SANQ000468", 0)
    sanq_growth = {m: sanq_monthly.get(m, sanq_flat) for m in all_months}
    fid_non_sipp_with_growth = {m: fid_non_sipp_total - sanq_flat + sanq_growth[m] for m in all_months}
    total_fidelity_growth = {m: fid_non_sipp_with_growth[m] + paul_sipp_growth_vals[m] + susan_sipp_growth_vals[m]
                             for m in all_months}

    # ── 5. Fidelity accounts total (all including SIPPs) ─────────────────────
    fid_total_row_num = cur_row
    fid_all_merged = {**hist("fidelity_all"), **total_fidelity_growth}
    cur_row = write_total_row(cur_row, "Fidelity accounts", fid_all_merged, C_SUM_T)

    # Update Paul SIPP row values with growth
    for _r in range(3, fid_total_row_num):
        if ws.cell(row=_r, column=1).value == "  SIPP Savings – Fidelity (2000001606)":
            for col_idx, m in enumerate(all_months, 4):
                ws.cell(row=_r, column=col_idx).value = int(round(paul_sipp_growth_vals[m]))
            break
    # Update Susan SIPP row values with growth
    for _r in range(3, fid_total_row_num):
        if ws.cell(row=_r, column=1).value == "  SIPP Savings – Fidelity (2000001604)":
            for col_idx, m in enumerate(all_months, 4):
                ws.cell(row=_r, column=col_idx).value = int(round(susan_sipp_growth_vals[m]))
            break
    # Update paul_total using history-corrected arriva and expleo values
    arriva_merged = projection_with_hist_override(summary_data["arriva"], "arriva")
    expleo_merged_upd = projection_with_hist_override(summary_data["expleo_sw"], "expleo")
    paul_total = {m: paul_sipp_growth_vals[m] + arriva_merged.get(m, 0) + expleo_merged_upd.get(m, 0)
                  for m in all_months}
    for _r in range(3, fid_total_row_num):
        if ws.cell(row=_r, column=1).value == "Total Paul's pensions":
            for col_idx, m in enumerate(all_months, 4):
                ws.cell(row=_r, column=col_idx).value = int(round(paul_total[m]))
            break

    # Update Susan total using history-corrected pension + SIPP growth values
    for _r in range(3, fid_total_row_num):
        if ws.cell(row=_r, column=1).value == "Total Susan's pensions":
            for col_idx, m in enumerate(all_months, 4):
                sp_merged = sum(
                    projection_with_hist_override(summary_data["susan_pensions"].get(k, {}),
                                  SUSAN_HIST_KEYS.get(k)).get(m, 0)
                    for k in summary_data["susan_pensions"]
                )
                ws.cell(row=_r, column=col_idx).value = int(round(sp_merged + susan_sipp_growth_vals[m]))
            break

    # ── 8. TOTAL FAMILY WEALTH — at the very bottom ───────────────────────────
    grand_row = cur_row
    cur_row += 1
    paul_non_sipp = {m: summary_data["arriva"][m] + summary_data["expleo_sw"][m] for m in all_months}
    grand_by_month = {m: total_fidelity_growth[m] + paul_non_sipp.get(m, 0)
                      + susan_excl_sipp.get(m, 0) + assets_by_month.get(m, 0)
                      for m in all_months}
    # Build historic grand total from history rows — covers ALL historic periods
    # (not just display months) so Yearly Increase can reference prior year values
    hist_grand = {}
    all_hist_periods = sorted(set().union(*[set(s.keys()) for s in history.values()]))
    for m in all_hist_periods:
        fid_h    = hist("fidelity_all").get(m, 0)
        arriva_h = hist("arriva").get(m, 0)
        expleo_h = hist("expleo").get(m, 0)
        capita_h = hist("capita").get(m, 0)
        rmpp_h   = hist("rmpp").get(m, 0)
        coll_h   = hist("collective").get(m, 0)
        cb_h     = hist("cash_balance").get(m, 0)
        avc_h    = hist("avc").get(m, 0)
        mf_h     = hist("mf").get(m, 0)
        house_h  = hist("house").get(m, 0)
        # Liam ISA & Investment Account is inside fid_h (All Fidelity accounts) — not added separately
        # Jayne ISA is also inside fid_h — not added separately
        susan_non_fid_h = capita_h + rmpp_h + coll_h + cb_h + avc_h
        hist_grand[m] = (fid_h + arriva_h + expleo_h + susan_non_fid_h +
                         mf_h + house_h)

    all_grand = {**hist_grand, **grand_by_month}

    c = ws.cell(row=grand_row, column=1, value="TOTAL FAMILY WEALTH")
    c.font = F(bold=True, color="FFFFFF", size=11); c.fill = P_SUM_H
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[grand_row].height = 22
    for col in range(1, 4):
        ws.cell(row=grand_row, column=col).fill = P_SUM_H
    for col_idx, m in enumerate(sum_months, 4):
        v = all_grand.get(m, 0)
        cell = ws.cell(row=grand_row, column=col_idx, value=int(round(v)) if v else None)
        cell.number_format = '#,##0'; cell.font = F(bold=True, color="FFFFFF")
        cell.fill = P_SUM_H; cell.alignment = Alignment(horizontal="right", vertical="center")

    # ── 9. Total Income and Accumulations — below TOTAL FAMILY WEALTH ───────────────────────
    # Values written now as placeholders; updated with Excel formulas once
    # Income (tot_r) and Accumulative (tot_acc_r) row numbers are known.
    fid_income_monthly = {m: int(round(fid_pivot[m].sum())) if (not fid_pivot.empty and m in fid_pivot.columns) else 0
                          for m in all_months}
    acc_monthly_increase = {m: 0 for m in all_months}
    for acc_hh, funds_hh in acc_holdings.items():
        for fund_hh, fd_hh in funds_hh.items():
            pa = fd_hh.get("price_appreciation", {})
            for mk, inc in pa.items():
                if mk in all_months:
                    acc_monthly_increase[mk] = acc_monthly_increase.get(mk, 0) + inc
    total_inc_acc = {m: fid_income_monthly.get(m, 0) + acc_monthly_increase.get(m, 0)
                     for m in all_months}
    total_inc_acc_row = cur_row  # remember for later formula update
    cur_row = write_total_row(cur_row, "Total Income and Accumulations", total_inc_acc, C_SUM_T)

    # ── Calculations section (placed AFTER all assets rows) ───────────────────
    calc_start = cur_row + 1  # one blank row after Total Income and Accumulations

    C_CALC_H = "17202A"
    P_CALC_H = P(C_CALC_H)
    P_CALC_A = P("EAECF0")

    ws.merge_cells(start_row=calc_start, start_column=1, end_row=calc_start, end_column=n_sum_cols)
    ch = ws.cell(row=calc_start, column=1, value="Summary Calculations")
    ch.font = F(bold=True, color="FFFFFF", size=12); ch.fill = P_CALC_H
    ch.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[calc_start].height = 24

    calc_hdr = calc_start + 1
    for col_idx, m in enumerate(sum_months, 4):
        is_hist = m in hist_months
        c = ws.cell(row=calc_hdr, column=col_idx,
                    value=m.strftime("%b %Y") if not (col_idx <= 3) else None)
        c.font = F(bold=True, color="AAAAAA" if is_hist else "FFFFFF")
        c.fill = P("BFC9CA") if is_hist else P_CALC_H
        c.alignment = Alignment(horizontal="right", vertical="center")
    for col in (1, 2, 3):
        ws.cell(row=calc_hdr, column=col).fill = P_CALC_H
    ws.row_dimensions[calc_hdr].height = 16

    def mcol(m): return get_column_letter(4 + sum_months.index(m)) if m in sum_months else None

    def calc_row_fn(r, label, formulas_by_month, alt=False, bold=False, font_color="000000"):
        fill = P_CALC_A if alt else None
        font = F(bold=bold, color=font_color)
        c = ws.cell(row=r, column=1, value=label)
        c.font = font; c.alignment = Alignment(horizontal="left", vertical="center")
        c.border = Border(bottom=Side(style="thin", color="CCCCCC"))
        if fill: c.fill = fill
        for col in (2, 3):
            bc = ws.cell(row=r, column=col)
            bc.border = Border(bottom=Side(style="thin", color="CCCCCC"))
            if fill: bc.fill = fill
        for col_idx, m in enumerate(sum_months, 4):
            formula = formulas_by_month.get(m)
            cell = ws.cell(row=r, column=col_idx)
            if formula is None:
                cell.value = None
            elif isinstance(formula, str) and formula.startswith("="):
                # Write as proper Excel formula — never as text
                cell.data_type = "f"
                cell.value = formula
                cell.number_format = '#,##0;(#,##0);"-"'
            elif isinstance(formula, (int, float)):
                cell.value = int(round(formula))
                cell.number_format = '#,##0;(#,##0);"-"'
            else:
                cell.value = formula
            cell.font = font
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.border = Border(bottom=Side(style="thin", color="CCCCCC"))
            if fill: cell.fill = fill
        return r + 1

    # Scan summary rows to find key row numbers
    row_refs = {}
    for _r in range(3, calc_start):
        v = ws.cell(row=_r, column=1).value
        if v: row_refs[str(v).strip()] = _r

    gw_row      = row_refs.get("TOTAL FAMILY WEALTH", grand_row)
    fid_row_sum = row_refs.get("Total Fidelity")
    paul_row_s  = row_refs.get("Total Paul's pensions")
    susan_row_s = row_refs.get("Total Susan's pensions")
    assets_row  = row_refs.get("Total Other Assets")
    house_row   = next((r for k, r in row_refs.items() if k.startswith("House")), None)
    car_row     = next((r for k, r in row_refs.items() if "Cars" in k), None)
    arriva_row  = next((r for k, r in row_refs.items() if "Arriva" in k), None)
    expleo_row  = next((r for k, r in row_refs.items() if "Expleo" in k), None)

    calc_cur = calc_hdr + 1

    # 1. Total — all assets incl cash and property (= TOTAL FAMILY WEALTH)
    calc_cur = calc_row_fn(calc_cur, "Total",
        {m: f"={mcol(m)}{gw_row}" for m in all_months},
        alt=True, bold=True, font_color="17202A")

    # 2. Investments & Cash — Total minus House and Car
    def inv_formula(m):
        parts = [f"{mcol(m)}{gw_row}"]
        if house_row: parts.append(f"-{mcol(m)}{house_row}")
        if car_row:   parts.append(f"-{mcol(m)}{car_row}")
        return "=" + "".join(parts)
    calc_cur = calc_row_fn(calc_cur, "Investments & Cash",
        {m: inv_formula(m) for m in sum_months}, alt=False)

    # 3. Yearly Increase — this month's TOTAL FAMILY WEALTH minus same month last year
    # Prior year values come from all_grand dict (which includes historic data)
    yearly = {}
    for i, m in enumerate(sum_months):
        m_minus_12 = m - 12
        prior_val = all_grand.get(m_minus_12, 0)
        if prior_val:
            # Use direct cell reference for current month, subtract hardcoded prior value
            yearly[m] = f"={mcol(m)}{gw_row}-{int(prior_val)}"
    calc_cur = calc_row_fn(calc_cur, "Yearly Increase", yearly, alt=True)

    # 4. PS Pension — Paul: Fidelity SIPP + Arriva + Expleo SW
    paul_fid_r = next((r for k, r in row_refs.items() if "SIPP Savings (Paul)" in k or "2000001606" in k), None)
    def ps_formula(m):
        parts = []
        if paul_fid_r:  parts.append(f"{mcol(m)}{paul_fid_r}")
        if arriva_row:  parts.append(f"{mcol(m)}{arriva_row}")
        if expleo_row:  parts.append(f"{mcol(m)}{expleo_row}")
        return ("=" + "+".join(parts)) if parts else None
    calc_cur = calc_row_fn(calc_cur, "PS Pension (Paul)", {m: ps_formula(m) for m in sum_months}, alt=False)

    # 5. SS Pension — Susan: total pensions row (incl SIPP)
    calc_cur = calc_row_fn(calc_cur, "SS Pension (Susan)",
        {m: f"={mcol(m)}{susan_row_s}" for m in sum_months} if susan_row_s else {},
        alt=True)

    # 6. Monthly Change — vs previous month
    mc_formulas = {}
    for i, m in enumerate(sum_months):
        if i > 0:
            mc_formulas[m] = f"={mcol(m)}{gw_row}-{mcol(sum_months[i-1])}{gw_row}"
    calc_cur = calc_row_fn(calc_cur, "Monthly Change", mc_formulas, alt=False)

    # 7. Av Monthly Change — average monthly change since start of all data
    av_formulas = {}
    for i, m in enumerate(sum_months):
        if i >= 2:
            av_formulas[m] = f"=ROUND(({mcol(m)}{gw_row}-{mcol(sum_months[0])}{gw_row})/{i},0)"
    calc_cur = calc_row_fn(calc_cur, "Av Monthly Change", av_formulas, alt=True)

    # 8. Monthly Investment Increase — income generated from investments (fid_pivot)
    monthly_inv = {}
    for m in sum_months:
        inc = round(fid_pivot[m].sum()) if m in fid_pivot.columns else 0
        if inc: monthly_inv[m] = inc
    calc_cur = calc_row_fn(calc_cur, "Monthly Investment Increase", monthly_inv,
                           alt=False, font_color="145A32")

    # ── Section 1: Spending Summary ────────────────────────────────────────────
    n_spend_cols = 3 + len(spend_months)
    R = calc_cur + 0  # no blank spacer

    # Section title
    ws.merge_cells(start_row=R+1, start_column=1, end_row=R+1, end_column=n_spend_cols)
    title = ws.cell(row=R+1, column=1, value="Amex & Barclays Spend")
    title.font = F(bold=True, color="FFFFFF", size=12)
    title.fill = P(C_NAVY)
    title.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[R+1].height = 24

    # Column headers — col1=Category, col2=blank, col3=Total, col4+=months (aligned with income table)
    headers = ["Category", "", "Total"] + spend_month_labels
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=R+2, column=col, value=h)
        if col <= 3:
            c.font = HDR_FONT
            c.fill = HDR_FILL
        else:
            m = spend_months[col - 4]
            c.font = HDR_FONT if m in actual_set else EST_FONT
            c.fill = HDR_FILL if m in actual_set else EST_HDR_FILL
        c.alignment = Alignment(horizontal="left" if col == 1 else "right", vertical="center")
    ws.row_dimensions[R+2].height = 18

    # Data rows
    for row_idx, cat in enumerate(CATEGORIES, R+3):
        is_alt = (row_idx % 2 == 0)
        font  = BOLD_FONT

        c = ws.cell(row=row_idx, column=1, value=cat)
        c.font = font
        c.alignment = Alignment(horizontal="left", vertical="center")
        c.border = Border(bottom=Side(style="thin", color="D9E1F2"))

        # col 2 blank
        ws.cell(row=row_idx, column=2).border = Border(bottom=Side(style="thin", color="D9E1F2"))

        fc = get_column_letter(4)
        lc = get_column_letter(3 + len(spend_months))
        tc = ws.cell(row=row_idx, column=3,
                     value=f"=SUM({fc}{row_idx}:{lc}{row_idx})")
        tc.number_format = NUM_FMT
        tc.font = font
        tc.alignment = Alignment(horizontal="right", vertical="center")
        tc.fill = TOT_FILL
        tc.border = Border(bottom=Side(style="thin", color="D9E1F2"))

        for col_idx, m in enumerate(spend_months, 4):
            val = spend_pivot.loc[cat, m] if cat in spend_pivot.index else 0
            fill = col_fill(m, is_alt, ALT_FILL, ALT_FILL)
            cell = ws.cell(row=row_idx, column=col_idx,
                           value=int(round(val)) or None)
            cell.number_format = NUM_FMT
            cell.font = val_font(m)
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.border = Border(bottom=Side(style="thin", color="D9E1F2"))
            if fill: cell.fill = fill

    # Footer: total spend
    foot_row = R + 3 + len(CATEGORIES)
    exp_rows = [R + 3 + i for i in range(len(CATEGORIES))]
    ws.cell(row=foot_row, column=1, value="TOTAL").font = BOLD_FONT
    ws.cell(row=foot_row, column=1).fill = FOOT_FILL
    ws.cell(row=foot_row, column=1).alignment = Alignment(horizontal="left", vertical="center")
    ws.cell(row=foot_row, column=2).fill = FOOT_FILL  # blank col
    for col_idx in range(3, 4 + len(spend_months)):
        cl = get_column_letter(col_idx)
        refs = "+".join(f"{cl}{r}" for r in exp_rows)
        cell = ws.cell(row=foot_row, column=col_idx, value=f"={refs}")
        cell.number_format = NUM_FMT
        cell.font = BOLD_FONT
        cell.fill = FOOT_FILL
        cell.alignment = Alignment(horizontal="right", vertical="center")


    # ── Note row below TOTAL ────────────────────────────────────────────────────
    note_row = foot_row + 1
    note_cell = ws.cell(row=note_row, column=1,
                        value="Susan's lease car additional miles at 30.3p.  Contract to End July 2027")
    note_cell.font = Font(italic=True, color="808080", size=9)
    note_cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[note_row].height = 14
    # Merge across all spend columns so it reads as one line
    ws.merge_cells(start_row=note_row, start_column=1,
                   end_row=note_row, end_column=3 + len(all_months))

    # ── Expense Reimbursements table (underneath Amex & Barclays Spend) ────────
    REIMB_START_ROW = note_row + 2  # 1 blank row gap below the lease-car note
    fc_r = get_column_letter(4)
    lc_r = get_column_letter(3 + len(all_months))

    # Title row — spans full width like the spend table
    ws.merge_cells(start_row=REIMB_START_ROW, start_column=1,
                   end_row=REIMB_START_ROW, end_column=3 + len(all_months))
    rt = ws.cell(row=REIMB_START_ROW, column=1, value="Expense Reimbursements")
    rt.font = F(bold=True, color="FFFFFF", size=12)
    rt.fill = P("8E44AD")  # purple — distinct from navy spend table
    rt.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[REIMB_START_ROW].height = 22

    # Header row — Category | (blank) | Total | Jan..Dec, matching spend table
    header_row = REIMB_START_ROW + 1
    hc = ws.cell(row=header_row, column=1, value="Source")
    hc.font = F(bold=True, color="FFFFFF", size=10)
    hc.fill = P("8E44AD")
    hc.alignment = Alignment(horizontal="left", vertical="center")
    ws.cell(row=header_row, column=2).fill = P("8E44AD")
    htot = ws.cell(row=header_row, column=3, value="Total")
    htot.font = F(bold=True, color="FFFFFF", size=10)
    htot.fill = P("8E44AD")
    htot.alignment = Alignment(horizontal="right", vertical="center")
    for i, m in enumerate(all_months):
        col = 4 + i
        c = ws.cell(row=header_row, column=col, value=m.strftime("%b %Y"))
        c.font = F(bold=True, color="FFFFFF", size=10)
        c.fill = P("8E44AD")
        c.alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[header_row].height = 20

    # Build month -> amount map per source
    reimb_list = reimbursements or []
    by_source = {"Royal Mail": {}, "Expleo": {}}
    for r in reimb_list:
        memo_upper = r["memo"].upper()
        source = "Royal Mail" if "ROYAL MAIL" in memo_upper else "Expleo"
        m = r["month"]
        by_source[source][m] = by_source[source].get(m, 0) + r["amount"]

    # Two data rows: Royal Mail, Expleo
    for ri, source in enumerate(["Royal Mail", "Expleo"]):
        row_n = header_row + 1 + ri
        bg = "F4ECF7" if ri % 2 == 1 else "FFFFFF"
        c = ws.cell(row=row_n, column=1, value=source)
        c.font = F(size=10)
        c.fill = P(bg)
        c.alignment = Alignment(horizontal="left", vertical="center")
        c.border = Border(bottom=Side(style="thin", color="E8DAEF"))
        ws.cell(row=row_n, column=2).fill = P(bg)
        ws.cell(row=row_n, column=2).border = Border(bottom=Side(style="thin", color="E8DAEF"))

        # Month columns
        for i, m in enumerate(all_months):
            col = 4 + i
            val = by_source[source].get(m, 0)
            cell = ws.cell(row=row_n, column=col)
            if val:
                cell.value = val
                cell.number_format = NUM_FMT
            cell.font = F(size=10)
            cell.fill = P(bg)
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.border = Border(bottom=Side(style="thin", color="E8DAEF"))

        # Total column
        tc = ws.cell(row=row_n, column=3, value=f"=SUM({fc_r}{row_n}:{lc_r}{row_n})")
        tc.number_format = NUM_FMT
        tc.font = F(bold=True, size=10)
        tc.fill = P(bg)
        tc.alignment = Alignment(horizontal="right", vertical="center")
        tc.border = Border(bottom=Side(style="thin", color="E8DAEF"))

        ws.row_dimensions[row_n].height = 16

    # TOTAL row
    reimb_row = header_row + 3
    ws.cell(row=reimb_row, column=1, value="TOTAL").font = F(bold=True, size=10)
    ws.cell(row=reimb_row, column=1).fill = P("D2B4DE")
    ws.cell(row=reimb_row, column=2).fill = P("D2B4DE")
    tot_total = ws.cell(row=reimb_row, column=3,
                         value=f"=SUM(C{header_row+1}:C{header_row+2})")
    tot_total.number_format = NUM_FMT
    tot_total.font = F(bold=True, size=10)
    tot_total.fill = P("D2B4DE")
    tot_total.alignment = Alignment(horizontal="right", vertical="center")
    for i in range(len(all_months)):
        col = 4 + i
        col_l = get_column_letter(col)
        cm = ws.cell(row=reimb_row, column=col,
                      value=f"=SUM({col_l}{header_row+1}:{col_l}{header_row+2})")
        cm.number_format = NUM_FMT
        cm.font = F(bold=True, size=10)
        cm.fill = P("D2B4DE")
        cm.alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[reimb_row].height = 18

    # Note
    reimb_note_row = reimb_row + 1
    ws.merge_cells(start_row=reimb_note_row, start_column=1,
                   end_row=reimb_note_row, end_column=3 + len(all_months))
    rn = ws.cell(row=reimb_note_row, column=1,
                 value="Reimbursements from Royal Mail or Expleo for expenses incurred on the family's "
                       "behalf — recovered from employers, excluded from family spending totals.")
    rn.font = Font(italic=True, color="808080", size=8)
    rn.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[reimb_note_row].height = 28

    # ── Spacer rows ────────────────────────────────────────────────────────────
    fid_start_row = reimb_note_row + 2
    tot_r = fid_start_row  # fallback; overwritten when income section is written



    # ── Section 2: Income ─────────────────────────────────────────────────────
    if not fid_pivot.empty:
        # Build combined income rows: Salary first, then Fidelity accounts
        fid_accounts = list(fid_pivot.index)
        # Account + Units + Total + months
        n_fid_cols = 3 + len(fid_months)

        # Section title
        ws.merge_cells(start_row=fid_start_row, start_column=1,
                       end_row=fid_start_row, end_column=n_fid_cols)
        ft = ws.cell(row=fid_start_row, column=1, value="Income")
        ft.font = F(bold=True, color="FFFFFF", size=12)
        ft.fill = FID_FILL
        ft.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[fid_start_row].height = 24

        # Column headers
        fid_hdr_row = fid_start_row + 1
        # Income table: col1=Account, col2=Units, col3=Total, col4+=months
        fid_headers = ["Account", "Units", "Total"] + fid_month_labels
        for col, h in enumerate(fid_headers, 1):
            c = ws.cell(row=fid_hdr_row, column=col, value=h)
            if col <= 3:
                c.font = FID_HFONT
                c.fill = FID_FILL
            else:
                m = fid_months[col - 4]
                c.font = FID_HFONT if m in actual_set else EST_FONT
                c.fill = FID_FILL if m in actual_set else EST_HDR_FILL
            c.alignment = Alignment(horizontal="left" if col == 1 else "right", vertical="center")
        ws.row_dimensions[fid_hdr_row].height = 18

        # ── Salary row ────────────────────────────────────────────────────────
        sal_row = fid_start_row + 2
        sal_label = ws.cell(row=sal_row, column=1, value="Salary")
        sal_label.font = SAL_FONT
        sal_label.fill = SAL_FILL
        sal_label.alignment = Alignment(horizontal="left", vertical="center")
        sal_label.border = Border(bottom=Side(style="thin", color="A9DFBF"))

        # Units col (col 2) — blank for salary
        su = ws.cell(row=sal_row, column=2, value=None)
        su.fill = SAL_FILL
        su.border = Border(bottom=Side(style="thin", color="A9DFBF"))

        fc = get_column_letter(4)
        lc = get_column_letter(3 + len(fid_months))
        sal_tot = ws.cell(row=sal_row, column=3,
                          value=f"=SUM({fc}{sal_row}:{lc}{sal_row})")
        sal_tot.number_format = NUM_FMT
        sal_tot.font = SAL_FONT
        sal_tot.fill = SAL_FILL
        sal_tot.alignment = Alignment(horizontal="right", vertical="center")
        sal_tot.border = Border(bottom=Side(style="thin", color="A9DFBF"))

        for col_idx, m in enumerate(fid_months, 4):
            val = spend_pivot.loc["Salary", m] if "Salary" in spend_pivot.index else 0
            is_est = m not in actual_set
            cell = ws.cell(row=sal_row, column=col_idx,
                           value=int(round(val)) or None)
            cell.number_format = NUM_FMT
            cell.font = SAL_FONT if not is_est else F(bold=True, color="778866")
            cell.fill = SAL_FILL if not is_est else P("EEF5E8")
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.border = Border(bottom=Side(style="thin", color="A9DFBF"))

        # ── Per-person sections ───────────────────────────────────────────────
        fc = get_column_letter(4)
        lc = get_column_letter(3 + len(fid_months))
        current_row = sal_row + 1
        account_total_rows = []
        equity_data_rows = []

        for person, person_accs in acc_fund_map.items():
            # Person sub-header — spans all cols
            ph = ws.cell(row=current_row, column=1, value=person)
            ph.font = F(bold=True, color="FFFFFF", size=9)
            ph.fill = P("1E8449")
            ph.alignment = Alignment(horizontal="left", vertical="center")
            for col in range(2, 4 + len(fid_months)):
                ws.cell(row=current_row, column=col).fill = P("1E8449")
            current_row += 1

            for acc, fund_df in person_accs.items():
                label = ACCOUNT_LABELS.get(acc, acc)
                funds_in_acc = list(fund_df.index)
                n_fund_rows = len(funds_in_acc)

                acc_r = current_row
                account_total_rows.append(acc_r)

                # Col 1: account label
                c = ws.cell(row=acc_r, column=1, value=label)
                c.font = F(bold=True, color="145A32")
                c.fill = FIDT_FILL
                c.alignment = Alignment(horizontal="left", vertical="center")
                c.border = Border(bottom=Side(style="thin", color="A9DFBF"))

                # Col 2: units — blank for account header row
                uc = ws.cell(row=acc_r, column=2, value=None)
                uc.fill = FIDT_FILL
                uc.border = Border(bottom=Side(style="thin", color="A9DFBF"))

                # Col 3: total
                tc = ws.cell(row=acc_r, column=3,
                             value=f"=SUM({fc}{acc_r+1}:{lc}{acc_r+n_fund_rows})")
                tc.number_format = NUM_FMT
                tc.font = F(bold=True, color="145A32")
                tc.fill = FIDT_FILL
                tc.alignment = Alignment(horizontal="right", vertical="center")
                tc.border = Border(bottom=Side(style="thin", color="A9DFBF"))

                for col_idx, m in enumerate(fid_months, 4):
                    val = fid_pivot.loc[acc, m] if acc in fid_pivot.index else 0
                    is_est = m not in actual_set
                    cell = ws.cell(row=acc_r, column=col_idx,
                                   value=int(round(val)) or None)
                    cell.number_format = NUM_FMT
                    cell.font = F(bold=True, color="145A32") if not is_est else F(bold=True, color="558855")
                    cell.fill = FIDT_FILL if not is_est else P("C8E6C9")
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    cell.border = Border(bottom=Side(style="thin", color="A9DFBF"))

                current_row += 1

                # Fund rows — col1=indented name, col2=units, col3=total, col4+=months
                for fund_offset, fund in enumerate(funds_in_acc):
                    fr = current_row
                    fill = FIDA_FILL if (fund_offset % 2 == 0) else None

                    # Col 1: fund name indented
                    c = ws.cell(row=fr, column=1, value=f"  {fund}")
                    c.font = BODY_FONT
                    c.alignment = Alignment(horizontal="left", vertical="center")
                    c.border = Border(bottom=Side(style="thin", color="D5F5E3"))
                    if fill: c.fill = fill

                    # Col 2: units held — try exact match then normalised match
                    units = holdings.get((acc, fund), 0)
                    if not units:
                        # Strip all spaces for comparison
                        def norm(s): return s.replace(" ", "").upper()
                        units = next(
                            (v for (a, f), v in holdings.items()
                             if a == acc and norm(f) == norm(fund)),
                            0
                        )
                    uc = ws.cell(row=fr, column=2,
                                 value=round(units, 2) if units > 0 else None)
                    uc.font = BODY_FONT
                    uc.number_format = '#,##0.##'
                    uc.alignment = Alignment(horizontal="right", vertical="center")
                    uc.border = Border(bottom=Side(style="thin", color="D5F5E3"))
                    if fill: uc.fill = fill

                    # Col 3: total
                    tc = ws.cell(row=fr, column=3,
                                 value=f"=SUM({fc}{fr}:{lc}{fr})")
                    tc.number_format = NUM_FMT
                    tc.font = BODY_FONT
                    tc.alignment = Alignment(horizontal="right", vertical="center")
                    tc.border = Border(bottom=Side(style="thin", color="D5F5E3"))
                    if fill: tc.fill = fill

                    # Col 4+: month values
                    for col_idx, m in enumerate(fid_months, 4):
                        val = fund_df.loc[fund, m] if fund in fund_df.index else 0
                        is_est = m not in actual_set
                        fill2 = (FIDA_FILL if (fund_offset % 2 == 0) else None) if not is_est else (P("DCF0DC") if (fund_offset % 2 == 0) else P("EBF5EB"))
                        cell = ws.cell(row=fr, column=col_idx,
                                       value=int(round(val)) or None)
                        cell.number_format = NUM_FMT
                        cell.font = BODY_FONT if not is_est else EST_BODY
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                        cell.border = Border(bottom=Side(style="thin", color="D5F5E3"))
                        if fill2: cell.fill = fill2

                    current_row += 1

                # ── Equity dividends — write inline under Paul SIPP (2000001606) ──
                if acc == "2000001606":
                    EQUITY_DIVIDENDS_INLINE = [
                        ("  RELX PLC", {"2026-06": 413, "2026-09": 168}),
                        ("  Sage Group PLC", {"2026-06": 178}),
                        ("  Auto Trader Group", {"2026-09": 315}),
                        ("  Weir Group", {"2026-11": 78}),
                    ]
                    for eq_offset, (eq_label, eq_divs) in enumerate(EQUITY_DIVIDENDS_INLINE):
                        eq_row = current_row
                        equity_data_rows.append(eq_row)
                        fill_eq = FIDA_FILL if (eq_offset % 2 == 0) else None
                        c = ws.cell(row=eq_row, column=1, value=eq_label)
                        c.font = F(color="145A32", italic=True, size=10)
                        c.alignment = Alignment(horizontal="left", vertical="center")
                        c.border = Border(bottom=Side(style="thin", color="D5F5E3"))
                        if fill_eq: c.fill = fill_eq
                        ws.cell(row=eq_row, column=2).border = Border(bottom=Side(style="thin", color="D5F5E3"))
                        if fill_eq: ws.cell(row=eq_row, column=2).fill = fill_eq

                        # Total column (col 3)
                        tc_eq = ws.cell(row=eq_row, column=3, value=f"=SUM({fc}{eq_row}:{lc}{eq_row})")
                        tc_eq.number_format = NUM_FMT
                        tc_eq.font = F(color="145A32", italic=True, size=10)
                        tc_eq.alignment = Alignment(horizontal="right", vertical="center")
                        tc_eq.border = Border(bottom=Side(style="thin", color="D5F5E3"))
                        if fill_eq: tc_eq.fill = fill_eq
                        for col_idx, m in enumerate(fid_months, 4):
                            val = eq_divs.get(str(m), 0)
                            cell = ws.cell(row=eq_row, column=col_idx)
                            if val:
                                cell.value = val
                                cell.number_format = NUM_FMT
                            cell.font = F(color="145A32", italic=True, size=10)
                            cell.alignment = Alignment(horizontal="right", vertical="center")
                            cell.border = Border(bottom=Side(style="thin", color="D5F5E3"))
                            if fill_eq: cell.fill = fill_eq
                        ws.row_dimensions[eq_row].height = 15
                        current_row += 1

        # Note row
        note_eq = current_row
        nc = ws.cell(row=note_eq, column=1,
                     value="  * Sep/Nov equity dividend amounts are estimates based on prior year interim payments")
        nc.font = Font(italic=True, color="999999", size=8)
        nc.alignment = Alignment(horizontal="left", vertical="center")
        ws.merge_cells(start_row=note_eq, start_column=1,
                       end_row=note_eq, end_column=3 + len(fid_months))
        ws.row_dimensions[note_eq].height = 12
        current_row += 1

        # June provisional note
        note_jun = current_row
        nj = ws.cell(row=note_jun, column=1,
                     value="  * June 2026 figures are PROVISIONAL — based on partial transaction data (60-day export, 10 Jun) "
                           "and May-value estimates for fund income/salary. Will be revised once full June data is provided.")
        nj.font = Font(italic=True, bold=True, color="B7791F", size=8)
        nj.fill = P("FFF8E1")
        nj.alignment = Alignment(horizontal="left", vertical="center")
        ws.merge_cells(start_row=note_jun, start_column=1,
                       end_row=note_jun, end_column=3 + len(fid_months))
        ws.row_dimensions[note_jun].height = 14
        current_row += 1

        # Grand total — Salary + all fund accounts + equity dividends
        tot_r = current_row
        c = ws.cell(row=tot_r, column=1, value="Total Income")
        c.font = F(bold=True, color="145A32")
        c.fill = FIDT_FILL
        c.alignment = Alignment(horizontal="left", vertical="center")
        ws.cell(row=tot_r, column=2).fill = FIDT_FILL
        all_income_rows = [sal_row] + account_total_rows + equity_data_rows
        for col_idx in range(3, 4 + len(fid_months)):
            cl = get_column_letter(col_idx)
            refs = "+".join(f"{cl}{rr}" for rr in all_income_rows)
            cell = ws.cell(row=tot_r, column=col_idx, value=f"={refs}")
            cell.number_format = NUM_FMT
            cell.font = F(bold=True, color="145A32")
            cell.fill = FIDT_FILL
            cell.alignment = Alignment(horizontal="right", vertical="center")

    # ── Section 3: Accumulative Holdings ─────────────────────────────────────
    # Estimate monthly growth for Acc funds using the income yield from Inc equivalents
    # Yield proxy: monthly income per account / Inc fund value → apply to Acc fund value

    # Build yield map from Inc fund income data
    # Key: normalised fund base name → monthly income rate
    ACC_INC_EQUIV = {
        "Aegon High Yield Bond B Acc":               "Aegon High Yield Bond B Inc",
        "Schroder High Yield Opportunities Fund Z Acc": "Schroder High Yield Opportunities Fund Z Inc",
        "WS Guinness Global Energy Fund I Acc":       None,  # no Inc equivalent — use 6% annual
        "Man High Yield Opportunities Fund Prof D Acc": "Man High Yield Opportunities Fund Prof D Inc",
    }

    # Estimate annual yield rates from actual income / May 26 value
    def est_annual_yield(inc_fund_name, acc_value):
        """Estimate annual yield for an Acc fund from its Inc equivalent's income rate."""
        if inc_fund_name is None:
            return 0.06  # 6% default for unknown
        # Sum total 2026 income across all accounts for this fund from fid_pivot
        total_inc = 0
        if inc_fund_name in fid_pivot.index:
            total_inc = fid_pivot.loc[inc_fund_name, actual_months].sum()
        # Annualise: divide by months of data, multiply by 12
        months_of_data = len(actual_months)
        if months_of_data > 0 and total_inc > 0:
            # Find total Inc fund value from AccountSummary
            inc_total_val = sum(
                v for (a, f), v_dict in {}.items()
            )
            # Fall back: use the Acc fund value as proxy
            monthly_rate = total_inc / acc_value if acc_value > 0 else 0
            return monthly_rate * 12
        return 0.07  # 7% default

    C_ACC_H = "0E4D6B"   # teal-navy header
    C_ACC_A = "E8F4F8"   # light teal alt
    C_ACC_T = "A8D5E2"   # teal total

    ACC_FILL  = P(C_ACC_H)
    ACCA_FILL = P(C_ACC_A)
    ACCT_FILL = P(C_ACC_T)
    ACC_HFONT = F(bold=True, color="FFFFFF")
    ACC_TFONT = F(bold=True, color="0E4D6B")

    acc3_start = tot_r + 2

    n_acc_cols = 3 + len(fid_months)   # Label + Units + Total + months (same as income table)
    ws.merge_cells(start_row=acc3_start, start_column=1,
                   end_row=acc3_start, end_column=n_acc_cols)
    t_acc = ws.cell(row=acc3_start, column=1, value="Accumulative Holdings")
    t_acc.font = F(bold=True, color="FFFFFF", size=12)
    t_acc.fill = ACC_FILL
    t_acc.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[acc3_start].height = 24

    acc_hdr_row = acc3_start + 1
    for col, h in enumerate(["Account / Fund", "Units", "Total"] + fid_month_labels, 1):
        c = ws.cell(row=acc_hdr_row, column=col, value=h)
        if col <= 3:
            c.font = ACC_HFONT; c.fill = ACC_FILL
        else:
            m = fid_months[col - 4]
            c.font = ACC_HFONT if m in actual_set else EST_FONT
            c.fill = ACC_FILL if m in actual_set else EST_HDR_FILL
        c.alignment = Alignment(horizontal="left" if col == 1 else "right", vertical="center")
    ws.row_dimensions[acc_hdr_row].height = 18

    acc_cur_row = acc_hdr_row + 1
    acc_account_total_rows = []

    for person in FAMILY_ORDER:
        person_accs = [acc for acc, owner in ACCOUNT_OWNER.items() if owner == person]
        for acc in person_accs:
            if acc not in acc_holdings:
                continue
            funds = acc_holdings[acc]
            label = f"{ACCOUNT_LABELS.get(acc, acc)} ({person})"
            n_fund_rows = len(funds)

            # Account header row
            acc_ar = acc_cur_row
            acc_account_total_rows.append(acc_ar)

            c = ws.cell(row=acc_ar, column=1, value=label)
            c.font = ACC_TFONT; c.fill = ACCT_FILL
            c.alignment = Alignment(horizontal="left", vertical="center")
            c.border = Border(bottom=Side(style="thin", color=C_ACC_T))

            ws.cell(row=acc_ar, column=2).fill = ACCT_FILL  # units blank for account
            ws.cell(row=acc_ar, column=2).border = Border(bottom=Side(style="thin", color=C_ACC_T))

            fc_a = get_column_letter(4)
            lc_a = get_column_letter(3 + len(fid_months))
            tc = ws.cell(row=acc_ar, column=3,
                         value=f"=SUM({fc_a}{acc_ar+1}:{lc_a}{acc_ar+n_fund_rows})")
            tc.number_format = NUM_FMT; tc.font = ACC_TFONT; tc.fill = ACCT_FILL
            tc.alignment = Alignment(horizontal="right", vertical="center")
            tc.border = Border(bottom=Side(style="thin", color=C_ACC_T))

            # Sum of account values for each month (from fund rows below)
            for col_idx, m in enumerate(fid_months, 4):
                cell = ws.cell(row=acc_ar, column=col_idx,
                               value=f"=SUM({get_column_letter(col_idx)}{acc_ar+1}:{get_column_letter(col_idx)}{acc_ar+n_fund_rows})")
                cell.number_format = NUM_FMT; cell.font = ACC_TFONT; cell.fill = ACCT_FILL
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.border = Border(bottom=Side(style="thin", color=C_ACC_T))

            acc_cur_row += 1

            # Fund rows — use actual price×units from transaction history
            for fund_offset, (fund, fund_data) in enumerate(sorted(funds.items())):
                fr = acc_cur_row
                is_alt = (fund_offset % 2 == 0)
                fill = ACCA_FILL if is_alt else None
                units = fund_data["units"]
                # Use price_appreciation (ignores new unit purchases, just price-driven gain)
                month_vals = fund_data.get("price_appreciation", fund_data.get("monthly_values", {}))
                may_val = fund_data["value"]

                # For future months not in price history, extrapolate using known annual yield
                # Annual yields from Fidelity (guaranteed non-negative)
                FUND_YIELDS = {
                    "Aegon High Yield Bond B Acc":                  0.0732,
                    "Schroder High Yield Opportunities Fund Z Acc": 0.0769,
                    "WS Guinness Global Energy Fund I Acc":         0.0235,
                }
                annual_yield = FUND_YIELDS.get(fund, 0.05)
                monthly_growth = (1 + annual_yield) ** (1/12) - 1

                if month_vals:
                    last_known = max(month_vals.keys())
                    last_val = month_vals[last_known]
                    for m in fid_months:
                        if m not in month_vals:
                            offset = (m - last_known).n
                            month_vals[m] = round(last_val * ((1 + monthly_growth) ** offset))

                c = ws.cell(row=fr, column=1, value=f"  {fund}")
                c.font = BODY_FONT; c.alignment = Alignment(horizontal="left", vertical="center")
                c.border = Border(bottom=Side(style="thin", color="B2E0EC"))
                if fill: c.fill = fill

                uc = ws.cell(row=fr, column=2, value=round(units, 2) if units else None)
                uc.number_format = '#,##0.##'; uc.font = BODY_FONT
                uc.alignment = Alignment(horizontal="right", vertical="center")
                uc.border = Border(bottom=Side(style="thin", color="B2E0EC"))
                if fill: uc.fill = fill

                tc2 = ws.cell(row=fr, column=3,
                              value=f"=SUM({fc_a}{fr}:{lc_a}{fr})")
                tc2.number_format = NUM_FMT; tc2.font = BODY_FONT
                tc2.alignment = Alignment(horizontal="right", vertical="center")
                tc2.border = Border(bottom=Side(style="thin", color="B2E0EC"))
                if fill: tc2.fill = fill

                for col_idx, m in enumerate(fid_months, 4):
                    val = month_vals.get(m, 0)
                    is_est = m not in actual_set
                    cell = ws.cell(row=fr, column=col_idx, value=int(round(val)) if val else None)
                    cell.number_format = NUM_FMT
                    cell.font = BODY_FONT if not is_est else EST_BODY
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    cell.border = Border(bottom=Side(style="thin", color="B2E0EC"))
                    if fill: cell.fill = fill

                acc_cur_row += 1

    # Grand total row — shows INCREASE each month (new units × price), not total value
    tot_acc_r = acc_cur_row
    c = ws.cell(row=tot_acc_r, column=1, value="TOTAL Accumulative (price appreciation)")
    c.font = F(bold=True, color="FFFFFF"); c.fill = ACC_FILL
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.cell(row=tot_acc_r, column=2).fill = ACC_FILL
    ws.cell(row=tot_acc_r, column=3).fill = ACC_FILL

    for col_idx, m in enumerate(fid_months, 4):
        monthly_increase = sum(
            fd_h.get("price_appreciation", {}).get(m, 0)
            for funds_h in acc_holdings.values()
            for fd_h in funds_h.values()
        )
        cell = ws.cell(row=tot_acc_r, column=col_idx,
                       value=int(round(monthly_increase)) if monthly_increase else None)
        cell.number_format = NUM_FMT; cell.font = F(bold=True, color="FFFFFF")
        cell.fill = ACC_FILL; cell.alignment = Alignment(horizontal="right", vertical="center")
    acc_cur_row += 1

    # ── Update Total Income and Accumulations with Excel formulas now that row numbers are known ──
    # = Total Income (tot_r) + TOTAL Accumulative (tot_acc_r)
    col_letters = [get_column_letter(4 + i) for i in range(len(all_months))]
    for col_idx, m in enumerate(all_months, 4):
        cl = get_column_letter(col_idx)
        cell = ws.cell(row=total_inc_acc_row, column=col_idx,
                       value=f"={cl}{tot_r}+{cl}{tot_acc_r}")
        cell.number_format = NUM_FMT
        cell.font = F(bold=True, color="042C53")
        cell.fill = P(C_SUM_T)
        cell.alignment = Alignment(horizontal="right", vertical="center")
    # Col C = annual total (sum of all 12 months)
    sum_range = "+".join(f"{cl}{total_inc_acc_row}" for cl in col_letters)
    c_total = ws.cell(row=total_inc_acc_row, column=3, value=f"={sum_range}")
    c_total.number_format = NUM_FMT
    c_total.font = F(bold=True, color="042C53")
    c_total.fill = P(C_SUM_T)
    c_total.alignment = Alignment(horizontal="right", vertical="center")
    fid3_start = acc_cur_row + 1

    C_FID3_H = "4A235B"   # deep purple header
    C_FID3_A = "F5EEF8"   # light purple alt rows
    C_FID3_T = "D2B4DE"   # purple total

    FID3_FILL  = P(C_FID3_H)
    FID3A_FILL = P(C_FID3_A)
    FID3T_FILL = P(C_FID3_T)
    FID3_HFONT = F(bold=True, color="FFFFFF")
    FID3_TFONT = F(bold=True, color="4A235B")

    n_fid3_cols = 3 + len(spend_months)
    ws.merge_cells(start_row=fid3_start, start_column=1,
                   end_row=fid3_start, end_column=n_fid3_cols)
    t3 = ws.cell(row=fid3_start, column=1, value="Fidelity")
    t3.font = F(bold=True, color="FFFFFF", size=12)
    t3.fill = FID3_FILL
    t3.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[fid3_start].height = 24

    fid3_hdr = fid3_start + 1
    for col, h in enumerate(["Category", "", "Total"] + spend_month_labels, 1):
        c = ws.cell(row=fid3_hdr, column=col, value=h)
        if col <= 3:
            c.font = FID3_HFONT
            c.fill = FID3_FILL
        else:
            m = spend_months[col - 4]
            c.font = FID3_HFONT if m in actual_set else EST_FONT
            c.fill = FID3_FILL if m in actual_set else EST_HDR_FILL
        c.alignment = Alignment(horizontal="left" if col == 1 else "right", vertical="center")
    ws.row_dimensions[fid3_hdr].height = 18

    fid3_data_row = fid3_start + 2
    fill = FID3A_FILL
    c = ws.cell(row=fid3_data_row, column=1, value="Fidelity card payments")
    c.font = BODY_FONT; c.fill = fill
    c.alignment = Alignment(horizontal="left", vertical="center")
    c.border = Border(bottom=Side(style="thin", color="D2B4DE"))
    ws.cell(row=fid3_data_row, column=2).fill = fill
    ws.cell(row=fid3_data_row, column=2).border = Border(bottom=Side(style="thin", color="D2B4DE"))

    fc = get_column_letter(4)
    lc = get_column_letter(3 + len(spend_months))
    tc = ws.cell(row=fid3_data_row, column=3,
                 value=f"=SUM({fc}{fid3_data_row}:{lc}{fid3_data_row})")
    tc.number_format = NUM_FMT; tc.font = FID3_TFONT; tc.fill = FID3T_FILL
    tc.alignment = Alignment(horizontal="right", vertical="center")
    tc.border = Border(bottom=Side(style="thin", color="D2B4DE"))

    for col_idx, m in enumerate(spend_months, 4):
        val = spend_pivot.loc["Fidelity", m] if "Fidelity" in spend_pivot.index else 0
        is_est = m not in actual_set
        cell = ws.cell(row=fid3_data_row, column=col_idx,
                       value=int(round(val)) or None)
        cell.number_format = NUM_FMT
        cell.font = BODY_FONT if not is_est else EST_BODY
        cell.fill = FID3A_FILL if not is_est else P("ECE8F4")
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.border = Border(bottom=Side(style="thin", color="D2B4DE"))

    # ── Total Income + Salary row ──────────────────────────────────────────────
    combo_row = fid3_data_row + 1
    c = ws.cell(row=combo_row, column=1, value="Total Income & Salary")
    c.font = F(bold=True, color="FFFFFF"); c.fill = FID3_FILL
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.cell(row=combo_row, column=2).fill = FID3_FILL  # blank

    # Income table: col1=label, col2=units, col3=Total, col4+=months (same as spend now)
    # Reference tot_r directly — all cols now aligned
    for col_idx in range(3, 4 + len(spend_months)):
        cl = get_column_letter(col_idx)
        cell = ws.cell(row=combo_row, column=col_idx, value=f"={cl}{tot_r}")
        cell.number_format = NUM_FMT
        cell.font = F(bold=True, color="FFFFFF"); cell.fill = FID3_FILL
        cell.alignment = Alignment(horizontal="right", vertical="center")

    # ── Investment Risk Metrics Table ──────────────────────────────────────────
    # Classify all Fidelity holdings into: Shares, Income Funds, Non-Income Funds
    # Columns: Group | Paul SIPP | Paul ISA | Joint Acct | Susan SIPP | Susan ISA | Jayne JISA | Liam ISA | Liam Inv | TOTAL
    # Rows: Shares, Income Funds, Non-Income Funds, TOTAL

    METRIC_ACCS_ORDER = [
        ('2000001606', 'Paul SIPP'),
        ('SANX002282', 'Paul ISA'),
        ('SANQ000468', 'Joint Acct'),
        ('2000001604', 'Susan SIPP'),
        ('SANX002617', 'Susan ISA'),
        ('SANX002936', 'Jayne JISA'),
        ('AS10303823', 'Liam ISA'),
        ('AG10131710', 'Liam Inv Acct'),
    ]

    HOLDING_GROUPS = {
        'Aegon High Yield Bond B Inc':                   'Income Funds',
        'Aegon High Yield Bond B Acc':                   'Income Funds',
        'Schroder High Yield Opportunities Fund Z Inc':  'Income Funds',
        'Schroder High Yield Opportunities Fund Z Acc':  'Income Funds',
        'Man High Yield Opportunities Fund Prof D Inc':  'Income Funds',
        'Man High Yield Opportunities Fund Prof D Acc':  'Income Funds',
        'WS Guinness Global Energy Fund I Acc':          'Non-Income Funds',
        'AUTOTRADER GROUP PLC,ORD GBP0.01(AUTO)':       'Shares',
        'AVIVA,ORD GBP0.328947368(AV.)':                 'Shares',
        'RELX PLC,ORD GBP0.1444(REL)':                  'Shares',
        'THE SAGE GROUP PLC,GBP0.01051948(SGE)':         'Shares',
        'WEIR GROUP,ORD GBP0.125(WEIR)':                 'Shares',
        'LEGAL & GENERAL GROUP,ORD GBP0.025(LGEN)':      'Shares',
    }
    GROUPS_ORDER = ['Shares', 'Income Funds', 'Non-Income Funds']

    # Build data: group → account → value
    from collections import defaultdict as _dd
    metric_data = {g: {a: 0.0 for a, _ in METRIC_ACCS_ORDER} for g in GROUPS_ORDER}
    metric_data['_cash'] = {a: 0.0 for a, _ in METRIC_ACCS_ORDER}

    import csv as _csv2, io as _io2, os as _os2
    # Find AccountSummary path — passed in via summary_data
    _acct_path = summary_data.get('_account_summary_path', '')
    if not _acct_path or not _os2.path.exists(_acct_path):
        for _p in ['AccountSummary.csv', '/home/claude/AccountSummary.csv']:
            if _os2.path.exists(_p):
                _acct_path = _p
                break
    with open(_acct_path, encoding='utf-8-sig') as _f:
        _content = _f.read()
    _lines = _content.replace('\r','').split('\n')
    _hi = max(i for i, l in enumerate(_lines) if l.startswith('Type,Holdings,Account number'))
    for _row in _csv2.DictReader(_io2.StringIO('\n'.join(_lines[_hi:]))):
        _acc = _row.get('Account number','').strip()
        if _acc not in {a for a,_ in METRIC_ACCS_ORDER}: continue
        _t = _row.get('Type','').strip()
        _val = float(_row.get('Value (£)','0').replace(',','') or 0)
        if _t == 'Asset':
            _fund = _row.get('Holdings','').strip()
            _grp = HOLDING_GROUPS.get(_fund)
            if _grp:
                metric_data[_grp][_acc] += _val
        elif _t == 'Account':
            # Total account value — used to derive cash (account total - sum of assets)
            metric_data['_cash'][_acc] = _val

    # Cash = account total - sum of classified assets
    metric_cash = {}
    for acc, _ in METRIC_ACCS_ORDER:
        asset_sum = sum(metric_data[g][acc] for g in GROUPS_ORDER)
        metric_cash[acc] = max(0, metric_data['_cash'][acc] - asset_sum)

    # ── Write table ─────────────────────────────────────────────────────────────
    C_MET_H = "2C3E50"    # dark slate header
    C_MET_A = "F2F3F4"    # light grey alt
    C_MET_T = "D5D8DC"    # grey total
    C_MET_S = "E8F4F8"    # light blue for Shares
    C_MET_I = "EAF5EA"    # light green for Income
    C_MET_N = "FEF9E7"    # light yellow for Non-Income

    GROUP_COLOURS = {
        'Shares':           C_MET_S,
        'Income Funds':     C_MET_I,
        'Non-Income Funds': C_MET_N,
    }

    met_start = combo_row + 3 if 'combo_row' in dir() else ws.max_row + 3
    n_acc_cols = len(METRIC_ACCS_ORDER)
    n_met_cols = 1 + n_acc_cols + 1  # Group label + accounts + Total

    # Section title
    ws.merge_cells(start_row=met_start, start_column=1,
                   end_row=met_start, end_column=n_met_cols)
    t = ws.cell(row=met_start, column=1, value="Investment Risk Metrics")
    t.font = F(bold=True, color="FFFFFF", size=12)
    t.fill = P(C_MET_H)
    t.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[met_start].height = 24

    # Column headers
    hdr_row = met_start + 1
    ws.cell(row=hdr_row, column=1, value="Group").font = F(bold=True, color="FFFFFF", size=10)
    ws.cell(row=hdr_row, column=1).fill = P(C_MET_H)
    ws.cell(row=hdr_row, column=1).alignment = Alignment(horizontal="left", vertical="center")
    for ci, (acc, label) in enumerate(METRIC_ACCS_ORDER, 2):
        c = ws.cell(row=hdr_row, column=ci, value=label)
        c.font = F(bold=True, color="FFFFFF", size=9)
        c.fill = P(C_MET_H)
        c.alignment = Alignment(horizontal="right", vertical="center")
    tot_hdr = ws.cell(row=hdr_row, column=2 + n_acc_cols, value="TOTAL")
    tot_hdr.font = F(bold=True, color="FFFFFF", size=10)
    tot_hdr.fill = P(C_MET_H)
    tot_hdr.alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[hdr_row].height = 18

    # Data rows — value rows
    val_rows = {}
    cur_met = hdr_row + 1
    grand_met_total = sum(
        sum(metric_data[g][acc] for acc, _ in METRIC_ACCS_ORDER)
        for g in GROUPS_ORDER
    ) + sum(metric_cash.values())

    all_groups = GROUPS_ORDER + ['Cash']
    for gi, grp in enumerate(all_groups):
        row_fill = GROUP_COLOURS.get(grp, C_MET_A)
        r = cur_met
        val_rows[grp] = r
        # Label
        lbl = ws.cell(row=r, column=1, value=grp)
        lbl.font = F(bold=False, color="000000", size=10)
        lbl.fill = P(row_fill)
        lbl.alignment = Alignment(horizontal="left", vertical="center")
        # Account values
        row_total = 0
        for ci, (acc, _) in enumerate(METRIC_ACCS_ORDER, 2):
            val = metric_data[grp][acc] if grp in metric_data else metric_cash.get(acc, 0)
            if grp == 'Cash':
                val = metric_cash.get(acc, 0)
            row_total += val
            cell = ws.cell(row=r, column=ci)
            if val > 0:
                cell.value = round(val)
                cell.number_format = '#,##0'
            cell.fill = P(row_fill)
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.font = F(size=10)
        # Row total
        tc = ws.cell(row=r, column=2 + n_acc_cols)
        if row_total > 0:
            tc.value = round(row_total)
            tc.number_format = '#,##0'
        tc.fill = P(row_fill)
        tc.alignment = Alignment(horizontal="right", vertical="center")
        tc.font = F(bold=True, size=10)
        ws.row_dimensions[r].height = 16
        cur_met += 1

    # % of total rows
    pct_start = cur_met
    cur_met += 1  # blank row
    ws.row_dimensions[cur_met - 1].height = 6

    for gi, grp in enumerate(all_groups):
        row_fill = GROUP_COLOURS.get(grp, C_MET_A)
        r = pct_start + gi
        lbl = ws.cell(row=r, column=1, value=f"{grp} %")
        lbl.font = F(italic=True, color="555555", size=9)
        lbl.fill = P(row_fill)
        lbl.alignment = Alignment(horizontal="left", vertical="center")
        row_total = sum(metric_data[grp][acc] for acc, _ in METRIC_ACCS_ORDER) if grp in metric_data else sum(metric_cash.values())
        for ci, (acc, _) in enumerate(METRIC_ACCS_ORDER, 2):
            val = metric_data[grp][acc] if grp in metric_data else metric_cash.get(acc, 0)
            if grp == 'Cash':
                val = metric_cash.get(acc, 0)
            acc_total_val = metric_data['_cash'].get(acc, 0)
            if acc_total_val > 0:
                pct = val / acc_total_val * 100
                cell = ws.cell(row=r, column=ci)
                if pct > 0:
                    cell.value = round(pct, 1)
                    cell.number_format = '0.0"%"'
                cell.fill = P(row_fill)
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.font = F(italic=True, size=9, color="555555")
        # Overall % of portfolio
        overall_pct = row_total / grand_met_total * 100 if grand_met_total > 0 else 0
        tc = ws.cell(row=r, column=2 + n_acc_cols)
        if overall_pct > 0:
            tc.value = round(overall_pct, 1)
            tc.number_format = '0.0"%"'
        tc.fill = P(row_fill)
        tc.alignment = Alignment(horizontal="right", vertical="center")
        tc.font = F(bold=True, italic=True, size=9, color="555555")
        ws.row_dimensions[r].height = 14
        cur_met += 1

    # Grand total row
    cur_met += 1
    gt_row = cur_met
    ws.cell(row=gt_row, column=1, value="TOTAL INVESTMENTS").font = F(bold=True, color="FFFFFF", size=10)
    ws.cell(row=gt_row, column=1).fill = P(C_MET_H)
    ws.cell(row=gt_row, column=1).alignment = Alignment(horizontal="left", vertical="center")
    for ci, (acc, _) in enumerate(METRIC_ACCS_ORDER, 2):
        acc_total_val = metric_data['_cash'].get(acc, 0)
        cell = ws.cell(row=gt_row, column=ci)
        if acc_total_val > 0:
            cell.value = round(acc_total_val)
            cell.number_format = '#,##0'
        cell.fill = P(C_MET_H)
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.font = F(bold=True, color="FFFFFF", size=10)
    tc = ws.cell(row=gt_row, column=2 + n_acc_cols)
    fid_total_val = sum(metric_data['_cash'].get(acc, 0) for acc, _ in METRIC_ACCS_ORDER)
    if fid_total_val > 0:
        tc.value = round(fid_total_val)
        tc.number_format = '#,##0'
    tc.fill = P(C_MET_H)
    tc.alignment = Alignment(horizontal="right", vertical="center")
    tc.font = F(bold=True, color="FFFFFF", size=10)
    ws.row_dimensions[gt_row].height = 18

    # Note row
    note_met = gt_row + 1
    nc = ws.cell(row=note_met, column=1,
                 value="Values as at AccountSummary export date. Cash = account total minus classified assets. Totals match Fidelity accounts section above.")
    nc.font = Font(italic=True, color="888888", size=8)
    nc.alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells(start_row=note_met, start_column=1,
                   end_row=note_met, end_column=n_met_cols)
    ws.row_dimensions[note_met].height = 12

    # ── Targets Table ───────────────────────────────────────────────────────────
    C_TGT_H  = "1A3A5C"   # dark navy header
    C_TGT_ON = "E9F7EF"   # green — on/above target
    C_TGT_OF = "FDEDEC"   # red   — below target
    C_TGT_NR = "FEF9E7"   # amber — within 10% of target
    C_TGT_LB = "EBF5FB"   # blue  — label rows

    tgt_start = note_met + 3
    # Columns: Metric | Actual | Target | Status | Notes
    TGT_COLS = ["Metric", "Actual", "Target", "Status", "Notes"]
    TGT_WIDTHS = [38, 16, 16, 12, 40]

    # Title
    ws.merge_cells(start_row=tgt_start, start_column=1, end_row=tgt_start, end_column=5)
    tt = ws.cell(row=tgt_start, column=1, value="2026 Targets")
    tt.font = F(bold=True, color="FFFFFF", size=12)
    tt.fill = P(C_TGT_H)
    tt.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[tgt_start].height = 24

    # Headers
    hdr_tgt = tgt_start + 1
    for ci, (h, w) in enumerate(zip(TGT_COLS, TGT_WIDTHS), 1):
        c = ws.cell(row=hdr_tgt, column=ci, value=h)
        c.font = F(bold=True, color="FFFFFF", size=10)
        c.fill = P(C_TGT_H)
        c.alignment = Alignment(horizontal="left" if ci == 1 else "right", vertical="center")
    ws.row_dimensions[hdr_tgt].height = 18

    # ── Compute actuals ────────────────────────────────────────────────────────
    # Income per month — use hardcoded income history for Jan-Apr, live pivot for May+
    inc_hist_data = load_income_history()
    hist_income_jan_apr = sum(
        sum(v for v in month_vals.values() if pd.Period(m,'M') < pd.Period('2026-05','M'))
        for acc_id, month_vals in inc_hist_data.items()
        for m in [str(p) for p in month_vals]
    )
    # Recalculate properly: sum Jan-Apr from hardcoded history
    jan_apr_income = 0
    for acc_id, month_vals in inc_hist_data.items():
        for period, val in month_vals.items():
            jan_apr_income += val
    # May+ income from live fid_pivot
    may_dec_income = 0
    if not fid_pivot.empty:
        for m in [pd.Period(f'2026-{mo:02d}','M') for mo in range(5, 13)]:
            if m in fid_pivot.columns:
                may_dec_income += fid_pivot[m].sum()
            else:
                # Estimate from May actual
                may_col = pd.Period('2026-05','M')
                if may_col in fid_pivot.columns:
                    may_dec_income += fid_pivot[may_col].sum()
    # Add salary (hardcoded Jan-Apr) + May salary from spend_pivot
    salary_jan_apr = sum(load_spend_history().get('Salary', {}).values())
    salary_may = spend_pivot.loc['Salary', pd.Period('2026-05','M')] if 'Salary' in spend_pivot.index and pd.Period('2026-05','M') in spend_pivot.columns else 0
    # Equity dividends annual
    equity_annual = 413 + 168 + 178 + 315 + 78  # RELX + Sage + AutoTrader + Weir

    # Build from Total Income row in the spreadsheet (more reliable after formula evaluation)
    # Use fid_pivot directly for accuracy
    fid_annual = 0
    if not fid_pivot.empty:
        for m in [pd.Period(f'2026-{mo:02d}','M') for mo in range(1, 13)]:
            if m in fid_pivot.columns:
                fid_annual += fid_pivot[m].sum()
            else:
                # Estimate missing months from nearest known
                nearest = max((c for c in fid_pivot.columns if c <= m), default=None)
                if nearest:
                    fid_annual += fid_pivot[nearest].sum()
    # Add history months not in fid_pivot
    for acc_id, month_vals in inc_hist_data.items():
        for period, val in month_vals.items():
            if period not in fid_pivot.columns:
                fid_annual += val
    salary_annual = sum(load_spend_history().get('Salary', {}).values()) + salary_may * 8
    total_annual_income = fid_annual + salary_annual + equity_annual
    income_avg_pm = round(total_annual_income / 12)

    # 2. Paul SIPP 25% drawdown
    paul_sipp_val_tgt = 0
    for _r in ws.iter_rows():
        if '2000001606' in str(_r[0].value or '') and 'SIPP' in str(_r[0].value or ''):
            paul_sipp_val_tgt = float(_r[7].value or 0)  # May
            break
    paul_drawdown = round(paul_sipp_val_tgt * 0.25)

    # 3. Susan SIPP 25% drawdown
    susan_sipp_val_tgt = 0
    for _r in ws.iter_rows():
        if '2000001604' in str(_r[0].value or '') and 'SIPP' in str(_r[0].value or ''):
            susan_sipp_val_tgt = float(_r[7].value or 0)
            break
    susan_drawdown = round(susan_sipp_val_tgt * 0.25)

    # 4. ISA values combined
    paul_isa_val = susan_isa_val = 0
    for _r in ws.iter_rows():
        if str(_r[0].value or '').strip() == 'Investment ISA (Paul)':
            paul_isa_val = float(_r[7].value or 0)
        if str(_r[0].value or '').strip() == 'Investment ISA (Susan)':
            susan_isa_val = float(_r[7].value or 0)
    isa_combined = round(paul_isa_val + susan_isa_val)

    # 5. Growth % (Shares + Non-Income Funds) and Income % of total invested
    shares_val = inc_fund_val = non_inc_val = 0
    for _r in ws.iter_rows():
        v = str(_r[0].value or '').strip()
        if v == 'Shares':
            shares_val = float(_r[9].value or 0)
        elif v == 'Income Funds':
            inc_fund_val = float(_r[9].value or 0)
        elif v == 'Non-Income Funds':
            non_inc_val = float(_r[9].value or 0)
    total_invested = shares_val + inc_fund_val + non_inc_val
    growth_pct  = round((shares_val + non_inc_val) / total_invested * 100, 1) if total_invested else 0
    income_pct  = round(inc_fund_val / total_invested * 100, 1) if total_invested else 0

    # 6. Fidelity service fees — sum from live TransactionHistory
    import csv as _csv3, io as _io3
    import sys as _sys3
    _th_path = next((a for a in _sys3.argv[1:] if 'TransactionHistory' in a or 'transaction' in a.lower()), 'TransactionHistory.csv')
    with open(_th_path) as _f3:
        _lines3 = _f3.read().replace('\r','').split('\n')
    _start3 = next(i for i,l in enumerate(_lines3) if l.startswith('Order date'))
    _rows3 = list(_csv3.DictReader(_io3.StringIO('\n'.join(_lines3[_start3:]))))
    svc_fees_live = sum(abs(float(r.get('Amount','0') or 0))
                        for r in _rows3
                        if r.get('Transaction type','').strip() == 'Service Fee'
                        and r.get('Status','').strip() == 'Completed')
    # Annualise from Jan-May actual (5 months)
    svc_fees_annual = round(svc_fees_live / 5 * 12)

    def status_cell(ws, r, c, actual, target, higher_is_better=True, is_pct=False):
        """Write RAG status cell."""
        if target == 0:
            ws.cell(row=r, column=c, value="—")
            return "—"
        pct_of_tgt = actual / target
        if higher_is_better:
            if pct_of_tgt >= 1.0:   rag, txt = C_TGT_ON, "✓ On Target"
            elif pct_of_tgt >= 0.9: rag, txt = C_TGT_NR, "⚠ Near Target"
            else:                    rag, txt = C_TGT_OF, "✗ Below Target"
        else:  # lower is better (e.g. costs)
            if pct_of_tgt <= 1.0:   rag, txt = C_TGT_ON, "✓ Within Budget"
            elif pct_of_tgt <= 1.1: rag, txt = C_TGT_NR, "⚠ Slightly Over"
            else:                    rag, txt = C_TGT_OF, "✗ Over Budget"
        cell = ws.cell(row=r, column=c, value=txt)
        cell.fill = P(rag)
        cell.font = F(bold=True, size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        return rag

    def tgt_row(ws, r, metric, actual, target, status_rag, note,
                actual_fmt="£{:,.0f}", target_fmt="£{:,.0f}", alt=False):
        bg = status_rag if status_rag not in (True, False) else (C_TGT_LB if alt else "FFFFFF")
        # Metric label
        c = ws.cell(row=r, column=1, value=metric)
        c.font = F(size=10); c.fill = P(C_TGT_LB)
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        # Actual
        act_cell = ws.cell(row=r, column=2, value=actual_fmt.format(actual) if actual_fmt else actual)
        act_cell.font = F(bold=True, size=11)
        act_cell.fill = P(status_rag if isinstance(status_rag, str) and len(status_rag)==6 else "FFFFFF")
        act_cell.alignment = Alignment(horizontal="right", vertical="center")
        # Target
        tgt_cell = ws.cell(row=r, column=3, value=target_fmt.format(target) if target_fmt else target)
        tgt_cell.font = F(size=10, color="444444"); tgt_cell.fill = P("F8F9FA")
        tgt_cell.alignment = Alignment(horizontal="right", vertical="center")
        # Note
        nc = ws.cell(row=r, column=5, value=note)
        nc.font = F(italic=True, size=9, color="666666"); nc.fill = P(C_TGT_LB)
        nc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws.row_dimensions[r].height = 22

    # ── Write target rows ──────────────────────────────────────────────────────
    tgt_data_start = hdr_tgt + 1

    # Row 1: Income per month
    r = tgt_data_start
    rag = status_cell(ws, r, 4, income_avg_pm, 10000, higher_is_better=True)
    tgt_row(ws, r,
            "Income per month (avg over 12 months)",
            income_avg_pm, 10000, rag,
            f"Annual total £{int(total_annual_income):,} ÷ 12. Includes salary, Fidelity fund income & equity dividends.")

    # Row 2: Paul SIPP 25% drawdown
    r += 1
    PAUL_DRAWDOWN_TAKEN = 0
    rag2 = status_cell(ws, r, 4, PAUL_DRAWDOWN_TAKEN + paul_drawdown, 269000, higher_is_better=True)
    tgt_row(ws, r,
            "Paul Pension — 25% tax-free drawdown available",
            paul_drawdown, 269000, rag2,
            f"SIPP value £{int(paul_sipp_val_tgt):,} × 25% = £{paul_drawdown:,}. Taken to date: £0.")

    # Row 3: Susan SIPP 25% drawdown
    r += 1
    SUSAN_DRAWDOWN_TAKEN = 0
    rag3 = status_cell(ws, r, 4, SUSAN_DRAWDOWN_TAKEN + susan_drawdown, 269000, higher_is_better=True)
    tgt_row(ws, r,
            "Susan Pension — 25% tax-free drawdown available",
            susan_drawdown, 269000, rag3,
            f"SIPP value £{int(susan_sipp_val_tgt):,} × 25% = £{susan_drawdown:,}. Taken to date: £0.")

    # Row 4: ISA values
    r += 1
    rag4 = status_cell(ws, r, 4, isa_combined, 1000000, higher_is_better=True)
    tgt_row(ws, r,
            "ISA Values (Paul ISA + Susan ISA combined)",
            isa_combined, 1000000, rag4,
            f"Paul ISA £{int(paul_isa_val):,} + Susan ISA £{int(susan_isa_val):,}.")

    # Row 5: Growth funds % — target is to REACH 60%, so higher is better
    r += 1
    rag5 = status_cell(ws, r, 4, growth_pct, 60, higher_is_better=True)
    tgt_row(ws, r,
            "Growth funds % of total invested (Shares + Non-Income Funds)",
            growth_pct, 60, rag5,
            f"Shares £{int(shares_val):,} + Non-Income £{int(non_inc_val):,} = £{int(shares_val+non_inc_val):,} of £{int(total_invested):,}.",
            actual_fmt="{:.1f}%", target_fmt="{:.0f}%")

    # Row 6: Income fund % — target is to stay AT/BELOW 40%, lower is better
    r += 1
    rag6 = status_cell(ws, r, 4, income_pct, 40, higher_is_better=False)
    tgt_row(ws, r,
            "Income funds % of total invested",
            income_pct, 40, rag6,
            f"Income Funds £{int(inc_fund_val):,} of £{int(total_invested):,} total invested. Currently {income_pct:.1f}% — target is to reduce to 40%.",
            actual_fmt="{:.1f}%", target_fmt="{:.0f}%")

    # Row 7: Fidelity service costs
    r += 1
    rag7 = status_cell(ws, r, 4, svc_fees_annual, 3000, higher_is_better=False)
    tgt_row(ws, r,
            "Fidelity annual service costs",
            svc_fees_annual, 3000, rag7,
            f"Jan–May actual £{svc_fees_live:,.2f} × 12/5 = £{svc_fees_annual:,} annualised estimate.")

    # Single comprehensive sheet
    ws.title = "Wealth Summary"
    ws.sheet_properties.tabColor = "1F4E79"

    # ── Column widths ──────────────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 12
    # Historic month cols: narrower
    for i, m in enumerate(sum_months):
        col_letter = get_column_letter(4 + i)
        ws.column_dimensions[col_letter].width = 9 if m in hist_months else 11

    # Freeze panes: col A–C (labels) and row 1–2 (title + month headers) always visible
    ws.freeze_panes = "D3"

    # ── Move 'Investment Risk Metrics' and '2026 Targets' to a separate sheet ──
    section_rows = {}
    for row in ws.iter_rows():
        v = str(row[0].value or '').strip()
        if v in ('Investment Risk Metrics', '2026 Targets'):
            section_rows[v] = row[0].row

    if 'Investment Risk Metrics' in section_rows:
        move_start = section_rows['Investment Risk Metrics']
        move_end = ws.max_row

        ws_targets = wb.create_sheet("Targets")
        ws_targets.sheet_properties.tabColor = "E67E22"

        from copy import copy as _copy
        dst_row = 1
        for sr in range(move_start, move_end + 1):
            for sc in range(1, ws.max_column + 1):
                src_cell = ws.cell(row=sr, column=sc)
                dst_cell = ws_targets.cell(row=dst_row, column=sc)
                dst_cell.value = src_cell.value
                if src_cell.data_type == 'f':
                    dst_cell.data_type = 'f'
                if src_cell.has_style:
                    try:
                        dst_cell.font          = _copy(src_cell.font)
                        dst_cell.fill          = _copy(src_cell.fill)
                        dst_cell.alignment     = _copy(src_cell.alignment)
                        dst_cell.number_format = src_cell.number_format
                        dst_cell.border        = _copy(src_cell.border)
                    except Exception:
                        pass
            if sr in ws.row_dimensions:
                ws_targets.row_dimensions[dst_row].height = ws.row_dimensions[sr].height
            dst_row += 1

        # Copy column widths
        for col_letter, col_dim in ws.column_dimensions.items():
            ws_targets.column_dimensions[col_letter].width = col_dim.width

        # Copy merged cells within the moved range
        for merged in list(ws.merged_cells.ranges):
            if merged.min_row >= move_start and merged.max_row <= move_end:
                offset = move_start - 1
                try:
                    ws_targets.merge_cells(
                        start_row=merged.min_row - offset, start_column=merged.min_col,
                        end_row=merged.max_row - offset,   end_column=merged.max_col)
                except Exception:
                    pass

        ws_targets.freeze_panes = "D3"

        # Remove the moved rows from the main sheet (delete from bottom up not needed —
        # delete_rows handles the range in one call)
        ws.delete_rows(move_start, move_end - move_start + 1)

    # Force Excel to recalculate all formulas when the file is opened
    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True

    # Post-process: ensure every cell with a string starting "=" is stored as a formula
    # (openpyxl sometimes stores formula strings as text data_type="s" instead of "f")
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    cell.data_type = "f"

    # ── Insert "This file is..." note row at top of Wealth Summary ─────────────
    import re as _re
    for sheet in wb.worksheets:
        # Capture merge ranges BEFORE insert (insert_rows does not shift these)
        old_merges = [str(m) for m in sheet.merged_cells.ranges]
        for m in list(sheet.merged_cells.ranges):
            try:
                sheet.unmerge_cells(str(m))
            except KeyError:
                sheet.merged_cells.ranges.discard(m)

        sheet.insert_rows(1)

        # Formula text needs manual row-number adjustment (insert_rows shifts
        # cell positions but not the formula strings themselves)
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    def _bump(m):
                        return m.group(1) + str(int(m.group(2)) + 1)
                    cell.value = _re.sub(r"([A-Z]{1,3})(\d+)", _bump, cell.value)
                    cell.data_type = "f"

        # Re-apply merges, shifted down by 1 row
        for m_str in old_merges:
            # m_str like "A47:O47"
            match = _re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", m_str)
            if match:
                c1, r1, c2, r2 = match.groups()
                from openpyxl.utils import column_index_from_string
                sheet.merge_cells(start_row=int(r1)+1, start_column=column_index_from_string(c1),
                                  end_row=int(r2)+1, end_column=column_index_from_string(c2))

        # Shift freeze panes down by 1 row (D3 -> D4)
        if sheet.freeze_panes:
            fp = sheet.freeze_panes
            col_part = ''.join(c for c in fp if c.isalpha())
            row_part = ''.join(c for c in fp if c.isdigit())
            if row_part:
                sheet.freeze_panes = f"{col_part}{int(row_part)+1}"

    # Write the note row
    note_text = "This file is 'Spending Summary · XLSX' from downloads"
    for sheet in wb.worksheets:
        c = sheet.cell(row=1, column=1, value=note_text)
        c.font = Font(italic=True, color="7F4000", size=10, name="Arial")
        c.fill = PatternFill("solid", fgColor="FFF3CD")
        c.alignment = Alignment(horizontal="left", vertical="center")
        sheet.row_dimensions[1].height = 16
        # Extend yellow shading across all columns and merge
        for col in range(2, sheet.max_column + 1):
            sheet.cell(row=1, column=col).fill = PatternFill("solid", fgColor="FFF3CD")
        sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=sheet.max_column)

    preserve_manual_sheets(wb, output_path)
    wb.save(output_path)


def copy_sheet_into(src_ws, dst_ws):
    """Cell-by-cell copy of values + styles + merges + dimensions between
    workbooks (openpyxl has no cross-workbook copy_worksheet)."""
    from copy import copy as _copy
    for row in src_ws.iter_rows():
        for cell in row:
            if cell.__class__.__name__ == 'MergedCell':
                continue
            dst = dst_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                dst.font = _copy(cell.font)
                dst.fill = _copy(cell.fill)
                dst.border = _copy(cell.border)
                dst.alignment = _copy(cell.alignment)
                dst.number_format = cell.number_format
                dst.protection = _copy(cell.protection)
    for m in src_ws.merged_cells.ranges:
        dst_ws.merge_cells(str(m))
    for col, dim in src_ws.column_dimensions.items():
        dst_ws.column_dimensions[col].width = dim.width
    for r, dim in src_ws.row_dimensions.items():
        dst_ws.row_dimensions[r].height = dim.height


def preserve_manual_sheets(new_wb, output_path):
    """This builder regenerates the whole workbook every run — but the user
    keeps hand-maintained tabs (e.g. 'Payslip Summary', 'Retirement Income
    Plan') alongside the generated ones, and a rebuild used to silently drop
    them (lost tabs reported 2026-07-12, restored from the 2026-07-02 copy).
    Carry over every sheet in the existing file whose name this run didn't
    generate, so manual tabs survive rebuilds."""
    if not os.path.exists(output_path):
        return
    try:
        old_wb = load_workbook(output_path)
    except Exception as e:
        print(f"WARNING: could not read existing {output_path} to preserve "
              f"manual tabs ({e}) — generated tabs only this run.", file=sys.stderr)
        return
    generated = set(new_wb.sheetnames)
    for name in old_wb.sheetnames:
        if name in generated:
            continue
        copy_sheet_into(old_wb[name], new_wb.create_sheet(name))
        print(f"Preserved manual tab: {name}")


# ── Main ───────────────────────────────────────────────────────────────────────
def main():
    args = sys.argv[1:]
    amex_path    = args[0] if len(args) > 0 else "activity.csv"
    bar_path     = args[1] if len(args) > 1 else "data.csv"
    fid_path     = args[2] if len(args) > 2 else "TransactionHistory.csv"
    summary_path = args[3] if len(args) > 3 else "AccountSummary.csv"
    output_path  = args[4] if len(args) > 4 else "spending_summary.xlsx"
    # Optional: a pending-orders export (every row has Completion date ==
    # 'Pending') — net units from pending Buy/Sell orders are layered on top
    # of the settled holdings figure. Not required; omit to skip.
    pending_path = args[5] if len(args) > 5 else None

    for path in (amex_path, bar_path, fid_path):
        if not os.path.exists(path):
            print(f"Error: file not found — {path}")
            sys.exit(1)

    print(f"  Amex:     {amex_path}")
    print(f"  Barclays: {bar_path}")
    print(f"  Fidelity: {fid_path}")
    if pending_path:
        print(f"  Pending:  {pending_path}")
    print(f"  Output:   {output_path}\n")

    print("Loading Amex...")
    amex_df = load_amex(amex_path)
    print(f"  {len(amex_df)} transactions\n")

    print("Loading Barclays...")
    bar_df = load_barclays(bar_path)
    print(f"  {len(bar_df)} transactions\n")

    print("Loading Fidelity income...")
    fid_df = load_fidelity_income(fid_path)
    print(f"  {len(fid_df)} income entries\n")

    print("Building holdings...")
    settled_holdings = build_holdings(fid_path, summary_path)
    holdings = build_holdings(fid_path, summary_path, pending_path)
    src = "AccountSummary.csv" if os.path.exists(summary_path) else "transaction history"
    print(f"  {len(holdings)} positions (from {src})")
    if pending_path and os.path.exists(pending_path):
        changed = sorted(k for k in set(settled_holdings) | set(holdings)
                          if round(settled_holdings.get(k, 0), 2) != round(holdings.get(k, 0), 2))
        if changed:
            print(f"  {len(changed)} position(s) adjusted by pending orders:")
            for acc, fund in changed:
                before = settled_holdings.get((acc, fund), 0)
                after = holdings.get((acc, fund), 0)
                print(f"    {acc} — {fund}: {before:,.2f} -> {after:,.2f}")
        else:
            print("  No pending Buy/Sell orders affected current holdings")
    print()

    print("Building pivots...")
    spend_pivot, spend_months       = build_spending_pivot(amex_df, bar_df)
    fid_pivot, fid_months           = build_fidelity_pivot(fid_df)
    acc_fund_map, acc_fund_months   = build_account_fund_pivot(fid_df)

    # Always show full Jan–Dec 2026 — transaction files may only cover partial year
    # but history fills Jan–Apr and projections cover future months
    JAN_2026 = pd.Period("2026-01", "M")
    DEC_2026 = pd.Period("2026-12", "M")
    all_months = [JAN_2026 + i for i in range(12)]

    # Actual months = those covered by transaction files (spend or fidelity data)
    tx_months = sorted(set(spend_months) | set(fid_months))
    MAY_2026 = pd.Period("2026-05", "M")
    actual_months = [m for m in tx_months if m != MAY_2026]
    partial_months = [m for m in tx_months if m == MAY_2026]

    spend_months = fid_months = acc_fund_months = all_months
    spend_pivot = spend_pivot.reindex(columns=all_months + ["Total"], fill_value=0)
    if not fid_pivot.empty:
        fid_pivot = fid_pivot.reindex(columns=all_months + ["Total"], fill_value=0)
        fid_pivot["Total"] = fid_pivot[all_months].sum(axis=1)
    for person in acc_fund_map:
        for acc in acc_fund_map[person]:
            df = acc_fund_map[person][acc]
            df = df.reindex(columns=all_months + ["Total"], fill_value=0)
            df["Total"] = df[all_months].sum(axis=1)
            acc_fund_map[person][acc] = df

    # Future months = all months not covered by transaction files
    last_tx = max(tx_months) if tx_months else MAY_2026
    future_months = partial_months[:]  # May first (partial, scale up)
    m = last_tx + 1
    while m <= DEC_2026:
        future_months.append(m)
        m += 1

    # Also add any months in all_months before first tx month that aren't partial
    print(f"  Actuals: {len(actual_months)} months ({actual_months[0] if actual_months else 'none'} – {actual_months[-1] if actual_months else 'none'})")
    print(f"  Estimating: {len(future_months)} months ({future_months[0] if future_months else 'none'} – {future_months[-1] if future_months else 'none'})")

    # Full year Jan–Dec (actuals + future, always 12 months)
    full_months = all_months

    # ── Inject missing stocks BEFORE estimation ───────────────────────────────
    # Stocks held (from AccountSummary) but with no 2026 income transactions
    # Uses confirmed dividend pence-per-share × units for accurate estimates
    STOCK_INJECT = {
        "LEGAL & GENERAL GROUP,ORD GBP0.025(LGEN)": [
            (6, 15.36),   # Final: pay Jun 4 (ex Apr 23)
            (9, 6.12),    # Interim: pay Sep 25 (ex Aug 20)
        ],
    }
    FUND_NAME_MAP = {
        "LEGAL & GENERAL GROUP,ORD GBP0.025(LGEN)": "LEGAL & GENERAL GROUP, ORD GBP0.025 (LGEN)",
    }
    # Track injected fund names so estimate_future_months can skip them
    injected_funds = set()

    for person in acc_fund_map:
        for acc in list(acc_fund_map[person].keys()):
            for summary_name, payments in STOCK_INJECT.items():
                income_name = FUND_NAME_MAP.get(summary_name, summary_name)
                fund_df = acc_fund_map[person][acc]
                if income_name in fund_df.index and fund_df.loc[income_name].sum() > 0:
                    continue
                units = holdings.get((acc, summary_name), 0)
                if units <= 0:
                    continue
                new_row = {m: 0 for m in full_months + ["Total"]}
                for pay_month, ppm in payments:
                    for m in full_months:
                        if m.month == pay_month and m.year == 2026:
                            new_row[m] = round(units * ppm / 100)
                new_row["Total"] = sum(new_row[m] for m in full_months)
                new_df = pd.DataFrame([new_row], index=[income_name])
                new_df.index.name = "fund"
                fund_df = fund_df.reindex(columns=full_months + ["Total"], fill_value=0)
                acc_fund_map[person][acc] = pd.concat([fund_df, new_df]).sort_index()
                injected_funds.add(income_name)
                print(f"  Injected {income_name} → {acc} ({units:,.0f} units, "
                      f"Jun=£{round(units*payments[0][1]/100):,}, "
                      f"Sep=£{round(units*payments[1][1]/100):,})")

    # Apply estimates to all pivots (injected funds are skipped)
    spend_pivot    = estimate_future_months(spend_pivot, actual_months, future_months)
    if not fid_pivot.empty:
        fid_pivot  = estimate_future_months(fid_pivot, actual_months, future_months)
    for person in acc_fund_map:
        for acc in acc_fund_map[person]:
            acc_fund_map[person][acc] = estimate_future_months(
                acc_fund_map[person][acc], actual_months, future_months,
                skip_funds=injected_funds)

    # ── June provisional estimate fallback ────────────────────────────────────
    # June 2026 TransactionHistory only covers the last 60 days (from ~mid-Apr),
    # so some monthly fund distributions paid later in June aren't captured yet.
    # Fall back to the May value as a placeholder estimate — flagged for revision
    # once the rest of June's transaction data is provided.
    JUN_2026 = pd.Period("2026-06", "M")
    MAY_2026_fb = pd.Period("2026-05", "M")
    june_estimated_funds = set()
    for person in acc_fund_map:
        for acc, fund_df in acc_fund_map[person].items():
            if JUN_2026 in fund_df.columns and MAY_2026_fb in fund_df.columns:
                for fund in fund_df.index:
                    if fund in injected_funds:
                        continue
                    jun_val = fund_df.loc[fund, JUN_2026]
                    may_val = fund_df.loc[fund, MAY_2026_fb]
                    if (pd.isna(jun_val) or jun_val == 0) and may_val and may_val > 0:
                        fund_df.loc[fund, JUN_2026] = may_val
                        june_estimated_funds.add((acc, fund))
    if june_estimated_funds:
        print(f"  June provisional estimates applied to {len(june_estimated_funds)} fund rows (using May values)")

    # June Salary estimate — payslip not yet received, use Apr (normal monthly)
    # since May included a one-off £42,432 "UK EXCL CM AND W" payment
    APR_2026_fb = pd.Period("2026-04", "M")
    if "Salary" in spend_pivot.index and JUN_2026 in spend_pivot.columns and APR_2026_fb in spend_pivot.columns:
        jun_sal = spend_pivot.loc["Salary", JUN_2026]
        apr_sal = spend_pivot.loc["Salary", APR_2026_fb]
        if (pd.isna(jun_sal) or jun_sal == 0) and apr_sal and apr_sal > 0:
            spend_pivot.loc["Salary", JUN_2026] = apr_sal
            print(f"  June Salary estimated as £{apr_sal:,.0f} (Apr value — May included one-off £42,432 payment)")

    # Rebuild fid_pivot account totals after estimation + injection
    for person in acc_fund_map:
        for acc in acc_fund_map[person]:
            fund_df = acc_fund_map[person][acc]
            acc_total = fund_df[full_months].sum()
            if acc in fid_pivot.index:
                for m in full_months:
                    # Only overwrite months NOT covered by hardcoded income history
                    # History covers Jan–May 2026 (HIST_CUTOFF = 2026-06)
                    hist_cutoff = pd.Period("2026-06", "M")
                    if m >= hist_cutoff:
                        fid_pivot.loc[acc, m] = acc_total[m]
                # Always update Total
                fid_pivot.loc[acc, "Total"] = fid_pivot.loc[acc, full_months].sum()

    all_months = full_months
    spend_months = fid_months = acc_fund_months = all_months

    print(f"  Fidelity: {len(fid_pivot)} accounts\n")

    print("Writing Excel...")
    # Actual months for write_excel = Jan–May (history + transaction file months)
    # Future months = Jun–Dec
    JAN_2026 = pd.Period("2026-01", "M")
    MAY_2026_m = pd.Period("2026-05", "M")
    DEC_2026_m = pd.Period("2026-12", "M")
    actual_for_excel = [JAN_2026 + i for i in range(6)]   # Jan–Jun (AccountSummary 10 Jun 2026)
    future_for_excel = [JAN_2026 + i for i in range(6, 12)]  # Jul–Dec
    full_months = actual_for_excel + future_for_excel  # Jan–Dec

    summary_data = build_summary_data(summary_path, full_months)
    acc_holdings = build_acc_holdings(summary_path, fid_path, inc_income_df=fid_df)
    # Expense Reimbursements: Royal Mail or Expleo credits under £1,000 (excl. salary)
    REIMBURSEMENT_KEYWORDS = ["ROYAL MAIL", "EXPLEO"]

    # One-time backfill from expenses.csv (Jan-Jun 2026 full year feed) — these
    # months are no longer covered by the live data.csv (60-day window), so
    # hardcode them here. Going forward, data.csv provides ongoing updates.
    REIMBURSEMENT_BACKFILL = [
        {"date": pd.Timestamp("2026-02-04"), "amount": 36.80,
         "memo": "ROYAL MAIL            \t1612 2000284071 K BGC"},
        {"date": pd.Timestamp("2026-03-06"), "amount": 68.24,
         "memo": "EXPLEO UK LIMITED     \t1160304622 BGC\t"},
        {"date": pd.Timestamp("2026-03-20"), "amount": 163.03,
         "memo": "EXPLEO UK LIMITED     \t1161123617 BGC\t"},
        {"date": pd.Timestamp("2026-04-07"), "amount": 342.97,
         "memo": "EXPLEO UK LIMITED     \t1162170846 BGC\t"},
    ]

    reimbursements = []
    seen_keys = set()
    for entry in REIMBURSEMENT_BACKFILL:
        key = (entry["date"].date(), round(entry["amount"], 2))
        if key not in seen_keys:
            seen_keys.add(key)
            reimbursements.append({
                "date": entry["date"],
                "month": entry["date"].to_period("M"),
                "amount": entry["amount"],
                "memo": entry["memo"].strip()
            })

    # Live extraction from data.csv — covers ongoing/future updates
    try:
        bar_df_r = pd.read_csv("data.csv")
        bar_df_r["Amount"] = pd.to_numeric(bar_df_r["Amount"], errors="coerce").fillna(0)
        bar_df_r["Date"] = pd.to_datetime(bar_df_r["Date"], dayfirst=True, errors="coerce")
        bar_df_r = bar_df_r.dropna(subset=["Date"])
        for _, r in bar_df_r.iterrows():
            memo = str(r.get("Memo", "")).upper()
            amt = r["Amount"]
            if amt > 0 and amt < 1000 and any(k in memo for k in REIMBURSEMENT_KEYWORDS):
                if "EUKPT" in memo:
                    continue
                key = (r["Date"].date(), round(amt, 2))
                if key not in seen_keys:
                    seen_keys.add(key)
                    reimbursements.append({
                        "date": r["Date"],
                        "month": r["Date"].to_period("M"),
                        "amount": amt,
                        "memo": str(r.get("Memo", "")).strip()
                    })
    except Exception:
        pass

    write_excel(spend_pivot, actual_for_excel, future_for_excel, fid_pivot,
                acc_fund_map, holdings, summary_data, acc_holdings, output_path,
                reimbursements=reimbursements)
    print(f"Done → {output_path}\n")

    # Console preview of Fidelity section
    if not fid_pivot.empty:
        col_w, num_w = 32, 10
        hdr = f"{'Account':<{col_w}} {'Total':>{num_w}}" + "".join(
            f"  {m.strftime('%b %y'):>{num_w}}" for m in fid_months)
        print("Fidelity Income by Account")
        print("-" * len(hdr))
        print(hdr)
        print("-" * len(hdr))
        for acc in fid_pivot.index:
            label = ACCOUNT_LABELS.get(acc, acc)
            total = int(round(fid_pivot.loc[acc, "Total"]))
            monthly = [int(round(fid_pivot.loc[acc, m])) for m in fid_months]
            vals = f"{total:>{num_w},}" + "".join(f"  {v:>{num_w},}" for v in monthly)
            print(f"{label:<{col_w}} {vals}")
        gtotal = int(round(fid_pivot["Total"].sum()))
        gmonthly = [int(round(sum(fid_pivot[m]))) for m in fid_months]
        print("-" * len(hdr))
        gvals = f"{gtotal:>{num_w},}" + "".join(f"  {v:>{num_w},}" for v in gmonthly)
        print(f"{'TOTAL':<{col_w}} {gvals}")


if __name__ == "__main__":
    main()
