"""The Wealth Summary's visual palette: colours, the two factories that build
fills and fonts (P, F), the named fills/fonts themselves, the fml formula helper
and HIST_MAP.

Hoisted out of write_excel on 2026-07-19 — these are pure constants with no
dependence on the run's data, so leaving them inside the function would have
forced every extracted phase to take twenty-odd style arguments. Moved verbatim.
"""
from openpyxl.styles import Font, PatternFill

NUM_FMT = '#,##0;(#,##0);"-"'

def fml(cell, formula):
    """Write an Excel formula to a cell, ensuring it is stored as a formula not text."""
    cell.data_type = "f"
    cell.value = formula
    cell.number_format = '#,##0;(#,##0);"-"'

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


# These two depend on the RUN's data (which months are actual vs estimated), so
# unlike the palette above they can't be plain constants. Factories keep the call
# sites in the phase modules unchanged — col_fill(m, is_alt, base, base_alt) and
# val_font(m, bold, sal) still take exactly the arguments they always did.
def make_col_fill(actual_set):
    def col_fill(m, is_alt, base_fill, base_alt):
        """Return appropriate fill for actual vs estimated column."""
        if m in actual_set:
            return base_alt if is_alt else None
        else:
            return EST_ALT if is_alt else EST_FILL
    return col_fill


def make_val_font(actual_set):
    def val_font(m, bold=False, sal=False):
        if sal: return SAL_FONT
        if m not in actual_set: return EST_BODY
        return BOLD_FONT if bold else BODY_FONT
    return val_font
