#!/usr/bin/env python3
"""
Builds a protected, self-calculating Excel workbook for a bank loan account with
quarterly compounding of interest and up to three effective-date-driven interest
rates (original rate + 2 revisions).

Design principles
-----------------
1. Interest accrues on the DAILY OUTSTANDING BALANCE (actual days / basis) and is
   CAPITALISED at every quarter end. The capitalised closing balance becomes the
   balance brought forward ("B/F") of the next quarter -> quarterly compounding.
2. A rate revision is entered with an EFFECTIVE-FROM date. Every accrual segment is
   split across the rate periods it overlaps, so a revision can never reach back and
   recompute a period that ended before its effective date. Past quarters are
   therefore arithmetically frozen.
3. Every cell the user may type into is unlocked and pale yellow. Everything else is
   locked by sheet protection (password) and the calculation engine lives on a
   very-hidden, protected worksheet. Workbook structure is locked as well.
4. A "Locked / reconciled up to" date acts as a hard cut-off: data validation refuses
   any transaction date on/before it and refuses a rate revision dated on/before it.

Usage:
    python3 build_loan_workbook.py <output.xlsx> [--example]
"""

from __future__ import annotations

import datetime as dt
import sys

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.workbook.protection import WorkbookProtection
from openpyxl.worksheet.datavalidation import DataValidation

# --------------------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------------------
PASSWORD = "MMS-loan-2026"          # sheet / workbook protection password
QUARTERS = 24                        # number of quarter blocks generated (6 years)
ENTRY_ROWS = 10                      # repayment input rows available in each quarter

SHEET_MAIN = "Loan Statement"
SHEET_SUM = "Quarter Summary"
SHEET_HELP = "How to Use"
SHEET_ENG = "Engine"

# Ledger geometry ----------------------------------------------------------------------
ROW_TITLE = 1
ROW_RATE_TITLE = 2
ROW_LEGEND = 3
ROW_PANEL_HDR = 5
ROW_P1 = 6                           # first row of the three setup panels
ROW_P_LAST = 11                      # last row of the setup panels
ROW_TBL_HDR = 13                     # ledger column-header row
ROW_LEDGER = 14                      # first ledger row
BLOCK = 2 + ENTRY_ROWS + 1           # qtr header + B/F + entries + quarter-end row
LAST_ROW = ROW_LEDGER + QUARTERS * BLOCK - 1

# Columns ------------------------------------------------------------------------------
C_SL, C_DATE, C_INST, C_PART = 1, 2, 3, 4
C_DEBIT, C_CREDIT, C_BAL, C_RATE = 5, 6, 7, 8
C_DAYS, C_INT, C_ACCR, C_BANK = 9, 10, 11, 12
C_VAR, C_STAT = 13, 14
C_EFF, C_QS, C_QE, C_RTXT = 15, 16, 17, 18      # hidden helper columns
C_BALH, C_ACCRH = 19, 20                         # hidden running balance / accrual chain

L = {c: get_column_letter(c) for c in range(1, 21)}

# Formats ------------------------------------------------------------------------------
DATE_FMT = "DD-MM-YYYY"
DATE_Z = "DD-MM-YYYY;;;"        # blank when the cell is 0 (setup not done yet)
MONEY = '#,##0.00;[Red]-#,##0.00;"\u2013"'
MONEY_T = '#,##0.00;[Red]-#,##0.00'
PCT = '0.00"%";[Red]-0.00"%";"\u2013"'
DAYS_FMT = '#,##0;;"\u2013"'

# Colours ------------------------------------------------------------------------------
CLR_INPUT = "FFF2CC"        # pale yellow  -> typeable
CLR_CALC = "EAF1F8"         # pale blue    -> calculated & locked
CLR_TITLE = "1F3864"
CLR_HDR = "1F4E78"
CLR_QBAR = "2E75B6"
CLR_BF = "E2EFDA"
CLR_INT = "FCE4D6"
CLR_LOCK = "D0CECE"
CLR_PANEL = "305496"
CLR_NOTE = "FFF9E6"

F_WHITE_B = Font(bold=True, color="FFFFFF", size=10)
F_TITLE = Font(bold=True, color="FFFFFF", size=14)
F_SUB = Font(bold=True, color="FFFFFF", size=10)
F_LBL = Font(bold=True, size=9, color="1F3864")
F_BODY = Font(size=10)
F_BODY_B = Font(bold=True, size=10)
F_SMALL = Font(size=8, italic=True, color="595959")

THIN = Side(style="thin", color="BFBFBF")
MED = Side(style="medium", color="BF8F00")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BOX_IN = Border(left=MED, right=MED, top=MED, bottom=MED)

CENTER = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center")
RIGHT = Alignment(horizontal="right", vertical="center")
WRAP_C = Alignment(horizontal="center", vertical="center", wrap_text=True)
WRAP_L = Alignment(horizontal="left", vertical="top", wrap_text=True)

UNLOCKED = Protection(locked=False)
LOCKED = Protection(locked=True)

# Setup-panel input cells (named ranges point here) ------------------------------------
CELL_BF_DATE = "H6"
CELL_BF_AMT = "H7"
CELL_BASIS = "H8"
CELL_QMODE = "H9"
CELL_LOCKED = "H10"
CELL_R1_DATE, CELL_R1_PCT = "K7", "L7"
CELL_R2_DATE, CELL_R2_PCT = "K8", "L8"
CELL_R3_DATE, CELL_R3_PCT = "K9", "L9"

QMODE_CAL = "Calendar quarter-ends"
QMODE_ANN = "3 months from B/F date"


# --------------------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------------------
def rate_factor(d1: str, d2: str) -> str:
    """Excel expression: fraction of a year's interest earned between dates d1 and d2,
    split across the (up to) three rate periods that the interval overlaps."""
    parts = []
    for n in (1, 2, 3):
        parts.append(
            f"Rate_{n}*MAX(0,MIN({d2},R{n}_To)-MAX({d1},R{n}_From))"
        )
    return "(" + "+".join(parts) + ")/100/Day_Basis"


def qtr_rows(q: int) -> tuple[int, int, int, int, int]:
    """(header row, B/F row, first entry row, last entry row, quarter-end row)"""
    start = ROW_LEDGER + (q - 1) * BLOCK
    return start, start + 1, start + 2, start + 1 + ENTRY_ROWS, start + 2 + ENTRY_ROWS


def put(ws, row, col, value=None, *, fmt=None, font=None, fill=None, align=None,
        border=BOX, unlocked=False):
    c = ws.cell(row=row, column=col)
    if value is not None:
        c.value = value
    if fmt:
        c.number_format = fmt
    c.font = font or F_BODY
    if fill:
        c.fill = PatternFill("solid", fgColor=fill)
    if align:
        c.alignment = align
    if border:
        c.border = border
    c.protection = UNLOCKED if unlocked else LOCKED
    return c


def inp(ws, row, col, **kw):
    kw.setdefault("fill", CLR_INPUT)
    kw.setdefault("border", BOX_IN)
    kw["unlocked"] = True
    return put(ws, row, col, **kw)


def calc(ws, row, col, formula, **kw):
    kw.setdefault("fill", CLR_CALC)
    return put(ws, row, col, formula, **kw)


# --------------------------------------------------------------------------------------
# Engine sheet (very hidden): rate periods, quarter calendar, live totals
# --------------------------------------------------------------------------------------
def build_engine(wb):
    ws = wb.create_sheet(SHEET_ENG)
    M = f"'{SHEET_MAIN}'!"

    rows = [
        ("Big_Date", "=DATE(2999,12,31)"),
        ("R1_From", f"={M}${CELL_R1_DATE[0]}${CELL_R1_DATE[1:]}"),
        ("R1_To", f'=IF({M}${CELL_R2_DATE[0]}${CELL_R2_DATE[1:]}="",Big_Date,{M}${CELL_R2_DATE[0]}${CELL_R2_DATE[1:]})'),
        ("R2_From", f'=IF({M}${CELL_R2_DATE[0]}${CELL_R2_DATE[1:]}="",Big_Date,{M}${CELL_R2_DATE[0]}${CELL_R2_DATE[1:]})'),
        ("R2_To", f'=IF({M}${CELL_R3_DATE[0]}${CELL_R3_DATE[1:]}="",Big_Date,{M}${CELL_R3_DATE[0]}${CELL_R3_DATE[1:]})'),
        ("R3_From", f'=IF({M}${CELL_R3_DATE[0]}${CELL_R3_DATE[1:]}="",Big_Date,{M}${CELL_R3_DATE[0]}${CELL_R3_DATE[1:]})'),
        ("R3_To", "=Big_Date"),
        ("Rate_1", f'=IF({M}${CELL_R1_PCT[0]}${CELL_R1_PCT[1:]}="",0,{M}${CELL_R1_PCT[0]}${CELL_R1_PCT[1:]})'),
        ("Rate_2", f'=IF({M}${CELL_R2_PCT[0]}${CELL_R2_PCT[1:]}="",0,{M}${CELL_R2_PCT[0]}${CELL_R2_PCT[1:]})'),
        ("Rate_3", f'=IF({M}${CELL_R3_PCT[0]}${CELL_R3_PCT[1:]}="",0,{M}${CELL_R3_PCT[0]}${CELL_R3_PCT[1:]})'),
    ]
    for i, (name, formula) in enumerate(rows, start=1):
        ws.cell(row=i, column=1, value=name).font = F_BODY
        ws.cell(row=i, column=2, value=formula)

    # live figures (display only - the ledger never depends on TODAY())
    live = [
        (12, "Bal_Today", "=IFERROR(LOOKUP(TODAY(),Eff_Dates,Bal_Col),BF_Amount)"),
        (13, "Last_Event", "=IFERROR(LOOKUP(TODAY(),Eff_Dates,Eff_Dates),BF_Date)"),
        (14, "Accr_Today",
         '=IF(OR(Day_Basis=0,Day_Basis=""),0,IFERROR(LOOKUP(TODAY(),Eff_Dates,Accr_Col),0)'
         "+$B$12*" + rate_factor("$B$13", "TODAY()") + ")"),
        (15, "Out_Today", "=$B$12+$B$14"),
        (16, "Total_Repaid", f"=SUM('{SHEET_MAIN}'!$F${ROW_LEDGER}:$F${LAST_ROW})"),
        (17, "Total_Interest", f"=SUM('{SHEET_MAIN}'!$E${ROW_LEDGER}:$E${LAST_ROW})"),
        (18, "Variance_Rows",
         f'=COUNTIF(\'{SHEET_MAIN}\'!$M${ROW_LEDGER}:$M${LAST_ROW},">0.004")'
         f'+COUNTIF(\'{SHEET_MAIN}\'!$M${ROW_LEDGER}:$M${LAST_ROW},"<-0.004")'),
    ]
    for r, name, formula in live:
        ws.cell(row=r, column=1, value=name).font = F_BODY
        ws.cell(row=r, column=2, value=formula)

    # quarter calendar: D=q, E=start, F=end, G=rate text
    for h, col in (("Q#", 4), ("Start", 5), ("End", 6), ("Rates in quarter", 7)):
        ws.cell(row=1, column=col, value=h).font = F_BODY_B

    # Until the B/F date is filled in, every quarter date stays 0 so the ledger shows
    # dashes instead of 1900-dates or errors.
    first_qe = (
        '=IF(BF_Date="",0,'
        f'IF(Qtr_Mode="{QMODE_CAL}",'
        "IF(EOMONTH(BF_Date,MOD(3-MOD(MONTH(BF_Date),3),3))=BF_Date,"
        "EOMONTH(BF_Date,3),EOMONTH(BF_Date,MOD(3-MOD(MONTH(BF_Date),3),3))),"
        "EDATE(BF_Date,3)))"
    )
    for q in range(1, QUARTERS + 1):
        r = q + 1
        ws.cell(row=r, column=4, value=q)
        if q == 1:
            ws.cell(row=r, column=5, value='=IF(BF_Date="",0,BF_Date)')
            ws.cell(row=r, column=6, value=first_qe)
        else:
            ws.cell(row=r, column=5, value=f"=$F${r - 1}")
            ws.cell(row=r, column=6,
                    value=f'=IF($F${r - 1}=0,0,IF(Qtr_Mode="{QMODE_CAL}",'
                          f"EOMONTH($F${r - 1},3),EDATE($F${r - 1},3)))")
        ws.cell(row=r, column=5).number_format = DATE_FMT
        ws.cell(row=r, column=6).number_format = DATE_FMT

        seg = []
        for n in (1, 2, 3):
            ov = f"MAX(0,MIN($F${r},R{n}_To)-MAX($E${r},R{n}_From))"
            seg.append(f'IF({ov}>0,TEXT(Rate_{n},"0.00")&"% for "&{ov}&'
                       f'IF({ov}=1," day   "," days   "),"")')
        ws.cell(row=r, column=7, value="=TRIM(" + "&".join(seg) + ")")

    ws.sheet_state = "veryHidden"
    ws.protection.password = PASSWORD
    ws.protection.sheet = True
    ws.protection.enable()
    return ws


# --------------------------------------------------------------------------------------
# Main ledger sheet
# --------------------------------------------------------------------------------------
def build_main(wb):
    ws = wb.active
    ws.title = SHEET_MAIN
    E = f"'{SHEET_ENG}'!"

    widths = {1: 6, 2: 12.5, 3: 15, 4: 44, 5: 14, 6: 14, 7: 16, 8: 10.5,
              9: 7.5, 10: 14, 11: 15, 12: 15, 13: 12, 14: 12.5}
    for col, w in widths.items():
        ws.column_dimensions[L[col]].width = w
    for col in (C_EFF, C_QS, C_QE, C_RTXT, C_BALH, C_ACCRH):
        ws.column_dimensions[L[col]].hidden = True

    # ---- title rows ------------------------------------------------------------------
    ws.merge_cells(start_row=ROW_TITLE, start_column=1, end_row=ROW_TITLE, end_column=C_STAT)
    t = put(ws, ROW_TITLE, 1, "BANK LOAN ACCOUNT  \u2013  QUARTERLY COMPOUND INTEREST STATEMENT",
            font=F_TITLE, fill=CLR_TITLE, align=CENTER, border=None)
    ws.row_dimensions[ROW_TITLE].height = 30

    rate_title = (
        '=IF(OR(Rate_1=0,BF_Date=""),'
        '"\u25b2  START HERE :  fill the Balance B/F date & amount and the original interest '
        'rate in the panel below  \u25bc",'
        '"INTEREST RATE ON THIS FILE :   "&TEXT(Rate_1,"0.00")&"% p.a. w.e.f. "'
        '&TEXT(R1_From,"DD-MM-YYYY")'
        '&IF(Rate_2>0,"      |      "&TEXT(Rate_2,"0.00")&"% p.a. w.e.f. "&TEXT(R2_From,"DD-MM-YYYY"),"")'
        '&IF(Rate_3>0,"      |      "&TEXT(Rate_3,"0.00")&"% p.a. w.e.f. "&TEXT(R3_From,"DD-MM-YYYY"),"")'
        '&"            BASIS : actual days / "&Day_Basis&"   \u00b7   COMPOUNDED QUARTERLY")'
    )
    ws.merge_cells(start_row=ROW_RATE_TITLE, start_column=1, end_row=ROW_RATE_TITLE, end_column=C_STAT)
    put(ws, ROW_RATE_TITLE, 1, rate_title, font=F_SUB, fill=CLR_QBAR, align=CENTER, border=None)
    ws.row_dimensions[ROW_RATE_TITLE].height = 20

    legend = ('\u25a0 PALE YELLOW + gold border = the ONLY cells you can type in.      '
              '\u25a0 PALE BLUE / WHITE = locked formulas.      '
              '\u25a0 GREY = period locked & reconciled (no entry possible).      '
              'Interest is charged on the daily outstanding balance and capitalised every quarter-end.')
    ws.merge_cells(start_row=ROW_LEGEND, start_column=1, end_row=ROW_LEGEND, end_column=C_STAT)
    put(ws, ROW_LEGEND, 1, legend, font=Font(size=9, italic=True, color="7F6000"),
        fill=CLR_NOTE, align=LEFT, border=None)
    ws.row_dimensions[ROW_LEGEND].height = 16
    ws.row_dimensions[4].height = 6

    # ---- panel headers ---------------------------------------------------------------
    for c1, c2, text in ((1, 5, "LOAN ACCOUNT PARTICULARS"),
                         (6, 9, "LOAN BASIS  &  LOCK CONTROL"),
                         (10, 14, "INTEREST RATE REGISTER  \u2013  3 SLOTS, EFFECTIVE-DATE DRIVEN")):
        ws.merge_cells(start_row=ROW_PANEL_HDR, start_column=c1, end_row=ROW_PANEL_HDR, end_column=c2)
        put(ws, ROW_PANEL_HDR, c1, text, font=F_WHITE_B, fill=CLR_PANEL, align=CENTER)

    # ---- left panel: account particulars ---------------------------------------------
    left = [("Loan A/C No.", None), ("Borrower name", None), ("Type of facility", None),
            ("Address", None), ("Sanctioned limit (BDT)", MONEY_T)]
    for i, (label, fmt) in enumerate(left):
        r = ROW_P1 + i
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        put(ws, r, 1, label, font=F_LBL, align=LEFT)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
        inp(ws, r, 3, fmt=fmt, align=RIGHT if fmt else LEFT)
        for c in (4, 5):
            put(ws, r, c, fill=CLR_INPUT, border=BOX_IN, unlocked=True)

    r = ROW_P_LAST
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    put(ws, r, 1, "Statement period", font=F_LBL, align=LEFT)
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
    calc(ws, r, 3,
         f'=IF(BF_Date="","\u2013 not set \u2013",TEXT(BF_Date,"DD-MM-YYYY")&"  to  "'
         f'&TEXT({E}$F${QUARTERS + 1},"DD-MM-YYYY"))',
         align=LEFT)
    for c in (4, 5):
        put(ws, r, c, fill=CLR_CALC)

    # ---- middle panel: basis & lock control ------------------------------------------
    mid = [
        ("Balance B/F date", CELL_BF_DATE, DATE_FMT),
        ("Balance B/F amount (BDT)", CELL_BF_AMT, MONEY_T),
        ("Interest basis (days/year)", CELL_BASIS, "0"),
        ("Quarter-end convention", CELL_QMODE, None),
        ("Locked / reconciled up to", CELL_LOCKED, DATE_FMT),
    ]
    for i, (label, cell, fmt) in enumerate(mid):
        r = ROW_P1 + i
        ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=7)
        put(ws, r, 6, label, font=F_LBL, align=LEFT)
        ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=9)
        inp(ws, r, 8, fmt=fmt, align=CENTER)
        put(ws, r, 9, fill=CLR_INPUT, border=BOX_IN, unlocked=True)

    r = ROW_P_LAST
    ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=7)
    calc(ws, r, 6, '="Outstanding as on "&TEXT(TODAY(),"DD-MM-YYYY")&" (incl. accrued)"',
         font=F_LBL, align=LEFT)
    ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=9)
    calc(ws, r, 8, f"={E}$B$15", fmt=MONEY, font=F_BODY_B, align=CENTER)
    put(ws, r, 9, fill=CLR_CALC)

    # ---- right panel: rate register ---------------------------------------------------
    hdrs = ["Slot", "Effective from", "% per annum", "In force up to", "Lock status"]
    for i, h in enumerate(hdrs):
        put(ws, ROW_P1, 10 + i, h, font=Font(bold=True, size=9), fill="D9E1F2", align=WRAP_C)

    slots = [("Rate 1 \u2013 original", CELL_R1_DATE, CELL_R1_PCT),
             ("Rate 2 \u2013 1st revision", CELL_R2_DATE, CELL_R2_PCT),
             ("Rate 3 \u2013 2nd revision", CELL_R3_DATE, CELL_R3_PCT)]
    for i, (label, dcell, pcell) in enumerate(slots):
        r = ROW_P1 + 1 + i
        put(ws, r, 10, label, font=F_LBL, align=LEFT)
        if i == 0:
            calc(ws, r, 11, '=IF(BF_Date="",0,BF_Date)', fmt=DATE_Z, align=CENTER)
        else:
            inp(ws, r, 11, fmt=DATE_FMT, align=CENTER)
        inp(ws, r, 12, fmt='0.00"%"', align=CENTER)
        nxt = {0: CELL_R2_DATE, 1: CELL_R3_DATE, 2: None}[i]
        if nxt:
            calc(ws, r, 13,
                 f'=IF(${nxt[0]}${nxt[1:]}="","till date",TEXT(${nxt[0]}${nxt[1:]}-1,"DD-MM-YYYY"))',
                 fmt=None, align=CENTER)
        else:
            calc(ws, r, 13, f'=IF(${pcell[0]}${pcell[1:]}="","\u2013","till date")', align=CENTER)
        calc(ws, r, 14,
             f'=IF(${pcell[0]}${pcell[1:]}="","(slot free)",'
             f'IF(AND(Locked_To<>"",${dcell[0]}${dcell[1:]}<=Locked_To),"\U0001f512 LOCKED","EDITABLE"))',
             align=CENTER, font=Font(bold=True, size=9))

    for i, note in enumerate((
        "A revision may only be dated AFTER the \u2018Locked / reconciled up to\u2019 date. "
        "Interest already accrued under an earlier rate is never recalculated.",
        "Within a quarter that contains a revision date, interest is split day-by-day between "
        "the old and the new rate automatically.",
    )):
        r = ROW_P1 + 4 + i
        ws.merge_cells(start_row=r, start_column=10, end_row=r, end_column=14)
        put(ws, r, 10, note, font=F_SMALL, fill=CLR_NOTE, align=WRAP_L)
    ws.row_dimensions[ROW_P1 + 4].height = 22
    ws.row_dimensions[ROW_P1 + 5].height = 22
    ws.row_dimensions[12].height = 8

    # ---- ledger header ----------------------------------------------------------------
    heads = [
        (C_SL, "Sl."), (C_DATE, "Date"), (C_INST, "Cheque /\nInstrument"),
        (C_PART, "Particulars"), (C_DEBIT, "Debit\n(interest charged)"),
        (C_CREDIT, "Credit\n(repayment)"), (C_BAL, "Balance\n(outstanding)"),
        (C_RATE, "Rate\napplied"), (C_DAYS, "Days"), (C_INT, "Interest for\nthe period"),
        (C_ACCR, "Accrued interest\n(not yet charged)"), (C_BANK, "Balance\nper Bank"),
        (C_VAR, "Variance"), (C_STAT, "Status"),
    ]
    for col, text in heads:
        put(ws, ROW_TBL_HDR, col, text, font=F_WHITE_B, fill=CLR_HDR, align=WRAP_C)
    ws.row_dimensions[ROW_TBL_HDR].height = 34
    ws.freeze_panes = ws.cell(row=ROW_LEDGER, column=1)

    # ---- quarter blocks ---------------------------------------------------------------
    for q in range(1, QUARTERS + 1):
        hdr, bf, e1, e2, ir = qtr_rows(q)
        er = q + 1                                   # engine row for this quarter

        # hidden helpers on the header row
        put(ws, hdr, C_QS, f"={E}$E${er}", fmt=DATE_FMT, border=None)
        put(ws, hdr, C_QE, f"={E}$F${er}", fmt=DATE_FMT, border=None)
        put(ws, hdr, C_RTXT, f"={E}$G${er}", border=None)
        # keep the hidden date column strictly non-decreasing across blocks
        put(ws, hdr, C_EFF, f"={E}$E${er}", fmt=DATE_FMT, border=None)

        # quarter bar
        ws.merge_cells(start_row=hdr, start_column=1, end_row=hdr, end_column=C_STAT)
        put(ws, hdr, 1,
            f'=IF($P${hdr}=0,"QUARTER {q}          \u2013  enter the Balance B/F date above to '
            f'activate this quarter  \u2013",'
            f'"QUARTER {q}          "&TEXT($P${hdr},"DD-MM-YYYY")&"   to   "'
            f'&TEXT($Q${hdr},"DD-MM-YYYY")&"          "'
            f'&IF(AND(Locked_To<>"",$Q${hdr}<=Locked_To),'
            f'"\U0001f512 LOCKED \u2013 reconciled, no entry allowed","\u25b6 OPEN for entry")'
            f'&"          RATE(S) APPLIED :  "&$R${hdr})',
            font=F_WHITE_B, fill=CLR_QBAR, align=LEFT)
        ws.row_dimensions[hdr].height = 18

        # ---- balance brought forward row ----
        put(ws, bf, C_SL, "B/F", font=F_BODY_B, fill=CLR_BF, align=CENTER)
        put(ws, bf, C_DATE, f"=$P${hdr}", fmt=DATE_Z, font=F_BODY_B, fill=CLR_BF, align=CENTER)
        put(ws, bf, C_INST, fill=CLR_BF)
        part = ('=IF($B{r}=0,"","Opening balance brought forward as on "&TEXT($B{r},"DD-MM-YYYY"))'
                if q == 1 else
                '="Balance carried forward from quarter '
                + str(q - 1) + ' \u2013 interest capitalised"')
        put(ws, bf, C_PART, part.format(r=bf), font=F_BODY_B, fill=CLR_BF, align=LEFT)
        put(ws, bf, C_DEBIT, fill=CLR_BF, fmt=MONEY)
        put(ws, bf, C_CREDIT, fill=CLR_BF, fmt=MONEY)
        # NOTE: the row above a B/F row is the merged quarter bar, so the previous
        # quarter's capitalisation row is bf-2, not bf-1.
        put(ws, bf, C_BALH, "=BF_Amount" if q == 1 else f"=$S${bf - 2}",
            fmt=MONEY, border=None)
        put(ws, bf, C_ACCRH, 0, fmt=MONEY, border=None)
        put(ws, bf, C_BAL, f"=$S${bf}", fmt=MONEY, font=F_BODY_B, fill=CLR_BF, align=RIGHT)
        put(ws, bf, C_RATE, fill=CLR_BF, fmt=PCT)
        put(ws, bf, C_DAYS, 0 if q == 1 else f"=$O${bf}-$O${bf - 2}",
            fmt=DAYS_FMT, fill=CLR_BF, align=CENTER)
        put(ws, bf, C_INT, 0, fmt=MONEY, fill=CLR_BF, align=RIGHT)
        put(ws, bf, C_ACCR, 0, fmt=MONEY, fill=CLR_BF, align=RIGHT)
        inp(ws, bf, C_BANK, fmt=MONEY, align=RIGHT)
        put(ws, bf, C_VAR, f'=IF($L${bf}="","",ROUND($S${bf}-$L${bf},2))',
            fmt=MONEY, fill=CLR_BF, align=RIGHT)
        put(ws, bf, C_STAT, f'=IF(AND(Locked_To<>"",$O${bf}<=Locked_To),"LOCKED","OPEN")',
            fill=CLR_BF, align=CENTER, font=Font(size=8, bold=True))
        put(ws, bf, C_EFF, f"=$P${hdr}", fmt=DATE_FMT, border=None)

        # ---- repayment entry rows ----
        for r in range(e1, e2 + 1):
            calc(ws, r, C_SL, f'=IF($F{r}="","",COUNT($F${ROW_LEDGER}:$F{r}))',
                 fill=None, align=CENTER, font=Font(size=9))
            inp(ws, r, C_DATE, fmt=DATE_FMT, align=CENTER)
            inp(ws, r, C_INST, align=LEFT)
            inp(ws, r, C_PART, align=LEFT)
            put(ws, r, C_DEBIT, fmt=MONEY, fill=None)
            inp(ws, r, C_CREDIT, fmt=MONEY, align=RIGHT)
            put(ws, r, C_BALH, f'=$S{r - 1}-IF($F{r}="",0,$F{r})', fmt=MONEY, border=None)
            put(ws, r, C_ACCRH, f"=$T{r - 1}+$J{r}", fmt=MONEY, border=None)
            calc(ws, r, C_BAL, f'=IF($B{r}="","",$S{r})', fmt=MONEY, align=RIGHT)
            calc(ws, r, C_RATE,
                 f'=IF($I{r}<=0,"",{rate_factor(f"$O{r - 1}", f"$O{r}")}/$I{r}*Day_Basis*100)',
                 fmt=PCT, align=CENTER)
            calc(ws, r, C_DAYS, f"=$O{r}-$O{r - 1}", fmt=DAYS_FMT, align=CENTER)
            calc(ws, r, C_INT,
                 f'=IF($I{r}<=0,0,$S{r - 1}*{rate_factor(f"$O{r - 1}", f"$O{r}")})',
                 fmt=MONEY, align=RIGHT)
            calc(ws, r, C_ACCR, f'=IF($B{r}="","",$T{r})', fmt=MONEY, align=RIGHT)
            inp(ws, r, C_BANK, fmt=MONEY, align=RIGHT)
            calc(ws, r, C_VAR, f'=IF($L{r}="","",ROUND($S{r}-$L{r},2))', fmt=MONEY, align=RIGHT)
            calc(ws, r, C_STAT,
                 f'=IF(AND(Locked_To<>"",$O{r}<=Locked_To),"LOCKED","OPEN")',
                 align=CENTER, font=Font(size=8, bold=True))
            put(ws, r, C_EFF, f'=IF($B{r}="",$O{r - 1},$B{r})', fmt=DATE_FMT, border=None)

        # ---- quarter-end capitalisation row ----
        put(ws, ir, C_SL, "INT", font=F_BODY_B, fill=CLR_INT, align=CENTER)
        put(ws, ir, C_DATE, f"=$Q${hdr}", fmt=DATE_Z, font=F_BODY_B, fill=CLR_INT, align=CENTER)
        put(ws, ir, C_INST, fill=CLR_INT)
        put(ws, ir, C_PART,
            f'="Interest for the quarter debited & capitalised   \u00b7   rate(s) applied :  "&$R${hdr}',
            font=F_BODY_B, fill=CLR_INT, align=LEFT)
        put(ws, ir, C_DEBIT, f"=ROUND($T{ir - 1}+$J{ir},2)", fmt=MONEY,
            font=F_BODY_B, fill=CLR_INT, align=RIGHT)
        put(ws, ir, C_CREDIT, fmt=MONEY, fill=CLR_INT)
        put(ws, ir, C_BALH, f"=$S{ir - 1}+$E{ir}", fmt=MONEY, border=None)
        put(ws, ir, C_ACCRH, 0, fmt=MONEY, border=None)
        put(ws, ir, C_BAL, f"=$S{ir}", fmt=MONEY, font=F_BODY_B,
            fill=CLR_INT, align=RIGHT)
        put(ws, ir, C_RATE,
            f'=IF($I{ir}<=0,"",{rate_factor(f"$O{ir - 1}", f"$O{ir}")}/$I{ir}*Day_Basis*100)',
            fmt=PCT, fill=CLR_INT, align=CENTER)
        put(ws, ir, C_DAYS, f"=$O{ir}-$O{ir - 1}", fmt=DAYS_FMT, fill=CLR_INT, align=CENTER)
        put(ws, ir, C_INT,
            f'=IF($I{ir}<=0,0,$S{ir - 1}*{rate_factor(f"$O{ir - 1}", f"$O{ir}")})',
            fmt=MONEY, fill=CLR_INT, align=RIGHT)
        put(ws, ir, C_ACCR, 0, fmt=MONEY, fill=CLR_INT, align=RIGHT)
        inp(ws, ir, C_BANK, fmt=MONEY, align=RIGHT)
        put(ws, ir, C_VAR, f'=IF($L{ir}="","",ROUND($S{ir}-$L{ir},2))',
            fmt=MONEY, fill=CLR_INT, align=RIGHT)
        put(ws, ir, C_STAT, f'=IF(AND(Locked_To<>"",$O{ir}<=Locked_To),"LOCKED","OPEN")',
            fill=CLR_INT, align=CENTER, font=Font(size=8, bold=True))
        put(ws, ir, C_EFF, f"=$Q${hdr}", fmt=DATE_FMT, border=None)

    add_validations(ws)
    add_conditional_formats(ws)

    ws.print_title_rows = f"{ROW_TBL_HDR}:{ROW_TBL_HDR}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.protection.password = PASSWORD
    ws.protection.sheet = True
    ws.protection.enable()
    ws.protection.selectLockedCells = False
    ws.protection.selectUnlockedCells = False
    ws.protection.formatCells = True
    ws.protection.insertRows = True
    ws.protection.insertColumns = True
    ws.protection.deleteRows = True
    ws.protection.deleteColumns = True
    ws.protection.sort = True
    ws.protection.autoFilter = True
    return ws


# --------------------------------------------------------------------------------------
# Data validation
# --------------------------------------------------------------------------------------
def dv(ws, kind, sqref, *, formula1=None, formula2=None, operator=None,
       title="", msg="", prompt_title="", prompt=""):
    # Excel keeps validation formulas without the leading '=' sign.
    if isinstance(formula1, str) and formula1.startswith("="):
        formula1 = formula1[1:]
    if isinstance(formula2, str) and formula2.startswith("="):
        formula2 = formula2[1:]
    d = DataValidation(type=kind, operator=operator, formula1=formula1, formula2=formula2,
                       allow_blank=True, showErrorMessage=True, showInputMessage=bool(prompt),
                       errorStyle="stop")
    d.errorTitle, d.error = title, msg
    d.promptTitle, d.prompt = prompt_title, prompt
    ws.add_data_validation(d)
    d.add(sqref)
    return d


def add_validations(ws):
    dv(ws, "date", CELL_BF_DATE, operator="between",
       formula1="DATE(1990,1,1)", formula2="DATE(2100,12,31)",
       title="Enter a valid date",
       msg="Type the date the opening balance is brought forward from, e.g. 17-11-2024.",
       prompt_title="Balance B/F date", prompt="Date on which the opening balance stands.")
    dv(ws, "decimal", CELL_BF_AMT, operator="greaterThanOrEqual", formula1="0",
       title="Amount must be a number",
       msg="Enter the outstanding loan amount brought forward (0 or more).",
       prompt_title="Balance B/F amount", prompt="Opening outstanding principal.")
    dv(ws, "list", CELL_BASIS, formula1='"365,366,360"',
       title="Choose 365, 366 or 360", msg="Pick the day-count basis your bank uses.",
       prompt_title="Day basis", prompt="Days in a year used for interest (usually 365).")
    dv(ws, "list", CELL_QMODE, formula1=f'"{QMODE_CAL},{QMODE_ANN}"',
       title="Choose from the list",
       msg="Pick calendar quarter-ends (31 Mar / 30 Jun / 30 Sep / 31 Dec) or quarters counted "
           "from the B/F date.",
       prompt_title="Quarter-end convention", prompt="How quarter-end dates are generated.")
    dv(ws, "custom", CELL_LOCKED,
       formula1=f"=AND(ISNUMBER({CELL_LOCKED}),{CELL_LOCKED}>=BF_Date,{CELL_LOCKED}<=TODAY())",
       title="Invalid lock date",
       msg="The lock date must be a real date, on/after the B/F date and not in the future. "
           "It can only be moved forward as you reconcile with the bank.",
       prompt_title="Locked / reconciled up to",
       prompt="Everything dated on or before this date is frozen: no entries, no rate revisions.")

    dv(ws, "decimal", f"{CELL_R1_PCT}:{CELL_R3_PCT}", operator="between",
       formula1="0", formula2="100", title="Rate out of range",
       msg="Enter the annual rate in percent, e.g. 9.5 for 9.50% p.a.",
       prompt_title="Interest rate % p.a.", prompt="Annual rate in percent, e.g. 9.5")
    dv(ws, "custom", CELL_R2_DATE,
       formula1=f'=AND(ISNUMBER({CELL_R2_DATE}),{CELL_R2_DATE}>BF_Date,'
                f'{CELL_R2_DATE}>Locked_To,OR({CELL_R3_DATE}="",{CELL_R2_DATE}<{CELL_R3_DATE}))',
       title="Revision cannot be back-dated",
       msg="A revised rate must take effect AFTER the B/F date, AFTER the "
           "\u2018Locked / reconciled up to\u2019 date, and before the next revision. "
           "Periods already calculated at the earlier rate stay untouched.",
       prompt_title="1st revision effective from",
       prompt="Date the revised rate starts. Must be later than the locked-up-to date.")
    dv(ws, "custom", CELL_R3_DATE,
       formula1=f'=AND(ISNUMBER({CELL_R3_DATE}),{CELL_R3_DATE}>Locked_To,'
                f'{CELL_R2_DATE}<>"",{CELL_R3_DATE}>{CELL_R2_DATE})',
       title="Revision cannot be back-dated",
       msg="Fill slot 2 first. The 2nd revision must take effect after the 1st revision and "
           "after the \u2018Locked / reconciled up to\u2019 date.",
       prompt_title="2nd revision effective from",
       prompt="Date the second revised rate starts.")

    for q in range(1, QUARTERS + 1):
        hdr, bf, e1, e2, ir = qtr_rows(q)
        dv(ws, "custom", f"B{e1}:B{e2}",
           formula1=f'=AND(ISNUMBER($B{e1}),$B{e1}>=$O{e1 - 1},$B{e1}<=$Q${hdr},$B{e1}>Locked_To)',
           title="Date not allowed in this quarter",
           msg="The date must fall inside the quarter shown in the blue bar above this block, "
               "must not be earlier than the entry above it, and must be after the "
               "\u2018Locked / reconciled up to\u2019 date. Use the block of the correct quarter.",
           prompt_title="Repayment date",
           prompt="Value date of the repayment (must sit inside this quarter).")
        dv(ws, "custom", f"F{e1}:F{e2}",
           formula1=f'=AND(ISNUMBER($F{e1}),$F{e1}>0,ISNUMBER($B{e1}))',
           title="Enter the date first",
           msg="Type the repayment date in the Date column first, then a positive amount.",
           prompt_title="Repayment amount",
           prompt="Amount credited to the loan on that date.")
        dv(ws, "custom", f"L{bf}:L{ir}", formula1=f"=ISNUMBER($L{bf})",
           title="Numbers only",
           msg="Type the balance shown on the bank\u2019s statement so the Variance column can "
               "cross-check your figure.",
           prompt_title="Balance per Bank",
           prompt="Optional: bank\u2019s own balance for cross-checking.")


# --------------------------------------------------------------------------------------
# Conditional formatting
# --------------------------------------------------------------------------------------
def add_conditional_formats(ws):
    rng = f"A{ROW_LEDGER}:N{LAST_ROW}"
    grey = PatternFill("solid", bgColor=CLR_LOCK)
    ws.conditional_formatting.add(
        rng, FormulaRule(formula=[f'$N{ROW_LEDGER}="LOCKED"'], fill=grey, stopIfTrue=False))
    red = PatternFill("solid", bgColor="FFC7CE")
    ws.conditional_formatting.add(
        f"M{ROW_LEDGER}:M{LAST_ROW}",
        FormulaRule(formula=[f'AND($M{ROW_LEDGER}<>"",ABS($M{ROW_LEDGER})>0.004)'],
                    fill=red, font=Font(bold=True, color="9C0006")))
    ws.conditional_formatting.add(
        f"G{ROW_LEDGER}:G{LAST_ROW}",
        FormulaRule(formula=[f"$G{ROW_LEDGER}<-0.004"], font=Font(bold=True, color="9C0006")))
    for cell in (CELL_BF_DATE, CELL_BF_AMT, CELL_R1_PCT, CELL_BASIS, CELL_QMODE):
        ws.conditional_formatting.add(
            cell, FormulaRule(formula=[f'{cell}=""'], fill=red))


# --------------------------------------------------------------------------------------
# Quarter summary sheet
# --------------------------------------------------------------------------------------
def build_summary(wb):
    ws = wb.create_sheet(SHEET_SUM)
    M = f"'{SHEET_MAIN}'!"
    E = f"'{SHEET_ENG}'!"
    widths = {1: 6, 2: 13, 3: 13, 4: 16, 5: 16, 6: 15, 7: 16, 8: 38, 9: 15, 10: 12, 11: 11, 12: 11}
    for c, w in widths.items():
        ws.column_dimensions[get_column_letter(c)].width = w

    ws.merge_cells("A1:L1")
    put(ws, 1, 1, "QUARTER-BY-QUARTER SUMMARY  (all figures locked \u2013 driven by the Loan Statement)",
        font=F_TITLE, fill=CLR_TITLE, align=CENTER, border=None)
    ws.row_dimensions[1].height = 28
    ws.merge_cells("A2:L2")
    put(ws, 2, 1, f"={M}$A$2", font=F_SUB, fill=CLR_QBAR, align=CENTER, border=None)

    heads = ["Qtr", "From", "To", "Opening balance", "Repayments in quarter",
             "Interest capitalised", "Closing balance", "Rate(s) applied in the quarter",
             "Balance per Bank", "Variance", "Status", "Arithmetic check"]
    for i, h in enumerate(heads, start=1):
        put(ws, 4, i, h, font=F_WHITE_B, fill=CLR_HDR, align=WRAP_C)
    ws.row_dimensions[4].height = 32
    ws.freeze_panes = "A5"

    for q in range(1, QUARTERS + 1):
        hdr, bf, e1, e2, ir = qtr_rows(q)
        r = 4 + q
        put(ws, r, 1, q, align=CENTER, font=F_BODY_B)
        put(ws, r, 2, f"={M}$P${hdr}", fmt=DATE_Z, align=CENTER)
        put(ws, r, 3, f"={M}$Q${hdr}", fmt=DATE_Z, align=CENTER)
        put(ws, r, 4, f"={M}$S${bf}", fmt=MONEY, align=RIGHT)
        put(ws, r, 5, f"=SUM({M}$F${e1}:$F${e2})", fmt=MONEY, align=RIGHT)
        put(ws, r, 6, f"={M}$E${ir}", fmt=MONEY, align=RIGHT)
        put(ws, r, 7, f"={M}$S${ir}", fmt=MONEY, align=RIGHT, font=F_BODY_B)
        put(ws, r, 8, f"={M}$R${hdr}", align=LEFT, font=Font(size=9))
        put(ws, r, 9, f"={M}$L${ir}", fmt=MONEY, align=RIGHT)
        put(ws, r, 10, f"={M}$M${ir}", fmt=MONEY, align=RIGHT)
        put(ws, r, 11, f"={M}$N${ir}", align=CENTER, font=Font(size=8, bold=True))
        put(ws, r, 12,
            f'=IF(ROUND($D{r}-$E{r}+$F{r}-$G{r},2)=0,"OK","CHECK")',
            align=CENTER, font=Font(size=8, bold=True))

    tot = 4 + QUARTERS + 1
    put(ws, tot, 1, "TOTAL", font=F_WHITE_B, fill=CLR_PANEL, align=CENTER)
    for c in (2, 3, 4, 7, 8, 9, 10, 11, 12):
        put(ws, tot, c, fill=CLR_PANEL)
    put(ws, tot, 5, f"=SUM($E$5:$E${4 + QUARTERS})", fmt=MONEY, font=F_WHITE_B,
        fill=CLR_PANEL, align=RIGHT)
    put(ws, tot, 6, f"=SUM($F$5:$F${4 + QUARTERS})", fmt=MONEY, font=F_WHITE_B,
        fill=CLR_PANEL, align=RIGHT)

    panel = [
        ("Outstanding today (principal + accrued interest)", f"={E}$B$15", MONEY),
        ("Principal balance as per last event", f"={E}$B$12", MONEY),
        ("Interest accrued but not yet capitalised", f"={E}$B$14", MONEY),
        ("Total repaid to date", f"={E}$B$16", MONEY),
        ("Total interest charged to date", f"={E}$B$17", MONEY),
        ("Rows where your balance differs from the bank\u2019s", f"={E}$B$18", "0"),
    ]
    base = tot + 2
    ws.merge_cells(start_row=base, start_column=1, end_row=base, end_column=4)
    put(ws, base, 1, "LIVE POSITION  &  INTEGRITY CHECKS", font=F_WHITE_B,
        fill=CLR_PANEL, align=CENTER)
    for i, (label, formula, fmt) in enumerate(panel, start=1):
        r = base + i
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        put(ws, r, 1, label, font=F_LBL, align=LEFT)
        put(ws, r, 4, formula, fmt=fmt, align=RIGHT, font=F_BODY_B, fill=CLR_CALC)
    ws.conditional_formatting.add(
        f"D{base + 6}",
        FormulaRule(formula=[f"$D${base + 6}>0"], fill=PatternFill("solid", bgColor="FFC7CE"),
                    font=Font(bold=True, color="9C0006")))
    ws.conditional_formatting.add(
        f"L5:L{4 + QUARTERS}",
        FormulaRule(formula=['$L5="CHECK"'], fill=PatternFill("solid", bgColor="FFC7CE"),
                    font=Font(bold=True, color="9C0006")))
    ws.conditional_formatting.add(
        f"A5:L{4 + QUARTERS}",
        FormulaRule(formula=['$K5="LOCKED"'], fill=PatternFill("solid", bgColor=CLR_LOCK)))

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.protection.password = PASSWORD
    ws.protection.sheet = True
    ws.protection.enable()
    ws.protection.selectLockedCells = False
    return ws


# --------------------------------------------------------------------------------------
# Instruction sheet
# --------------------------------------------------------------------------------------
HELP = [
    ("H", "HOW THIS FILE WORKS"),
    ("P", "This workbook keeps a bank loan account on the same layout as a bank statement and "
          "charges interest the way the bank does: interest accrues on the daily outstanding "
          "balance and is added to the balance (capitalised) at every quarter-end, so the next "
          "quarter earns interest on interest \u2013 quarterly compounding."),
    ("H", "1.  WHAT YOU FILL IN  (pale yellow cells only)"),
    ("B", "Top-left panel \u2013 account particulars: A/C no., borrower, facility, address, limit."),
    ("B", "Middle panel \u2013 Balance B/F date and amount: the opening position of the file. "
          "Interest basis: 365 (or 360/366 if your bank says so). Quarter-end convention: "
          "calendar quarter-ends, or quarters counted three months at a time from the B/F date."),
    ("B", "Rate register (top-right) \u2013 slot 1 is the original rate and always runs from the "
          "B/F date. Slots 2 and 3 are revisions: type the date the new rate takes effect and "
          "the new percentage."),
    ("B", "In the ledger \u2013 for every repayment, type the Date and the Credit amount in the "
          "block of the quarter the repayment falls in. Cheque/instrument and particulars are "
          "optional. Balance per Bank is optional: fill it and the Variance column tells you "
          "instantly whether your file agrees with the bank."),
    ("H", "2.  WHAT THE FILE DOES ON ITS OWN"),
    ("B", "Generates every quarter-end date and shows each quarter in its own blue block."),
    ("B", "Counts the exact number of days between events and charges interest on the balance "
          "that was outstanding during those days."),
    ("B", "At each quarter-end it debits the accumulated interest, adds it to the balance, and "
          "carries the new figure forward as the next quarter\u2019s opening balance."),
    ("B", "Splits a quarter day-by-day when a rate revision takes effect inside that quarter, "
          "and prints the rates used on the quarter bar."),
    ("H", "3.  HOW EARLIER PERIODS ARE PROTECTED FROM A LATER RATE CHANGE"),
    ("B", "A revision is stored with an effective-from date. Interest for any day is computed at "
          "the rate in force on that day, so a rate that starts on 01-07-2026 mathematically "
          "cannot alter a quarter that ended on 30-06-2026."),
    ("B", "The cell \u2018Locked / reconciled up to\u2019 is the hard cut-off. Once you set it "
          "(normally the last quarter-end you have reconciled with the bank), the file refuses "
          "any transaction dated on or before it, and refuses any rate revision dated on or "
          "before it. Those quarters turn grey and are marked LOCKED."),
    ("B", "The title line at the top of every sheet always states the rate(s) on the file with "
          "their effective dates, and the Quarter Summary records the rate(s) used in each "
          "quarter \u2013 a permanent audit trail."),
    ("H", "4.  WHY YOU CANNOT DAMAGE THE CALCULATIONS"),
    ("B", "Every sheet is protected. Only the pale-yellow, gold-bordered cells accept typing; "
          "all formulas, dates, balances and totals are locked."),
    ("B", "Rows and columns cannot be inserted, deleted or sorted, and the workbook structure is "
          "locked, so sheets cannot be added, removed or unhidden."),
    ("B", "The calculation engine sits on a hidden, protected sheet, so the formulas cannot be "
          "reached even by accident."),
    ("B", "Data validation rejects impossible input: a repayment dated outside its quarter, out "
          "of sequence, or inside a locked period; an amount typed without a date; a rate outside "
          "0\u2013100%; a back-dated revision."),
    ("B", "The Variance column and the \u2018Arithmetic check\u2019 column on the Quarter Summary "
          "turn red the moment a figure stops agreeing with the bank or with the ledger, so any "
          "tampering is visible at a glance."),
    ("H", "5.  PRACTICAL NOTES"),
    ("B", "Ten repayment rows are provided per quarter and 24 quarters in total. Unused rows and "
          "future quarters simply show \u2018\u2013\u2019."),
    ("B", "Convention: a repayment reduces the balance from its value date, and interest accrued "
          "during a quarter is charged at the quarter-end. Day count = actual days between "
          "dates \u00f7 basis."),
    ("B", "Interest is rounded to 2 decimals only at the moment it is capitalised, exactly like a "
          "bank ledger."),
    ("B", "Sheet/workbook password: " + PASSWORD + " \u2013 keep it with the file owner. You only "
          "need it to change the structure (for example to add more repayment rows or quarters)."),
    ("N", "Ask for a rebuilt file if you need more than 3 rate slots, more rows per quarter, a "
          "longer tenure, or repayments split between interest and principal instead of "
          "quarter-end capitalisation."),
]


def build_help(wb):
    ws = wb.create_sheet(SHEET_HELP)
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 120
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:B1")
    put(ws, 1, 1, "BANK LOAN \u2013 QUARTERLY COMPOUND INTEREST FILE  \u00b7  USER GUIDE",
        font=F_TITLE, fill=CLR_TITLE, align=CENTER, border=None)
    ws.row_dimensions[1].height = 30

    r = 3
    for kind, text in HELP:
        if kind == "H":
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
            put(ws, r, 1, text, font=F_WHITE_B, fill=CLR_PANEL, align=LEFT, border=None)
            ws.row_dimensions[r].height = 18
        elif kind == "B":
            put(ws, r, 1, "\u2022", font=F_BODY_B, align=CENTER, border=None)
            c = put(ws, r, 2, text, font=F_BODY, align=WRAP_L, border=None)
            ws.row_dimensions[r].height = 15 * (1 + len(text) // 110)
        elif kind == "N":
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
            put(ws, r, 1, text, font=Font(size=10, italic=True, color="7F6000"),
                fill=CLR_NOTE, align=WRAP_L, border=None)
            ws.row_dimensions[r].height = 30
        else:
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
            put(ws, r, 1, text, font=F_BODY, align=WRAP_L, border=None)
            ws.row_dimensions[r].height = 15 * (1 + len(text) // 110)
        r += 2 if kind == "H" else 1

    ws.protection.password = PASSWORD
    ws.protection.sheet = True
    ws.protection.enable()
    ws.protection.selectLockedCells = False
    return ws


# --------------------------------------------------------------------------------------
# Named ranges
# --------------------------------------------------------------------------------------
def add_names(wb):
    main = f"'{SHEET_MAIN}'!"
    eng = f"'{SHEET_ENG}'!"
    names = {
        "BF_Date": f"{main}${CELL_BF_DATE[0]}${CELL_BF_DATE[1:]}",
        "BF_Amount": f"{main}${CELL_BF_AMT[0]}${CELL_BF_AMT[1:]}",
        "Day_Basis": f"{main}${CELL_BASIS[0]}${CELL_BASIS[1:]}",
        "Qtr_Mode": f"{main}${CELL_QMODE[0]}${CELL_QMODE[1:]}",
        "Locked_To": f"{main}${CELL_LOCKED[0]}${CELL_LOCKED[1:]}",
        "Eff_Dates": f"{main}$O${ROW_LEDGER}:$O${LAST_ROW}",
        "Bal_Col": f"{main}$S${ROW_LEDGER}:$S${LAST_ROW}",
        "Accr_Col": f"{main}$T${ROW_LEDGER}:$T${LAST_ROW}",
        "Big_Date": f"{eng}$B$1",
        "R1_From": f"{eng}$B$2", "R1_To": f"{eng}$B$3",
        "R2_From": f"{eng}$B$4", "R2_To": f"{eng}$B$5",
        "R3_From": f"{eng}$B$6", "R3_To": f"{eng}$B$7",
        "Rate_1": f"{eng}$B$8", "Rate_2": f"{eng}$B$9", "Rate_3": f"{eng}$B$10",
    }
    for name, ref in names.items():
        wb.defined_names.add(DefinedName(name, attr_text=ref))


# --------------------------------------------------------------------------------------
# Example data
# --------------------------------------------------------------------------------------
EXAMPLE_SETUP = {
    "C6": "1120000138481",
    "C7": "MD MOHSIN CHOWDHURY",
    "C8": "Term Loan \u2013 quarterly compounding",
    "C9": "Abc bay view, road 3, khulshi hills, Chittagong",
    "C10": 6000000,
    CELL_BF_DATE: dt.date(2024, 11, 17),
    CELL_BF_AMT: 5681000.00,
    CELL_BASIS: 365,
    CELL_QMODE: QMODE_CAL,
    CELL_LOCKED: dt.date(2025, 3, 31),
    CELL_R1_PCT: 9.00,
    CELL_R2_DATE: dt.date(2025, 7, 1),
    CELL_R2_PCT: 10.50,
    CELL_R3_DATE: dt.date(2026, 1, 1),
    CELL_R3_PCT: 11.25,
}

# quarter number -> list of (date, amount, instrument, particulars)
EXAMPLE_ENTRIES = {
    1: [(dt.date(2024, 12, 15), 500000, "CHQ 123456", "Repayment by transfer")],
    2: [(dt.date(2025, 1, 20), 300000, "CHQ 123457", "Repayment by transfer"),
        (dt.date(2025, 2, 28), 200000, "RTGS", "Part repayment")],
    3: [(dt.date(2025, 5, 15), 400000, "CHQ 123458", "Repayment by transfer")],
    4: [(dt.date(2025, 8, 10), 250000, "RTGS", "Part repayment (rate revised 01-07-2025)")],
    5: [(dt.date(2025, 11, 5), 350000, "CHQ 123459", "Repayment by transfer")],
    6: [(dt.date(2026, 2, 20), 300000, "RTGS", "Part repayment (rate revised 01-01-2026)")],
    7: [(dt.date(2026, 5, 15), 250000, "CHQ 123460", "Repayment by transfer")],
}


# quarter -> balance confirmed by the bank at that quarter-end (reconciliation demo)
EXAMPLE_BANK = {1: 5240662.36, 2: 4850255.14}


def fill_example(ws):
    for cell, value in EXAMPLE_SETUP.items():
        ws[cell] = value
    for q, entries in EXAMPLE_ENTRIES.items():
        _, _, e1, e2, _ = qtr_rows(q)
        for i, (d, amt, inst, part) in enumerate(entries):
            r = e1 + i
            assert r <= e2
            ws.cell(row=r, column=C_DATE, value=d)
            ws.cell(row=r, column=C_INST, value=inst)
            ws.cell(row=r, column=C_PART, value=part)
            ws.cell(row=r, column=C_CREDIT, value=amt)
    for q, bank_balance in EXAMPLE_BANK.items():
        _, _, _, _, ir = qtr_rows(q)
        ws.cell(row=ir, column=C_BANK, value=bank_balance)


# --------------------------------------------------------------------------------------
# Optional hard freeze of the opening position
# --------------------------------------------------------------------------------------
# Plain Excel cannot make a cell "write-once". If the opening figures and the original
# rate are known when the file is produced, they can instead be written as LOCKED cells:
# from then on only future rate revisions and new repayments can be entered, and nothing
# about the history can be retyped even by accident.
FROZEN_CELLS = ["C6", "D6", "E6", "C7", "D7", "E7", "C8", "D8", "E8", "C9", "D9", "E9",
                "C10", "D10", "E10",
                CELL_BF_DATE, "I6", CELL_BF_AMT, "I7", CELL_BASIS, "I8", CELL_QMODE, "I9",
                CELL_R1_PCT]


def freeze_setup(ws):
    frozen_fill = PatternFill("solid", fgColor="E7E6E6")
    for ref in FROZEN_CELLS:
        c = ws[ref]
        c.protection = LOCKED
        c.fill = frozen_fill
        c.border = BOX
    ws["A3"] = ('\u25a0 The opening position and the original rate are FROZEN in this build. '
                'Pale-yellow cells (repayments, rate revisions 2-3, bank balances, lock date) '
                'are the only cells that accept typing.')


def build(path: str, example: bool = False, freeze: bool = False):
    wb = Workbook()
    main = build_main(wb)
    build_summary(wb)
    build_help(wb)
    build_engine(wb)
    add_names(wb)
    if example:
        fill_example(main)
    if freeze:
        freeze_setup(main)
    wb.security = WorkbookProtection(workbookPassword=PASSWORD, lockStructure=True)
    wb.active = 0
    wb.save(path)
    return path


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "loan.xlsx"
    build(out, example="--example" in sys.argv, freeze="--freeze-setup" in sys.argv)
    print("written:", out)
