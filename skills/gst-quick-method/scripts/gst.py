#!/usr/bin/env -S uv run --script
# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl>=3.1", "formulas>=1.2"]
# ///
"""GST Quick Method workbook tool (quarterly filer, GST-only province, services).

Subcommands
  build        create a new, empty workbook (refuses to overwrite)
  add-sales    append invoices from a JSON file (dedupes on invoice #)
  add-capital  append capital purchases (ITC claims) from a JSON file
  rebuild      carry all inputs onto the current template (after a layout change)
  check        evaluate every formula (pure-Python engine, plus real Excel when
               reachable), cross-check each return line against an
               independent calculation, print the summary
  record-payment  record a quarter's filed date, amount paid, confirmation #
  set-business    set business name / GST account number on Settings
  path         print the configured workbook path for a fiscal year

The workbook keeps all math in Excel formulas; this tool only writes input
cells. Keep workbooks (financial data) outside any git repository.

Workbook location: env GST_WORKBOOK, else "workbook" in
~/.config/gst-quick-method/config.json -- a path template where {year} is
replaced by the fiscal year, e.g.
  {"workbook": "/mnt/c/accounting/{year}/GST Returns {year} - Quick Method.xlsx"}

Requirements for `check`: WSL with Windows Excel installed, and PowerShell 7
(pwsh.exe on PATH, or env GST_PWSH, or the default install location).
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

# ---------------------------------------------------------------- layout ---
SALES_FIRST, SALES_LAST = 5, 504          # 500 invoice rows
CAP_FIRST, CAP_LAST = 5, 104              # 100 capital-purchase rows
QCOLS = ["C", "D", "E", "F"]              # Summary: Q1..Q4
YCOL = "G"                                # Summary: fiscal-year total

MONEY = '"$"#,##0.00;[Red]-"$"#,##0.00'
DATE = "yyyy-mm-dd"
PCT = "0.0%"

INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")   # pale yellow = type here
HEAD_FILL = PatternFill("solid", fgColor="1F4E78")
SECTION_FILL = PatternFill("solid", fgColor="DDEBF7")
KEY_FILL = PatternFill("solid", fgColor="E2EFDA")
HEAD_FONT = Font(bold=True, color="FFFFFF")
INPUT_FONT = Font(color="0000FF")                       # blue = hard-coded input
BOLD = Font(bold=True)
GREY = Font(color="808080", italic=True)
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)

SETTINGS = [  # (row, label, name, value, format, note)
    (4, "Business name", "Business_Name", "", None, ""),
    (5, "GST/HST account (BN RT0001)", "GST_Number", "", "@", ""),
    (6, "Fiscal year start", "FY_Start", None, DATE, "Quarters are counted from this date"),
    (7, "GST rate", "GST_Rate", 0.05, PCT, "Rate charged on invoices (GST-only province: 5%)"),
    (8, "Quick Method remittance rate", "QM_Rate", 0.036, PCT, "Services, non-participating province"),
    (9, "Quick Method credit rate", "Credit_Rate", 0.01, PCT, "1% credit ..."),
    (10, "Credit threshold (GST-incl. sales)", "Credit_Threshold", 30000, MONEY, "... on the first $30,000 per fiscal year"),
]


DISCLAIMER = (
    "Not tax or accounting advice. This tool is provided as is, with no warranty. You are responsible for "
    "everything you file with CRA. Before filing, check the numbers against your own records, and confirm "
    "that the Quick Method applies to you and that the rates on the Settings sheet are current for your "
    "province and type of business. Rates and rules change. Not affiliated with or endorsed by the "
    "Canada Revenue Agency."
)

SUMMARY_LAYOUT = [  # (key, label, line # or "section")
    ("qnum", "Quarter #", None),
    ("start", "Period start", None),
    ("end", "Period end", None),
    ("due", "Filing and payment due", None),
    ("sec_sales", "SALES (from the Sales sheet)", "section"),
    ("count", "Invoices", None),
    ("net", "Sales excl. GST  (your income)", None),
    ("gst", "GST charged to clients", None),
    ("incl", "Sales incl. GST", None),
    ("sec_form", "CRA GST/HST RETURN -- fields exactly as on the CRA screen, top to bottom", "section"),
    ("f_start", "Payment Period Start Date", None),
    ("f_end", "Payment Period End Date", None),
    ("l101", "Sales and other revenues", "101"),
    ("l110", "Instalment & Annual Filer Payments", "110"),
    ("l111", "Rebates", "111"),
    ("l205", "GST/HST due on purchases of real property or purchases of Emission allowances", "205"),
    ("l405", "Other GST/HST to be Self-Assessed", "405"),
    ("house", "Total GST/HST new housing rebates included in line 108", ""),
    ("l109", "Net tax", "109"),
    ("l114", "Refund claimed", "114"),
    ("l103", "GST/HST Amounts Collected or that Became Collectible", "103"),
    ("l104", "GST/HST Adjustments", "104"),
    ("l106", "Input Tax Credits", "106"),
    ("l107", "Input Tax Credits Adjustments", "107"),
    ("sec_work", "WORKINGS (not typed into CRA)", "section"),
    ("l105", "Total GST/HST and adjustments (103 + 104)", "105"),
    ("l108", "Total ITCs and adjustments (106 + 107)", "108"),
    ("l112", "Total other credits (110 + 111)", "112"),
    ("l113a", "Balance (109 - 112)", "113A"),
    ("l113b", "Other debits (205 + 405)", "113B"),
    ("l113c", "Balance (113A + 113B)", "113C"),
    ("l115", "AMOUNT TO PAY CRA", "115"),
    ("sec_pay", "PAYMENT", "section"),
    ("filed", "Date filed / paid", None),
    ("paid", "Amount paid", None),
    ("conf", "CRA confirmation #", None),
    ("bal", "Balance still owing", None),
    ("sec_info", "FOR YOUR INCOME TAX", "section"),
    ("gain", "Quick Method gain (GST charged - line 105 + line 107), report as income", None),
]
FORM_KEYS = ("f_start", "f_end", "l101", "l110", "l111", "l205", "l405", "house", "l109", "l114",
             "l103", "l104", "l106", "l107")
INPUT_KEYS = ("l110", "l111", "l205", "l405", "house", "l104", "filed", "paid", "conf")
MONEY_KEYS = ("net", "gst", "incl", "l101", "l110", "l111", "l205", "l405", "house", "l109", "l114",
              "l103", "l104", "l106", "l107", "l105", "l108", "l112", "l113a", "l113b", "l113c",
              "l115", "paid", "bal", "gain")


_SF = lambda col: f"Sales!${col}${SALES_FIRST}:${col}${SALES_LAST}"  # noqa: E731
YTD_PANEL = [  # (row, label, formula, format) -- Summary!I:J; covers every invoice entered so far
    (5, "Gross income (incl. GST)", f"=SUM({_SF('H')})", MONEY),
    (6, "Net income (excl. GST)", f"=SUM({_SF('F')})", MONEY),
    (7, "GST charged to clients", f"=SUM({_SF('G')})", MONEY),
    (8, "Invoices", f'=COUNT({_SF("F")})', "0"),
    (9, "Through (latest invoice date)", f'=IF(COUNT({_SF("A")})=0,"",MAX({_SF("A")}))', DATE),
]


def summary_rows() -> dict:
    return {key: 5 + i for i, (key, _, _) in enumerate(SUMMARY_LAYOUT)}


def _hdr(ws, row, labels, widths=None):
    for i, text in enumerate(labels, start=1):
        c = ws.cell(row=row, column=i, value=text)
        c.fill, c.font, c.border = HEAD_FILL, HEAD_FONT, BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 32
    for i, w in enumerate(widths or [], start=1):
        ws.column_dimensions[chr(64 + i)].width = w


def _quarter_formula(date_ref, amount_ref=None):
    blank = f'OR({date_ref}="",{amount_ref}="")' if amount_ref else f'{date_ref}=""'
    return (f'=IF({blank},"",IF(OR({date_ref}<FY_Start,{date_ref}>EDATE(FY_Start,12)-1),'
            f'"OUT",INT(DATEDIF(FY_Start,{date_ref},"m")/3)+1))')


def build(path: Path, fy_start: dt.date, business: str, bn: str) -> None:
    if path.exists():
        sys.exit(f"refusing to overwrite existing file: {path}")
    wb = Workbook()

    # ---------------------------------------------------------- Summary --
    s = wb.active
    s.title = "Summary"
    # --------------------------------------------------------- Sales -----
    sl = wb.create_sheet("Sales")
    # --------------------------------------------------------- Capital ---
    cp = wb.create_sheet("Capital")
    # --------------------------------------------------------- Settings --
    st = wb.create_sheet("Settings")
    hw = wb.create_sheet("How to use")

    # Settings
    st["A1"] = "Settings"
    st["A1"].font = Font(bold=True, size=14)
    st["A2"] = "Yellow cells are inputs. Rates feed every formula in the workbook."
    st["A2"].font = GREY
    for row, label, name, value, fmt, note in SETTINGS:
        st.cell(row=row, column=1, value=label).font = BOLD
        c = st.cell(row=row, column=2, value=fy_start if name == "FY_Start" else value)
        if name == "Business_Name":
            c.value = business
        if name == "GST_Number":
            c.value = bn
        c.fill, c.font, c.border = INPUT_FILL, INPUT_FONT, BOX
        if fmt:
            c.number_format = fmt
        st.cell(row=row, column=3, value=note).font = GREY
        wb.defined_names[name] = DefinedName(name, attr_text=f"Settings!$B${row}")
    st["A12"] = "Fiscal year end"
    st["A12"].font = BOLD
    st["B12"] = "=EDATE(FY_Start,12)-1"
    st["B12"].number_format = DATE
    st["A13"] = "Filing frequency"
    st["A13"].font = BOLD
    st["B13"] = "Quarterly (due one month after each quarter ends)"
    st.column_dimensions["A"].width = 34
    st.column_dimensions["B"].width = 22
    st.column_dimensions["C"].width = 44

    # Sales ledger
    sl["A1"] = "Sales invoices"
    sl["A1"].font = Font(bold=True, size=14)
    sl["A2"] = ("Enter the invoice amount BEFORE GST (as on your accounting software's income report). "
                "GST, total and quarter are calculated. Red = duplicate invoice #, date outside the fiscal year, or a missing date/amount.")
    sl["A2"].font = GREY
    sl["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    sl.merge_cells("A2:J2")
    sl.row_dimensions[2].height = 32
    _hdr(sl, 4, ["Date", "Invoice #", "Customer", "Memo", "Project", "Amount excl. GST",
                 "GST", "Total incl. GST", "Quarter", "Notes"],
         [12, 10, 16, 32, 9, 16, 12, 16, 9, 30])
    for r in range(SALES_FIRST, SALES_LAST + 1):
        for col in "ABCDEFJ":
            c = sl[f"{col}{r}"]
            c.fill, c.font = INPUT_FILL, INPUT_FONT
        sl[f"A{r}"].number_format = DATE
        sl[f"F{r}"].number_format = MONEY
        sl[f"G{r}"] = f'=IF(F{r}="","",ROUND(F{r}*GST_Rate,2))'
        sl[f"H{r}"] = f'=IF(F{r}="","",F{r}+G{r})'
        sl[f"I{r}"] = _quarter_formula(f"A{r}", f"F{r}")  # no quarter until date and amount are both in
        for col in "GH":
            sl[f"{col}{r}"].number_format = MONEY
        sl[f"I{r}"].number_format = '"Q"0'
        sl[f"I{r}"].alignment = Alignment(horizontal="center")
    sl.freeze_panes = "A5"
    rng = f"A{SALES_FIRST}:J{SALES_LAST}"
    red = PatternFill("solid", fgColor="F8CBAD")
    sl.conditional_formatting.add(
        f"B{SALES_FIRST}:B{SALES_LAST}",
        FormulaRule(formula=[f'AND(B{SALES_FIRST}<>"",COUNTIF($B${SALES_FIRST}:$B${SALES_LAST},B{SALES_FIRST})>1)'], fill=red))
    sl.conditional_formatting.add(
        f"I{SALES_FIRST}:I{SALES_LAST}",
        FormulaRule(formula=[f'I{SALES_FIRST}="OUT"'], fill=red))
    sl.conditional_formatting.add(
        f"A{SALES_FIRST}:A{SALES_LAST}",
        FormulaRule(formula=[f'AND(A{SALES_FIRST}="",F{SALES_FIRST}<>"")'], fill=red))
    sl.conditional_formatting.add(
        f"F{SALES_FIRST}:F{SALES_LAST}",
        FormulaRule(formula=[f'AND(F{SALES_FIRST}="",A{SALES_FIRST}<>"")'], fill=red))
    dv = DataValidation(type="date", operator="greaterThan", formula1="DATE(2000,1,1)",
                        error="Enter a date", allow_blank=True)
    sl.add_data_validation(dv)
    dv.add(f"A{SALES_FIRST}:A{SALES_LAST}")
    sl.auto_filter.ref = f"A4:J{SALES_LAST}"
    del rng

    # Capital purchases
    cp["A1"] = "Capital purchases -- GST input tax credits (line 106)"
    cp["A1"].font = Font(bold=True, size=14)
    cp["A2"] = ("Quick Method: only capital property earns an ITC (computers, equipment, furniture...). "
                "Enter the GST actually shown on the receipt. ITC is claimed only if the asset is used "
                "primarily (>50%) in the business. Orange = GST paid exceeds 5% of cost (check for PST mixed in).")
    cp["A2"].font = GREY
    cp["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    cp.merge_cells("A2:I2")
    cp.row_dimensions[2].height = 45
    _hdr(cp, 4, ["Date", "Vendor", "Description", "Cost excl. taxes", "GST paid (receipt)",
                 "Business use >50%? (Y/N)", "ITC claimed", "Quarter", "Notes"],
         [12, 20, 34, 16, 16, 14, 14, 9, 30])
    for r in range(CAP_FIRST, CAP_LAST + 1):
        for col in "ABCDEFI":
            c = cp[f"{col}{r}"]
            c.fill, c.font = INPUT_FILL, INPUT_FONT
        cp[f"A{r}"].number_format = DATE
        for col in "DEG":
            cp[f"{col}{r}"].number_format = MONEY
        cp[f"F{r}"].alignment = Alignment(horizontal="center")
        cp[f"G{r}"] = f'=IF(E{r}="","",IF(UPPER(F{r})="Y",E{r},0))'
        cp[f"H{r}"] = _quarter_formula(f"A{r}")
        cp[f"H{r}"].number_format = '"Q"0'
        cp[f"H{r}"].alignment = Alignment(horizontal="center")
    yn = DataValidation(type="list", formula1='"Y,N"', allow_blank=True)
    cp.add_data_validation(yn)
    yn.add(f"F{CAP_FIRST}:F{CAP_LAST}")
    orange = PatternFill("solid", fgColor="FFD966")
    cp.conditional_formatting.add(
        f"E{CAP_FIRST}:E{CAP_LAST}",
        FormulaRule(formula=[f'AND(E{CAP_FIRST}<>"",D{CAP_FIRST}<>"",E{CAP_FIRST}>ROUND(D{CAP_FIRST}*GST_Rate,2)+0.01)'], fill=orange))
    cp.conditional_formatting.add(
        f"H{CAP_FIRST}:H{CAP_LAST}", FormulaRule(formula=[f'H{CAP_FIRST}="OUT"'], fill=red))
    cp.freeze_panes = "A5"

    # Summary (the return, one column per quarter)
    S, C = "Sales!", "Capital!"
    sq = f"{S}$I${SALES_FIRST}:$I${SALES_LAST}"
    cq = f"{C}$H${CAP_FIRST}:$H${CAP_LAST}"
    s["A1"] = "GST/HST Return -- Quick Method (GST34), quarterly"
    s["A1"].font = Font(bold=True, size=14)
    s["A2"] = '=IF(Business_Name="","",Business_Name&"   ")&IF(GST_Number="","","BN "&GST_Number&"   ")&"Fiscal year "&TEXT(FY_Start,"yyyy-mm-dd")&" to "&TEXT(EDATE(FY_Start,12)-1,"yyyy-mm-dd")'
    s["A2"].font = GREY
    s["A3"] = "Not tax advice: check every number against your own records before filing (see note at the bottom)."
    s["A3"].font = Font(bold=True, color="C00000")
    _hdr(s, 4, ["", "GST34 line", "Q1", "Q2", "Q3", "Q4", "Year"], [76, 11, 15, 15, 15, 15, 16])

    R = summary_rows()
    for key, label, kind in SUMMARY_LAYOUT:
        r = R[key]
        s[f"A{r}"] = label
        if kind == "section":
            for col in "ABCDEFG":
                s[f"{col}{r}"].fill = SECTION_FILL
            s[f"A{r}"].font = BOLD
        elif kind:
            s[f"B{r}"] = int(kind) if kind.isdigit() else kind  # numbers, so Excel doesn't flag "number as text"
            s[f"B{r}"].alignment = Alignment(horizontal="center")
            s[f"B{r}"].font = BOLD
    for i, col in enumerate(QCOLS, start=1):
        q = f"{col}${R['qnum']}"
        c_ = lambda k: f"{col}{R[k]}"  # noqa: E731
        f = {
            "qnum": i,
            "start": f"=EDATE(FY_Start,3*({q}-1))",
            "end": f"=EDATE(FY_Start,3*{q})-1",
            "due": f"=EOMONTH({c_('end')},1)",
            "count": f"=COUNTIFS({sq},{q})",
            "net": f"=SUMIFS({S}$F${SALES_FIRST}:$F${SALES_LAST},{sq},{q})",
            "gst": f"=SUMIFS({S}$G${SALES_FIRST}:$G${SALES_LAST},{sq},{q})",
            "incl": f"=SUMIFS({S}$H${SALES_FIRST}:$H${SALES_LAST},{sq},{q})",
            # --- CRA form, same order as the screen ---
            "f_start": f"={c_('start')}",
            "f_end": f"={c_('end')}",
            "l101": f"=ROUND({c_('incl')},0)",
            "l110": 0, "l111": 0, "l205": 0, "l405": 0, "house": 0, "l104": 0,
            "l109": f"={c_('l105')}-{c_('l108')}",
            "l114": f"=IF({c_('l113c')}<0,-{c_('l113c')},0)",
            "l103": f"={c_('l101')}*QM_Rate",
            "l106": f"=SUMIFS({C}$G${CAP_FIRST}:$G${CAP_LAST},{cq},{q})",
            # 1% of this quarter's line 101, limited to what is left of the yearly threshold
            "l107": (f"=ROUND(Credit_Rate*MIN({c_('l101')},MAX(0,Credit_Threshold-"
                     f"(SUM($C{R['l101']}:{c_('l101')})-{c_('l101')}))),2)"),
            # --- workings ---
            "l105": f"={c_('l103')}+{c_('l104')}",
            "l108": f"={c_('l106')}+{c_('l107')}",
            "l112": f"={c_('l110')}+{c_('l111')}",
            "l113a": f"={c_('l109')}-{c_('l112')}",
            # SUM(a,b) rather than a+b so it doesn't match line 112's shape (Excel would flag 113A as inconsistent)
            "l113b": f"=SUM({c_('l205')},{c_('l405')})",
            "l113c": f"={c_('l113a')}+{c_('l113b')}",
            "l115": f"=IF({c_('l113c')}>0,{c_('l113c')},0)",
            # amount owed is paid in cents; 103 is unrounded (x 3.6%), so round before subtracting
            "bal": f"=ROUND({c_('l115')},2)-{c_('paid')}",
            "gain": f"={c_('gst')}-{c_('l105')}+{c_('l107')}",
        }
        for key, val in f.items():
            s[c_(key)] = val
        for key in INPUT_KEYS:
            c = s[c_(key)]
            c.fill, c.font, c.border = INPUT_FILL, INPUT_FONT, BOX
    # Year column
    s[f"{YCOL}{R['start']}"] = "=FY_Start"
    s[f"{YCOL}{R['end']}"] = "=EDATE(FY_Start,12)-1"
    for key in ("count", *MONEY_KEYS):
        # explicit +, not SUM(C:F): SUM would make Excel flag "omits adjacent cells" (line # in column B)
        s[f"{YCOL}{R[key]}"] = "=" + "+".join(f"{c}{R[key]}" for c in QCOLS)
    # formats
    for key in ("start", "end", "due", "filed"):
        for col in QCOLS + [YCOL]:
            s[f"{col}{R[key]}"].number_format = DATE
    for key in ("f_start", "f_end"):
        for col in QCOLS:
            s[f"{col}{R[key]}"].number_format = "dd/mm/yyyy"  # as typed on the CRA screen
    for key in MONEY_KEYS:
        for col in QCOLS + [YCOL]:
            s[f"{col}{R[key]}"].number_format = MONEY
    for col in QCOLS:
        s[f"{col}{R['qnum']}"].font = GREY
        s[f"{col}{R['conf']}"].number_format = "@"
    s[f"A{R['qnum']}"].font = GREY
    for key in FORM_KEYS:
        for col in QCOLS + ([YCOL] if key not in ("f_start", "f_end") else []):
            cell = s[f"{col}{R[key]}"]
            if key not in INPUT_KEYS:
                cell.fill = KEY_FILL
            cell.border = BOX
    for key in ("l109", "l115"):
        for col in "ABCDEFG":
            s[f"{col}{R[key]}"].font = Font(bold=True, size=12)
    s.freeze_panes = "C5"
    # Year-to-date panel, on the side
    s.column_dimensions["H"].width = 3
    s.column_dimensions["I"].width = 30
    s.column_dimensions["J"].width = 16
    for col in "IJ":
        c = s[f"{col}4"]
        c.fill, c.font, c.border = HEAD_FILL, HEAD_FONT, BOX
    s["I4"] = "YEAR TO DATE"
    for row, label, formula, fmt in YTD_PANEL:
        s[f"I{row}"] = label
        s[f"J{row}"] = formula
        s[f"J{row}"].number_format = fmt
        for col in "IJ":
            s[f"{col}{row}"].border = BOX
        s[f"J{row}"].fill = KEY_FILL
    for row in (5, 6):
        for col in "IJ":
            s[f"{col}{row}"].font = Font(bold=True, size=12)
    note = R["gain"] + 2
    for i, text in enumerate([
        "The CRA block lists the fields with the same labels and order as the CRA screen (green = calculated, "
        "yellow = type if it applies, normally 0).",
        "Quick Method: 'Sales and other revenues' includes GST; 'Amounts Collected' = that x remittance rate; "
        "the 1% credit goes in 'Input Tax Credits Adjustments' (107) and stops after the first $30,000 of the year.",
        "'Total GST/HST new housing rebates included in line 108' is informational only and does not change the math.",
        "",
        DISCLAIMER,
    ]):
        s[f"A{note + i}"] = text
        s[f"A{note + i}"].font = GREY
    last = note + 4
    s[f"A{last}"].font = Font(bold=True, color="C00000")
    s[f"A{last}"].alignment = Alignment(wrap_text=True, vertical="top")
    s.merge_cells(f"A{last}:G{last}")
    s.row_dimensions[last].height = 48

    # How to use
    hw.column_dimensions["A"].width = 110
    for i, line in enumerate([
        "How to use",
        "",
        "1. Each quarter, add that quarter's invoices to the Sales sheet: date, invoice #, amount BEFORE GST.",
        "   (An income-account report, e.g. QuickBooks, shows amounts before GST -- copy them as-is.)",
        "2. Add any capital purchase (computer, equipment...) to the Capital sheet with the GST from the receipt",
        "   (GST only, not PST). With the Claude skill, paste a photo or screenshot of the receipt or invoice.",
        "3. Open Summary: the quarter's column holds every GST34 line (101, 103, 105-109) to type into CRA My Business Account.",
        "4. After paying, fill in date, amount and confirmation # in the PAYMENT block.",
        "",
        "Colours: yellow = input, green = numbers that go on the return, red = something to fix.",
        "The quarter of every row is calculated from its date; nothing needs to be moved between sections.",
        "A new fiscal year gets a new workbook (Settings > Fiscal year start).",
        "",
        DISCLAIMER,
    ], start=1):
        c = hw.cell(row=i, column=1, value=line)
        c.font = Font(bold=True, size=14) if i == 1 else Font()
        if line == DISCLAIMER:
            c.font = Font(bold=True, color="C00000")
            c.alignment = Alignment(wrap_text=True, vertical="top")
            hw.row_dimensions[i].height = 64

    wb.calculation.fullCalcOnLoad = True
    path.parent.mkdir(parents=True, exist_ok=True)
    _save(wb, path)
    print(f"built {path}")


# ------------------------------------------------------------- saving ---
IGNORED_ERRORS = ('<ignoredErrors><ignoredError sqref="A1:Z1000" numberStoredAsText="1" '
                  'formula="1" formulaRange="1"/></ignoredErrors>')
_AFTER_IGNORED = ("<smartTags", "<drawing", "<legacyDrawing", "<picture", "<oleObjects", "<controls",
                  "<webPublishItems", "<tableParts", "<extLst", "</worksheet>")


def _save(wb, path: Path) -> None:
    """Save, then tell Excel to skip its green-triangle checks (openpyxl cannot write ignoredErrors).

    The flagged patterns are deliberate: return-line rows use different formulas on purpose."""
    wb.save(path)
    tmp = path.with_name(path.name + ".tmp")
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("xl/worksheets/sheet") and item.filename.endswith(".xml"):
                xml = data.decode("utf-8")
                if "<ignoredErrors" not in xml:
                    pos = min(i for i in (xml.find(t) for t in _AFTER_IGNORED) if i >= 0)
                    xml = xml[:pos] + IGNORED_ERRORS + xml[pos:]
                data = xml.encode("utf-8")
            zout.writestr(item, data)
    tmp.replace(path)


# ----------------------------------------------------------- editing ---
def _guard(path: Path) -> None:
    if not path.exists():
        sys.exit(f"no such workbook: {path}")
    lock = path.with_name("~$" + path.name)
    if lock.exists():
        sys.exit(f"{path.name} is open in Excel (found {lock.name}); close it and retry")


def _backup(path: Path) -> Path:
    bdir = path.parent / "_backup"
    bdir.mkdir(exist_ok=True)
    dest = bdir / f"{path.stem}.{dt.datetime.now():%Y%m%d-%H%M%S}{path.suffix}"
    shutil.copy2(path, dest)
    return dest


def _date(v) -> dt.date:
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    s = str(v).strip()
    for fmt in ("%Y-%m-%d", "%m/%d/%Y"):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    sys.exit(f"bad date {v!r}: use YYYY-MM-DD or MM/DD/YYYY")


def _money(v) -> Decimal:
    d = Decimal(str(v).replace(",", "").replace("$", "").strip())
    if d != d.quantize(Decimal("0.01")):
        sys.exit(f"amount {v!r} has more than 2 decimals")
    return d


def _fy_start(wb) -> dt.date:
    return _date(wb["Settings"]["B6"].value)


def _quarter(d: dt.date, fy: dt.date):
    months = (d.year - fy.year) * 12 + d.month - fy.month - (1 if d.day < fy.day else 0)
    return months // 3 + 1 if 0 <= months < 12 else None


def _invoice_key(v) -> str:
    return str(v).strip().lstrip("#").removesuffix(".0")


def _read_block(ws, first, last, cols):
    out = []
    for r in range(first, last + 1):
        vals = {k: ws[f"{c}{r}"].value for k, c in cols.items()}
        if any(v not in (None, "") for v in vals.values()):
            out.append(vals)
    return out


SALES_COLS = {"date": "A", "invoice": "B", "customer": "C", "memo": "D", "project": "E",
              "amount": "F", "notes": "J"}
CAP_COLS = {"date": "A", "vendor": "B", "description": "C", "cost": "D", "gst": "E",
            "business": "F", "notes": "I"}


def _write_block(ws, first, last, cols, records, money_keys):
    if len(records) > last - first + 1:
        sys.exit("ledger is full; extend the template")
    for i in range(last - first + 1):
        r = first + i
        rec = records[i] if i < len(records) else {}
        for k, c in cols.items():
            v = rec.get(k)
            if k == "date" and v not in (None, ""):
                v = dt.datetime.combine(_date(v), dt.time())
            elif k in money_keys and v not in (None, ""):
                v = float(_money(v))
            elif k == "invoice" and v not in (None, ""):
                key = _invoice_key(v)
                v = int(key) if key.isdigit() else key
            ws[f"{c}{r}"].value = v if v != "" else None


def rebuild(path: Path) -> None:
    """Rebuild the workbook on the current template, carrying every input across."""
    _guard(path)
    old = load_workbook(path)
    st = old["Settings"]
    settings = {name: st[f"B{row}"].value for row, _, name, *_ in SETTINGS}
    sales = _read_block(old["Sales"], SALES_FIRST, SALES_LAST, SALES_COLS)
    cap = _read_block(old["Capital"], CAP_FIRST, CAP_LAST, CAP_COLS)
    os_ = old["Summary"]
    label_row = {os_[f"A{r}"].value: r for r in range(1, os_.max_row + 1) if os_[f"A{r}"].value}
    labels = {k: lbl for k, lbl, _ in SUMMARY_LAYOUT}
    inputs = {}
    for key in INPUT_KEYS:
        r = label_row.get(labels[key])
        if r is None:
            sys.exit(f"cannot find '{labels[key]}' on the old Summary; refusing to drop inputs")
        inputs[key] = [os_[f"{col}{r}"].value for col in QCOLS]

    tmp = path.with_name(path.stem + ".rebuild" + path.suffix)
    tmp.unlink(missing_ok=True)
    build(tmp, _date(settings["FY_Start"]), settings["Business_Name"] or "", settings["GST_Number"] or "")
    wb = load_workbook(tmp)
    for row, _, name, *_ in SETTINGS:
        wb["Settings"][f"B{row}"].value = settings[name]
    _write_block(wb["Sales"], SALES_FIRST, SALES_LAST, SALES_COLS, sales, {"amount"})
    _write_block(wb["Capital"], CAP_FIRST, CAP_LAST, CAP_COLS, cap, {"cost", "gst"})
    R = summary_rows()
    for key, vals in inputs.items():
        for col, v in zip(QCOLS, vals):
            wb["Summary"][f"{col}{R[key]}"].value = v
    wb.calculation.fullCalcOnLoad = True
    _save(wb, tmp)
    bak = _backup(path)
    tmp.replace(path)
    print(f"rebuilt {path}: {len(sales)} invoices, {len(cap)} capital rows, summary inputs kept "
          f"(backup {bak.name})")


def record_payment(path: Path, quarter: int, date: str, amount: str, conf: str | None) -> None:
    _guard(path)
    if quarter not in (1, 2, 3, 4):
        sys.exit("quarter must be 1-4")
    wb = load_workbook(path)
    col, R, s = QCOLS[quarter - 1], summary_rows(), wb["Summary"]
    d, amt = _date(date), _money(amount)
    s[f"{col}{R['filed']}"].value = dt.datetime.combine(d, dt.time())
    s[f"{col}{R['paid']}"].value = float(amt)
    if conf:
        s[f"{col}{R['conf']}"].value = str(conf)
    bak = _backup(path)
    wb.calculation.fullCalcOnLoad = True
    _save(wb, path)
    print(f"Q{quarter}: paid {amt:,.2f} on {d}" + (f", confirmation {conf}" if conf else "") + f" (backup {bak.name})")


def set_business(path: Path, name: str | None, bn: str | None) -> None:
    _guard(path)
    wb = load_workbook(path)
    st = wb["Settings"]
    if name is not None:
        st["B4"].value = name
    if bn is not None:
        st["B5"].value = bn
    bak = _backup(path)
    _save(wb, path)
    print(f"business: {st['B4'].value!r}, GST account: {st['B5'].value!r} (backup {bak.name})")


def add_sales(path: Path, rows_file: Path, dry_run: bool) -> None:
    _guard(path)
    new = json.loads(rows_file.read_text())
    wb = load_workbook(path)
    ws = wb["Sales"]
    fy = _fy_start(wb)
    existing = _read_block(ws, SALES_FIRST, SALES_LAST, SALES_COLS)
    seen = {_invoice_key(e["invoice"]): e for e in existing if e.get("invoice") not in (None, "")}
    added, skipped, conflicts = [], [], []
    for rec in new:
        for req in ("date", "invoice", "amount"):
            if rec.get(req) in (None, ""):
                sys.exit(f"row missing {req}: {rec}")
        key = _invoice_key(rec["invoice"])
        if key in seen:
            have = seen[key]
            if _money(have["amount"]) != _money(rec["amount"]) or _date(have["date"]) != _date(rec["date"]):
                conflicts.append(f"invoice {key}: workbook has {_date(have['date'])} {_money(have['amount']):,.2f}, "
                                 f"new data says {_date(rec['date'])} {_money(rec['amount']):,.2f}")
            skipped.append(rec)
            continue
        if _quarter(_date(rec["date"]), fy) is None:
            sys.exit(f"invoice {key} dated {rec['date']} is outside the fiscal year starting {fy}")
        seen[key] = rec
        added.append(rec)
    _print_rows("ADD", added, fy, "amount")
    print(f"SKIP {len(skipped)} invoice(s) already in the workbook"
          + (f": {', '.join(_invoice_key(r['invoice']) for r in skipped)}" if skipped else ""))
    if conflicts:
        print("CONFLICT -- already-entered invoices disagree with the new data; nothing written:")
        for c in conflicts:
            print("  " + c)
        sys.exit(1)
    if not added:
        print("nothing to add")
        return
    if dry_run:
        print("dry run: workbook not changed")
        return
    allrows = existing + added
    allrows.sort(key=lambda e: (_date(e["date"]), _invoice_key(e["invoice"]).zfill(12)))
    bak = _backup(path)
    _write_block(ws, SALES_FIRST, SALES_LAST, SALES_COLS, allrows, {"amount"})
    wb.calculation.fullCalcOnLoad = True
    _save(wb, path)
    print(f"saved {path} (backup {bak.name})")


def add_capital(path: Path, rows_file: Path, dry_run: bool) -> None:
    _guard(path)
    new = json.loads(rows_file.read_text())
    wb = load_workbook(path)
    ws = wb["Capital"]
    fy = _fy_start(wb)
    for rec in new:
        for req in ("date", "description", "gst"):
            if rec.get(req) in (None, ""):
                sys.exit(f"row missing {req}: {rec}")
        if _quarter(_date(rec["date"]), fy) is None:
            sys.exit(f"{rec['description']} dated {rec['date']} is outside the fiscal year starting {fy}")
        rec.setdefault("business", "Y")
        rec["business"] = str(rec["business"]).upper()
    _print_rows("ADD", new, fy, "gst")
    if dry_run:
        print("dry run: workbook not changed")
        return
    existing = _read_block(ws, CAP_FIRST, CAP_LAST, CAP_COLS)
    allrows = sorted(existing + new, key=lambda e: _date(e["date"]))
    bak = _backup(path)
    _write_block(ws, CAP_FIRST, CAP_LAST, CAP_COLS, allrows, {"cost", "gst"})
    wb.calculation.fullCalcOnLoad = True
    _save(wb, path)
    print(f"saved {path} (backup {bak.name})")


def _print_rows(tag, rows, fy, amount_key):
    total = Decimal(0)
    for rec in rows:
        d = _date(rec["date"])
        amt = _money(rec[amount_key])
        total += amt
        label = rec.get("invoice") or rec.get("description")
        who = rec.get("project") or rec.get("customer") or rec.get("vendor") or ""
        print(f"{tag} Q{_quarter(d, fy)} {d}  {str(label):>8}  {who:<10} {amt:>12,.2f}")
    if rows:
        print(f"{tag} {len(rows)} row(s), total {total:,.2f}")


# ------------------------------------------------------------- check ---
def _round(x: Decimal, places="0.01") -> Decimal:
    return x.quantize(Decimal(places), rounding=ROUND_HALF_UP)


def expected(wb) -> dict:
    """Independent re-implementation of the Summary sheet from the raw inputs."""
    st = wb["Settings"]
    fy = _fy_start(wb)
    gst_rate = Decimal(str(st["B7"].value))
    qm = Decimal(str(st["B8"].value))
    cr = Decimal(str(st["B9"].value))
    thr = Decimal(str(st["B10"].value))
    q = {i: dict(count=0, net=Decimal(0), gst=Decimal(0), incl=Decimal(0), l106=Decimal(0))
         for i in range(1, 5)}
    for e in _read_block(wb["Sales"], SALES_FIRST, SALES_LAST, SALES_COLS):
        if e["amount"] in (None, ""):
            continue
        qi = _quarter(_date(e["date"]), fy)
        net = _money(e["amount"])
        g = _round(net * gst_rate)
        q[qi]["count"] += 1
        q[qi]["net"] += net
        q[qi]["gst"] += g
        q[qi]["incl"] += net + g
    for e in _read_block(wb["Capital"], CAP_FIRST, CAP_LAST, CAP_COLS):
        if e["gst"] in (None, "") or str(e.get("business") or "").upper() != "Y":
            continue
        q[_quarter(_date(e["date"]), fy)]["l106"] += _money(e["gst"])
    s = wb["Summary"]
    R = summary_rows()
    cum = Decimal(0)
    for i, col in enumerate(QCOLS, start=1):
        d = q[i]
        for key in ("l110", "l111", "l205", "l405", "house", "l104"):
            d[key] = _money(s[f"{col}{R[key]}"].value or 0)
        d["l101"] = _round(d["incl"], "1")
        d["l103"] = d["l101"] * qm
        d["l105"] = d["l103"] + d["l104"]
        d["l107"] = _round(cr * min(d["l101"], max(Decimal(0), thr - cum)))
        cum += d["l101"]
        d["l108"] = d["l106"] + d["l107"]
        d["l109"] = d["l105"] - d["l108"]
        d["l112"] = d["l110"] + d["l111"]
        d["l113a"] = d["l109"] - d["l112"]
        d["l113b"] = d["l205"] + d["l405"]
        d["l113c"] = d["l113a"] + d["l113b"]
        d["l114"] = -d["l113c"] if d["l113c"] < 0 else Decimal(0)
        d["l115"] = d["l113c"] if d["l113c"] > 0 else Decimal(0)
        d["paid"] = _money(s[f"{col}{R['paid']}"].value or 0)
        d["bal"] = _round(d["l115"]) - d["paid"]
    return q


CHECK_KEYS = ("count", "net", "gst", "incl", "l101", "l110", "l111", "l205", "l405", "house", "l109",
              "l114", "l103", "l104", "l106", "l107", "l105", "l108", "l113c", "l115", "paid", "bal")

RECALC_PS1 = r"""param([string]$In, [string]$Out)
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open($In, 0, $true)
  $xl.CalculateFull()
  # report Excel's green-triangle warnings: 3 number-as-text, 4 inconsistent formula, 5 omits adjacent cells
  foreach ($ws in $wb.Worksheets) {
    foreach ($c in $ws.Range("A1:J60").Cells) {
      foreach ($k in 3,4,5) {
        if ($c.Errors.Item($k).Value) { Write-Output ("FLAG " + $ws.Name + "!" + $c.Address($false,$false) + " type " + $k) }
      }
    }
  }
  $wb.SaveCopyAs($Out)
  $wb.Close($false)
} finally {
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
"""
CONFIG = Path.home() / ".config" / "gst-quick-method" / "config.json"


def _is_wsl() -> bool:
    try:
        return "microsoft" in Path("/proc/version").read_text().lower()
    except OSError:
        return False


def _powershell() -> str | None:
    """PowerShell 7 if present, else Windows PowerShell 5.1 (both can drive Excel via COM)."""
    # the Windows executables only: a Linux/macOS `pwsh` (or a shell wrapper) cannot reach Excel
    cands = [os.environ.get("GST_PWSH"), shutil.which("pwsh.exe"), shutil.which("powershell.exe")]
    if _is_wsl():
        cands += ["/mnt/c/Program Files/PowerShell/7/pwsh.exe",
                  "/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe"]
    return next((c for c in cands if c and Path(c).exists()), None)


def _excel_temp() -> tuple[Path, callable] | None:
    """A temp dir Excel can open, plus a function turning our paths into Windows paths; None if no Windows."""
    if sys.platform == "win32":
        return Path(tempfile.gettempdir()), str
    if _is_wsl() and shutil.which("cmd.exe") and shutil.which("wslpath"):
        out = subprocess.run(["cmd.exe", "/c", "echo %TEMP%"], capture_output=True, text=True,
                             cwd="/mnt/c", check=True).stdout.strip()
        tmp = subprocess.run(["wslpath", "-u", out], capture_output=True, text=True, check=True).stdout.strip()
        to_win = lambda q: subprocess.run(["wslpath", "-w", str(q)], capture_output=True, text=True,  # noqa: E731
                                          check=True).stdout.strip()
        return Path(tmp), to_win
    return None


def workbook_path(year: int) -> Path:
    tmpl = os.environ.get("GST_WORKBOOK")
    if not tmpl and CONFIG.exists():
        tmpl = json.loads(CONFIG.read_text()).get("workbook")
    if not tmpl:
        sys.exit(f"no workbook configured: set GST_WORKBOOK or create {CONFIG} "
                 '(e.g. {"workbook": "/path/{year}/GST Returns {year} - Quick Method.xlsx"})')
    return Path(tmpl.replace("{year}", str(year)))


def _serial_to_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, (int, float)) and v > 0:
        return dt.date(1899, 12, 30) + dt.timedelta(days=int(v))
    return None


def calc_python(path: Path) -> tuple[dict, list]:
    """Evaluate every formula with the pure-Python `formulas` engine. Returns ({(SHEET, A1): value}, flags)."""
    import contextlib
    import io
    import warnings
    warnings.filterwarnings("ignore")
    import formulas
    with contextlib.redirect_stderr(io.StringIO()):
        sol = formulas.ExcelModel().loads(str(path)).finish().calculate()
    vals = {}
    for k, v in sol.items():
        ref = str(k).split("]", 1)[-1]          # SHEET'!A1  (ranges contain ':')
        if "'!" not in ref or ":" in ref:
            continue
        sheet, cell = ref.split("'!")
        x = v.value[0][0] if hasattr(v, "value") else v
        if hasattr(x, "item"):
            x = x.item()
        kind = type(x).__name__
        if kind == "Token":                      # the engine's marker for a blank cell
            x = None
        elif kind == "XlError":
            x = str(x)
        vals[(sheet.upper(), cell)] = x
    return vals, []


def calc_excel(path: Path):
    """Recalculate a copy in real Excel via COM. Returns ({(SHEET, A1): value}, warning flags) or None."""
    where, ps = _excel_temp(), _powershell()
    if not where or not ps:
        return None
    tmp = Path(tempfile.mkdtemp(prefix="gstcheck-", dir=where[0]))
    try:
        (tmp / "recalc.ps1").write_text(RECALC_PS1)
        shutil.copy2(path, tmp / "in.xlsx")
        win = where[1](tmp)
        res = subprocess.run([ps, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File",
                              win + "\\recalc.ps1", "-In", win + "\\in.xlsx", "-Out", win + "\\out.xlsx"],
                             capture_output=True, text=True, timeout=300)
        if res.returncode != 0 or not (tmp / "out.xlsx").exists():
            return None                          # no Excel installed (COM failed): Python engine only
        wb = load_workbook(tmp / "out.xlsx", data_only=True)
        vals = {(ws.title.upper(), c.coordinate): c.value
                for ws in wb.worksheets for row in ws.iter_rows() for c in row if c.value is not None}
        flags = [ln.strip()[5:] for ln in res.stdout.splitlines() if ln.startswith("FLAG ")]
        return vals, flags
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def _verify(name: str, vals: dict, flags: list, exp: dict, show: bool) -> list:
    """Compare one engine's results with the independent calculation; print the table if show."""
    R = summary_rows()
    labels = {k: (f"{n} " if n else "") + lbl for k, lbl, n in SUMMARY_LAYOUT}
    get = lambda sheet, a: vals.get((sheet.upper(), a))  # noqa: E731
    num = lambda sheet, a: Decimal(str(get(sheet, a) or 0))  # noqa: E731
    problems = [f"{name}: Excel warning triangle at {f}" for f in flags]
    for (sheet, cell), v in vals.items():
        if isinstance(v, str) and v.startswith("#"):
            problems.append(f"{name}: {sheet}!{cell} = {v}")
    for ws_name, col, last in (("Sales", "I", SALES_LAST), ("Capital", "H", CAP_LAST)):
        for r in range(5, last + 1):
            if get(ws_name, f"{col}{r}") == "OUT":
                problems.append(f"{ws_name} row {r}: date outside fiscal year")
    out = [f"{'':<56}" + "".join(f"{'Q' + str(i):>13}" for i in range(1, 5)) + f"{'Year':>14}"]
    for key in CHECK_KEYS:
        row = R[key]
        cells = []
        for i, col in enumerate(QCOLS, start=1):
            got, want = num("Summary", f"{col}{row}"), Decimal(exp[i][key])
            if abs(got - want) > Decimal("0.005"):
                problems.append(f"{name}: {labels[key]} Q{i}: workbook {got} != independent {want}")
            cells.append(f"{got:>13,.2f}" if key != "count" else f"{int(got):>13}")
        yv = num("Summary", f"G{row}")
        if abs(yv - sum(Decimal(exp[i][key]) for i in range(1, 5))) > Decimal("0.005"):
            problems.append(f"{name}: {labels[key]} Year total mismatch")
        out.append(f"{labels[key][:56]:<56}" + "".join(cells)
                   + (f"{yv:>14,.2f}" if key != "count" else f"{int(yv):>14}"))
    due = [_serial_to_date(get("Summary", f"{c}{R['due']}")) for c in QCOLS]
    out.append(f"{'Filing and payment due':<56}" + "".join(f"{(d.isoformat() if d else ''):>13}" for d in due))
    out.append("")
    for row, key, label in ((5, "incl", "Gross income (incl. GST)"), (6, "net", "Net income (excl. GST)"),
                            (7, "gst", "GST charged to clients")):
        got, want = num("Summary", f"J{row}"), sum(exp[i][key] for i in range(1, 5))
        if abs(got - want) > Decimal("0.005"):
            problems.append(f"{name}: YTD {label}: workbook {got} != independent {want}")
        out.append(f"YTD {label:<52}{got:>13,.2f}")
    thru = _serial_to_date(get("Summary", "J9"))
    out.append(f"YTD {'through (latest invoice)':<52}{thru.isoformat() if thru else '':>13}")
    if show:
        print("\n".join(out))
    return problems


def check(path: Path, engine: str) -> None:
    exp = expected(load_workbook(path))
    results = []
    if engine in ("auto", "python"):
        results.append(("Python engine", *calc_python(path)))
    if engine in ("auto", "excel"):
        xl = calc_excel(path)
        if xl:
            results.append(("Excel", *xl))
        elif engine == "excel":
            sys.exit("Excel is not reachable (needs Windows or WSL with Excel installed)")
    problems = []
    for i, (name, vals, flags) in enumerate(results):
        problems += _verify(name, vals, flags, exp, show=(i == 0))
    engines = " and ".join(n for n, *_ in results)
    if problems:
        print("\nPROBLEMS:")
        for p in problems:
            print("  " + p)
        sys.exit(1)
    extra = ("" if any(n == "Excel" for n, *_ in results) or engine == "python"
             else " (Excel not reachable: formulas checked in Python only)")
    verb = "match" if len(results) > 1 else "matches"
    print(f"\nOK: {engines} {verb} the independent calculation on every line; no formula errors"
          + (" or warnings" if any(n == "Excel" for n, *_ in results) else "") + f".{extra}")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    b = sub.add_parser("build")
    b.add_argument("workbook", type=Path)
    b.add_argument("--fy-start", default=f"{dt.date.today().year}-01-01")
    b.add_argument("--business", default="")
    b.add_argument("--bn", default="")
    for name in ("add-sales", "add-capital"):
        p = sub.add_parser(name)
        p.add_argument("workbook", type=Path)
        p.add_argument("rows", type=Path, help="JSON list of rows")
        p.add_argument("--dry-run", action="store_true")
    rp = sub.add_parser("record-payment", help="record a quarter's filing/payment on Summary")
    rp.add_argument("workbook", type=Path)
    rp.add_argument("--quarter", type=int, required=True)
    rp.add_argument("--date", required=True)
    rp.add_argument("--amount", required=True)
    rp.add_argument("--confirmation")
    sb = sub.add_parser("set-business", help="set business name and/or GST account number")
    sb.add_argument("workbook", type=Path)
    sb.add_argument("--name")
    sb.add_argument("--bn")
    pp = sub.add_parser("path", help="print the configured workbook path")
    pp.add_argument("year", type=int, nargs="?", default=dt.date.today().year)
    rb = sub.add_parser("rebuild", help="move all inputs onto the current template (after a layout change)")
    rb.add_argument("workbook", type=Path)
    c = sub.add_parser("check")
    c.add_argument("workbook", type=Path)
    c.add_argument("--engine", choices=["auto", "python", "excel"], default="auto",
                   help="auto: Python engine, plus real Excel when reachable (default)")
    a = ap.parse_args()
    if a.cmd == "build":
        build(a.workbook, _date(a.fy_start), a.business, a.bn)
    elif a.cmd == "add-sales":
        add_sales(a.workbook, a.rows, a.dry_run)
    elif a.cmd == "add-capital":
        add_capital(a.workbook, a.rows, a.dry_run)
    elif a.cmd == "record-payment":
        record_payment(a.workbook, a.quarter, a.date, a.amount, a.confirmation)
    elif a.cmd == "set-business":
        set_business(a.workbook, a.name, a.bn)
    elif a.cmd == "path":
        print(workbook_path(a.year))
    elif a.cmd == "rebuild":
        rebuild(a.workbook)
    else:
        check(a.workbook, a.engine)


if __name__ == "__main__":
    main()
