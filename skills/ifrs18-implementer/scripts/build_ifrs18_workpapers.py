#!/usr/bin/env python3
"""Build an IFRS 18 implementation working-paper file (.xlsx) with live formulas.

Usage
-----
  python build_ifrs18_workpapers.py --out IFRS18_Workpapers.xlsx              # demo data
  python build_ifrs18_workpapers.py --tb my_tb.csv --setup setup.json --out X.xlsx

The trial balance CSV follows assets/tb_template.csv (debit +, credit -).
The setup JSON overrides keys of demo_data.SETUP (entity, labels, tax rate, NCI, EPS inputs...).
Lines used in the TB's `ifrs18_line` column must exist in the line catalogue below
(sheet "Lists"); add new lines to CATALOGUE when an entity needs them.

After building, recalculate (LibreOffice) and confirm sheet 19_Checks shows "ALL CHECKS OK".
"""
import argparse
import csv
import json
import os
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import demo_data  # noqa: E402

# --------------------------------------------------------------------------------------
# Line catalogue: key (unique), display label, section, IFRS 18 reference
# --------------------------------------------------------------------------------------
OP, INV, FIN, TAX, DISC = "Operating", "Investing", "Financing", "Income taxes", "Discontinued operations"
OCINR, OCIR = "OCI - not reclassified", "OCI - reclassifiable"
NCA, CA, EQ, NCL, CL = "Non-current assets", "Current assets", "Equity", "Non-current liabilities", "Current liabilities"

CATALOGUE = [
    # Operating
    ("Revenue", "Revenue", OP, "IFRS 18.75(a)(i)"),
    ("Cost of sales", "Cost of sales", OP, "IFRS 18.75(a)(ii), 78-82"),
    ("Selling and distribution expenses", "Selling and distribution expenses", OP, "IFRS 18.78-82, B85"),
    ("Administrative expenses", "Administrative expenses", OP, "IFRS 18.78-82, B85"),
    ("Research and development expenses", "Research and development expenses", OP, "IFRS 18.78-82; IAS 38.126"),
    ("Employee benefits expense", "Employee benefits expense", OP, "IFRS 18.78-80 (by nature)"),
    ("Depreciation and amortisation", "Depreciation and amortisation", OP, "IFRS 18.78-80 (by nature)"),
    ("Raw materials and consumables used", "Raw materials and consumables used", OP, "IFRS 18.78-80 (by nature)"),
    ("Impairment losses on goodwill and remeasurement of disposal group",
     "Impairment losses on goodwill and remeasurement of disposal group", OP, "IFRS 18.75(a)(ii), 78-80, B79-B81"),
    ("Impairment loss on trade and other receivables and contract assets",
     "Impairment loss on trade and other receivables and contract assets", OP, "IFRS 18.75(b)(ii)"),
    ("Foreign exchange differences", "Foreign exchange differences", OP, "IFRS 18.24, B65-B67"),
    ("Other operating income", "Other operating income", OP, "IFRS 18.24, B24-B26"),
    ("Other operating expenses, including flood-related expenses",
     "Other operating expenses, including flood-related expenses", OP, "IFRS 18.24, B24-B26"),
    # Investing
    ("Share of profit of equity-accounted investees, net of tax",
     "Share of profit of equity-accounted investees, net of tax", INV, "IFRS 18.53(a), 75(a)(iii)"),
    ("Income from investment properties", "Income from investment properties", INV, "IFRS 18.53(c), B47"),
    ("Interest and other investing income", "Interest and other investing income", INV, "IFRS 18.53(b)-(c), 75(b), B46-B47"),
    ("Remeasurement to fair value of pre-existing interest in an acquiree",
     "Remeasurement to fair value of pre-existing interest in an acquiree", INV, "IFRS 18.24, B78"),
    ("Financial assets at FVTPL - net change in fair value",
     "Financial assets at FVTPL - net change in fair value", INV, "IFRS 18.53(c); IFRS 7.20(a)(i)"),
    # Financing
    ("Interest and other financing expenses", "Interest and other financing expenses", FIN, "IFRS 18.59-61, B49-B54"),
    # Tax and discontinued
    ("Income tax expense", "Income tax expense", TAX, "IFRS 18.67, 75(a)(iv)"),
    ("Profit (loss) from discontinued operation, net of tax",
     "Profit (loss) from discontinued operation, net of tax", DISC, "IFRS 18.68, 75(a)(v); IFRS 5.33(a)"),
    # OCI - not reclassified
    ("Revaluation of property, plant and equipment", "Revaluation of property, plant and equipment", OCINR, "IFRS 18.88(b)"),
    ("Remeasurements of defined benefit liability (asset)", "Remeasurements of defined benefit liability (asset)", OCINR, "IFRS 18.88(b)"),
    ("Equity investments at FVOCI - net change in fair value", "Equity investments at FVOCI - net change in fair value", OCINR, "IFRS 7.20(a)(vii)"),
    ("Equity-accounted investees - share of OCI (not reclassified)", "Equity-accounted investees - share of OCI", OCINR, "IFRS 18.89(a)"),
    ("Related tax - items not reclassified", "Related tax", OCINR, "IFRS 18.93-95"),
    # OCI - reclassifiable
    ("Foreign operations - foreign currency translation differences",
     "Foreign operations - foreign currency translation differences", OCIR, "IFRS 18.88(a); IAS 21.52(b)"),
    ("Net investment hedge - net loss", "Net investment hedge - net loss", OCIR, "IFRS 18.24"),
    ("Equity-accounted investees - share of OCI (reclassifiable)", "Share of OCI of equity-accounted investees", OCIR, "IFRS 18.89(a)"),
    ("Reclassification of foreign currency differences on loss of significant influence",
     "Reclassification of foreign currency differences on loss of significant influence", OCIR, "IFRS 18.90"),
    ("Cash flow hedges - effective portion of changes in fair value",
     "Cash flow hedges - effective portion of changes in fair value", OCIR, "IFRS 7.24C(b)(i)"),
    ("Cash flow hedges - reclassified to profit or loss", "Cash flow hedges - reclassified to profit or loss", OCIR, "IFRS 18.90"),
    ("Cost of hedging reserve - changes in fair value", "Cost of hedging reserve - changes in fair value", OCIR, "IFRS 18.24"),
    ("Cost of hedging reserve - reclassified to profit or loss", "Cost of hedging reserve - reclassified to profit or loss", OCIR, "IFRS 18.90"),
    ("Debt investments at FVOCI - net change in fair value", "Debt investments at FVOCI - net change in fair value", OCIR, "IFRS 7.20(a)(viii)"),
    ("Debt investments at FVOCI - reclassified to profit or loss", "Debt investments at FVOCI - reclassified to profit or loss", OCIR, "IFRS 18.90"),
    ("Related tax - items that are or may be reclassified", "Related tax", OCIR, "IFRS 18.93-95"),
    # SFP - non-current assets
    ("Property, plant and equipment", "Property, plant and equipment", NCA, "IFRS 18.103(a)"),
    ("Intangible assets (non-current)", "Intangible assets", NCA, "IFRS 18.103(c)"),
    ("Goodwill", "Goodwill", NCA, "IFRS 18.103(d)"),
    ("Biological assets", "Biological assets", NCA, "IFRS 18.103(h)"),
    ("Investment property", "Investment property", NCA, "IFRS 18.103(b)"),
    ("Equity-accounted investees", "Equity-accounted investees", NCA, "IFRS 18.103(g)"),
    ("Financial assets (non-current)", "Financial assets", NCA, "IFRS 18.103(e)"),
    ("Deferred tax assets", "Deferred tax assets", NCA, "IFRS 18.103(r), 98"),
    ("Employee benefits (asset)", "Employee benefits", NCA, "IFRS 18.24, B109-B110"),
    # SFP - current assets
    ("Intangible assets (current)", "Intangible assets", CA, "IFRS 18.103(c)"),
    ("Inventories", "Inventories", CA, "IFRS 18.103(i)"),
    ("Contract assets", "Contract assets", CA, "IFRS 18.24; IFRS 15.105"),
    ("Financial assets (current)", "Financial assets", CA, "IFRS 18.103(e)"),
    ("Current tax assets", "Current tax assets", CA, "IFRS 18.103(q)"),
    ("Trade and other receivables", "Trade and other receivables", CA, "IFRS 18.103(j)"),
    ("Prepayments", "Prepayments", CA, "IFRS 18.24, B109-B110"),
    ("Cash and cash equivalents", "Cash and cash equivalents", CA, "IFRS 18.103(k)"),
    ("Assets held for sale", "Assets held for sale", CA, "IFRS 18.103(l); IFRS 5.38"),
    # SFP - equity
    ("Share capital", "Share capital", EQ, "IFRS 18.104(b)"),
    ("Share premium", "Share premium", EQ, "IFRS 18.24, B111(f)"),
    ("Reserves", "Reserves", EQ, "IFRS 18.104(b)"),
    ("Retained earnings", "Retained earnings", EQ, "IFRS 18.24, B111(f)"),
    ("Non-controlling interests", "Non-controlling interests", EQ, "IFRS 18.104(a)"),
    # SFP - non-current liabilities
    ("Loans and borrowings (non-current)", "Loans and borrowings", NCL, "IFRS 18.103(o)"),
    ("Employee benefits (non-current liability)", "Employee benefits", NCL, "IFRS 18.24, B111(e)"),
    ("Trade and other payables (non-current)", "Trade and other payables", NCL, "IFRS 18.103(m)"),
    ("Deferred income (non-current)", "Deferred income", NCL, "IFRS 18.24"),
    ("Provisions (non-current)", "Provisions", NCL, "IFRS 18.103(n)"),
    ("Deferred tax liabilities", "Deferred tax liabilities", NCL, "IFRS 18.103(r), 98"),
    # SFP - current liabilities
    ("Bank overdraft", "Bank overdraft", CL, "IFRS 18.24"),
    ("Current tax liabilities", "Current tax liabilities", CL, "IFRS 18.103(q)"),
    ("Loans and borrowings (current)", "Loans and borrowings", CL, "IFRS 18.103(o)"),
    ("Employee benefits (current liability)", "Employee benefits", CL, "IFRS 18.24, B111(e)"),
    ("Trade and other payables (current)", "Trade and other payables", CL, "IFRS 18.103(m)"),
    ("Contract liabilities", "Contract liabilities", CL, "IFRS 18.24; IFRS 15.105"),
    ("Deferred income (current)", "Deferred income", CL, "IFRS 18.24"),
    ("Provisions (current)", "Provisions", CL, "IFRS 18.103(n)"),
    ("Liabilities directly associated with assets held for sale",
     "Liabilities directly associated with assets held for sale", CL, "IFRS 18.103(s); IFRS 5.38"),
]

# IAS 1 captions used in the previously published statement (comparatives) and their position
IAS1_LINES = [
    ("Revenue", "Above operating profit"),
    ("Cost of sales", "Above operating profit"),
    ("Selling and distribution expenses", "Above operating profit"),
    ("Administrative expenses", "Above operating profit"),
    ("Research and development expenses", "Above operating profit"),
    ("Impairment loss on trade receivables and contract assets", "Above operating profit"),
    ("Other income", "Above operating profit"),
    ("Other expenses", "Above operating profit"),
    ("Finance income", "Net finance costs"),
    ("Finance costs", "Net finance costs"),
    ("Share of profit of equity-accounted investees, net of tax", "Below net finance costs"),
    ("Income tax expense", "Income tax"),
    ("Profit (loss) from discontinued operation, net of tax", "Discontinued operation"),
]

# IFRS 18 line -> IAS 1 caption it replaces (pairs rows in the C3 reconciliation; extend per entity)
CAPTION_EQUIV = {
    "Impairment loss on trade and other receivables and contract assets": "Impairment loss on trade receivables and contract assets",
    "Other operating income": "Other income",
    "Other operating expenses, including flood-related expenses": "Other expenses",
    "Interest and other financing expenses": "Finance costs",
}

NATURES5 = ["Depreciation", "Amortisation", "Employee benefits", "Impairment losses (IAS 36)", "Inventory write-downs (IAS 2)"]
NATURE_LIST = NATURES5 + ["Raw materials and other costs", "Other remeasurements", ""]

# --------------------------------------------------------------------------------------
# Styling helpers
# --------------------------------------------------------------------------------------
NAVY = "1F3864"
F_TITLE = Font(name="Calibri", size=14, bold=True, color=NAVY)
F_SUB = Font(name="Calibri", size=10, italic=True, color="595959")
F_HDR = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
F_BOLD = Font(name="Calibri", size=10, bold=True)
F_NORM = Font(name="Calibri", size=10)
F_INPUT = Font(name="Calibri", size=10, color="0000FF")
F_LINK = Font(name="Calibri", size=10, color="0563C1", underline="single")
FILL_HDR = PatternFill("solid", fgColor=NAVY)
FILL_SEC = PatternFill("solid", fgColor="D9E1F2")
FILL_TOT = PatternFill("solid", fgColor="F2F2F2")
FILL_INPUT = PatternFill("solid", fgColor="FFF2CC")
FILL_KEY = PatternFill("solid", fgColor="E2EFDA")
THIN = Side(style="thin", color="BFBFBF")
B_ALL = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
B_TOP = Border(top=Side(style="thin", color="000000"))
B_TOT = Border(top=Side(style="thin", color="000000"), bottom=Side(style="double", color="000000"))
NUM = '#,##0;(#,##0);"-"'
NUM2 = '#,##0.00;(#,##0.00);"-"'
PCT = "0.0%"
WRAP = Alignment(wrap_text=True, vertical="top")


def q(sheet):
    return f"'{sheet}'"


def title(ws, text, sub, wpref):
    ws["A1"] = text
    ws["A1"].font = F_TITLE
    ws["A2"] = sub
    ws["A2"].font = F_SUB
    ws["A3"] = f"WP ref: {wpref}   |   Entity: "
    ws["A3"].font = F_SUB
    ws["B3"] = "='01_Setup'!C5"
    ws["B3"].font = F_SUB
    ws.sheet_view.showGridLines = False


def header(ws, row, labels, col=1):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=lab)
        c.font = F_HDR
        c.fill = FILL_HDR
        c.alignment = Alignment(wrap_text=True, vertical="center")
        c.border = B_ALL


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def put(ws, row, col, value, font=F_NORM, fmt=None, fill=None, border=None, align=None):
    c = ws.cell(row=row, column=col, value=value)
    c.font = font
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    if border:
        c.border = border
    if align:
        c.alignment = align
    return c


def inp(ws, row, col, value, fmt=None):
    return put(ws, row, col, value, font=F_INPUT, fmt=fmt, fill=FILL_INPUT, border=B_ALL)


def add_list_validation(ws, rng, formula):
    dv = DataValidation(type="list", formula1=formula, allow_blank=True)
    dv.error = "Choose a value from the list"
    ws.add_data_validation(dv)
    dv.add(rng)


# --------------------------------------------------------------------------------------
# Input loading
# --------------------------------------------------------------------------------------
def load_tb(path):
    rows = []
    with open(path, newline="", encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            rows.append((r["account"], r["description"], r["type"].strip().upper(),
                         float(r["cy"] or 0), float(r["py"] or 0), r.get("ias1_line", ""),
                         r.get("ifrs18_line", ""), r.get("nature", ""), r.get("mpm_item", ""),
                         r.get("transition_ref", "")))
    return rows


# --------------------------------------------------------------------------------------
# Builder
# --------------------------------------------------------------------------------------
class Builder:
    def __init__(self, tb, setup, cashflow, capitalised, explanations):
        self.tb = tb
        self.s = setup
        self.cf = cashflow
        self.cap = capitalised
        self.expl = explanations
        self.wb = Workbook()
        self.wb.remove(self.wb.active)
        self.cat = {k: (lab, sec, ref) for k, lab, sec, ref in CATALOGUE}
        self.used = {r[6] for r in tb}
        unknown = sorted(x for x in self.used if x and x not in self.cat)
        if unknown:
            print("WARNING - TB lines not in CATALOGUE (shown as 'NOT IN CATALOGUE' in 05_TB):", unknown)
        self.pos = {}      # line key -> (sheet, row) of its CY amount (col D) on a statement
        self.setup_ref = {}
        self.checks = []   # (description, formula_actual, formula_expected, wp)
        # sheet order
        names = ["00_Index", "01_Setup", "02_Plan", "03_Gap", "04_SMBA", "05_TB", "06_Rules",
                 "07_Classification", "08_SPLOCI", "09_Transition", "10_ExpNature", "11_Aggregation",
                 "12_MPM_ID", "13_MPM_Recon", "14_SFP", "15_CashFlow", "16_Disclosure", "17_Reclass",
                 "18_Issues", "19_Checks", "Lists"]
        self.ws = {n: self.wb.create_sheet(n) for n in names}
        self.n_tb = len(tb)
        self.TB_FIRST = 6
        self.TB_LAST = 6 + max(self.n_tb, 1) + 300   # spare rows for additions

    # ---------- range helpers ----------
    def tbr(self, col):
        return f"'05_TB'!${col}${self.TB_FIRST}:${col}${self.TB_LAST}"

    def lines_in(self, section):
        """Catalogue keys of a section that are used in the TB (keeps catalogue order)."""
        return [k for k, lab, sec, ref in CATALOGUE if sec == section and k in self.used]

    # ---------- 01 Setup ----------
    def setup(self):
        ws = self.ws["01_Setup"]
        title(ws, "Engagement set-up and key assumptions", "Blue cells on yellow are inputs; everything else is formula-driven.", "01")
        s = self.s
        rows = [
            ("entity", "Entity / group name", s["entity"], None),
            ("cy_label", "Current period label", s["cy_label"], None),
            ("py_label", "Comparative period label", s["py_label"], None),
            ("cy_end", "Current reporting date", s["cy_end"], None),
            ("py_end", "Comparative reporting date", s["py_end"], None),
            ("currency", "Presentation currency", s["currency"], None),
            ("units", "Level of rounding", s["units"], None),
            ("date_initial_application", "Date of initial application of IFRS 18 (C1-C2)", s["date_initial_application"], None),
            ("early_adoption", "Early adoption? (disclose if Yes - C1)", s["early_adoption"], None),
            ("smba_investing", "SMBA: investing in particular types of assets? (see 04_SMBA)", s["smba_investing"], None),
            ("smba_financing", "SMBA: providing financing to customers? (see 04_SMBA)", s["smba_financing"], None),
            ("expense_method", "Operating expenses presented by (IFRS 18.78)", s["expense_method"], None),
            ("tax_rate", "Statutory income tax rate (default for MPM tax effects)", s["tax_rate"], PCT),
            ("nci_profit_cy", "Profit attributable to NCI - current", s["nci_profit_cy"], NUM),
            ("nci_profit_py", "Profit attributable to NCI - comparative", s["nci_profit_py"], NUM),
            ("nci_oci_cy", "OCI attributable to NCI - current", s["nci_oci_cy"], NUM),
            ("nci_oci_py", "OCI attributable to NCI - comparative", s["nci_oci_py"], NUM),
            ("pref_div_cy", "Preference dividends (EPS numerator) - current", s["pref_div_cy"], NUM),
            ("pref_div_py", "Preference dividends (EPS numerator) - comparative", s["pref_div_py"], NUM),
            ("diluted_adj_cy", "Diluted earnings adjustment - current", s["diluted_adj_cy"], NUM),
            ("diluted_adj_py", "Diluted earnings adjustment - comparative", s["diluted_adj_py"], NUM),
            ("wans_basic_cy", "Weighted-average shares, basic (000) - current", s["wans_basic_cy"], NUM),
            ("wans_basic_py", "Weighted-average shares, basic (000) - comparative", s["wans_basic_py"], NUM),
            ("wans_diluted_cy", "Weighted-average shares, diluted (000) - current", s["wans_diluted_cy"], NUM),
            ("wans_diluted_py", "Weighted-average shares, diluted (000) - comparative", s["wans_diluted_py"], NUM),
            ("owners_share_mpm", "Share of MPM attributable to ordinary shareholders (IAS 33.73C)", s["owners_share_mpm"], PCT),
            ("materiality", "Overall materiality", s["materiality"], NUM),
            ("performance_materiality", "Performance materiality (drives 11_Aggregation flags)", s["performance_materiality"], NUM),
            ("trivial", "Clearly trivial threshold", s["trivial"], NUM),
            ("preparer", "Prepared by", s.get("preparer", ""), None),
            ("reviewer", "Reviewed by", s.get("reviewer", ""), None),
        ]
        header(ws, 4, ["Key", "Item", "Value", "Guidance"])
        guide = {
            "date_initial_application": "Annual periods beginning on or after 1 January 2027; retrospective application with restated comparatives.",
            "smba_investing": "Matter of fact, not an assertion. Conclude in 04_SMBA; it changes classification of investing items.",
            "smba_financing": "If Yes, apply IFRS 18.65 policy choice for financing-category items and reconsider PBFIT subtotal (IFRS 18.69(b), 73).",
            "expense_method": "Nature, Function or Mixed - whichever provides the most useful structured summary (IFRS 18.78, B80-B85).",
            "tax_rate": "IFRS 18.B141 permits statutory rate, reasonable pro-rata allocation or another more appropriate method - disclose which.",
        }
        for i, (key, lab, val, fmt) in enumerate(rows):
            r = 5 + i
            put(ws, r, 1, key, font=F_SUB)
            put(ws, r, 2, lab)
            inp(ws, r, 3, val, fmt)
            put(ws, r, 4, guide.get(key, ""), font=F_SUB, align=WRAP)
            self.setup_ref[key] = f"'01_Setup'!$C${r}"
        add_list_validation(ws, f"C{5+8}:C{5+10}", '"Yes,No"')
        add_list_validation(ws, f"C{5+11}", '"Nature,Function,Mixed"')
        widths(ws, [22, 58, 40, 90])
        ws.freeze_panes = "A5"

    def S(self, key):
        return self.setup_ref[key]

    # ---------- Lists ----------
    def lists(self):
        ws = self.ws["Lists"]
        ws["A1"] = "Reference lists (drive formulas and drop-downs - extend with care)"
        ws["A1"].font = F_BOLD
        header(ws, 3, ["IFRS 18 line key", "Display label", "Section / category", "IFRS 18 reference"])
        for i, (k, lab, sec, ref) in enumerate(CATALOGUE):
            for j, v in enumerate((k, lab, sec, ref)):
                put(ws, 4 + i, 1 + j, v)
        self.cat_rng = f"Lists!$A$4:$C${3 + len(CATALOGUE)}"
        self.cat_keys = f"Lists!$A$4:$A${3 + len(CATALOGUE)}"
        header(ws, 3, ["IAS 1 caption", "Position under IAS 1"], col=6)
        for i, (lab, grp) in enumerate(IAS1_LINES):
            put(ws, 4 + i, 6, lab)
            put(ws, 4 + i, 7, grp)
        self.ias1_keys = f"Lists!$F$4:$F${3 + len(IAS1_LINES)}"
        header(ws, 3, ["Nature tag"], col=9)
        for i, n in enumerate(NATURE_LIST[:-1]):
            put(ws, 4 + i, 9, n)
        self.nature_rng = f"Lists!$I$4:$I${3 + len(NATURE_LIST) - 1}"
        header(ws, 3, ["Status"], col=11)
        for i, v in enumerate(["Not started", "In progress", "Complete", "N/A"]):
            put(ws, 4 + i, 11, v)
        self.status_rng = "Lists!$K$4:$K$7"
        header(ws, 3, ["Transition ref"], col=13)
        for i, e in enumerate(self.expl):
            put(ws, 4 + i, 13, e[0])
        widths(ws, [55, 50, 24, 34, 2, 52, 24, 2, 30, 2, 14, 2, 14])

    # ---------- 05 Trial balance ----------
    def trial_balance(self):
        ws = self.ws["05_TB"]
        title(ws, "Trial balance and IFRS 18 mapping", "Debit +, credit -. Map every account: IAS 1 caption (comparatives as published), "
              "IFRS 18 line, nature tag (for IFRS 18.83), MPM reconciling item, transition reference.", "05")
        hdr = ["Account", "Description", "Type (PL/OCI/BS)", "=\"CY \"&'01_Setup'!C6", "=\"PY \"&'01_Setup'!C7",
               "IAS 1 caption (as previously presented)", "IFRS 18 line", "IFRS 18 category (formula)",
               "Nature tag (IFRS 18.83)", "MPM reconciling item", "Transition ref", "Mapping status (formula)"]
        header(ws, 5, hdr)
        for i, r in enumerate(self.tb):
            row = self.TB_FIRST + i
            acct, desc, typ, cy, py, ias1, l18, nat, mpm, tref = r
            put(ws, row, 1, acct, font=F_INPUT)
            put(ws, row, 2, desc, font=F_INPUT)
            put(ws, row, 3, typ, font=F_INPUT)
            put(ws, row, 4, cy, font=F_INPUT, fmt=NUM)
            put(ws, row, 5, py, font=F_INPUT, fmt=NUM)
            put(ws, row, 6, ias1 or None, font=F_INPUT)
            put(ws, row, 7, l18, font=F_INPUT)
            put(ws, row, 8, f'=IF(G{row}="","",IFERROR(VLOOKUP(G{row},{self.cat_rng},3,FALSE),"NOT IN CATALOGUE"))')
            put(ws, row, 9, nat or None, font=F_INPUT)
            put(ws, row, 10, mpm or None, font=F_INPUT)
            put(ws, row, 11, tref or None, font=F_INPUT)
            bs_cat = f'OR(H{row}="{NCA}",H{row}="{CA}",H{row}="{EQ}",H{row}="{NCL}",H{row}="{CL}")'
            oci_cat = f'LEFT(H{row},3)="OCI"'
            conflict = (f'OR(AND(C{row}="PL",OR({oci_cat},{bs_cat})),AND(C{row}="OCI",NOT({oci_cat})),'
                        f'AND(C{row}="BS",NOT({bs_cat})))')
            put(ws, row, 12, f'=IF(A{row}="","",IF(G{row}="","UNMAPPED",IF(H{row}="NOT IN CATALOGUE","CHECK LINE",'
                             f'IF(AND(C{row}="PL",F{row}=""),"NO IAS 1 CAPTION",IF({conflict},"TYPE/CATEGORY CONFLICT","OK")))))')
        last = self.TB_FIRST + self.n_tb - 1
        tot = last + 2
        put(ws, tot, 2, "Trial balance total (must be nil)", font=F_BOLD)
        for col in ("D", "E"):
            put(ws, tot, " DE".index(col) + 3, f"=SUM({col}{self.TB_FIRST}:{col}{last})", font=F_BOLD, fmt=NUM, border=B_TOT)
        self.tb_total = (f"'05_TB'!D{tot}", f"'05_TB'!E{tot}")
        put(ws, tot + 1, 2, "Note: insert new accounts inside the data block (above this total) so SUMIFS ranges pick them up.", font=F_SUB)
        add_list_validation(ws, f"C{self.TB_FIRST}:C{last}", '"PL,OCI,BS"')
        add_list_validation(ws, f"G{self.TB_FIRST}:G{last}", f"={self.cat_keys}")
        add_list_validation(ws, f"F{self.TB_FIRST}:F{last}", f"={self.ias1_keys}")
        add_list_validation(ws, f"I{self.TB_FIRST}:I{last}", f"={self.nature_rng}")
        ws.conditional_formatting.add(f"L{self.TB_FIRST}:L{last}",
                                      FormulaRule(formula=[f'AND(L{self.TB_FIRST}<>"OK",L{self.TB_FIRST}<>"")'],
                                                  fill=PatternFill("solid", fgColor="F8CBAD")))
        widths(ws, [9, 52, 9, 12, 12, 40, 52, 22, 26, 30, 9, 22])
        ws.freeze_panes = "C6"
        ws.auto_filter.ref = f"A5:L{last}"
        self.checks.append(("Trial balance balances - current period", self.tb_total[0], "0", "05_TB"))
        self.checks.append(("Trial balance balances - comparative period", self.tb_total[1], "0", "05_TB"))
        self.checks.append(("Accounts not fully mapped (count)", f'COUNTIFS({self.tbr("L")},"<>OK",{self.tbr("A")},"<>")', "0", "05_TB"))

    # ---------- 08 SPLOCI ----------
    def sploci(self):
        ws = self.ws["08_SPLOCI"]
        title(ws, "Statement of profit or loss and other comprehensive income (IFRS 18)",
              "One-statement approach (IFRS 18.12). All amounts pulled from 05_TB by IFRS 18 line. Comparatives restated (IFRS 18.C2).", "08")
        header(ws, 5, ["IFRS 18 reference", "Line item", "Category", "=\"CY \"&'01_Setup'!C6", "=\"PY \"&'01_Setup'!C7 &\" restated\"", "", "Line key"])
        cy, py = self.tbr("D"), self.tbr("E")
        r = 6
        self.spl_rows = {}

        def line(key, sign=-1):
            nonlocal r
            lab, sec, ref = self.cat[key]
            put(ws, r, 1, ref, font=F_SUB)
            put(ws, r, 2, lab)
            put(ws, r, 3, sec, font=F_SUB)
            s = "-" if sign < 0 else ""
            put(ws, r, 4, f"={s}SUMIFS({cy},{self.tbr('G')},$G{r})", fmt=NUM)
            put(ws, r, 5, f"={s}SUMIFS({py},{self.tbr('G')},$G{r})", fmt=NUM)
            put(ws, r, 7, key, font=F_SUB)
            self.pos[key] = ("08_SPLOCI", r)
            self.spl_rows[key] = r
            r += 1
            return r - 1

        def subtotal(label, ref, formula_cy, formula_py, strong=False):
            nonlocal r
            put(ws, r, 1, ref, font=F_SUB)
            put(ws, r, 2, label, font=F_BOLD)
            put(ws, r, 4, formula_cy, font=F_BOLD, fmt=NUM, border=B_TOT if strong else B_TOP, fill=FILL_TOT)
            put(ws, r, 5, formula_py, font=F_BOLD, fmt=NUM, border=B_TOT if strong else B_TOP, fill=FILL_TOT)
            r += 1
            return r - 1

        def section(text):
            nonlocal r
            put(ws, r, 2, text, font=F_BOLD, fill=FILL_SEC)
            for c in (1, 3, 4, 5):
                ws.cell(row=r, column=c).fill = FILL_SEC
            r += 1

        section("Continuing operations - Operating category")
        op_rows = []
        for k in self.lines_in(OP):
            op_rows.append(line(k))
            if k == "Cost of sales" and "Revenue" in self.spl_rows:
                rv, cs = self.spl_rows["Revenue"], self.spl_rows["Cost of sales"]
                subtotal("Gross profit", "IFRS 18.24, IE7 (additional subtotal)", f"=D{rv}+D{cs}", f"=E{rv}+E{cs}")
        self.r_op = subtotal("Operating profit", "IFRS 18.69(a)",
                             "=" + "+".join(f"D{x}" for x in op_rows), "=" + "+".join(f"E{x}" for x in op_rows))
        section("Investing category")
        inv_rows = [line(k) for k in self.lines_in(INV)]
        f_inv = lambda c: "".join(f"+{c}{x}" for x in inv_rows)
        self.r_pbfit = subtotal("Profit before financing and income taxes", "IFRS 18.69(b)",
                                f"=D{self.r_op}{f_inv('D')}", f"=E{self.r_op}{f_inv('E')}")
        section("Financing category")
        fin_rows = [line(k) for k in self.lines_in(FIN)]
        f_fin = lambda c: "".join(f"+{c}{x}" for x in fin_rows)
        self.r_pbt = subtotal("Profit before income taxes", "IFRS 18.24, IE7",
                              f"=D{self.r_pbfit}{f_fin('D')}", f"=E{self.r_pbfit}{f_fin('E')}")
        section("Income taxes category")
        tax_rows = [line(k) for k in self.lines_in(TAX)]
        self.r_cont = subtotal("Profit from continuing operations", "IFRS 18.24",
                               f"=D{self.r_pbt}" + "".join(f"+D{x}" for x in tax_rows),
                               f"=E{self.r_pbt}" + "".join(f"+E{x}" for x in tax_rows))
        section("Discontinued operations category")
        disc_rows = [line(k) for k in self.lines_in(DISC)]
        self.r_profit = subtotal("Profit for the period", "IFRS 18.69(c)",
                                 f"=D{self.r_cont}" + "".join(f"+D{x}" for x in disc_rows),
                                 f"=E{self.r_cont}" + "".join(f"+E{x}" for x in disc_rows), strong=True)
        r += 1
        section("Other comprehensive income")
        put(ws, r, 1, "IFRS 18.88(b)", font=F_SUB)
        put(ws, r, 2, "Items that will not be reclassified to profit or loss", font=F_BOLD)
        r += 1
        nr = [line(k) for k in self.lines_in(OCINR)]
        r_nr = subtotal("", "", "=" + ("+".join(f"D{x}" for x in nr) or "0"), "=" + ("+".join(f"E{x}" for x in nr) or "0"))
        put(ws, r, 1, "IFRS 18.88(a)", font=F_SUB)
        put(ws, r, 2, "Items that are or may be reclassified subsequently to profit or loss", font=F_BOLD)
        r += 1
        rc = [line(k) for k in self.lines_in(OCIR)]
        r_rc = subtotal("", "", "=" + ("+".join(f"D{x}" for x in rc) or "0"), "=" + ("+".join(f"E{x}" for x in rc) or "0"))
        self.r_oci = subtotal("Other comprehensive income for the period, net of tax", "IFRS 18.86(b)",
                              f"=D{r_nr}+D{r_rc}", f"=E{r_nr}+E{r_rc}")
        self.r_tci = subtotal("Total comprehensive income for the period", "IFRS 18.86(c)",
                              f"=D{self.r_profit}+D{self.r_oci}", f"=E{self.r_profit}+E{self.r_oci}", strong=True)
        r += 1
        section("Attribution")
        put(ws, r, 2, "Profit attributable to:", font=F_BOLD); r += 1
        put(ws, r, 1, "IFRS 18.76(a)", font=F_SUB); put(ws, r, 2, "Owners of the Company")
        put(ws, r, 4, f"=D{self.r_profit}-D{r+1}", fmt=NUM); put(ws, r, 5, f"=E{self.r_profit}-E{r+1}", fmt=NUM)
        self.r_owners = r; r += 1
        put(ws, r, 1, "IFRS 18.76(b)", font=F_SUB); put(ws, r, 2, "Non-controlling interests")
        put(ws, r, 4, f"={self.S('nci_profit_cy')}", fmt=NUM); put(ws, r, 5, f"={self.S('nci_profit_py')}", fmt=NUM)
        self.r_nci = r; r += 1
        put(ws, r, 2, "Total comprehensive income attributable to:", font=F_BOLD); r += 1
        put(ws, r, 1, "IFRS 18.87(a)", font=F_SUB); put(ws, r, 2, "Owners of the Company")
        put(ws, r, 4, f"=D{self.r_tci}-D{r+1}", fmt=NUM); put(ws, r, 5, f"=E{self.r_tci}-E{r+1}", fmt=NUM)
        self.r_tci_owners = r; r += 1
        put(ws, r, 1, "IFRS 18.87(b)", font=F_SUB); put(ws, r, 2, "Non-controlling interests")
        put(ws, r, 4, f"={self.S('nci_profit_cy')}+{self.S('nci_oci_cy')}", fmt=NUM)
        put(ws, r, 5, f"={self.S('nci_profit_py')}+{self.S('nci_oci_py')}", fmt=NUM)
        self.r_tci_nci = r; r += 2
        section("Earnings per share (IAS 33.66)")
        put(ws, r, 2, "Basic EPS")
        put(ws, r, 4, f"=(D{self.r_owners}-{self.S('pref_div_cy')})/{self.S('wans_basic_cy')}", fmt=NUM2)
        put(ws, r, 5, f"=(E{self.r_owners}-{self.S('pref_div_py')})/{self.S('wans_basic_py')}", fmt=NUM2); r += 1
        put(ws, r, 2, "Diluted EPS")
        put(ws, r, 4, f"=(D{self.r_owners}-{self.S('pref_div_cy')}+{self.S('diluted_adj_cy')})/{self.S('wans_diluted_cy')}", fmt=NUM2)
        put(ws, r, 5, f"=(E{self.r_owners}-{self.S('pref_div_py')}+{self.S('diluted_adj_py')})/{self.S('wans_diluted_py')}", fmt=NUM2); r += 2
        notes = [
            "Presentation notes:",
            "1. Line items are presented when needed for a useful structured summary (IFRS 18.24, 41-43); immaterial required lines may be omitted (IFRS 18.19, B1-B5).",
            "2. Mixed presentation of operating expenses (by function and by nature) - see 10_ExpNature for the IFRS 18.83 disclosure.",
            "3. OCI components shown before tax with aggregate tax per group (IFRS 18.93-95); tax per component disclosed in the tax note.",
            "4. Discontinued operation shown as a single post-tax amount (IFRS 18.75(a)(v)); analysis in the IFRS 5 note.",
        ]
        for n in notes:
            put(ws, r, 2, n, font=F_SUB); r += 1
        widths(ws, [30, 70, 22, 14, 16, 2, 40])
        ws.column_dimensions["G"].hidden = True
        ws.freeze_panes = "C6"
        # checks
        for cat, rr in ((OP, self.r_op),):
            self.checks.append((f"Operating profit per lines = SUMIFS on '{cat}' category (CY)", f"'08_SPLOCI'!D{rr}",
                                f'-SUMIFS({cy},{self.tbr("H")},"{cat}")', "08_SPLOCI"))
        self.checks.append(("Profit for the period = minus sum of PL accounts (CY)", f"'08_SPLOCI'!D{self.r_profit}",
                            f'-SUMIFS({cy},{self.tbr("C")},"PL")', "08_SPLOCI"))
        self.checks.append(("Profit for the period = minus sum of PL accounts (PY)", f"'08_SPLOCI'!E{self.r_profit}",
                            f'-SUMIFS({py},{self.tbr("C")},"PL")', "08_SPLOCI"))
        self.checks.append(("OCI = minus sum of OCI accounts (CY)", f"'08_SPLOCI'!D{self.r_oci}",
                            f'-SUMIFS({cy},{self.tbr("C")},"OCI")', "08_SPLOCI"))

    # ---------- 14 SFP ----------
    def sfp(self):
        ws = self.ws["14_SFP"]
        title(ws, "Statement of financial position (IFRS 18.96-105)",
              "Current/non-current presentation (IFRS 18.96). Goodwill presented separately (IFRS 18.103(d)). Equity rolls in profit and OCI.", "14")
        header(ws, 5, ["IFRS 18 reference", "Line item", "Section", "=\"CY \"&'01_Setup'!C6", "=\"PY \"&'01_Setup'!C7", "", "Line key"])
        cy, py = self.tbr("D"), self.tbr("E")
        r = 6
        tot = {}
        sp = "'08_SPLOCI'"
        roll = {  # equity lines that absorb current-period profit/OCI
            "Retained earnings": (f"+{sp}!D{self.r_owners}", f"+{sp}!E{self.r_owners}"),
            "Reserves": (f"+{sp}!D{self.r_tci_owners}-{sp}!D{self.r_owners}", f"+{sp}!E{self.r_tci_owners}-{sp}!E{self.r_owners}"),
            "Non-controlling interests": (f"+{sp}!D{self.r_tci_nci}", f"+{sp}!E{self.r_tci_nci}"),
        }
        for sec, sign in ((NCA, 1), (CA, 1), (EQ, -1), (NCL, -1), (CL, -1)):
            put(ws, r, 2, sec, font=F_BOLD, fill=FILL_SEC)
            r += 1
            first = r
            for k in self.lines_in(sec):
                lab, _, ref = self.cat[k]
                s = "-" if sign < 0 else ""
                extra = roll.get(k, ("", ""))
                put(ws, r, 1, ref, font=F_SUB)
                put(ws, r, 2, lab)
                put(ws, r, 3, sec, font=F_SUB)
                put(ws, r, 4, f"={s}SUMIFS({cy},{self.tbr('G')},$G{r}){extra[0]}", fmt=NUM)
                put(ws, r, 5, f"={s}SUMIFS({py},{self.tbr('G')},$G{r}){extra[1]}", fmt=NUM)
                put(ws, r, 7, k, font=F_SUB)
                self.pos[k] = ("14_SFP", r)
                r += 1
            last = r - 1
            put(ws, r, 2, f"Total {sec.lower()}", font=F_BOLD)
            put(ws, r, 4, f"=SUM(D{first}:D{last})", font=F_BOLD, fmt=NUM, border=B_TOP, fill=FILL_TOT)
            put(ws, r, 5, f"=SUM(E{first}:E{last})", font=F_BOLD, fmt=NUM, border=B_TOP, fill=FILL_TOT)
            tot[sec] = r
            r += 1
            if sec == CA:
                put(ws, r, 2, "Total assets", font=F_BOLD)
                put(ws, r, 4, f"=D{tot[NCA]}+D{tot[CA]}", font=F_BOLD, fmt=NUM, border=B_TOT)
                put(ws, r, 5, f"=E{tot[NCA]}+E{tot[CA]}", font=F_BOLD, fmt=NUM, border=B_TOT)
                self.r_ta = r
                r += 2
        put(ws, r, 2, "Total liabilities", font=F_BOLD)
        put(ws, r, 4, f"=D{tot[NCL]}+D{tot[CL]}", font=F_BOLD, fmt=NUM, border=B_TOP)
        put(ws, r, 5, f"=E{tot[NCL]}+E{tot[CL]}", font=F_BOLD, fmt=NUM, border=B_TOP)
        r += 1
        put(ws, r, 2, "Total equity and liabilities", font=F_BOLD)
        put(ws, r, 4, f"=D{tot[EQ]}+D{r-1}", font=F_BOLD, fmt=NUM, border=B_TOT)
        put(ws, r, 5, f"=E{tot[EQ]}+E{r-1}", font=F_BOLD, fmt=NUM, border=B_TOT)
        self.r_tel = r
        r += 2
        for n in ["Notes:",
                  "- A third statement of financial position at the start of the comparative period is required only if a retrospective change has a material "
                  "effect on it (IFRS 18.37-38). IFRS 18 reclassifications often affect only presentation - document the conclusion in 18_Issues.",
                  "- Review aggregation: separate line items for items with dissimilar characteristics (IFRS 18.41-43, B109-B111) - see 11_Aggregation.",
                  "- Retained earnings / reserves / NCI in 05_TB are balances before current-period profit and OCI; this sheet rolls them forward."]:
            put(ws, r, 2, n, font=F_SUB)
            r += 1
        widths(ws, [26, 58, 22, 14, 14, 2, 40])
        ws.column_dimensions["G"].hidden = True
        ws.freeze_panes = "C6"
        self.sfp_tot = tot
        self.checks.append(("SFP balances: total assets = total equity and liabilities (CY)", f"'14_SFP'!D{self.r_ta}", f"'14_SFP'!D{self.r_tel}", "14_SFP"))
        self.checks.append(("SFP balances: total assets = total equity and liabilities (PY)", f"'14_SFP'!E{self.r_ta}", f"'14_SFP'!E{self.r_tel}", "14_SFP"))

    # ---------- 07 Classification workings ----------
    def classification(self):
        ws = self.ws["07_Classification"]
        title(ws, "Classification workings - income and expenses by category",
              "Summarises 05_TB by IFRS 18 category and proves completeness of the five-category split (IFRS 18.47).", "07")
        header(ws, 5, ["Category", "=\"CY \"&'01_Setup'!C6", "=\"PY \"&'01_Setup'!C7", "No. of accounts", "Comment"])
        cy, py = self.tbr("D"), self.tbr("E")
        r = 6
        for c in (OP, INV, FIN, TAX, DISC):
            put(ws, r, 1, c)
            put(ws, r, 2, f'=-SUMIFS({cy},{self.tbr("H")},A{r})', fmt=NUM)
            put(ws, r, 3, f'=-SUMIFS({py},{self.tbr("H")},A{r})', fmt=NUM)
            put(ws, r, 4, f'=COUNTIFS({self.tbr("H")},A{r})', fmt="0")
            r += 1
        put(ws, r, 1, "Profit for the period", font=F_BOLD)
        put(ws, r, 2, f"=SUM(B6:B{r-1})", font=F_BOLD, fmt=NUM, border=B_TOT)
        put(ws, r, 3, f"=SUM(C6:C{r-1})", font=F_BOLD, fmt=NUM, border=B_TOT)
        self.checks.append(("Five categories sum to profit for the period (CY)", f"'07_Classification'!B{r}",
                            f"'08_SPLOCI'!D{self.r_profit}", "07_Classification"))
        r += 2
        put(ws, r, 1, "Classification by IFRS 18 line and by IAS 1 caption (comparative period) - see 09_Transition for the full bridge.", font=F_SUB)
        r += 2
        header(ws, r, ["IFRS 18 line", "Category", "CY", "PY", "Accounts mapped"])
        r += 1
        for k, lab, sec, ref in CATALOGUE:
            if k in self.used and sec in (OP, INV, FIN, TAX, DISC):
                put(ws, r, 1, k)
                put(ws, r, 2, sec, font=F_SUB)
                put(ws, r, 3, f'=-SUMIFS({cy},{self.tbr("G")},A{r})', fmt=NUM)
                put(ws, r, 4, f'=-SUMIFS({py},{self.tbr("G")},A{r})', fmt=NUM)
                put(ws, r, 5, f'=COUNTIFS({self.tbr("G")},A{r})', fmt="0")
                r += 1
        widths(ws, [60, 22, 14, 14, 16])

    # ---------- 09 Transition ----------
    def transition(self):
        ws = self.ws["09_Transition"]
        title(ws, "Transition reconciliation IAS 1 to IFRS 18 - comparative period (IFRS 18.C3)",
              "For each line item in the statement of profit or loss: amount previously presented under IAS 1 vs restated under IFRS 18. "
              "Net profit is unchanged (presentation only).", "09")
        py = self.tbr("E")
        pl18 = [k for k in self.lines_in(OP) + self.lines_in(INV) + self.lines_in(FIN) + self.lines_in(TAX) + self.lines_in(DISC)]
        ias1 = [lab for lab, g in IAS1_LINES]
        ias1_grp = dict(IAS1_LINES)
        # Part A: summary reconciliation
        r = 5
        put(ws, r, 1, "A. Line-item reconciliation (comparative period)", font=F_BOLD)
        r += 1
        header(ws, r, ["IAS 1 caption", "IFRS 18 line item", "IAS 1 amount", "Adjustment", "IFRS 18 amount", "Transition refs"])
        r += 1
        first = r
        refs_by_line = {}
        for row in self.tb:
            if row[9]:
                refs_by_line.setdefault(row[6], set()).add(row[9])
                refs_by_line.setdefault(row[5], set()).add(row[9])
        sp = "'08_SPLOCI'"

        def recon_row(i1, k18):
            nonlocal r
            put(ws, r, 1, i1 or "-")
            put(ws, r, 2, self.cat[k18][0] if k18 else "-")
            put(ws, r, 3, f'=-SUMIFS({py},{self.tbr("F")},A{r})' if i1 else 0, fmt=NUM)
            put(ws, r, 5, f"={sp}!E{self.spl_rows[k18]}" if k18 else 0, fmt=NUM)
            put(ws, r, 4, f"=E{r}-C{r}", fmt=NUM)
            refs = sorted(refs_by_line.get(k18, set()) | refs_by_line.get(i1, set()))
            put(ws, r, 6, ", ".join(refs), font=F_SUB)
            r += 1

        def sub(label, c_formula, e_formula, strong=False):
            nonlocal r
            put(ws, r, 1, label, font=F_BOLD)
            put(ws, r, 2, label, font=F_BOLD)
            put(ws, r, 3, c_formula, font=F_BOLD, fmt=NUM, border=B_TOT if strong else B_TOP, fill=FILL_TOT)
            put(ws, r, 5, e_formula, font=F_BOLD, fmt=NUM, border=B_TOT if strong else B_TOP, fill=FILL_TOT)
            put(ws, r, 4, f"=E{r}-C{r}", font=F_BOLD, fmt=NUM, border=B_TOT if strong else B_TOP, fill=FILL_TOT)
            r += 1
            return r - 1

        matched = set()

        def eq(k):
            c = CAPTION_EQUIV.get(k, k)
            return c if c in ias1 else None

        for k in self.lines_in(OP):
            i1 = eq(k)
            if i1:
                matched.add(i1)
            recon_row(i1, k)
        for i1 in ias1:  # IAS 1 operating captions with no IFRS 18 equivalent
            if ias1_grp[i1] == "Above operating profit" and i1 not in matched:
                recon_row(i1, None)
                matched.add(i1)
        ias1_op = "+".join(f'-SUMIFS({py},{self.tbr("F")},"{x}")' for x in ias1 if ias1_grp[x] == "Above operating profit")
        r_op = sub("Operating profit", f"={ias1_op}", f"={sp}!E{self.r_op}")
        for k in self.lines_in(INV):
            i1 = eq(k)
            if i1:
                matched.add(i1)
            recon_row(i1, k)
        sub("Profit before financing and income taxes", "=0", f"={sp}!E{self.r_pbfit}")
        for i1 in ias1:
            if ias1_grp[i1] == "Net finance costs" and i1 not in matched and not any(eq(k) == i1 for k in self.lines_in(FIN)):
                recon_row(i1, None)
                matched.add(i1)
        for k in self.lines_in(FIN):
            i1 = eq(k)
            if i1:
                matched.add(i1)
            recon_row(i1, k)
        for i1 in ias1:
            if ias1_grp[i1] == "Below net finance costs" and i1 not in matched:
                recon_row(i1, None)
                matched.add(i1)
        r_pbt = sub("Profit before income taxes", f"=SUM(C{first}:C{r-1})-C{r_op}", f"={sp}!E{self.r_pbt}")
        for k in self.lines_in(TAX) + self.lines_in(DISC):
            recon_row(eq(k), k)
        r_profit = sub("Profit for the period", f"=C{r_pbt}+SUM(C{r_pbt+1}:C{r-1})", f"={sp}!E{self.r_profit}", strong=True)
        put(ws, r, 1, "Gross profit is not reproduced here; the PBFIT row shows nil under IAS 1 because the subtotal did not exist.", font=F_SUB)
        self.checks.append(("Transition: profit unchanged IAS 1 vs IFRS 18 (PY)", f"'09_Transition'!C{r_profit}", f"'09_Transition'!E{r_profit}", "09_Transition"))
        r += 2
        # Part B: explanations
        put(ws, r, 1, "B. Explanation of adjustments (disclosed in the change-in-accounting-policy note)", font=F_BOLD)
        r += 1
        header(ws, r, ["Ref", "Topic", "Explanation", "", "", "Reference", "Accounts", "Gross amount moved (PY)"])
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
        r += 1
        for ref, topic, text, cite in self.expl:
            put(ws, r, 1, ref, font=F_BOLD, align=WRAP)
            put(ws, r, 2, topic, align=WRAP)
            inp(ws, r, 3, text).alignment = WRAP
            ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
            put(ws, r, 6, cite, font=F_SUB, align=WRAP)
            put(ws, r, 7, f'=COUNTIFS({self.tbr("K")},A{r})', fmt="0", align=WRAP)
            put(ws, r, 8, f'=SUMPRODUCT(ABS({py})*({self.tbr("K")}=A{r}))', fmt=NUM, align=WRAP)
            ws.row_dimensions[r].height = 15 * (len(text) // 45 + 1)
            r += 1
        r += 1
        # Part C: bridge matrix
        put(ws, r, 1, "C. Bridge matrix (PY): rows = IAS 1 captions, columns = IFRS 18 line items. Off-pattern cells are reclassifications.", font=F_BOLD)
        r += 1
        hdr_row = r
        put(ws, r, 1, "IAS 1 caption \\ IFRS 18 line", font=F_HDR, fill=FILL_HDR)
        for j, k in enumerate(pl18):
            c = put(ws, r, 2 + j, k, font=F_HDR, fill=FILL_HDR, align=Alignment(wrap_text=True, vertical="top", text_rotation=90))
        put(ws, r, 2 + len(pl18), "Total (IAS 1)", font=F_HDR, fill=FILL_HDR)
        ws.row_dimensions[r].height = 200
        r += 1
        top = r
        for i1 in ias1:
            put(ws, r, 1, i1)
            for j, k in enumerate(pl18):
                col = get_column_letter(2 + j)
                put(ws, r, 2 + j, f'=-SUMIFS({py},{self.tbr("F")},$A{r},{self.tbr("G")},{col}${hdr_row})', fmt=NUM)
            lc = get_column_letter(1 + len(pl18))
            put(ws, r, 2 + len(pl18), f"=SUM(B{r}:{lc}{r})", font=F_BOLD, fmt=NUM)
            r += 1
        put(ws, r, 1, "Total (IFRS 18)", font=F_BOLD)
        for j in range(len(pl18) + 1):
            col = get_column_letter(2 + j)
            put(ws, r, 2 + j, f"=SUM({col}{top}:{col}{r-1})", font=F_BOLD, fmt=NUM, border=B_TOT)
        tc = get_column_letter(2 + len(pl18))
        self.checks.append(("Bridge matrix grand total = profit for the period (PY)", f"'09_Transition'!{tc}{r}", f"'08_SPLOCI'!E{self.r_profit}", "09_Transition"))
        self.checks.append(("No P&L account without IAS 1 caption (PY bridge complete)",
                            f'COUNTIFS({self.tbr("C")},"PL",{self.tbr("F")},"")', "0", "09_Transition"))
        ws.conditional_formatting.add(f"B{top}:{get_column_letter(1+len(pl18))}{r-1}",
                                      CellIsRule(operator="notEqual", formula=["0"], fill=FILL_KEY))
        widths(ws, [44, 52, 14, 14, 14, 18, 10, 14] + [9] * max(0, len(pl18) - 7))
        ws.column_dimensions["B"].width = 52
        ws.freeze_panes = "A5"

    # ---------- 10 Expenses by nature ----------
    def exp_nature(self):
        ws = self.ws["10_ExpNature"]
        title(ws, "Operating expenses by nature - specified amounts (IFRS 18.83, B84)",
              "Required when any operating expenses are presented by function. Amounts recognised come from 05_TB nature tags; "
              "capitalised amounts are inputs (IFRS 18.83(a) total amount incurred).", "10")
        cy, py = self.tbr("D"), self.tbr("E")
        func_lines = [k for k in self.lines_in(OP) if any(r[6] == k and r[7] for r in self.tb)]
        cols = NATURES5 + ["Other natures (balance)", "Total line item"]
        r = 5
        self.nat_tot_cells = {}
        for per, rng, pcol in (("CY", cy, "D"), ("PY", py, "E")):
            blocks = {}
            for blk, label in (("rec", "Recognised as an expense in profit or loss"),
                               ("cap", "Capitalised into the carrying amount of assets (input)"),
                               ("inc", "Total amount incurred (recognised + capitalised) - IFRS 18.83(a)/(b)")):
                put(ws, r, 1, f"{per} - {label}", font=F_BOLD, fill=FILL_SEC)
                r += 1
                header(ws, r, ["Line item (operating category)"] + cols)
                r += 1
                first = r
                for k in func_lines:
                    put(ws, r, 1, k)
                    for j, n in enumerate(NATURES5):
                        if blk == "rec":
                            put(ws, r, 2 + j, f'=SUMIFS({rng},{self.tbr("G")},$A{r},{self.tbr("I")},"{n}")', fmt=NUM)
                        elif blk == "cap":
                            v = self.cap.get(k, {}).get(n, (0, 0))[0 if per == "CY" else 1]
                            inp(ws, r, 2 + j, v, NUM)
                        else:
                            col = get_column_letter(2 + j)
                            put(ws, r, 2 + j, f"={col}{blocks['rec'] + (r - first)}+{col}{blocks['cap'] + (r - first)}", fmt=NUM)
                    if blk == "rec":
                        put(ws, r, 8, f"=-'08_SPLOCI'!{pcol}{self.spl_rows[k]}", font=F_BOLD, fmt=NUM)
                        put(ws, r, 7, f"=H{r}-SUM(B{r}:F{r})", fmt=NUM)
                    r += 1
                put(ws, r, 1, "Total", font=F_BOLD)
                for j in range(len(cols)):
                    col = get_column_letter(2 + j)
                    if blk != "rec" and j >= 5:
                        continue
                    put(ws, r, 2 + j, f"=SUM({col}{first}:{col}{r-1})", font=F_BOLD, fmt=NUM, border=B_TOT)
                blocks[blk] = first
                if blk == "inc":
                    self.nat_tot_cells[per] = r
                r += 2
        put(ws, r, 1, "Required narrative (IFRS 18.82(b), 83(b)(i)): describe what each function line includes, and state that capitalised amounts are "
                      "included and identify the assets (e.g. inventory, development costs). Alternative: disclose only amounts recognised as expense.",
            font=F_SUB)
        r += 1
        put(ws, r, 1, "If investment property were measured at cost, depreciation and impairment in the investing category would also be identified (IFRS 18.83(b)(ii)).", font=F_SUB)
        widths(ws, [60, 14, 14, 16, 18, 18, 18, 16])
        ws.freeze_panes = "B5"

    # ---------- 13 MPM reconciliation ----------
    def mpm_recon(self):
        ws = self.ws["13_MPM_Recon"]
        title(ws, "MPM reconciliation, tax and NCI effects (IFRS 18.121-123, B136-B141) and MPM per share (IAS 33.73B-73C)",
              "Reconciling items are pulled from 05_TB column J (MPM reconciling item). Tax rate and NCI % per item are inputs.", "13")
        items = []
        for row in self.tb:
            if row[8] and row[8] not in items:
                items.append(row[8])
        line_of = {}
        for row in self.tb:
            if row[8] and row[8] not in line_of:
                line_of[row[8]] = self.cat.get(row[6], (row[6],))[0]
        r = 5
        put(ws, r, 1, "MPM name", font=F_BOLD)
        inp(ws, r, 2, "Adjusted operating profit")
        r += 1
        put(ws, r, 1, "Most directly comparable subtotal (IFRS 18.123(c))", font=F_BOLD)
        put(ws, r, 2, "Operating profit (IFRS 18.118)")
        r += 2
        self.mpm_tot = {}
        for per, rng, pcol, prate in (("CY", self.tbr("D"), "D", 0), ("PY", self.tbr("E"), "E", 1)):
            put(ws, r, 1, f"{per} reconciliation", font=F_BOLD, fill=FILL_SEC)
            r += 1
            header(ws, r, ["Reconciling item", "Line item in which presented (IFRS 18.B137(b))", "Amount", "Tax rate", "Income tax effect",
                           "NCI share", "Effect on NCI", "Basis of tax effect"])
            r += 1
            put(ws, r, 1, "Operating profit", font=F_BOLD)
            put(ws, r, 3, f"='08_SPLOCI'!{pcol}{self.r_op}", font=F_BOLD, fmt=NUM)
            start = r
            r += 1
            first = r
            for it in items:
                put(ws, r, 1, it)
                put(ws, r, 2, line_of[it], font=F_SUB)
                put(ws, r, 3, f'=SUMIFS({rng},{self.tbr("J")},A{r})', fmt=NUM)
                rate = 0 if "goodwill" in it.lower() else f"={self.S('tax_rate')}"
                inp(ws, r, 4, rate, PCT)
                put(ws, r, 5, f"=-C{r}*D{r}", fmt=NUM)
                inp(ws, r, 6, 0.1, PCT)
                put(ws, r, 7, f"=(C{r}+E{r})*F{r}", fmt=NUM)
                inp(ws, r, 8, "Statutory rate of the relevant jurisdiction (B141(a))" if rate else "Non-deductible - nil tax effect")
                r += 1
            put(ws, r, 1, "=B5", font=F_BOLD)
            put(ws, r, 3, f"=C{start}+SUM(C{first}:C{r-1})", font=F_BOLD, fmt=NUM, border=B_TOT)
            put(ws, r, 5, f"=SUM(E{first}:E{r-1})", font=F_BOLD, fmt=NUM, border=B_TOT)
            put(ws, r, 7, f"=SUM(G{first}:G{r-1})", font=F_BOLD, fmt=NUM, border=B_TOT)
            self.mpm_tot[per] = r
            r += 2
        put(ws, r, 1, "MPM per share (IAS 33.73B-73C) - only MPMs or IFRS 18 totals/subtotals may be the numerator", font=F_BOLD, fill=FILL_SEC)
        r += 1
        header(ws, r, ["", "", "CY", "PY"])
        r += 1
        put(ws, r, 1, "MPM attributable to ordinary shareholders")
        put(ws, r, 3, f"=C{self.mpm_tot['CY']}*{self.S('owners_share_mpm')}", fmt=NUM)
        put(ws, r, 4, f"=C{self.mpm_tot['PY']}*{self.S('owners_share_mpm')}", fmt=NUM)
        r += 1
        put(ws, r, 1, "Basic MPM per share")
        put(ws, r, 3, f"=C{r-1}/{self.S('wans_basic_cy')}", fmt=NUM2)
        put(ws, r, 4, f"=D{r-1}/{self.S('wans_basic_py')}", fmt=NUM2)
        r += 1
        put(ws, r, 1, "Diluted MPM per share")
        put(ws, r, 3, f"=C{r-2}/{self.S('wans_diluted_cy')}", fmt=NUM2)
        put(ws, r, 4, f"=D{r-2}/{self.S('wans_diluted_py')}", fmt=NUM2)
        r += 2
        put(ws, r, 1, "Single-note disclosure checklist for each MPM", font=F_BOLD, fill=FILL_SEC)
        r += 1
        header(ws, r, ["Requirement", "Reference", "Done? (Y/N/NA)", "Note / evidence"])
        r += 1
        reqs = [
            ("All MPM information in a single note", "IFRS 18.121"),
            ("Statement that MPMs provide management's view and are not necessarily comparable with similarly labelled measures of other entities", "IFRS 18.122"),
            ("Description of the aspect of performance communicated and why the MPM provides useful information", "IFRS 18.123(a)"),
            ("How the MPM is calculated", "IFRS 18.123(b)"),
            ("Reconciliation to the most directly comparable subtotal in IFRS 18.118 or total/subtotal required by IFRS", "IFRS 18.123(c), B136-B140"),
            ("Income tax effect and effect on NCI for each reconciling item", "IFRS 18.123(d)"),
            ("How the income tax effect was determined", "IFRS 18.123(e), B141"),
            ("Changes in MPMs (added, ceased, calculation changed): explanation, reasons, restated comparatives", "IFRS 18.124"),
            ("MPM labels and descriptions are clear, not misleading; faithful representation", "IFRS 18.B116, B119"),
            ("Comparative information for each MPM", "IFRS 18.31"),
        ]
        for t, ref in reqs:
            put(ws, r, 1, t, align=WRAP)
            put(ws, r, 2, ref, font=F_SUB)
            inp(ws, r, 3, "")
            inp(ws, r, 4, "")
            r += 1
        add_list_validation(ws, f"C{r-len(reqs)}:C{r-1}", '"Y,N,NA"')
        widths(ws, [62, 52, 13, 10, 14, 10, 13, 48])
        for per in ("CY", "PY"):
            rr = self.mpm_tot[per]
            col = "D" if per == "CY" else "E"
            self.checks.append((f"MPM = operating profit + reconciling items ({per})", f"'13_MPM_Recon'!C{rr}",
                                f"'08_SPLOCI'!{col}{self.r_op}+SUMIFS({self.tbr('D' if per=='CY' else 'E')},{self.tbr('J')},\"<>\")", "13_MPM_Recon"))

    # ---------- 15 Cash flow ----------
    def cashflow(self):
        ws = self.ws["15_CashFlow"]
        title(ws, "Statement of cash flows - IAS 7 as amended by IFRS 18",
              "Indirect method starts from OPERATING PROFIT (IAS 7.18(b), 20). Interest and dividend classification per IAS 7.33A-34D.", "15")
        header(ws, 5, ["Line item", "IAS 7 reference", "=\"CY \"&'01_Setup'!C6", "=\"PY \"&'01_Setup'!C7"])
        cf = self.cf
        r = 6

        def sec(t):
            nonlocal r
            put(ws, r, 1, t, font=F_BOLD, fill=FILL_SEC)
            r += 1

        def rows(items, ref=""):
            nonlocal r
            first = r
            for lab, a, b in items:
                put(ws, r, 1, lab)
                put(ws, r, 2, ref, font=F_SUB)
                inp(ws, r, 3, a, NUM)
                inp(ws, r, 4, b, NUM)
                r += 1
            return first, r - 1

        def tot(label, fc, fd, strong=False):
            nonlocal r
            put(ws, r, 1, label, font=F_BOLD)
            put(ws, r, 3, fc, font=F_BOLD, fmt=NUM, border=B_TOT if strong else B_TOP, fill=FILL_TOT)
            put(ws, r, 4, fd, font=F_BOLD, fmt=NUM, border=B_TOT if strong else B_TOP, fill=FILL_TOT)
            r += 1
            return r - 1

        sec("Cash flows from operating activities")
        put(ws, r, 1, "Operating profit (linked to 08_SPLOCI)", font=F_BOLD)
        put(ws, r, 2, "IAS 7.18(b) (amended)", font=F_SUB)
        put(ws, r, 3, f"='08_SPLOCI'!D{self.r_op}", fmt=NUM)
        put(ws, r, 4, f"='08_SPLOCI'!E{self.r_op}", fmt=NUM)
        r_op = r
        r += 1
        put(ws, r, 1, "Adjustments for:", font=F_SUB)
        r += 1
        a1, a2 = rows(cf["op_adjustments"], "IAS 7.20(b)")
        t1 = tot("", f"=C{r_op}+SUM(C{a1}:C{a2})", f"=D{r_op}+SUM(D{a1}:D{a2})")
        put(ws, r, 1, "Changes in:", font=F_SUB)
        r += 1
        w1, w2 = rows(cf["wc_changes"], "IAS 7.20(a)")
        t2 = tot("Cash generated from operating activities before tax", f"=C{t1}+SUM(C{w1}:C{w2})", f"=D{t1}+SUM(D{w1}:D{w2})")
        x1, _ = rows([("Income taxes paid", *cf["tax_paid"])], "IAS 7.35")
        t3 = tot("Net cash from operating activities - continuing operations", f"=C{t2}+C{x1}", f"=D{t2}+D{x1}")
        d1, _ = rows([("Net cash from operating activities - discontinued operation", *cf["disc_ops_operating"])], "IFRS 5.33(c)")
        t_op = tot("Net cash from operating activities", f"=C{t3}+C{d1}", f"=D{t3}+D{d1}")
        sec("Cash flows from investing activities")
        i1, i2 = rows(cf["investing"], "IAS 7.16, 34A")
        t_inv = tot("Net cash used in investing activities", f"=SUM(C{i1}:C{i2})", f"=SUM(D{i1}:D{i2})")
        sec("Cash flows from financing activities")
        f1, f2 = rows(cf["financing"], "IAS 7.17, 34A")
        t_fin = tot("Net cash from (used in) financing activities", f"=SUM(C{f1}:C{f2})", f"=SUM(D{f1}:D{f2})")
        t_net = tot("Net increase (decrease) in cash and cash equivalents", f"=C{t_op}+C{t_inv}+C{t_fin}", f"=D{t_op}+D{t_inv}+D{t_fin}")
        o1, _ = rows([("Cash and cash equivalents at 1 January (net of overdraft)", *cf["opening_cash"])], "")
        fx1, _ = rows([("Effect of movements in exchange rates on cash held", *cf["fx_on_cash"])], "IAS 7.28")
        t_close = tot("Cash and cash equivalents at 31 December (net of overdraft)", f"=C{t_net}+C{o1}+C{fx1}", f"=D{t_net}+D{o1}+D{fx1}", strong=True)
        cash_cy = f"SUMIFS({self.tbr('D')},{self.tbr('G')},\"Cash and cash equivalents\")+SUMIFS({self.tbr('D')},{self.tbr('G')},\"Bank overdraft\")"
        cash_py = f"SUMIFS({self.tbr('E')},{self.tbr('G')},\"Cash and cash equivalents\")+SUMIFS({self.tbr('E')},{self.tbr('G')},\"Bank overdraft\")"
        self.checks.append(("Cash flow closing cash = SFP cash less overdraft (CY)", f"'15_CashFlow'!C{t_close}", cash_cy, "15_CashFlow"))
        self.checks.append(("Cash flow closing cash = SFP cash less overdraft (PY)", f"'15_CashFlow'!D{t_close}", cash_py, "15_CashFlow"))
        self.checks.append(("Cash flow opening cash (CY) = closing cash (PY)", f"'15_CashFlow'!C{o1}", f"'15_CashFlow'!D{t_close}", "15_CashFlow"))
        r += 1
        put(ws, r, 1, "Interest and dividend classification matrix (IAS 7.33A-34D as amended) - confirm against the Standard for SMBA entities", font=F_BOLD, fill=FILL_SEC)
        r += 1
        header(ws, r, ["Cash flow", "Entity without an SMBA", "Entity with an SMBA", "Entity's classification (input)"])
        r += 1
        for a, b, c, d in [
            ("Interest paid", "Financing", "Same category as the related expense in the statement of profit or loss (operating or financing)", "Financing"),
            ("Interest received", "Investing", "Same category as the related income (operating or investing)", "Investing"),
            ("Dividends received", "Investing", "Same category as the related income (operating or investing)", "Investing"),
            ("Dividends paid", "Financing", "Financing", "Financing"),
        ]:
            put(ws, r, 1, a)
            put(ws, r, 2, b)
            put(ws, r, 3, c, align=WRAP)
            inp(ws, r, 4, d)
            r += 1
        put(ws, r, 1, "Policy choice for interest and dividends was removed for entities without an SMBA; disclose the change in the transition note.", font=F_SUB)
        widths(ws, [62, 34, 16, 34])
        ws.freeze_panes = "B6"

    # ---------- Narrative / checklist sheets ----------
    def index(self):
        ws = self.ws["00_Index"]
        ws.sheet_view.showGridLines = False
        put(ws, 1, 1, "IFRS 18 Presentation and Disclosure in Financial Statements - Implementation working papers", font=F_TITLE)
        put(ws, 2, 1, "Generated by the ifrs18-implementer skill. Live formulas - change 05_TB / inputs and everything recalculates.", font=F_SUB)
        put(ws, 3, 1, "Entity:", font=F_BOLD)
        put(ws, 3, 2, "='01_Setup'!C5")
        put(ws, 4, 1, "Overall check status:", font=F_BOLD)
        put(ws, 4, 2, "='19_Checks'!B4", font=F_BOLD)
        header(ws, 6, ["WP ref", "Sheet (click)", "Purpose", "Key IFRS references", "Status", "Prepared (initials/date)", "Reviewed (initials/date)"])
        wps = [
            ("01", "01_Setup", "Entity details, adoption date, SMBA and expense-method conclusions, EPS/NCI/materiality inputs", "IFRS 18.C1-C2; IAS 8"),
            ("02", "02_Plan", "Implementation plan: Assess - Redesign - Implement - Report", "IFRS 18.C1"),
            ("03", "03_Gap", "Impact / gap analysis IAS 1 practice vs IFRS 18", "IFRS 18 (whole)"),
            ("04", "04_SMBA", "Specified main business activity assessment", "IFRS 18.49-51, B30-B41"),
            ("05", "05_TB", "Trial balance and account mapping (IAS 1 caption, IFRS 18 line, nature, MPM, transition ref)", "IFRS 18.47"),
            ("06", "06_Rules", "Classification rule book by item type and SMBA status", "IFRS 18.47-68, B42-B77"),
            ("07", "07_Classification", "Category totals and completeness of the five-category split", "IFRS 18.47"),
            ("08", "08_SPLOCI", "Statement of profit or loss and OCI with required subtotals and EPS", "IFRS 18.69, 75-76, 86-95; IAS 33"),
            ("09", "09_Transition", "IAS 1 to IFRS 18 reconciliation of comparatives, explanations, bridge matrix", "IFRS 18.C3, C6"),
            ("10", "10_ExpNature", "Five specified expenses by nature for each function line", "IFRS 18.82-85, B84"),
            ("11", "11_Aggregation", "Aggregation / disaggregation and 'other' labelling review", "IFRS 18.41-45, B15-B26"),
            ("12", "12_MPM_ID", "Identification of management-defined performance measures", "IFRS 18.117-120, B113-B122"),
            ("13", "13_MPM_Recon", "MPM reconciliation with tax and NCI effects; MPM per share", "IFRS 18.121-125, B136-B141; IAS 33.73B-73C"),
            ("14", "14_SFP", "Statement of financial position line items and balance check", "IFRS 18.96-105"),
            ("15", "15_CashFlow", "Statement of cash flows from operating profit; interest/dividends", "IAS 7.18, 20, 33A-34D"),
            ("16", "16_Disclosure", "IFRS 18 presentation and disclosure checklist", "IFRS 18 & consequential amendments"),
            ("17", "17_Reclass", "Reclassification entries IAS 1 caption to IFRS 18 line (no measurement change)", "IFRS 18.C2"),
            ("18", "18_Issues", "Judgements, open issues and review notes", "IFRS 18.B30; IAS 8"),
            ("19", "19_Checks", "Control checks - must all be OK before sign-off", "-"),
        ]
        for i, (ref, sh, purpose, cite) in enumerate(wps):
            r = 7 + i
            put(ws, r, 1, ref)
            c = put(ws, r, 2, f'=HYPERLINK("#\'{sh}\'!A1","{sh}")', font=F_LINK)
            put(ws, r, 3, purpose, align=WRAP)
            put(ws, r, 4, cite, font=F_SUB, align=WRAP)
            inp(ws, r, 5, "Not started")
            inp(ws, r, 6, "")
            inp(ws, r, 7, "")
        add_list_validation(ws, f"E7:E{6+len(wps)}", f"={self.status_rng}")
        r = 8 + len(wps)
        for n in ["Colour code: blue text on yellow = input; black = formula; green cells in the bridge matrix = reclassifications.",
                  "Sign convention in 05_TB: debit positive, credit negative. Statements present income as positive and expenses as negative.",
                  "Recognition and measurement do not change under IFRS 18 - net profit and net assets must be unchanged by adoption.",
                  "Sources: IFRS 18 (IASB, April 2024) and consequential amendments; KPMG 'Guide to annual financial statements - Initial application of IFRS 18' "
                  "(Oct 2026); KPMG East Africa 'IFRS 18: A New Era of Financial Reporting' (2026). Demo figures adapted from the KPMG guide."]:
            put(ws, r, 1, n, font=F_SUB)
            r += 1
        widths(ws, [8, 22, 70, 40, 14, 20, 20])

    def plan(self):
        ws = self.ws["02_Plan"]
        title(ws, "Implementation plan and timeline", "Assumes 31 December year end. 2026 comparatives must be captured on the IFRS 18 basis.", "02")
        header(ws, 5, ["#", "Phase", "Task", "Output / WP", "Owner", "Start", "Due", "Status", "% weight", "Done weight"])
        tasks = [
            ("1 Assess", "Read IFRS 18, consequential amendments (IAS 7, IAS 33, IAS 34, IAS 8) and ACAR/regulator adoption status", "03_Gap"),
            ("1 Assess", "Map every P&L account to an IFRS 18 category and line", "05_TB"),
            ("1 Assess", "Test for a specified main business activity (per entity and for the group)", "04_SMBA"),
            ("1 Assess", "Inventory public communications for candidate MPMs (results releases, investor decks, annual report front half)", "12_MPM_ID"),
            ("1 Assess", "Impact assessment report to audit committee", "03_Gap"),
            ("2 Redesign", "Design new statement of profit or loss layout (subtotals, expense method, line items)", "08_SPLOCI"),
            ("2 Redesign", "Redesign chart of accounts / ERP mapping (nature tags, category, FX by source, derivative purpose)", "05_TB, 17_Reclass"),
            ("2 Redesign", "Revisit aggregation and labels of 'other' items; separate goodwill line on SFP", "11_Aggregation, 14_SFP"),
            ("2 Redesign", "Redesign statement of cash flows (start from operating profit; interest/dividend classification)", "15_CashFlow"),
            ("3 Implement", "MPM governance, policy and controls (definition, approval, consistency, tax/NCI computation)", "13_MPM_Recon"),
            ("3 Implement", "System and data changes; dry-run quarterly close on IFRS 18 basis", "05_TB"),
            ("3 Implement", "Train finance, investor relations and audit committee", "-"),
            ("3 Implement", "Interim reporting: IFRS 18 headings/subtotals and MPM disclosures in condensed interim FS (IAS 34 amendments)", "16_Disclosure"),
            ("4 Report", "Restate comparatives and prepare IAS 1 to IFRS 18 reconciliation", "09_Transition"),
            ("4 Report", "Prepare expenses-by-nature note and MPM note", "10_ExpNature, 13_MPM_Recon"),
            ("4 Report", "Complete disclosure checklist; auditor walk-through of MPMs (now within audit scope)", "16_Disclosure"),
        ]
        for i, (ph, t, wp) in enumerate(tasks):
            r = 6 + i
            put(ws, r, 1, i + 1)
            put(ws, r, 2, ph)
            put(ws, r, 3, t, align=WRAP)
            put(ws, r, 4, wp, font=F_SUB)
            inp(ws, r, 5, "")
            inp(ws, r, 6, "")
            inp(ws, r, 7, "")
            inp(ws, r, 8, "Not started")
            inp(ws, r, 9, 1, "0")
            put(ws, r, 10, f'=IF(H{r}="Complete",I{r},IF(H{r}="In progress",I{r}/2,0))', fmt="0.0")
        last = 5 + len(tasks)
        add_list_validation(ws, f"H6:H{last}", f"={self.status_rng}")
        put(ws, last + 2, 3, "Overall progress", font=F_BOLD)
        put(ws, last + 2, 10, f"=SUM(J6:J{last})/SUM(I6:I{last})", font=F_BOLD, fmt=PCT)
        r = last + 4
        put(ws, r, 2, "Key dates (31 December year end)", font=F_BOLD, fill=FILL_SEC)
        for i, (d, t) in enumerate([("1 Jan 2026", "Comparative period begins - 2026 transactions need classifying under IFRS 18 categories"),
                                    ("Q4 2026", "Decide and design: settle classification, SMBA and MPM judgements; map systems"),
                                    ("31 Dec 2026", "Comparatives close - 2026 statement of profit or loss restated on IFRS 18 basis"),
                                    ("1 Jan 2027", "IFRS 18 mandatory (IFRS 18.C1) - IAS 1 replaced")]):
            put(ws, r + 1 + i, 2, d, font=F_BOLD)
            put(ws, r + 1 + i, 3, t)
        widths(ws, [4, 14, 70, 26, 14, 12, 12, 14, 9, 11])
        ws.freeze_panes = "A6"

    def gap(self):
        ws = self.ws["03_Gap"]
        title(ws, "Impact and gap analysis - IAS 1 practice vs IFRS 18", "Complete 'Current practice' and 'Gap' for the entity; rating drives priority.", "03")
        header(ws, 5, ["Area", "IFRS 18 requirement", "Reference", "Current practice (IAS 1)", "Gap? (Y/N)", "Impact (H/M/L)", "Action", "Owner", "WP"])
        rows = [
            ("P&L structure", "Classify all income and expenses into operating, investing, financing, income taxes and discontinued operations", "IFRS 18.47"),
            ("Subtotals", "Present operating profit and profit before financing and income taxes (unless exception applies)", "IFRS 18.69, 73"),
            ("SMBA", "Assess whether the entity invests in assets or provides financing to customers as a main business activity", "IFRS 18.49-51, B30-B41"),
            ("Investing category", "Income/expenses from equity-accounted investments, cash and cash equivalents, and assets generating returns individually and largely independently", "IFRS 18.53-58"),
            ("Financing category", "Income/expenses from liabilities that involve only raising finance; interest and effects of interest-rate changes on other liabilities", "IFRS 18.59-66"),
            ("FX differences", "Classify in the same category as the items that gave rise to them", "IFRS 18.B65-B67"),
            ("Derivatives and hedges", "Designated hedges follow the hedged item; economic hedges follow the risk managed (with exceptions)", "IFRS 18.B70-B76"),
            ("Operating expense analysis", "Present by nature, function or mixed - most useful structured summary; avoid arbitrary mixes", "IFRS 18.78-81, B80-B85"),
            ("Expenses by nature note", "If any function presentation: disclose depreciation, amortisation, employee benefits, impairment, inventory write-downs per line", "IFRS 18.83-85"),
            ("MPMs", "Identify MPMs; single note with reconciliation, tax and NCI effects; now audited", "IFRS 18.117-125"),
            ("Aggregation", "Group on shared characteristics; disaggregate items with dissimilar characteristics when material", "IFRS 18.41-43, B15-B23"),
            ("Labelling 'other'", "Use informative labels; explain composition of 'other' aggregations when material", "IFRS 18.B24-B26"),
            ("Roles of PFS vs notes", "Primary statements = useful structured summary; notes = additional material information", "IFRS 18.15-17, B8-B14"),
            ("SFP line items", "Separate goodwill line; reassess other line items and current/non-current", "IFRS 18.96-105"),
            ("Statement of changes in equity", "Components and reconciliations largely carried over from IAS 1", "IFRS 18.107-111"),
            ("Statement of cash flows", "Indirect method from operating profit; interest/dividend classification fixed", "IAS 7.18, 33A-34D"),
            ("EPS", "Additional per-share amounts only using IFRS 18 totals/subtotals or MPMs", "IAS 33.73B-73C"),
            ("Interim reporting", "Condensed interim FS: IFRS 18 headings/subtotals; MPM disclosures", "IAS 34 (amended); IFRS 18.C5"),
            ("Transition", "Retrospective; reconcile comparative P&L line items IAS 1 vs IFRS 18", "IFRS 18.C2-C3"),
            ("Systems and CoA", "Chart of accounts tags for category, nature, MPM items, FX source, derivative purpose", "-"),
            ("Controls and governance", "MPM policy, approval, consistency with public communications; SOX/ICFR updates", "-"),
            ("Local regulation", "Confirm adoption by the national standard setter / regulator and any local formats (e.g. ACAR for CIFRS in Cambodia)", "-"),
        ]
        for i, (a, b, c) in enumerate(rows):
            r = 6 + i
            put(ws, r, 1, a, font=F_BOLD, align=WRAP)
            put(ws, r, 2, b, align=WRAP)
            put(ws, r, 3, c, font=F_SUB, align=WRAP)
            for col in range(4, 10):
                inp(ws, r, col, "")
        last = 5 + len(rows)
        add_list_validation(ws, f"E6:E{last}", '"Y,N"')
        add_list_validation(ws, f"F6:F{last}", '"H,M,L"')
        put(ws, last + 2, 1, "High-impact gaps", font=F_BOLD)
        put(ws, last + 2, 2, f'=COUNTIFS(E6:E{last},"Y",F6:F{last},"H")', fmt="0")
        widths(ws, [24, 62, 22, 40, 9, 10, 40, 14, 14])
        ws.freeze_panes = "B6"

    def smba(self):
        ws = self.ws["04_SMBA"]
        title(ws, "Specified main business activity (SMBA) assessment", "A matter of fact supported by evidence, not an assertion (IFRS 18.B30-B41). Reassess each period.", "04")
        r = 5
        self.smba_concl = {}
        for key, head, inds in [
            ("inv", "A. Investing in particular types of assets (e.g. investment property, financial assets) as a main business activity", [
                "Entity's business model is to generate returns from assets that produce a return individually and largely independently (e.g. investment entity, property investor)",
                "A reportable segment (IFRS 8) comprises a single business activity of investing in such assets",
                "Management uses a performance measure / subtotal based on returns from these investments as an important indicator of operating performance",
                "Income from these assets is a significant share of total income and is reported to the CODM as core performance",
                "Group perspective: activity of a subsidiary is a main business activity of the group as a whole (consolidated view)",
            ]),
            ("fin", "B. Providing financing to customers as a main business activity (e.g. banks, finance companies, finance lessors, manufacturers' finance arms)", [
                "Entity provides loans/finance leases to customers as part of its business model",
                "A reportable segment comprises a single business activity of providing financing to customers",
                "Net interest margin or similar subtotal is an important indicator of operating performance",
                "Financing income is a significant share of total income and is reported to the CODM as core performance",
                "Group perspective: financing activity of a subsidiary is a main business activity of the group",
            ]),
        ]:
            put(ws, r, 1, head, font=F_BOLD, fill=FILL_SEC)
            r += 1
            header(ws, r, ["Indicator", "Present? (Y/N)", "Evidence (document reference)", "Weight"])
            r += 1
            first = r
            for t in inds:
                put(ws, r, 1, t, align=WRAP)
                inp(ws, r, 2, "N")
                inp(ws, r, 3, "")
                inp(ws, r, 4, 1, "0")
                ws.row_dimensions[r].height = 30
                r += 1
            add_list_validation(ws, f"B{first}:B{r-1}", '"Y,N"')
            put(ws, r, 1, "Indicator score (weighted share of indicators present)", font=F_BOLD)
            put(ws, r, 2, f'=SUMPRODUCT((B{first}:B{r-1}="Y")*D{first}:D{r-1})/SUM(D{first}:D{r-1})', fmt=PCT, font=F_BOLD)
            r += 1
            put(ws, r, 1, "Preliminary indication (reviewer must confirm; judgement, not arithmetic)", font=F_BOLD)
            put(ws, r, 2, f'=IF(B{r-1}>=0.5,"Indicates SMBA - document judgement","No SMBA indicated")', font=F_BOLD)
            r += 1
            put(ws, r, 1, "Conclusion carried to 01_Setup", font=F_BOLD)
            put(ws, r, 2, f"={self.S('smba_investing' if key == 'inv' else 'smba_financing')}", font=F_BOLD, fill=FILL_KEY)
            r += 2
        for n in ["Consequences of an SMBA:",
                  "- Investing SMBA: income and expenses that would otherwise be investing are classified in operating (except equity-accounted investees, which stay investing) (IFRS 18.55-58).",
                  "- Financing-to-customers SMBA: income/expenses from liabilities relating to providing financing to customers are operating; accounting policy choice for other financing liabilities (IFRS 18.65); "
                  "cash and cash equivalents income may be operating; PBFIT subtotal may not be required (IFRS 18.69(b), 73).",
                  "- Statement of cash flows: interest and dividend classification follows the statement of profit or loss classification.",
                  "- A change in SMBA conclusion is applied prospectively from the date of change; disclose (verify paragraph in IFRS 18.51-52)."]:
            put(ws, r, 1, n, font=F_SUB)
            r += 1
        widths(ws, [100, 32, 40, 8])

    def rules(self):
        ws = self.ws["06_Rules"]
        title(ws, "Classification rule book - income and expenses by item type",
              "Use when mapping 05_TB. Where an entity has an SMBA the right-hand columns apply. Confirm against the Standard for complex cases.", "06")
        header(ws, 5, ["Item type", "No SMBA", "SMBA - investing in assets", "SMBA - financing customers", "Reference", "Notes / judgement"])
        rows = [
            ("Revenue from goods and services; cost of sales; selling, admin, R&D", OP, OP, OP, "IFRS 18.52", "Operating is the residual category"),
            ("Impairment of trade receivables / contract assets (IFRS 9)", OP, OP, OP, "IFRS 18.75(b)(ii)", ""),
            ("Gain/loss on disposal of PPE and intangibles", OP, OP, OP, "IFRS 18.B79", ""),
            ("Fair value changes of biological assets; government grants", OP, OP, OP, "IFRS 18.52", ""),
            ("Share of profit/loss of equity-accounted associates and JVs", INV, INV, INV, "IFRS 18.53(a), 55(a), B43", "Always investing, even with an SMBA"),
            ("Gains/losses on disposal of subsidiaries, associates, JVs", INV, "Operating/Investing - assess", INV, "IFRS 18.53(a), B43-B45", "Includes remeasurement of previously held interests"),
            ("Income from cash and cash equivalents", INV, INV, "Operating (see IFRS 18.56)", "IFRS 18.53(b), 56", "Verify policy for SMBA entities"),
            ("Investment property: rental income, FV changes, incremental directly attributable costs", INV, OP, INV, "IFRS 18.53(c), 55, B47", "Maintenance not incremental may be operating (judgement)"),
            ("Financial assets (debt/equity investments): interest, dividends, FV changes, impairment, derecognition gains", INV, OP, OP, "IFRS 18.53(c), B46-B47", ""),
            ("Interest expense on loans, bonds, other liabilities that involve only raising finance", FIN, FIN, "Operating (financing customers) / policy choice", "IFRS 18.59(a), 61, 64-65", ""),
            ("Interest on lease liabilities, unwind of provision discounts, net interest on DB liability", FIN, FIN, FIN, "IFRS 18.59(b), 60-61, B52-B54", "Only interest and effects of interest-rate changes; other income/expenses on these liabilities are operating"),
            ("FX differences", "Same category as the item giving rise", "Same category as the item", "Same category as the item", "IFRS 18.B65-B67", "Undue cost or effort relief - operating; intra-group balances - operating"),
            ("Derivatives designated as hedging instruments", "Category of hedged item", "Category of hedged item", "Category of hedged item", "IFRS 18.B70-B71", "Ineffectiveness follows effective portion"),
            ("Derivatives used to manage identified risks (not designated)", "Category affected by the risk", "Category affected by the risk", "Category affected by the risk", "IFRS 18.B70, B72-B73", "If gross-up required or undue cost - operating"),
            ("Derivatives not used to manage identified risks", INV, OP, OP, "IFRS 18.B74-B76", "Financing if relates to raising finance only - verify"),
            ("Hybrid contracts with host liabilities", "Per host / embedded derivative", "Per host", "Per host", "IFRS 18.B77", ""),
            ("Income tax expense (IAS 12)", TAX, TAX, TAX, "IFRS 18.67", ""),
            ("Discontinued operations (IFRS 5)", DISC, DISC, DISC, "IFRS 18.68", ""),
            ("Net monetary position gain/loss (hyperinflation, IAS 29)", OP, OP, OP, "IFRS 18.B68-B69", "Verify"),
        ]
        for i, row in enumerate(rows):
            for j, v in enumerate(row):
                put(ws, 6 + i, 1 + j, v, font=F_SUB if j >= 4 else F_NORM, align=WRAP)
        widths(ws, [58, 22, 26, 30, 24, 50])
        ws.freeze_panes = "B6"

    def aggregation(self):
        ws = self.ws["11_Aggregation"]
        title(ws, "Aggregation, disaggregation and labelling review (IFRS 18.41-45, B15-B26)",
              "Flags lines labelled 'other' or above performance materiality for review. Inputs record the judgement.", "11")
        header(ws, 5, ["Statement", "Line item", "CY amount", "Label contains 'other'?", "Above performance materiality?", "System prompt",
                       "Shared characteristics? (Y/N)", "Further disaggregated in notes? (note ref)", "Label informative? (Y/N)", "Reviewer comment"])
        r = 6
        pm = self.S("performance_materiality")
        for key, (sheet, row) in self.pos.items():
            lab = self.cat[key][0]
            put(ws, r, 1, "SPLOCI" if sheet == "08_SPLOCI" else "SFP", font=F_SUB)
            put(ws, r, 2, lab)
            put(ws, r, 3, f"='{sheet}'!D{row}", fmt=NUM)
            put(ws, r, 4, f'=IF(ISNUMBER(SEARCH("other",B{r})),"Yes","No")')
            put(ws, r, 5, f'=IF(ABS(C{r})>={pm},"Yes","No")')
            put(ws, r, 6, f'=IF(AND(D{r}="Yes",E{r}="Yes"),"Explain composition / relabel (B24-B26)",IF(E{r}="Yes","Check dissimilar characteristics (B15-B23)",""))', font=F_SUB)
            for c in range(7, 11):
                inp(ws, r, c, "")
            r += 1
        add_list_validation(ws, f"G6:G{r-1}", '"Y,N"')
        add_list_validation(ws, f"I6:I{r-1}", '"Y,N"')
        ws.conditional_formatting.add(f"F6:F{r-1}", FormulaRule(formula=['LEFT(F6,7)="Explain"'], fill=PatternFill("solid", fgColor="F8CBAD")))
        widths(ws, [9, 58, 13, 12, 13, 40, 14, 22, 12, 34])
        ws.freeze_panes = "C6"

    def mpm_id(self):
        ws = self.ws["12_MPM_ID"]
        title(ws, "Identification of management-defined performance measures (IFRS 18.117-120)",
              "A measure is an MPM only if ALL of (a)-(c) are Yes and it is NOT excluded under (d). Presumption in IFRS 18.119-120 can be rebutted only with evidence.", "12")
        header(ws, 5, ["Candidate measure", "Where communicated", "(a) Subtotal of income and expenses?", "(b) Used in public communications outside the FS?",
                       "(c) Communicates management's view of an aspect of performance of the entity as a whole?",
                       "(d) Excluded - listed in IFRS 18.118 or required by IFRS?", "Conclusion", "Most directly comparable subtotal", "Rationale"])
        cands = [
            ("Adjusted operating profit", "Results press release; investor presentation", "Y", "Y", "Y", "N", "Operating profit",
             "Excludes infrequent impairments and flood costs; also numerator of return on capital (B117)."),
            ("EBITDA = operating profit before depreciation, amortisation and IAS 36 impairments", "Investor presentation", "Y", "Y", "Y", "Y", "-",
             "Listed subtotal (operating profit before D&A and impairments in scope of IAS 36) - not an MPM if calculated exactly so."),
            ("Adjusted EBITDA (also excludes share-based payment)", "Analyst calls", "Y", "Y", "Y", "N", "Operating profit",
             "Adjustment beyond the listed subtotal makes it an MPM."),
            ("Gross profit", "Annual report", "Y", "Y", "Y", "Y", "-", "Listed in IFRS 18.118 - not an MPM."),
            ("Revenue growth %", "Press release", "N", "Y", "Y", "N", "-", "Not a subtotal of income and expenses."),
            ("Free cash flow", "Investor presentation", "N", "Y", "Y", "N", "-", "Cash flow measure - not income and expenses."),
            ("Return on capital (ratio)", "Annual report", "N", "Y", "Y", "N", "-", "A ratio is not an MPM; its numerator/denominator may be (B117)."),
            ("Adjusted net profit (internal board pack only)", "Internal reporting", "Y", "N", "Y", "N", "-", "Not used in public communications."),
        ]
        for i, row in enumerate(cands):
            r = 6 + i
            put(ws, r, 1, row[0], align=WRAP)
            inp(ws, r, 2, row[1])
            for j in range(4):
                inp(ws, r, 3 + j, row[2 + j])
            put(ws, r, 7, f'=IF(AND(C{r}="Y",D{r}="Y",E{r}="Y",F{r}="N"),"MPM - disclose in 13_MPM_Recon","Not an MPM")', font=F_BOLD)
            inp(ws, r, 8, row[6])
            inp(ws, r, 9, row[7]).alignment = WRAP
            ws.row_dimensions[r].height = 42
        last = 5 + len(cands)
        add_list_validation(ws, f"C6:F{last}", '"Y,N"')
        ws.conditional_formatting.add(f"G6:G{last}", FormulaRule(formula=['LEFT(G6,3)="MPM"'], fill=FILL_KEY))
        r = last + 2
        put(ws, r, 1, "Number of MPMs identified", font=F_BOLD)
        put(ws, r, 2, f'=COUNTIF(G6:G{last},"MPM*")', font=F_BOLD, fmt="0")
        r += 2
        put(ws, r, 1, "Subtotals listed in IFRS 18.118 / B124 that are NOT MPMs (verify list against the Standard):", font=F_BOLD)
        for i, t in enumerate(["Gross profit (revenue less cost of sales) and similar subtotals",
                               "Operating profit before depreciation, amortisation and impairments within the scope of IAS 36",
                               "Operating profit and income and expenses from all investments accounted for using the equity method",
                               "Operating profit and income and expenses from investments accounted for using the equity method (subset) - check B124",
                               "Profit or loss before income taxes; profit or loss from continuing operations",
                               "Totals and subtotals specifically required by IFRS Accounting Standards (e.g. operating profit, PBFIT, profit, OCI, TCI)"]):
            put(ws, r + 1 + i, 1, "- " + t, font=F_SUB)
        widths(ws, [52, 30, 13, 15, 22, 18, 30, 24, 60])
        ws.freeze_panes = "B6"

    def disclosure(self):
        ws = self.ws["16_Disclosure"]
        title(ws, "IFRS 18 presentation and disclosure checklist", "Respond Y / N / NA; reference the note and working paper. Paragraph references should be confirmed against the issued Standard.", "16")
        header(ws, 5, ["Area", "Requirement", "Reference", "Response (Y/N/NA)", "Note ref", "WP ref", "Comment"])
        items = [
            ("General", "Complete set of financial statements and titles; notes; comparative information", "IFRS 18.10-11, 31-35"),
            ("General", "Roles: primary statements provide a useful structured summary; notes provide material additional information", "IFRS 18.15-17, B8-B14"),
            ("General", "Material information not obscured; immaterial requirements not presented", "IFRS 18.19-20, B1-B5"),
            ("General", "Aggregation and disaggregation based on shared / non-shared characteristics", "IFRS 18.41-43, B15-B23"),
            ("General", "Informative labels; explanation of items labelled 'other'", "IFRS 18.B24-B26"),
            ("General", "Offsetting only when required or permitted", "IFRS 18.44-45"),
            ("General", "Third statement of financial position when retrospective change is material to it", "IFRS 18.37-38"),
            ("Profit or loss", "Five categories applied; SMBA assessed and documented", "IFRS 18.47-51"),
            ("Profit or loss", "Operating profit, PBFIT and profit or loss presented", "IFRS 18.69-74"),
            ("Profit or loss", "Required line items (revenue, interest revenue, operating expenses, EAI share, tax, discontinued ops, impairments)", "IFRS 18.75"),
            ("Profit or loss", "Allocation of profit to owners and NCI", "IFRS 18.76"),
            ("Profit or loss", "Operating expenses by nature/function/mixed - most useful structured summary", "IFRS 18.78-81, B80-B85"),
            ("Profit or loss", "Description of what each function line includes", "IFRS 18.82"),
            ("Profit or loss", "Five specified amounts by nature for each function line (single note)", "IFRS 18.83-85, B84"),
            ("Profit or loss", "Additional subtotals fit within the category structure and are not displayed with more prominence", "IFRS 18.24, B78"),
            ("OCI", "OCI totals, attribution, two groups (reclassified / not reclassified)", "IFRS 18.86-88"),
            ("OCI", "Share of OCI of equity-accounted investees split into the two groups", "IFRS 18.89"),
            ("OCI", "Reclassification adjustments presented or disclosed", "IFRS 18.90-92"),
            ("OCI", "Income tax relating to each component of OCI", "IFRS 18.93-95"),
            ("SFP", "Current/non-current (or liquidity) presentation and classification criteria", "IFRS 18.96-101"),
            ("SFP", "Required line items incl. separate goodwill line", "IFRS 18.103-105"),
            ("SCE", "Statement of changes in equity components and reconciliations", "IFRS 18.107-111"),
            ("Notes", "Systematic order and cross-referencing between statements and notes", "IFRS 18.113-115, B112"),
            ("Notes", "Share capital, reserves, dividends and capital-related disclosures carried over from IAS 1", "IFRS 18.126-132 (verify)"),
            ("MPM", "Single note covering all MPM requirements", "IFRS 18.121-125"),
            ("MPM", "Reconciliation, tax effect, NCI effect, method of determining tax effect", "IFRS 18.123, B136-B141"),
            ("MPM", "Changes to MPMs explained, comparatives restated", "IFRS 18.124"),
            ("IAS 7", "Indirect method reconciliation starts from operating profit", "IAS 7.18(b), 20"),
            ("IAS 7", "Interest and dividend cash flows classified per amended requirements", "IAS 7.33A-34D"),
            ("IAS 33", "Additional per-share amounts only on permitted numerators; equal prominence rules", "IAS 33.73B-73C"),
            ("IAS 34", "Condensed interim statements use IFRS 18 headings and subtotals; MPM disclosures in interim reports", "IAS 34 (amended); IFRS 18.C5"),
            ("IAS 8", "Basis of preparation, compliance statement, going concern (now in IAS 8)", "IAS 8 (amended)"),
            ("Transition", "Retrospective application with restated comparatives", "IFRS 18.C2"),
            ("Transition", "Reconciliation of each P&L line item IAS 1 vs IFRS 18 for the immediately preceding comparative period", "IFRS 18.C3"),
            ("Transition", "Early application disclosed (if applicable)", "IFRS 18.C1"),
            ("Transition", "IAS 8.28 change-in-policy disclosures as modified by IFRS 18 transition reliefs", "IFRS 18.C4; IAS 8.28"),
        ]
        for i, (a, b, c) in enumerate(items):
            r = 6 + i
            put(ws, r, 1, a, font=F_BOLD)
            put(ws, r, 2, b, align=WRAP)
            put(ws, r, 3, c, font=F_SUB)
            for col in range(4, 8):
                inp(ws, r, col, "")
        last = 5 + len(items)
        add_list_validation(ws, f"D6:D{last}", '"Y,N,NA"')
        put(ws, last + 2, 2, "Items answered", font=F_BOLD)
        put(ws, last + 2, 4, f'=COUNTA(D6:D{last})-COUNTBLANK(D6:D{last})&" of {len(items)}"')
        put(ws, last + 3, 2, "Items answered 'N' (non-compliance)", font=F_BOLD)
        put(ws, last + 3, 4, f'=COUNTIF(D6:D{last},"N")', fmt="0")
        widths(ws, [14, 80, 26, 12, 10, 10, 30])
        ws.freeze_panes = "B6"

    def reclass(self):
        ws = self.ws["17_Reclass"]
        title(ws, "Reclassification entries - IAS 1 caption to IFRS 18 line (comparative period)",
              "Presentation-only entries for the restated comparatives / CoA remapping. No effect on profit or net assets.", "17")
        header(ws, 5, ["Account", "Description", "IAS 1 caption", "IFRS 18 line", "PY amount (Dr+/Cr-)", "Reclass needed?",
                       "Debit caption", "Credit caption", "Amount", "Transition ref"])
        r = 6
        first = r
        for i, row in enumerate(self.tb):
            if row[2] != "PL":
                continue
            tr = self.TB_FIRST + i
            put(ws, r, 1, f"='05_TB'!A{tr}")
            put(ws, r, 2, f"='05_TB'!B{tr}")
            put(ws, r, 3, f"='05_TB'!F{tr}")
            put(ws, r, 4, f"='05_TB'!G{tr}")
            put(ws, r, 5, f"='05_TB'!E{tr}", fmt=NUM)
            put(ws, r, 6, f'=IF(C{r}=D{r},"No","Yes")')
            put(ws, r, 7, f'=IF(F{r}="No","",IF(E{r}>0,D{r},C{r}))')
            put(ws, r, 8, f'=IF(F{r}="No","",IF(E{r}>0,C{r},D{r}))')
            put(ws, r, 9, f'=IF(F{r}="No",0,ABS(E{r}))', fmt=NUM)
            put(ws, r, 10, f"='05_TB'!K{tr}&\"\"")
            r += 1
        put(ws, r, 2, "Total reclassified (gross)", font=F_BOLD)
        put(ws, r, 9, f"=SUM(I{first}:I{r-1})", font=F_BOLD, fmt=NUM, border=B_TOT)
        put(ws, r + 1, 2, "Accounts reclassified", font=F_BOLD)
        put(ws, r + 1, 9, f'=COUNTIF(F{first}:F{r-1},"Yes")', font=F_BOLD, fmt="0")
        ws.conditional_formatting.add(f"F{first}:F{r-1}", CellIsRule(operator="equal", formula=['"Yes"'], fill=FILL_KEY))
        widths(ws, [9, 46, 36, 46, 13, 10, 40, 40, 12, 8])
        ws.freeze_panes = "C6"
        ws.auto_filter.ref = f"A5:J{r-1}"

    def issues(self):
        ws = self.ws["18_Issues"]
        title(ws, "Judgements, open issues and review notes", "Record each significant judgement (also feeds the IAS 8 / IFRS 18 judgement disclosures).", "18")
        header(ws, 5, ["#", "Area", "Issue / judgement", "Reference", "Conclusion", "Raised by", "Cleared by", "Status"])
        seed = [
            ("SMBA", "Does the group invest in assets or provide financing to customers as a main business activity?", "IFRS 18.B30-B41"),
            ("Expense method", "Nature, function or mixed - which gives the most useful structured summary?", "IFRS 18.78, B80-B85"),
            ("Investment property costs", "Are maintenance costs 'incremental and directly attributable' (investing) or operating?", "IFRS 18.B47"),
            ("FX", "Can FX differences be traced to source items without undue cost or effort?", "IFRS 18.B65-B67"),
            ("Derivatives", "Purpose of each derivative class (designated hedge / economic hedge / trading)", "IFRS 18.B70-B76"),
            ("MPM", "Are all public-communication subtotals identified; is the presumption rebutted for any?", "IFRS 18.117-120"),
            ("MPM tax effect", "Method for tax effect of reconciling items (statutory rate / allocation / other)", "IFRS 18.B141"),
            ("Third SFP", "Is a third statement of financial position required?", "IFRS 18.37-38"),
            ("'Other' items", "Are material 'other' aggregations explained or relabelled?", "IFRS 18.B24-B26"),
        ]
        for i, (a, b, c) in enumerate(seed):
            r = 6 + i
            put(ws, r, 1, i + 1)
            put(ws, r, 2, a, font=F_BOLD)
            put(ws, r, 3, b, align=WRAP)
            put(ws, r, 4, c, font=F_SUB)
            for col in range(5, 8):
                inp(ws, r, col, "")
            inp(ws, r, 8, "Not started")
        add_list_validation(ws, "H6:H200", f"={self.status_rng}")
        widths(ws, [4, 22, 62, 20, 50, 12, 12, 13])
        ws.freeze_panes = "C6"

    def checks_sheet(self):
        ws = self.ws["19_Checks"]
        title(ws, "Control checks", "Every check must show OK (tolerance 0.5) before the file is signed off.", "19")
        put(ws, 4, 1, "Overall status", font=F_BOLD)
        header(ws, 6, ["#", "Check", "Actual", "Expected", "Difference", "Status", "WP"])
        r = 7
        for i, (desc, act, exp, wp) in enumerate(self.checks):
            put(ws, r, 1, i + 1)
            put(ws, r, 2, desc)
            put(ws, r, 3, f"={act}", fmt=NUM)
            put(ws, r, 4, f"={exp}", fmt=NUM)
            put(ws, r, 5, f"=C{r}-D{r}", fmt=NUM)
            put(ws, r, 6, f'=IF(ABS(E{r})<0.5,"OK","ERROR")', font=F_BOLD)
            put(ws, r, 7, f'=HYPERLINK("#\'{wp}\'!A1","{wp}")', font=F_LINK)
            r += 1
        put(ws, 4, 2, f'=IF(COUNTIF(F7:F{r-1},"OK")={len(self.checks)},"ALL CHECKS OK",COUNTIF(F7:F{r-1},"ERROR")&" CHECK(S) FAILING")', font=F_BOLD)
        ws.conditional_formatting.add(f"F7:F{r-1}", CellIsRule(operator="equal", formula=['"ERROR"'], fill=PatternFill("solid", fgColor="F8CBAD")))
        ws.conditional_formatting.add(f"F7:F{r-1}", CellIsRule(operator="equal", formula=['"OK"'], fill=FILL_KEY))
        widths(ws, [4, 70, 14, 14, 12, 10, 20])

    def build(self, out):
        self.setup()
        self.lists()
        self.trial_balance()
        self.sploci()
        self.sfp()
        self.classification()
        self.transition()
        self.exp_nature()
        self.mpm_recon()
        self.cashflow()
        self.index()
        self.plan()
        self.gap()
        self.smba()
        self.rules()
        self.aggregation()
        self.mpm_id()
        self.disclosure()
        self.reclass()
        self.issues()
        self.checks_sheet()
        self.ws["Lists"].sheet_state = "visible"
        for ws in self.wb.worksheets:
            ws.sheet_properties.tabColor = {"0": NAVY, "1": "2E75B6"}.get(ws.title[0], "7F7F7F")
            ws.page_setup.orientation = "landscape"
            ws.page_setup.fitToWidth = 1
            ws.sheet_properties.pageSetUpPr.fitToPage = True
            ws.page_setup.fitToHeight = 0
        self.wb.active = 0
        self.wb.save(out)
        return out


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default="IFRS18_Workpapers.xlsx")
    ap.add_argument("--tb", help="Trial balance CSV (see assets/tb_template.csv). Omit for demo data.")
    ap.add_argument("--setup", help="JSON file overriding setup inputs")
    ap.add_argument("--export-template", help="Write the demo TB as a CSV template to this path and exit")
    a = ap.parse_args()
    if a.export_template:
        with open(a.export_template, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["account", "description", "type", "cy", "py", "ias1_line", "ifrs18_line", "nature", "mpm_item", "transition_ref"])
            w.writerows(demo_data.TB)
        print("Template written:", a.export_template)
        return
    tb = load_tb(a.tb) if a.tb else demo_data.TB
    setup = dict(demo_data.SETUP)
    if a.setup:
        with open(a.setup, encoding="utf-8") as f:
            setup.update(json.load(f))
    cf = demo_data.CASHFLOW
    cap = demo_data.CAPITALISED if not a.tb else {}
    expl = demo_data.TRANSITION_EXPLANATIONS
    out = Builder(tb, setup, cf, cap, expl).build(a.out)
    print("Workbook written:", out)
    if a.tb:
        print("NOTE: 15_CashFlow inputs and 09_Transition explanations still hold demo values - replace them with the entity's.")


if __name__ == "__main__":
    main()
