#!/usr/bin/env python3
"""Review a Fixed Asset Register (FAR) + draft PP&E chart of accounts and build an
opening-balance (OB) workbook.

Usage:
    python3 -I review_far.py INPUT.xlsx OUTPUT.xlsx [--ob-date 2027-01-01]
        [--sheet "FA Details"] [--summary SUMMARY] [--extra-findings findings.json]

The input workbook is read only. Output sheets:
    Read Me | Findings | Revised COA | OB TB | OB Journal | OB Register | Exceptions
"""
import argparse
import datetime as dt
import json
import re
import sys
from collections import Counter, defaultdict

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

# ---------------------------------------------------------------- input layout
COL = dict(no=1, status=2, coa=3, tag=4, old=5, main=6, ref=7, date=8, desc=9,
           cost=10, method=11, life=12, resid=13, open_ad=18, end_ad=31, nbv=32)
FIRST_DATA_ROW = 11

# ---------------------------------------------------------------- revised COA
# abbr: (old_code, old_name, new_code, new_name, ad_code, life_policy, change_note)
H = "header"
COA = [
    (None, "121010", "Freehold and ELC Land", "121010", "Land", None, None, "Header – not postable"),
    ("FRL", "121011", "Freehold Land", "121011", "Freehold Land", None, "Not depreciated", "No change"),
    ("LCD", "121012", "Land Clearing & Development-ELC", "121012", "Land Clearing & Development – ELC", "121912",
     "Shorter of useful life and remaining ELC term", "Name standardised; AD contra added"),
    ("OLI", "121013", "Other Land Improvements - ELC", "121013", "Land Improvements – ELC (Roads, Drainage, Dams)", "121913",
     "5–20 yrs, capped at remaining ELC term", "Renamed to describe content; AD contra added"),
    (None, "121020", "Buildings on ELC Land", "121020", "Buildings", None, None, "Header – not postable"),
    ("OFB", "121021", "Office Buildings", "121021", "Office Buildings", "121921", "20 yrs", "AD contra added; life policy fixed at 20 yrs"),
    ("STH", "121022", "Staff Houses", "121022", "Staff Houses", "121922", "20 yrs", "AD contra added; detail uses 5 yrs – review"),
    ("WHS", "121023", "Warehouses and Stores", "121023", "Warehouses & Stores", "121923", "20 yrs", "AD contra added"),
    ("FAB", "121024", "Processing Facility Buildings", "121024", "Processing Facility Buildings", "121924", "20 yrs", "AD contra added"),
    ("IMP", "121030", "Impr.Lhold Property,Plants&Buil", "121030", "Leasehold Improvements – Property, Plant & Buildings", "121930",
     "Shorter of useful life and lease term", "Truncated name corrected; AD contra added"),
    (None, "121040", "Machinery and Heavy Equipments", "121040", "Machinery & Heavy Equipment", None, None, "Header – not postable"),
    ("AGM", "121041", "Agricultural Machinery", "121041", "Agricultural Machinery", "121941", "5 yrs", "AD contra added"),
    ("RPM", "121042", "Rubber Processing Machine", "121042", "Rubber Processing Machinery", "121942", "10 yrs", "AD contra added"),
    ("GPE", "121043", "Generators and Power Equipment", "121043", "Generators & Power Equipment", "121943", "5 yrs", "AD contra added"),
    ("OEM", "121044", "Other Heavy Equipment", "121044", "Other Heavy Equipment", "121944", "5 yrs",
     "AD contra added; appliances (A/C, fridges, washers) reclassed to 121052"),
    (None, "121050", "Office Furniture and Fixture", "121050", "Furniture, Fixtures & Office Equipment", None, None,
     "Converted to header; split into 121051 / 121052"),
    ("OFF", "121050", "Office Furniture and Fixture", "121051", "Furniture & Fixtures (incl. shelving/racking)", "121951", "5 yrs",
     "New sub-account (was 121050)"),
    ("OEQ", None, "(new)", "121052", "Office Equipment & Appliances", "121952", "3–5 yrs", "NEW – receives A/C, fridges, washers, counters"),
    ("ITE", "121060", "IT System and Equipment", "121060", "IT Systems & Equipment", "121960", "3 yrs", "AD contra added"),
    ("VEH", "121070", "Vehicles and Trucks", "121070", "Motor Vehicles & Trucks", "121970", "5 yrs", "AD contra added"),
    (None, "121080", "Biological Assets", "121080", "Bearer Plants (IAS 16)", None, None,
     "Renamed – bearer plants are PP&E, not IAS 41 biological assets"),
    ("RUB", "121081", "Bearer Plants - Rubber", "121081", "Bearer Plants – Rubber (Mature)", "121981", "Productive life (e.g. 25 yrs) – not GDT rate",
     "Policy changed from 'GDT Rate' to accounting useful life"),
    ("CAS", "121082", "Bearer Plants - Cashew", "121082", "Bearer Plants – Cashew (Mature)", "121982", "Productive life – not GDT rate",
     "Policy changed from 'GDT Rate' to accounting useful life"),
    ("OAP", "121087", "Other Agriculture Plants", "121087", "Bearer Plants – Other Tree Crops (Mature)", "121987", "Productive life",
     "Renamed; seedlings/seeds not yet planted belong in inventory or immature"),
    ("IRT", "121088", "Immatured Rubber Tree", "121088", "Bearer Plants – Rubber (Immature)", None, "Not depreciated until mature",
     "Renamed; no depreciation"),
    ("IOT", "121089", "Immatured Other Tree", "121089", "Bearer Plants – Other (Immature)", None, "Not depreciated until mature",
     "Renamed; no depreciation"),
    ("LAN", "121090", "Landscape and Infrastructure", "121090", "Site & Utility Infrastructure (Water, Power, Landscape)", "121990",
     "5–10 yrs", "Broadened – receives water towers/tanks/wells, poles, solar from 121098"),
    ("CWP", None, "(new)", "121097", "Capital Work-in-Progress", None, "Not depreciated", "NEW – costs of assets not yet available for use"),
    ("OFA", "121098", "Other Fixed Assets", "121098", "Other Fixed Assets", "121998", "3–5 yrs",
     "Restricted to genuinely miscellaneous items"),
    (None, None, "(new)", "121900", "Accumulated Depreciation – PP&E", None, None, "NEW header – contra-asset accounts 1219xx (one per class)"),
    ("IMPAIR", None, "(new)", "121999", "Accumulated Impairment – PP&E", None, None, "NEW – damaged / impaired assets (IAS 36 / s.27)"),
    ("LIS", "121083", "Livestocks", "123010", "Biological Assets – Livestock (IAS 41)", None, "Fair value less costs to sell",
     "MOVED out of PP&E – IAS 41 / s.34 asset"),
]
CLASS = {r[0]: r for r in COA if r[0]}

TREATMENT = {
    "Confirmed Assets": "Recognise",
    "Unsupported Assets": "Recognise – pending support",
    "Damaged Assets": "Recognise & impair",
    "NonExistent Assets": "Exclude",
}

RECLASS_RULES = [  # (from_class, regex on description, to_class, reason)
    ("OFA", r"water ?tank|ស៊ីទែន|water ?tower|ប៉មទឹក|well|អណ្តូង|បង្គោលភ្លើង|solar|lightning|ខ្សែអគ្គសនី|pole", "LAN",
     "Water/power/site infrastructure"),
    ("OFA", r"shel|rack|bed frame|គ្រែ", "OFF", "Shelving/racking & furniture"),
    ("OFA", r"ទូរទឹកកក|fridge|air ?condition", "OEQ", "Appliance"),
    ("OEM", r"air ?condition|ម៉ាស៊ីនត្រជាក់|fridge|ទូរទឹកកក|washer|bill counter", "OEQ", "Appliance, not heavy equipment"),
]
BUILDINGS = {"OFB", "STH", "WHS", "FAB"}

# ---------------------------------------------------------------- styles
F = "Arial"
f_norm = Font(name=F, size=10)
f_bold = Font(name=F, size=10, bold=True)
f_title = Font(name=F, size=14, bold=True)
f_sub = Font(name=F, size=10, italic=True, color="595959")
f_input = Font(name=F, size=10, color="0000FF")
f_head = Font(name=F, size=10, bold=True, color="FFFFFF")
fill_head = PatternFill("solid", start_color="1F4E78")
fill_tot = PatternFill("solid", start_color="DDEBF7")
fill_flag = PatternFill("solid", start_color="FFF2CC")
fill_new = PatternFill("solid", start_color="E2EFDA")
thin = Side(style="thin", color="BFBFBF")
box = Border(top=thin, bottom=thin, left=thin, right=thin)
NUM = '#,##0.00;(#,##0.00);"-"'
DATE = "dd-mmm-yyyy"


def header(ws, row, labels, widths=None):
    for i, lab in enumerate(labels, 1):
        c = ws.cell(row, i, lab)
        c.font, c.fill, c.border = f_head, fill_head, box
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[L(i)].width = w
    ws.row_dimensions[row].height = 32
    ws.freeze_panes = ws.cell(row + 1, 1)


def title(ws, text, ob_label, sub=None):
    ws["A1"], ws["A1"].font = text, f_title
    ws["A2"], ws["A2"].font = f"Opening balance as at {ob_label}", f_bold
    if sub:
        ws["A3"], ws["A3"].font = sub, f_sub


def num(x):
    return float(x) if isinstance(x, (int, float)) and not isinstance(x, bool) else 0.0


# ---------------------------------------------------------------- read source
def read_far(path, sheet, ob_date):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet]
    assets, sec = [], None
    for r in range(FIRST_DATA_ROW, ws.max_row + 1):
        g = lambda k: ws.cell(r, COL[k]).value
        a = g("no")
        if isinstance(a, str) and re.match(r"^\d{6}", a.strip()):
            sec = ws.cell(r, 4).value
            continue
        if not g("status") and not g("coa"):
            continue
        assets.append(dict(src_row=r, section=sec, status=g("status"), coa=g("coa") or sec, tag=g("tag"),
                           main=g("main"), ref=g("ref"), date=g("date"), desc=g("desc"),
                           cost=num(g("cost")), method=g("method"), life=g("life"),
                           end_ad=num(g("end_ad")), nbv=num(g("nbv"))))
    unrecorded = {}
    return wb, assets, unrecorded


def read_unrecorded(wb, summary):
    out = {}
    if summary not in wb.sheetnames:
        return out
    ws = wb[summary]
    for r in range(1, ws.max_row + 1):
        code, val = ws.cell(r, 3).value, ws.cell(r, 9).value
        if isinstance(code, str) and len(code) == 3 and isinstance(val, (int, float)) and val:
            out[code] = float(val)
    return out


# ---------------------------------------------------------------- checks
def exceptions_for(a, ob_date, tag_counts):
    ex = []
    d = a["date"]
    if d is None:
        ex.append("Acquisition date missing")
    elif isinstance(d, (int, float)):
        ex.append("Acquisition date is year only")
    elif isinstance(d, dt.datetime) and d.date() >= ob_date:
        ex.append(f"Acquisition date {d:%d-%b-%Y} is on/after OB date")
    if a["life"] in (None, ""):
        ex.append("Useful life missing")
    if not a["method"]:
        ex.append("Depreciation method missing")
    if abs(a["cost"] - a["end_ad"] - a["nbv"]) > 0.05:
        ex.append(f"NBV ≠ Cost − AD (diff {a['cost'] - a['end_ad'] - a['nbv']:,.2f})")
    t = a["tag"]
    if t in (None, ""):
        ex.append("Tag missing")
    elif not re.search(r"[A-Za-z0-9]", str(t)):
        ex.append(f"Tag is placeholder text '{t}'")
    elif tag_counts[t] > 1:
        ex.append("Duplicate tag")
    if a["cost"] > 0 and abs(a["nbv"]) <= 0.01 and a["status"] != "NonExistent Assets":
        ex.append("Fully depreciated – review useful life (IAS 16.51)")
    if a["coa"] in BUILDINGS and isinstance(a["life"], (int, float)) and a["life"] < 10:
        ex.append(f"Building depreciated over {a['life']:g} yrs vs 20-yr policy")
    if a["coa"] == "OLI" and isinstance(a["life"], (int, float)) and a["life"] < 5:
        ex.append(f"Land improvement over {a['life']:g} yrs – below 5-yr minimum")
    if a["cost"] >= 100000 and not a["tag"]:
        ex.append("Large untagged line (≥100k) – tag, split or support")
    return ex


def proposed_class(a):
    desc = str(a["desc"] or "").lower()
    for frm, rx, to, why in RECLASS_RULES:
        if a["coa"] == frm and re.search(rx, desc, re.I):
            return to, f"Proposed reclass {frm}→{to}: {why} – verify"
    return a["coa"], ""


# ---------------------------------------------------------------- build
def build(args):
    ob_date = dt.date.fromisoformat(args.ob_date)
    cutoff = ob_date - dt.timedelta(days=1)
    ob_label, cut_label = ob_date.strftime("%d %b %Y"), cutoff.strftime("%d %b %Y")
    src_wb, assets, _ = read_far(args.input, args.sheet, ob_date)
    unrecorded = read_unrecorded(src_wb, args.summary)
    tag_counts = Counter(a["tag"] for a in assets if a["tag"])
    for a in assets:
        a["exceptions"] = exceptions_for(a, ob_date, tag_counts)
        a["new_class"], a["reclass_note"] = proposed_class(a)
        a["treatment"] = TREATMENT.get(a["status"], "Review")

    wb = openpyxl.Workbook()
    readme = wb.active
    readme.title = "Read Me"

    # ---------------- OB Register
    reg = wb.create_sheet("OB Register")
    title(reg, "Fixed Asset Register – Opening Balance", ob_label,
          f"Source: {args.sheet} sheet of the draft FAR; balances at cut-off {cut_label}. Blue = values copied from source; black = formulas.")
    cols = ["No.", "OB Date", "Status", "OB Treatment", "Old Class", "Old Account", "Proposed Class",
            "New Account", "New Account Name", "AD Account", "Reclass Note", "Tag", "Reference", "Acq. Date",
            "Description", "Useful Life", f"Cost USD", f"Accum. Depn {cut_label}", f"NBV {cut_label} (source)",
            "OB Cost", "OB Accum. Depn", "OB Impairment", "OB NBV", "Exceptions", "Source Row"]
    header(reg, 5, cols, [6, 12, 18, 22, 8, 10, 9, 10, 34, 10, 34, 16, 18, 12, 40, 8, 14, 14, 14, 14, 14, 13, 14, 50, 8])
    r0 = 6
    for i, a in enumerate(assets):
        r = r0 + i
        old = CLASS.get(a["coa"]) or (None,) * 8
        new = CLASS.get(a["new_class"]) or (None,) * 8
        vals = [i + 1, ob_date, a["status"], a["treatment"], a["coa"], old[1], a["new_class"], new[3], new[4], new[5],
                a["reclass_note"], a["tag"], a["ref"], a["date"], a["desc"], a["life"], a["cost"], a["end_ad"], a["nbv"],
                f'=IF(D{r}="Exclude",0,Q{r})', f'=IF(D{r}="Exclude",0,R{r})',
                f'=IF(D{r}="Recognise & impair",T{r}-U{r},0)', f"=T{r}-U{r}-V{r}",
                "; ".join(a["exceptions"]), a["src_row"]]
        for c, v in enumerate(vals, 1):
            cell = reg.cell(r, c, v)
            cell.font = f_input if c in (17, 18, 19) else f_norm
            if c in (17, 18, 19, 20, 21, 22, 23):
                cell.number_format = NUM
            if c in (2, 14) and isinstance(v, (dt.date, dt.datetime)):
                cell.number_format = DATE
        if a["reclass_note"]:
            reg.cell(r, 7).fill = reg.cell(r, 11).fill = fill_new
        if a["exceptions"]:
            reg.cell(r, 24).fill = fill_flag
    rl = r0 + len(assets) - 1
    tr = rl + 1
    reg.cell(tr, 15, "Total").font = f_bold
    for c in range(17, 24):
        cell = reg.cell(tr, c, f"=SUBTOTAL(9,{L(c)}{r0}:{L(c)}{rl})")
        cell.font, cell.number_format, cell.fill = f_bold, NUM, fill_tot
    reg.auto_filter.ref = f"A5:{L(len(cols))}{rl}"
    RG = lambda col: f"'OB Register'!${col}${r0}:${col}${rl}"

    # ---------------- OB TB
    tb = wb.create_sheet("OB TB")
    title(tb, "PP&E Opening Trial Balance by Revised Account", ob_label,
          f"All figures are SUMIFS over 'OB Register'. NonExistent assets are excluded; Damaged assets are fully impaired.")
    tcols = ["New Account", "Account Name", "AD Account", "Cost – Confirmed", "Cost – Unsupported", "Cost – Damaged",
             "Cost – NonExistent", "Gross Cost (register)", "Less: Excluded", "OB Cost", "Gross Accum. Depn",
             "Less: AD Excluded", "OB Accum. Depn", "OB Impairment", "OB NBV", "Unrecorded (memo – not in OB)", "Check"]
    header(tb, 5, tcols, [11, 44, 11] + [15] * 13 + [8])
    r = 6
    first = r
    tb_rows = {}
    for row in COA:
        abbr, _, _, code, name, ad, _, _ = row
        if abbr in (None, "IMPAIR"):
            continue
        S = lambda col, extra="": f"SUMIFS({RG(col)},{RG('H')},$A{r}{extra})"
        st = lambda s: f",{RG('C')},\"{s}\""
        vals = [code, name, ad,
                "=" + S("Q", st("Confirmed Assets")), "=" + S("Q", st("Unsupported Assets")),
                "=" + S("Q", st("Damaged Assets")), "=" + S("Q", st("NonExistent Assets")),
                f"=SUM(D{r}:G{r})", f"=-G{r}", f"=H{r}+I{r}",
                "=" + S("R"), "=-" + S("R", st("NonExistent Assets")), f"=K{r}+L{r}",
                "=" + S("V"), f"=J{r}-M{r}-N{r}", unrecorded.get(abbr, 0), f"=ROUND(O{r}-{S('W')},2)"]
        for c, v in enumerate(vals, 1):
            cell = tb.cell(r, c, v)
            cell.font = f_input if c == 16 else f_norm
            if c >= 4:
                cell.number_format = NUM
        if abbr in ("LIS",):
            tb.cell(r, 2).fill = fill_flag
        tb_rows[abbr] = r
        r += 1
    last = r - 1
    tb.cell(r, 2, "Total PP&E (incl. biological)").font = f_bold
    for c in range(4, 18):
        cell = tb.cell(r, c, f"=SUM({L(c)}{first}:{L(c)}{last})")
        cell.font, cell.number_format, cell.fill = f_bold, NUM, fill_tot
    tot = r
    r += 2
    src_cost = sum(a["cost"] for a in assets)
    src_ad = sum(a["end_ad"] for a in assets)
    tb.cell(r, 2, "Reconciliation to source detail").font = f_bold
    for lab, srcv, col in (("Gross cost per source", src_cost, "H"), ("Gross accum. depn per source", src_ad, "K")):
        r += 1
        tb.cell(r, 2, lab)
        c = tb.cell(r, 3, round(srcv, 2)); c.font, c.number_format = f_input, NUM
        c.comment = Comment(f"Hard-coded: sum of column J / AE of '{args.sheet}' asset rows in source workbook.", "reviewer")
        d = tb.cell(r, 4, f"=ROUND({col}{tot}-C{r},2)"); d.number_format = NUM
        tb.cell(r, 5, "← difference (should be 0)").font = f_sub
    r += 2
    ne_nbv = sum(a["nbv"] for a in assets if a["treatment"] == "Exclude")
    corr = sum(a["cost"] - a["end_ad"] - a["nbv"] for a in assets if a["treatment"] != "Exclude")
    tb.cell(r, 2, "Bridge: NBV per source register → OB NBV").font = f_bold
    bridge = [("NBV per source register (sum of NBV column)", sum(a["nbv"] for a in assets), True),
              ("Less: NBV of NonExistent assets excluded", -ne_nbv, True),
              ("Less: impairment of Damaged assets", f"=-N{tot}", False),
              ("Add: rows where source NBV ≠ Cost − AD (OB uses Cost − AD)", corr, True)]
    b0 = r + 1
    for lab, v, hard in bridge:
        r += 1
        tb.cell(r, 2, lab)
        c = tb.cell(r, 3, v); c.number_format = NUM
        if hard:
            c.font = f_input
            c.comment = Comment("Hard-coded by reviewer script from source asset rows.", "reviewer")
    r += 1
    tb.cell(r, 2, "OB NBV per bridge").font = f_bold
    c = tb.cell(r, 3, f"=SUM(C{b0}:C{r-1})"); c.number_format, c.font = NUM, f_bold
    d = tb.cell(r, 4, f"=ROUND(C{r}-O{tot},2)"); d.number_format = NUM
    tb.cell(r, 5, "← difference to OB NBV (should be 0)").font = f_sub
    tb_tot = tot

    # ---------------- OB Journal
    jn = wb.create_sheet("OB Journal")
    title(jn, f"Opening Journal Entry – PP&E", ob_label,
          "Dr cost / Cr accumulated depreciation & impairment; balancing figure to opening equity. Links to 'OB TB'.")
    header(jn, 5, ["Account", "Account Name", "Debit USD", "Credit USD", "Narrative"], [12, 50, 16, 16, 60])
    r = 6
    jfirst = r
    for row in COA:
        abbr, _, _, code, name, ad, _, _ = row
        if abbr in tb_rows:
            tr_ = tb_rows[abbr]
            jn.cell(r, 1, code); jn.cell(r, 2, name)
            jn.cell(r, 3, f"='OB TB'!J{tr_}").number_format = NUM
            jn.cell(r, 5, f"OB cost as at {ob_label}")
            r += 1
    for row in COA:
        abbr, _, _, code, name, ad, _, _ = row
        if abbr in tb_rows and ad:
            tr_ = tb_rows[abbr]
            jn.cell(r, 1, ad); jn.cell(r, 2, f"Accum. Depn – {name}")
            jn.cell(r, 4, f"='OB TB'!M{tr_}").number_format = NUM
            jn.cell(r, 5, f"OB accumulated depreciation as at {ob_label}")
            r += 1
    jn.cell(r, 1, "121999"); jn.cell(r, 2, "Accumulated Impairment – PP&E")
    jn.cell(r, 4, f"='OB TB'!N{tb_tot}").number_format = NUM
    jn.cell(r, 5, "Damaged assets written down to nil recoverable amount"); r += 1
    jn.cell(r, 1, "Equity"); jn.cell(r, 2, "Opening Balance Equity / Retained Earnings (code to assign)")
    jn.cell(r, 4, f"=SUM(C{jfirst}:C{r-1})-SUM(D{jfirst}:D{r-1})").number_format = NUM
    jn.cell(r, 5, "Balancing figure – PP&E net book value carried into opening equity"); r += 1
    jlast = r - 1
    jn.cell(r, 2, "Total").font = f_bold
    for c in (3, 4):
        cell = jn.cell(r, c, f"=SUM({L(c)}{jfirst}:{L(c)}{jlast})")
        cell.font, cell.number_format, cell.fill = f_bold, NUM, fill_tot
    r += 1
    jn.cell(r, 2, "Check: Debit − Credit (must be 0)")
    jn.cell(r, 3, f"=ROUND(C{r-1}-D{r-1},2)").number_format = NUM
    r += 2
    jn.cell(r, 2, "Pre-OB adjustments (memo – post in the closing year, before cut-over)").font = f_bold
    for lab, f in (("Derecognise NonExistent assets – NBV charged to retained earnings",
                    f"='OB TB'!G{tb_tot}+'OB TB'!L{tb_tot}"),
                   ("Impairment of Damaged assets", f"='OB TB'!N{tb_tot}"),
                   ("Unrecorded assets awaiting approved valuation (not in OB)", f"='OB TB'!P{tb_tot}")):
        r += 1
        jn.cell(r, 2, lab); jn.cell(r, 3, f).number_format = NUM

    # ---------------- Revised COA
    rc = wb.create_sheet("Revised COA", 1)
    title(rc, "Revised Chart of Accounts – Property, Plant & Equipment", ob_label,
          "Green rows = new or moved accounts. Every postable account carries OB as at the cut-over date.")
    header(rc, 5, ["Class", "Old Code", "Old Name", "New Code", "New Account Name", "Type", "AD Contra",
                   "Useful-life / Measurement Policy", "Change", "OB Flag", "OB NBV USD"],
           [8, 10, 32, 10, 46, 12, 10, 34, 52, 24, 16])
    r = 6
    for row in COA:
        abbr, oc, on, nc, nn, ad, pol, note = row
        typ = "Header" if abbr is None else ("Contra" if abbr == "IMPAIR" else ("Non-PP&E" if abbr == "LIS" else "Cost"))
        vals = [abbr, oc, on, nc, nn, typ, ad, pol, note,
                "" if typ == "Header" else f"OB as at {ob_label}",
                (f"='OB TB'!O{tb_rows[abbr]}" if abbr in tb_rows else (f"=-'OB TB'!N{tb_tot}" if abbr == "IMPAIR" else None))]
        for c, v in enumerate(vals, 1):
            cell = rc.cell(r, c, v)
            cell.font = f_bold if typ == "Header" else f_norm
            if c == 11:
                cell.number_format = NUM
        if oc is None or abbr == "LIS" or "NEW" in (note or ""):
            for c in range(1, 12):
                rc.cell(r, c).fill = fill_new
        r += 1
    r += 1
    rc.cell(r, 1, "Accumulated depreciation contra accounts (1219xx) mirror each cost account's last 3 digits.").font = f_sub

    # ---------------- Exceptions
    ex = wb.create_sheet("Exceptions")
    title(ex, "Asset-level Exceptions", ob_label, "One line per issue. Resolve or document before the OB is locked.")
    header(ex, 5, ["Register No.", "Source Row", "Status", "Class", "Tag", "Description", "Cost USD", "NBV USD", "Issue"],
           [10, 10, 18, 8, 16, 44, 14, 14, 60])
    r = 6
    ex_counter = Counter()
    for i, a in enumerate(assets):
        for e in a["exceptions"]:
            key = re.sub(r"\(diff.*|'.*'|\d{2}-\w{3}-\d{4}|[\d.]+ yrs", "", e).strip()
            ex_counter[key] += 1
            for c, v in enumerate([i + 1, a["src_row"], a["status"], a["coa"], a["tag"], a["desc"], a["cost"], a["nbv"], e], 1):
                cell = ex.cell(r, c, v); cell.font = f_norm
                if c in (7, 8):
                    cell.number_format = NUM
            r += 1
    ex.auto_filter.ref = f"A5:I{max(r - 1, 6)}"

    # ---------------- Findings
    fd = wb.create_sheet("Findings", 1)
    title(fd, "Review Findings – Draft Fixed-Asset COA & Register", ob_label)
    header(fd, 5, ["#", "Area", "Finding", "Evidence / Amount USD", "Recommendation", "Priority"], [5, 18, 60, 40, 60, 10])
    by_status = defaultdict(lambda: [0, 0.0, 0.0, 0.0])
    for a in assets:
        s = by_status[a["status"]]
        s[0] += 1; s[1] += a["cost"]; s[2] += a["end_ad"]; s[3] += a["nbv"]
    findings = []
    if args.extra_findings:
        findings += json.load(open(args.extra_findings, encoding="utf-8"))
    ne, dm, us = by_status["NonExistent Assets"], by_status["Damaged Assets"], by_status["Unsupported Assets"]
    findings += [
        ["Register – status", "Non-existent assets are still carried in the register.",
         f"{ne[0]} items; cost {ne[1]:,.2f}; NBV {ne[3]:,.2f}",
         "Derecognise before cut-over (Dr retained earnings / loss for NBV); excluded from OB in this workbook.", "High"],
        ["Register – status", "Unsupported assets (no source documents) form the largest group.",
         f"{us[0]} items; cost {us[1]:,.2f}; NBV {us[3]:,.2f}",
         "Recognise provisionally; obtain documents or independent valuation; disclose estimation uncertainty.", "High"],
        ["Register – status", "Damaged assets carried at cost with no impairment.",
         f"{dm[0]} items; NBV {dm[3]:,.2f}",
         "Impair to recoverable amount (nil assumed) via 121999, or dispose.", "Medium"],
    ]
    labels = {
        "Acquisition date missing": ("Data – dates", "Restore acquisition dates from invoices/JVs; needed for depreciation start.", "Medium"),
        "Acquisition date is year only": ("Data – dates", "Year-only dates (e.g. 2017) used; record a full date (use 1 Jul of year if unknown, documented).", "Low"),
        "Acquisition date  is on/after OB date": ("Data – dates", "An asset dated on/after the OB date cannot be in the opening balance; correct the date or remove.", "High"),
        "Useful life missing": ("Policy – useful life", "Assign useful life per revised policy; depreciation cannot be validated without it.", "High"),
        "Depreciation method missing": ("Policy – method", "Set method (SL) for every depreciable asset.", "Low"),
        "NBV ≠ Cost − AD": ("Data – arithmetic", "Fix formula rows where NBV does not equal cost less accumulated depreciation.", "High"),
        "Tag missing": ("Data – tagging", "Tag every recognised asset; instruction #11 requires sequential unique tags.", "Medium"),
        "Tag is placeholder text": ("Data – tagging", "Replace Khmer status notes in the tag column (e.g. 'not found', 'dismantled') with real tags or a Status.", "Medium"),
        "Duplicate tag": ("Data – tagging", "Make tags unique.", "Medium"),
        "Fully depreciated – review useful life (IAS 16.51)": ("Policy – useful life", "Zero-NBV assets still in use indicate lives are too short; review annually.", "Low"),
        "Building depreciated over  vs 20-yr policy": ("Policy – useful life", "Buildings depreciated over 5 yrs contradict the 20-yr policy; reassess prospectively (IAS 8 change in estimate).", "Medium"),
        "Land improvement over  – below 5-yr minimum": ("Policy – useful life", "Roads/drainage over 3 yrs is aggressive; align to 5–20 yrs capped at ELC term.", "Low"),
        "Large untagged line (≥100k) – tag, split or support": ("Data – lump sums", "Tag large items (e.g. road segments); balancing 'remaining not found' / lump-sum lines must be split to assets or written off.", "High"),
    }
    for key, n in ex_counter.most_common():
        area, rec, pr = labels.get(key, ("Data", "Review.", "Low"))
        findings.append([area, re.sub(r"\s+", " ", key).replace("date is on", "date on").replace("over vs", "over <10 yrs vs").replace("over –", "over <5 yrs –"),
                         f"{n} asset lines (see Exceptions sheet)", rec, pr])
    rec_n = sum(1 for a in assets if a["reclass_note"])
    findings.append(["COA – classification", "Items posted to the wrong class (water/power infrastructure and shelving in 'Other Fixed Assets'; appliances in 'Other Heavy Equipment').",
                     f"{rec_n} lines proposed for reclass (green in OB Register)", "Confirm and post reclasses to the revised accounts at OB.", "Medium"])
    order = {"High": 0, "Medium": 1, "Low": 2}
    findings.sort(key=lambda f: order.get(f[4], 3))
    for i, f in enumerate(findings, 1):
        for c, v in enumerate([i] + f, 1):
            cell = fd.cell(5 + i, c, v)
            cell.font, cell.alignment = f_norm, Alignment(wrap_text=True, vertical="top")
        if f[4] == "High":
            fd.cell(5 + i, 6).fill = fill_flag

    # ---------------- Read Me
    title(readme, "Fixed-Asset COA Review & Opening Balance Pack", ob_label)
    lines = [
        ("Prepared", dt.date.today().strftime("%d %b %Y")),
        ("Source workbook", args.input.split("/")[-1].split("-", 1)[-1]),
        ("Cut-off (last day of old year)", cut_label),
        ("Opening balance date", ob_label),
        ("", ""),
        ("Sheets", ""),
        ("Findings", "Prioritised review points on the draft COA and register"),
        ("Revised COA", "Old → new account mapping, AD contra accounts, policies, OB flag and OB NBV"),
        ("OB TB", "Opening balances per revised account (cost, AD, impairment, NBV) + reconciliation to source"),
        ("OB Journal", "Opening journal entry with balancing check, plus pre-OB adjustments memo"),
        ("OB Register", "Every asset marked with OB date, treatment and new account; OB figures are formulas"),
        ("Exceptions", "Asset-level data issues to clear before the OB is locked"),
        ("", ""),
        ("OB treatment", "Confirmed → recognise | Unsupported → recognise pending support | Damaged → recognise & impair to nil | NonExistent → exclude"),
        ("Legend", "Blue text = values copied from source; black = formulas; green fill = new/moved account or proposed reclass; yellow = flag"),
        ("Note", "Proposed reclasses and policies are review recommendations – confirm with management before posting."),
    ]
    for i, (k, v) in enumerate(lines, 4):
        readme.cell(i, 1, k).font = f_bold
        readme.cell(i, 2, v).font = f_norm
    readme.column_dimensions["A"].width = 32
    readme.column_dimensions["B"].width = 110

    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False
    wb.save(args.output)
    print(json.dumps(dict(assets=len(assets), exceptions=sum(ex_counter.values()),
                          source_cost=round(src_cost, 2), source_ad=round(src_ad, 2),
                          by_status={k: [v[0], round(v[1], 2), round(v[3], 2)] for k, v in by_status.items()},
                          reclass=rec_n, unrecorded=unrecorded), ensure_ascii=False, indent=1))


if __name__ == "__main__":
    p = argparse.ArgumentParser()
    p.add_argument("input"); p.add_argument("output")
    p.add_argument("--ob-date", default="2027-01-01")
    p.add_argument("--sheet", default="FA Details")
    p.add_argument("--summary", default="SUMMARY")
    p.add_argument("--extra-findings")
    build(p.parse_args())
