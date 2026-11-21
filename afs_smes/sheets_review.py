"""Review sheets: live Checks, analytical review (Analysis) and the disclosure Checklist."""
from openpyxl.formatting.rule import CellIsRule
from openpyxl.worksheet.datavalidation import DataValidation

from .sheets_data import FIRST, status_formatting, AMBER
from .taxonomy import FACES
from .xl import (B_BOTTOM, CENTER, CY, F_BOLD, F_GREY, F_HEAD, F_INPUT, F_NORMAL, F_SUB, NUM, NUM2, PCT, PY, WRAP, face,
                 header_row, lead, put, sheet_title, widths)


def build_checks(book, ws):
    sheet_title(ws, "Checks - live reconciliations", "Each difference must be within the tolerance set on the Entity sheet. Recalculates with the workbook.")
    widths(ws, [5, 100, 18, 10])
    t, m, g = book.tb_last, book.map_last, book.gl_last
    base = [
        ("TB CY: final closing balances sum to zero (debits = credits)", f"SUM(TB!O{FIRST}:O{t})"),
        ("TB CY: opening balances sum to zero", f"SUM(TB!J{FIRST}:J{t})"),
        ("TB PY: closing balances sum to zero", f"SUM(TB!I{FIRST}:I{t})"),
        ("TB CY: opening + debits - credits = closing for every account", f"SUMPRODUCT(ABS(TB!P{FIRST}:P{t}))"),
        ("TB PY: opening + debits - credits = closing for every account", f"SUMPRODUCT(ABS(TB!Q{FIRST}:Q{t}))"),
        ("GL to TB CY: GL net movement = TB movement for every account", f"SUMPRODUCT(ABS(TB!S{FIRST}:S{t}))"),
        ("GL to TB PY: GL net movement = TB movement for every account", f"SUMPRODUCT(ABS(TB!U{FIRST}:U{t}))"),
        ("GL: total debits = total credits (both years)", f"SUM(GL!H{FIRST}:H{g})-SUM(GL!I{FIRST}:I{g})"),
        ("GL: every line belongs to a mapped account (count of unmapped lines)", f'COUNTIF(GL!K{FIRST}:K{g},"UNMAPPED")'),
        ("Mapping: every account has a valid AFS code (count of unmapped accounts)", f'COUNTIF(Mapping!L{FIRST}:L{m},"UNMAPPED")'),
        ("AJE: adjusting journals balance", "AJE!E3-AJE!F3"),
        ("AJE: every adjustment uses an account in the trial balance (count of unknown accounts)",
         f'COUNTIF(AJE!C{FIRST}:C{FIRST + 59},"ACCOUNT NOT IN TB")'),
        ("Lead: all mapped balances are carried to the lead schedule (CY)", f"SUM(Lead!F{FIRST}:F{book.lead_last})-SUM(TB!O{FIRST}:O{t})"),
    ]
    put(ws, "B3", "Overall result", F_BOLD)
    header_row(ws, 5, ["#", "Check", "Difference / count", "Status"], height=20)
    r = 5
    for i, (desc, f) in enumerate(base + book.checks, start=1):
        r += 1
        put(ws, f"A{r}", i, align=CENTER)
        put(ws, f"B{r}", desc, align=WRAP)
        put(ws, f"C{r}", "=" + f, fmt=NUM2)
        put(ws, f"D{r}", f'=IF(ABS(C{r})<={book.ent["tolerance"]},"PASS","FAIL")', F_BOLD, align=CENTER)
    put(ws, "C3", f'=IF(COUNTIF(D6:D{r},"FAIL")=0,"ALL "&COUNTA(D6:D{r})&" CHECKS PASS",COUNTIF(D6:D{r},"FAIL")&" CHECK(S) FAILED")', F_BOLD)
    status_formatting(ws, f"D6:D{r}")
    r += 2
    put(ws, f"B{r}", f'="Mapping rows awaiting reviewer confirmation (status REVIEW): "&COUNTIF(Mapping!L{FIRST}:L{m},"REVIEW")', F_GREY)
    put(ws, f"B{r + 1}", "Source-data checks performed before the workbook was built are on the DQ_Log sheet.", F_GREY)
    ws.freeze_panes = "A6"
    book.checks_last = r - 2


def build_analysis(book, ws):
    sheet_title(ws, "Analytical review", "Ratios and year-on-year movements, calculated from the statements. Explain flagged movements before finalising.")
    widths(ws, [62, 16, 16, 16, 12, 10])
    E = book.ent
    c = lambda sheet, key, fy: book.cell(sheet, key, fy)
    put(ws, "A4", "Key ratios", F_SUB)
    header_row(ws, 5, ["Ratio", "CY", "PY", "Change", "", ""], height=20)
    put(ws, "B5", book.year(CY), F_HEAD, "0")
    put(ws, "C5", book.year(PY), F_HEAD, "0")
    inv = lambda fy: c("SoFP", "INV", fy)
    ratios = [
        ("Gross profit margin", lambda fy: f"IF({c('SoCI','REV',fy)}=0,0,{c('SoCI','GP',fy)}/{c('SoCI','REV',fy)})", PCT),
        ("Operating profit margin", lambda fy: f"IF({c('SoCI','REV',fy)}=0,0,{c('SoCI','OP',fy)}/{c('SoCI','REV',fy)})", PCT),
        ("Net profit margin", lambda fy: f"IF({c('SoCI','REV',fy)}=0,0,{c('SoCI','PROFIT',fy)}/{c('SoCI','REV',fy)})", PCT),
        ("Return on equity (profit / closing equity)", lambda fy: f"IF({c('SoFP','T_EQ',fy)}=0,0,{c('SoCI','PROFIT',fy)}/{c('SoFP','T_EQ',fy)})", PCT),
        ("Effective tax rate", lambda fy: f"IF({c('SoCI','PBT',fy)}=0,0,-{c('SoCI','TAX',fy)}/{c('SoCI','PBT',fy)})", PCT),
        ("Current ratio (times)", lambda fy: f"IF({c('SoFP','T_CL',fy)}=0,0,{c('SoFP','T_CA',fy)}/{c('SoFP','T_CL',fy)})", "0.00"),
        ("Quick ratio (times)", lambda fy: f"IF({c('SoFP','T_CL',fy)}=0,0,({c('SoFP','T_CA',fy)}-{inv(fy)})/{c('SoFP','T_CL',fy)})", "0.00"),
        ("Receivable days (trade receivables, net / revenue x 365)",
         lambda fy: f"IF({c('SoCI','REV',fy)}=0,0,{lead('close', fy, ['TOR.TRADE', 'TOR.ALLOW'])}/{c('SoCI','REV',fy)}*365)", "0"),
        ("Inventory days (inventories / cost of sales x 365)", lambda fy: f"IF({c('SoCI','COS',fy)}=0,0,-{inv(fy)}/{c('SoCI','COS',fy)}*365)", "0"),
        ("Payable days (trade payables / cost of sales x 365)",
         lambda fy: f"IF({c('SoCI','COS',fy)}=0,0,{lead('close', fy, ['TOP.TRADE'])}/{c('SoCI','COS',fy)}*365)", "0"),
        ("Debt to equity (borrowings / equity)",
         lambda fy: f"IF({c('SoFP','T_EQ',fy)}=0,0,({c('SoFP','BORR_NC',fy)}+{c('SoFP','BORR_C',fy)})/{c('SoFP','T_EQ',fy)})", "0.00"),
        ("Interest cover (operating profit / finance costs, times)",
         lambda fy: f"IF({c('SoCI','FINCOST',fy)}=0,0,-{c('SoCI','OP',fy)}/{c('SoCI','FINCOST',fy)})", "0.0"),
        ("Cash conversion (net cash from operations / profit for the year)",
         lambda fy: f"IF({c('SoCI','PROFIT',fy)}=0,0,{c('SoCF','OPER',fy)}/{c('SoCI','PROFIT',fy)})", "0.00"),
    ]
    r = 5
    for cap, fn, fmt in ratios:
        r += 1
        put(ws, f"A{r}", cap)
        put(ws, f"B{r}", "=" + fn(CY), fmt=fmt)
        put(ws, f"C{r}", "=" + fn(PY), fmt=fmt)
        put(ws, f"D{r}", f"=B{r}-C{r}", fmt=fmt)
    r += 2
    put(ws, f"A{r}", "Year-on-year movement by line item", F_SUB)
    r += 1
    header_row(ws, r, ["Line item", "CY", "PY", "Change", "% change", "Flag"], height=20)
    start = r + 1
    for code, caption, stmt, sec in FACES:
        if code not in book.active_faces:
            continue
        r += 1
        sign = "" if sec in ("NCA", "CA") else "-"
        put(ws, f"A{r}", f"{caption} ({stmt}{'' if stmt == 'SoCI' else ', ' + sec})")
        if code == "RE":
            put(ws, f"B{r}", "=" + c("SoFP", "RE", CY), fmt=NUM)
            put(ws, f"C{r}", "=" + c("SoFP", "RE", PY), fmt=NUM)
        else:
            put(ws, f"B{r}", "=" + sign + face("close", CY, code), fmt=NUM)
            put(ws, f"C{r}", "=" + sign + face("close", PY, code), fmt=NUM)
        put(ws, f"D{r}", f"=B{r}-C{r}", fmt=NUM)
        put(ws, f"E{r}", f"=IF(C{r}=0,0,D{r}/ABS(C{r}))", fmt=PCT)
        put(ws, f"F{r}", f'=IF(AND(ABS(D{r})>={E["materiality"]},OR(C{r}=0,ABS(E{r})>={E["threshold"]})),"REVIEW","")', F_BOLD, align=CENTER)
    ws.conditional_formatting.add(f"F{start}:F{r}", CellIsRule(operator="equal", formula=['"REVIEW"'], fill=AMBER))
    ws.freeze_panes = "B6"


CHECKLIST = [
    ("3", "Statement of explicit and unreserved compliance with the IFRS for SMEs Accounting Standard", "basis"),
    ("3", "Going concern assessment; material uncertainties disclosed where they exist", "basis"),
    ("3", "Complete set of statements, each clearly identified: entity name, period, presentation currency, level of rounding", "SoFP, SoCI, SoCE, SoCF"),
    ("3", "Comparative information for all amounts, and for narrative where relevant", "All statements"),
    ("3", "Domicile, legal form, registered office, nature of operations and principal activities", "general"),
    ("4", "Minimum line items on the face of the SoFP; current/non-current distinction", "SoFP"),
    ("4", "Sub-classifications of line items (PPE classes, receivables, inventories, payables, provisions, equity)", "ppe, tor, inv, top, prov, equity"),
    ("4", "Share capital: shares issued, par value, reconciliation of the number of shares, rights and restrictions; description of each reserve", "equity"),
    ("5", "Total comprehensive income presented in one (or two) statement(s) with the minimum line items; items of OCI by nature", "SoCI"),
    ("5", "Analysis of expenses by nature or by function", "expenses"),
    ("6", "Reconciliation of each component of equity: profit or loss, OCI, transactions with owners; effects of retrospective adjustments", "SoCE"),
    ("7", "Cash flows classified as operating, investing and financing; interest, dividends and income tax shown separately", "SoCF"),
    ("7", "Components of cash and cash equivalents reconciled to the SoFP; restricted balances; non-cash transactions", "cash"),
    ("7", "Reconciliation of changes in liabilities arising from financing activities", "borr"),
    ("7", "Supplier finance arrangements", "top"),
    ("8", "Notes presented systematically and cross-referenced; basis of preparation; material accounting policy information", "basis, policies"),
    ("8", "Judgements in applying accounting policies and key sources of estimation uncertainty", "judgements"),
    ("10", "Changes in accounting policies, changes in estimates and corrections of prior period errors", "basis, SoCE"),
    ("11", "Financial instruments: carrying amounts by category; terms of borrowings; collateral; defaults; income, expense, gains and losses", "fininst, borr, finance, tor"),
    ("12", "Fair value measurement: techniques and inputs for items measured at fair value", "fininst, ppe"),
    ("13", "Inventories: policy and cost formula, carrying amounts by class, amount expensed, write-downs, pledges", "inv, policies"),
    ("14-15", "Investments in associates and jointly controlled entities: policy, carrying amounts, share of results", "invassoc"),
    ("16", "Investment property: measurement basis, valuation methods, reconciliation", "invprop"),
    ("17", "PPE: measurement bases, depreciation methods and rates, reconciliation of carrying amounts, pledges, commitments, revaluation disclosures", "ppe, policies"),
    ("18", "Intangible assets: useful lives, amortisation method, reconciliation, line item containing amortisation, R&D expensed", "intang, pbt"),
    ("20", "Leases: future minimum lease payments by maturity band, lease expense, description of significant arrangements", "commitments"),
    ("21", "Provisions: reconciliation, nature, timing and uncertainties; contingent liabilities and contingent assets", "prov, contingencies"),
    ("22", "Equity and liabilities: classification of instruments issued; distributions to owners", "equity, dividends"),
    ("23", "Revenue: disaggregation, contract balances, promises (performance obligations), significant judgements", "revenue"),
    ("24", "Government grants: nature and amounts recognised, unfulfilled conditions", "N/A unless grants received"),
    ("25", "Borrowing costs: finance costs recognised (all expensed)", "finance"),
    ("27", "Impairment of assets: losses and reversals by class of asset and line item", "tor, inv, ppe"),
    ("28", "Employee benefits: short-term benefits, post-employment plans, other long-term and termination benefits", "pbt, ebo"),
    ("29", "Income tax: components of tax expense, tax on OCI items, reconciliation to accounting profit, deferred tax balances and movements", "tax, dtax"),
    ("30", "Foreign currency: exchange differences in profit or loss, functional and presentation currency", "pbt, basis"),
    ("32", "Events after the reporting period: date of authorisation and who authorised; non-adjusting events", "events, approval"),
    ("33", "Related parties: parent and controlling party, key management personnel compensation, transactions and balances", "related"),
    ("35 / App. A", "Transition to the third edition (or first-time adoption): explanation and reconciliations", "basis"),
]


def build_checklist(book, ws):
    sheet_title(ws, "Disclosure checklist - IFRS for SMEs Accounting Standard (third edition)",
                "Summary by Section. Complete the status column; confirm detailed paragraph requirements against the Standard itself.")
    widths(ws, [11, 96, 30, 12, 40])
    header_row(ws, 5, ["Section", "Requirement (summary)", "Where addressed", "Status", "Reviewer comment"], height=20)
    r = 5
    for sec, req, where in CHECKLIST:
        r += 1
        keys = [k.strip() for k in where.split(",")]
        if all(k in book.notes for k in keys):
            where_txt, status = "Note " + ", ".join(str(book.notes[k]) for k in keys), "Drafted"
        elif any(k in book.notes for k in keys):
            where_txt, status = "Note " + ", ".join(str(book.notes[k]) for k in keys if k in book.notes), "Drafted"
        elif where[0].isupper() and not where.startswith("N/A"):
            where_txt, status = where, "Drafted"
        else:
            where_txt, status = "Not applicable - no balances" if not where.startswith("N/A") else where, "N/A"
        put(ws, f"A{r}", sec, align=CENTER)
        put(ws, f"B{r}", req, align=WRAP)
        put(ws, f"C{r}", where_txt, align=WRAP)
        put(ws, f"D{r}", status, F_INPUT, align=CENTER)
        ws[f"E{r}"].font = F_INPUT
        ws.row_dimensions[r].height = 15 * (1 + len(req) // 100)
    dv = DataValidation(type="list", formula1='"Drafted,Complete,N/A,Outstanding"', allow_blank=True)
    dv.add(f"D6:D{r}")
    ws.add_data_validation(dv)
    put(ws, f"B{r + 2}", '"Drafted" means the note exists in this workbook with figures linked; narrative in yellow still has to be completed.', F_GREY)
    ws.freeze_panes = "A6"
