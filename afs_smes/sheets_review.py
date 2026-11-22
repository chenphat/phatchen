"""Review sheets: live Checks, analytical review (Analysis) and the disclosure Checklist."""
from openpyxl.formatting.rule import CellIsRule
from openpyxl.worksheet.datavalidation import DataValidation

from .sheets_data import FIRST, status_formatting, AMBER
from .taxonomy import FACES
from .xl import (B_BOTTOM, CENTER, CY, F_BOLD, F_GREY, F_HEAD, F_INPUT, F_NORMAL, F_SUB, NUM, NUM2, PCT, PY, WRAP, face,
                 header_row, lead, put, sheet_title, widths)


def build_checks(book, ws):
    sheet_title(ws, "Checks - live reconciliations", "Each difference must be within the tolerance on the Entity sheet. Statement-to-note checks are in the presentation currency.")
    widths(ws, [5, 100, 18, 10])
    t, m, g = book.tb_last, book.map_last, book.gl_last
    base = [
        ("TB CY: unadjusted trial balance - total debits = total credits (opening, movement, closing)",
         f"ABS(TB!J{t + 2})+ABS(TB!L{t + 2})+ABS(TB!N{t + 2})"),
        ("TB PY: trial balance - total debits = total credits (opening, movement, closing)",
         f"ABS(TB!D{t + 2})+ABS(TB!F{t + 2})+ABS(TB!H{t + 2})"),
        ("ATB CY: adjusted closing balances sum to zero (debits = credits)", f"SUM(ATB!O{FIRST}:O{t})"),
        ("ATB CY: opening balances sum to zero", f"SUM(ATB!J{FIRST}:J{t})"),
        ("ATB PY: closing balances sum to zero", f"SUM(ATB!I{FIRST}:I{t})"),
        ("ATB CY: opening + debits - credits = closing for every account", f"SUMPRODUCT(ABS(ATB!P{FIRST}:P{t}))"),
        ("ATB PY: opening + debits - credits = closing for every account", f"SUMPRODUCT(ABS(ATB!Q{FIRST}:Q{t}))"),
        ("GL to TB CY: GL net movement = TB movement for every account", f"SUMPRODUCT(ABS(ATB!S{FIRST}:S{t}))"),
        ("GL to TB PY: GL net movement = TB movement for every account", f"SUMPRODUCT(ABS(ATB!U{FIRST}:U{t}))"),
        ("GL: total debits = total credits (both years)", f"SUM(GL!H{FIRST}:H{g})-SUM(GL!I{FIRST}:I{g})"),
        ("GL: every line belongs to a mapped account (count of unmapped lines)", f'COUNTIF(GL!K{FIRST}:K{g},"UNMAPPED")'),
        ("Mapping: every account has a valid AFS code (count of unmapped accounts)", f'COUNTIF(Mapping!L{FIRST}:L{m},"UNMAPPED")'),
        ("AJE: adjusting journals balance", "AJE!E3-AJE!F3"),
        ("AJE: every adjustment uses an account in the trial balance (count of unknown accounts)",
         f'COUNTIF(AJE!C{FIRST}:C{FIRST + 59},"ACCOUNT NOT IN TB")'),
        ("Lead: all mapped balances are carried to the lead schedule (CY)", f"SUM(Lead!F{FIRST}:F{book.lead_last})-SUM(ATB!O{FIRST}:O{t})"),
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
    sheet_title(ws, "Comparative analysis - two years",
                "Ratios (functional currency) and year-on-year movements with common-size percentages (presentation currency). "
                "Explain flagged movements before finalising.")
    widths(ws, [62, 17, 17, 17, 11, 10, 11, 11])
    E = book.ent
    c = lambda sheet, key, fy: book.cell(sheet, key, fy)
    put(ws, "A4", "Key ratios", F_SUB)
    header_row(ws, 5, ["Ratio", "CY", "PY", "Change"], height=20)
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
    book.ratio_rows = []
    for cap, fn, fmt in ratios:
        r += 1
        put(ws, f"A{r}", cap)
        put(ws, f"B{r}", "=" + fn(CY), fmt=fmt)
        put(ws, f"C{r}", "=" + fn(PY), fmt=fmt)
        put(ws, f"D{r}", f"=B{r}-C{r}", fmt=fmt)
        book.ratio_rows.append((r, fmt))
    r += 2
    put(ws, f"A{r}", f'="Year-on-year movement by line item - "&{E["pres_label"]}', F_SUB)
    r += 1
    header_row(ws, r, ["Line item", "CY", "PY", "Change", "% change", "Flag", "Common size CY", "Common size PY"], height=30)
    put(ws, f"B{r}", book.year(CY), F_HEAD, "0")
    put(ws, f"C{r}", book.year(PY), F_HEAD, "0")
    start = r + 1
    P = lambda sheet, key, fy: book.cell(sheet, key, fy, "P")
    lines = [(code, caption, stmt, sec) for code, caption, stmt, sec in FACES if code in book.active_faces or code == "RE"]
    extra = {"CA": ("TA", "Total assets", "SoFP"), "EQ": ("T_EQ", "Total equity", "SoFP"), "CL": ("TL", "Total liabilities", "SoFP"),
             "PL": ("PROFIT", "Profit for the year", "SoCI")}
    out, prev = [], None
    for code, caption, stmt, sec in lines:
        if prev and sec != prev and prev in extra:
            out.append(extra[prev] + (True,))
        out.append((code, f"{caption} ({stmt}{'' if stmt == 'SoCI' else ', ' + sec})", stmt, False))
        if code == "RESV":
            out.append(("TR", "Foreign currency translation reserve (SoFP, EQ)", "SoFP", False))
        prev = sec
    out.append(extra["PL"] + (True,))
    for key, caption, stmt, bold in out:
        r += 1
        font = F_BOLD if bold else F_NORMAL
        base = "TA" if stmt == "SoFP" else "REV"
        put(ws, f"A{r}", caption, font)
        put(ws, f"B{r}", "=" + P(stmt, key, CY), font, NUM)
        put(ws, f"C{r}", "=" + P(stmt, key, PY), font, NUM)
        put(ws, f"D{r}", f"=B{r}-C{r}", font, NUM)
        put(ws, f"E{r}", f"=IF(C{r}=0,0,D{r}/ABS(C{r}))", font, PCT)
        put(ws, f"F{r}", f'=IF(AND(ABS(D{r})>={book.mat["perf"]},OR(C{r}=0,ABS(E{r})>={E["threshold"]})),"REVIEW","")', F_BOLD, align=CENTER)
        put(ws, f"G{r}", f"=IF({P(stmt, base, CY)}=0,0,B{r}/{P(stmt, base, CY)})", font, PCT)
        put(ws, f"H{r}", f"=IF({P(stmt, base, PY)}=0,0,C{r}/{P(stmt, base, PY)})", font, PCT)
    ws.conditional_formatting.add(f"F{start}:F{r}", CellIsRule(operator="equal", formula=['"REVIEW"'], fill=AMBER))
    put(ws, f"A{r + 2}", "Flag: change of at least performance materiality (Materiality sheet) and at least the % threshold on the Entity "
                         "sheet. Common size: % of total assets (SoFP lines) or of revenue (SoCI lines).", F_GREY)
    ws.freeze_panes = "B6"


CHECKLIST = [
    # (paragraphs, requirement, where addressed)
    ("3.3", "Explicit and unreserved statement of compliance with the IFRS for SMEs Accounting Standard", "basis"),
    ("3.8-3.9", "Going concern assessment; material uncertainties, or the basis used if not a going concern", "basis"),
    ("3.14, 3.20", "Comparative information for all amounts, and for narrative where relevant", "All statements"),
    ("3.17", "Complete set: SoFP, SoCI (one or two statements), SoCE, SoCF and notes", "SoFP, SoCI, SoCE, SoCF"),
    ("3.23", "Each statement identified: entity name, individual or group, reporting date and period, presentation currency, rounding", "All statements"),
    ("3.24", "Domicile, legal form, country of incorporation, registered office; nature of operations and principal activities", "general"),
    ("4.2-4.3", "Minimum line items on the face of the SoFP; additional line items, headings and subtotals where relevant", "SoFP"),
    ("4.4-4.8", "Current and non-current assets and liabilities presented as separate classifications", "SoFP"),
    ("4.11", "Sub-classifications: PPE classes; receivables (related parties / others); inventories; payables (trade, related parties, deferred income, accruals); provisions; classes of equity", "ppe, tor, inv, top, prov, equity"),
    ("4.12", "Share capital: shares authorised, issued and fully paid, par value, reconciliation of shares outstanding, rights and restrictions; description of each reserve", "equity"),
    ("4.14", "Binding sale agreement for a major disposal of assets", "N/A unless applicable"),
    ("5.2, 5.5", "Total comprehensive income in one or two statements with the minimum line items; OCI items by nature, grouped by whether reclassified", "SoCI"),
    ("5.11", "Analysis of expenses by nature or by function (cost of sales shown separately under the function method)", "SoCI, expenses"),
    ("6.3", "SoCE: total comprehensive income; retrospective adjustments; reconciliation of each component of equity (profit, OCI, share issues, dividends)", "SoCE"),
    ("6.6", "Dividends proposed or declared before authorisation but not recognised, and amount per share", "dividends"),
    ("7.3, 7.7", "Cash flows classified as operating, investing and financing; operating cash flows by the indirect or direct method", "SoCF"),
    ("7.14-7.17", "Interest and dividends received and paid, and income tax cash flows, presented separately and classified consistently", "SoCF"),
    ("7.18", "Non-cash investing and financing transactions excluded from the SoCF and disclosed", "cash"),
    ("7.19A", "Reconciliation of opening and closing liabilities arising from financing activities", "borr"),
    ("7.19C", "Supplier finance arrangements: terms, carrying amounts, amounts already paid to suppliers, range of due dates", "top"),
    ("7.20-7.21", "Components of cash and cash equivalents reconciled to the SoFP; balances not available for use", "cash"),
    ("8.2-8.4", "Notes presented systematically, cross-referenced to the statements, in the normal order", "Notes"),
    ("8.5", "Material accounting policy information", "policies"),
    ("8.6-8.7", "Judgements in applying accounting policies; key sources of estimation uncertainty with the carrying amounts affected", "judgements"),
    ("10.13-10.14, 10.18, 10.23", "Changes in accounting policies, changes in accounting estimates and corrections of prior period errors", "basis, SoCE"),
    ("A1-A3", "Early application of the third edition disclosed; transition applied retrospectively with the permitted reliefs", "basis"),
    ("11.40-11.41", "Financial instruments: measurement bases; carrying amounts by category", "policies, fininst"),
    ("11.42, 11.46-11.47", "Terms and conditions of long-term debt; financial assets pledged as collateral; defaults and breaches on loans payable", "borr"),
    ("11.43, 11.43B", "Analysis of the age of trade receivables by reference to due date: gross amount and impairment", "tor"),
    ("11.43A", "Maturity analysis of financial liabilities: remaining contractual undiscounted cash flows", "fininst"),
    ("11.48", "Income, expense, gains and losses by category; total interest income and expense; impairment loss by class of financial asset", "finance, tor"),
    ("12.28-12.31", "Fair value measurement: carrying amounts, level in the fair value hierarchy, valuation techniques and inputs", "ppe, fininst"),
    ("13.22", "Inventories: policy and cost formula, carrying amounts by class, amount expensed, impairment losses, pledges", "inv, policies"),
    ("14.12, 15.19", "Investments in associates and jointly controlled entities: policy, carrying amounts, fair value where quoted", "invassoc"),
    ("16.10-16.11", "Investment property at fair value: valuer involvement, restrictions, obligations, reconciliation", "invprop"),
    ("17.31", "PPE by class: measurement bases, depreciation methods and rates, gross amount and accumulated depreciation, reconciliation of the carrying amount", "ppe, policies"),
    ("17.32-17.33", "PPE pledged as security; commitments to acquire PPE; revalued assets: date, valuer, cost-model carrying amount, revaluation surplus", "ppe, commitments, equity"),
    ("18.27-18.29", "Intangible assets: useful lives, amortisation method, gross amount, line item containing amortisation, reconciliation; R&D expensed", "intang, pbt"),
    ("20.13, 20.16", "Lessee: future minimum lease payments by time band, lease expense, description of significant leasing arrangements", "commitments"),
    ("21.14", "Provisions by class: reconciliation, nature, expected timing, uncertainties, expected reimbursements", "prov"),
    ("21.15-21.16", "Contingent liabilities and contingent assets", "contingencies"),
    ("23.82", "Revenue disaggregated into categories relevant to the entity's financial performance", "revenue"),
    ("23.83-23.84", "Opening and closing trade receivables, contract assets and contract liabilities; revenue recognised from opening contract liabilities", "revenue, tor"),
    ("23.85-23.87", "Promises in contracts with customers: when fulfilled, payment terms, nature of goods or services, returns and warranties; progress measurement; judgements", "revenue"),
    ("24.6", "Government grants: nature and amounts, unfulfilled conditions, other government assistance", "N/A unless grants received"),
    ("25.3", "Finance costs and total interest expense (no further disclosure required)", "finance"),
    ("27.32-27.33", "Impairment losses and reversals by class of asset, and the line items in which they are included", "inv, ppe, intang"),
    ("28.40-28.43", "Employee benefits: defined contribution expense; defined benefit plans; other long-term and termination benefits", "pbt, ebo"),
    ("29.38-29.40", "Income tax: major components of tax expense; tax on OCI items; reconciliation to accounting profit at the applicable rate; deferred tax by type of temporary difference", "tax, dtax"),
    ("30.25-30.26", "Exchange differences recognised in profit or loss; presentation and functional currency", "pbt, basis"),
    ("32.9-32.10", "Date of authorisation for issue and who authorised; non-adjusting events after the reporting period", "events, approval"),
    ("33.5, 33.7", "Name of the parent and ultimate controlling party; key management personnel compensation in total", "related"),
    ("33.9-33.10", "Related party transactions by category: amounts, outstanding balances, terms, allowances and expense for uncollectable balances", "related"),
    ("35.12", "First-time adoption: explanation of the transition and reconciliations", "N/A unless first-time adopter"),
]


def build_checklist(book, ws):
    sheet_title(ws, "Disclosure checklist - IFRS for SMEs Accounting Standard (third edition)",
                "Paragraph references are to the third edition (February 2025). Requirements are summarised: read the paragraph itself before marking a row Complete.")
    widths(ws, [16, 96, 30, 12, 40])
    header_row(ws, 5, ["Paragraph", "Requirement (summary)", "Where addressed", "Status", "Reviewer comment"], height=20)
    r = 5
    for sec, req, where in CHECKLIST:
        r += 1
        keys = [k.strip() for k in where.split(",")]
        if where.startswith("N/A"):
            where_txt, status = where, "N/A"
        else:
            stmts = [k for k in keys if k[0].isupper()]
            nums = sorted(book.notes[k] for k in keys if k in book.notes)
            parts = stmts + (["Note " + ", ".join(map(str, nums))] if nums else [])
            where_txt, status = ("; ".join(parts), "Drafted") if parts else ("Not applicable - no balances", "N/A")
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
