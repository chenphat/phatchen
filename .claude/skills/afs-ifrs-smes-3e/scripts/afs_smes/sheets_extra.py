"""Governance and working papers: Materiality, Policies, Findings, Directors' report and the Dashboard."""
from openpyxl.chart import BarChart, DoughnutChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.formatting.rule import CellIsRule, DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

from .sheets_data import AMBER, FIRST, GREEN, RED, status_formatting
from .sheets_notes import POLICIES
from .taxonomy import EXPENSE_FUNCTIONS, FACES, NATURES
from .xl import (B_BOTTOM, B_TOP, B_TOTAL, CENTER, CY, DATE, FONT, F_BOLD, F_GREY, F_INPUT, F_LINK, F_NORMAL, F_SUB,
                 F_TITLE, FILL_INPUT, NUM, PCT, PY, RIGHT, WRAP, glm, header_row, lead, nat, put, sheet_title, widths)

# validated categorical palette (fixed order) and reserved status colours
SERIES = ["2A78D6", "EB6834", "1BAF7A", "EDA100", "E87BA4", "008300"]
STATUS = {"High": "D03B3B", "Medium": "EC835A", "Low": "FAB219", "Info": "898781"}
NAVY = "14284B"
INK, MUTED, GRID = "0B0B0B", "898781", "E1E0D9"


# ----------------------------------------------------------------------------------------------- materiality
def build_materiality(book, ws):
    E = book.ent
    P = lambda sheet, key, fy=CY: book.cell(sheet, key, fy, "P")
    sheet_title(ws, "Materiality assessment - presentation and disclosure",
                "Information is material if omitting, misstating or obscuring it could reasonably be expected to influence the decisions "
                "of primary users (paragraph 3.16). Materiality depends on nature or magnitude, or both. Blue cells are judgements.")
    widths(ws, [4, 52, 20, 14, 20, 20, 22, 46])
    put(ws, "B4", "A. Quantitative thresholds", F_SUB)
    header_row(ws, 5, ["", "Benchmark", "Benchmark amount", "% applied", "Materiality", "Prior year amount", "", "Why this benchmark might be used"],
               height=30)
    put(ws, "C4", f'="Amounts in "&{E["pres_label"]}', F_GREY)
    bench = [
        ("Profit before tax", "SoCI", "PBT", 0.05, "Profit-oriented entity with stable earnings; users focus on profit."),
        ("Revenue", "SoCI", "REV", 0.01, "Earnings are volatile or near break-even; activity level is the focus."),
        ("Total assets", "SoFP", "TA", 0.01, "Asset-based or capital-intensive entity; lenders focus on the asset base."),
        ("Total equity", "SoFP", "T_EQ", 0.02, "Users focus on net assets, solvency or distributable reserves."),
    ]
    for i, (name, sheet, key, pct, why) in enumerate(bench):
        r = 6 + i
        put(ws, f"B{r}", name)
        put(ws, f"C{r}", f"=ABS({P(sheet, key)})", F_LINK, NUM)
        put(ws, f"D{r}", pct, F_INPUT, "0.0%")
        put(ws, f"E{r}", f"=C{r}*D{r}", fmt=NUM)
        put(ws, f"F{r}", f"=ABS({P(sheet, key, PY)})", F_LINK, NUM)
        put(ws, f"H{r}", why, F_GREY, align=WRAP)
        ws.row_dimensions[r].height = 28
    put(ws, "B11", "Benchmark selected", F_BOLD)
    put(ws, "C11", "Profit before tax", F_INPUT, fill=FILL_INPUT)
    dv = DataValidation(type="list", formula1="$B$6:$B$9", allow_blank=False)
    dv.add("C11")
    ws.add_data_validation(dv)
    put(ws, "H11", "[Record why this benchmark best reflects what the primary users of these financial statements focus on.]", F_INPUT,
        align=WRAP, fill=FILL_INPUT)
    ws.row_dimensions[11].height = 28
    rows = [
        (12, "Overall materiality", None, "=INDEX($E$6:$E$9,MATCH($C$11,$B$6:$B$9,0))",
         "Amounts above this are presumed material by size."),
        (13, "Performance materiality", 0.75, "=E12*D13",
         "Working threshold for individual line items and policies, lower than overall materiality to allow for aggregation."),
        (14, "Clearly trivial threshold", 0.05, "=E12*D14", "Differences below this are not accumulated or reported."),
    ]
    for r, name, pct, f, why in rows:
        put(ws, f"B{r}", name, F_BOLD)
        if pct is not None:
            put(ws, f"D{r}", pct, F_INPUT, "0.0%")
        put(ws, f"E{r}", f, F_BOLD, NUM, border=B_TOP if r == 12 else None)
        put(ws, f"H{r}", why, F_GREY, align=WRAP)
        ws.row_dimensions[r].height = 28
    put(ws, "F12", f'="= "&TEXT(E12*{E["pres_scale"]}/{E["fx_close_cy"]},"#,##0")&" "&{E["currency"]}', F_GREY)
    book.mat = {"overall": "Materiality!$E$12", "perf": "Materiality!$E$13", "trivial": "Materiality!$E$14"}

    put(ws, "B16", "B. Assessment by line item", F_SUB)
    put(ws, "B17", "A line is material by size at or above performance materiality. Nature can make a smaller item material: record it in the "
                   "blue column (for example related-party items, covenants, regulatory limits, items that turn a profit into a loss).", F_GREY)
    header_row(ws, 18, ["", "Line item", "Current year", "% of overall materiality", "Material by size?",
                        "Material by nature? (Yes / No)", "Conclusion", "Basis for the judgement on nature"], height=42)
    r = 18
    first = r + 1
    for code, caption, stmt, sec in FACES:
        if code not in book.active_faces and code != "RE":
            continue
        r += 1
        put(ws, f"B{r}", f"{caption} ({stmt})")
        put(ws, f"C{r}", f"=ABS({P(stmt, code)})", F_LINK, NUM)
        put(ws, f"D{r}", f"=IF({book.mat['overall']}=0,0,C{r}/{book.mat['overall']})", fmt="0%")
        put(ws, f"E{r}", f'=IF(C{r}>={book.mat["perf"]},"Yes","No")', align=CENTER)
        put(ws, f"F{r}", "No", F_INPUT, align=CENTER)
        put(ws, f"G{r}", f'=IF(OR(E{r}="Yes",F{r}="Yes"),"Material - present and disclose",'
                         f'IF(C{r}<{book.mat["trivial"]},"Clearly trivial","Not material - may be aggregated"))')
        ws[f"H{r}"].font = F_INPUT
    dv2 = DataValidation(type="list", formula1='"Yes,No"', allow_blank=False)
    dv2.add(f"F{first}:F{r}")
    ws.add_data_validation(dv2)
    ws.conditional_formatting.add(f"E{first}:F{r}", CellIsRule(operator="equal", formula=['"Yes"'], fill=AMBER))
    ws.conditional_formatting.add(f"D{first}:D{r}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=5,
                                                              color="86B6EF"))
    book.mat_rows = (first, r)
    r += 2
    put(ws, f"B{r}", "C. Conclusion", F_SUB)
    ws.merge_cells(f"B{r + 1}:H{r + 1}")
    put(ws, f"B{r + 1}", "[Summarise the conclusion: the benchmark and thresholds applied, the items judged material by nature, how material "
                         "items are presented and disclosed, and that immaterial information does not obscure material information "
                         "(paragraph 3.15A).]", F_INPUT, align=WRAP, fill=FILL_INPUT)
    ws.row_dimensions[r + 1].height = 34
    ws.freeze_panes = "A6"
    ws.sheet_view.showGridLines = False


# ----------------------------------------------------------------------------------------------- policies
POLICY_META = {
    # head: (reference, options permitted, selected, options for the drop-down, amount, always disclose)
    "Revenue from contracts with customers": ("Section 23 (23.3)", "Five-step model - no alternative.", "Five-step revenue model", None, ("SoCI", "REV"), False),
    "Property, plant and equipment": ("17.15-17.15D", "Cost model or revaluation model, chosen for each class.",
                                      "Revaluation model for land; cost model for other classes",
                                      ["Cost model for all classes", "Revaluation model for land; cost model for other classes"],
                                      ("SoFP", "PPE"), False),
    "Intangible assets": ("18.18-18.20", "Cost less amortisation only; finite life, not above ten years if it cannot be established reliably.",
                          "Cost model, straight-line amortisation", None, ("SoFP", "INTANG"), False),
    "Impairment of non-financial assets": ("Section 27", "No alternative.", "Recoverable amount test when indicators exist", None, ("SoFP", "PPE"), False),
    "Inventories": ("13.4, 13.18", "Cost formula: first-in, first-out or weighted average. LIFO is not permitted.", "weighted average cost",
                    ["weighted average cost", "first-in, first-out (FIFO)"], ("SoFP", "INV"), False),
    "Financial instruments": ("11.13-11.21", "Part I for basic instruments (amortised cost); Part II for other instruments (fair value).",
                              "Amortised cost; impairment on objective evidence", None, ("SoFP", "TOR"), False),
    "Cash and cash equivalents": ("7.2", "Overdrafts repayable on demand are included when integral to cash management.",
                                  "Cash, deposits of three months or less, less overdrafts", None, ("SoFP", "CASH"), False),
    "Leases": ("20.4, 20.15", "Classification as finance or operating lease by substance - no alternative.", "Operating leases: straight-line expense",
               None, ("NAT", "RENT"), False),
    "Provisions": ("21.4, 21.7", "No alternative.", "Best estimate of the amount to settle", None, ("SoFP", "PROV_C"), False),
    "Employee benefits": ("28.3, 28.24", "Actuarial gains and losses of defined benefit plans in profit or loss or in OCI.",
                          "Short-term benefits expensed as service is rendered", None, ("NAT", "EMP"), False),
    "Borrowing costs": ("25.2", "No alternative - all borrowing costs are expensed.", "Expensed as incurred", None, ("SoCI", "FINCOST"), False),
    "Income tax": ("Section 29", "No alternative.", "Current and deferred tax (temporary difference approach)", None, ("SoCI", "TAX"), False),
    "Foreign currency translation": ("30.2-30.26", "Functional currency is a matter of fact; any presentation currency may be chosen.",
                                     "Functional currency translated to the presentation currency", None, ("SoFP", "TA"), True),
    "Dividends": ("22.17", "No alternative.", "Recognised when declared", None, ("SoFP", "RE"), False),
}

PRESENTATION_CHOICES = [
    ("Statement of comprehensive income", "5.2", "Single statement, or separate income statement and statement of comprehensive income.",
     "Single statement of comprehensive income (sheet SoCI)."),
    ("Analysis of expenses", "5.11", "By nature or by function, whichever is reliable and more relevant.",
     "By function on the face (sheet SoCI); by nature in the notes. Sheet SoCI_Nature gives the alternative."),
    ("Statement of income and retained earnings", "3.18, 6.4", "Permitted in place of the SoCI and SoCE only if equity changes arise solely from "
     "profit or loss, dividends, error corrections and policy changes.", "Not used: equity also changed through share issues, a revaluation and translation differences."),
    ("Operating cash flows", "7.7", "Indirect or direct method.", "Indirect method."),
    ("Interest paid and interest received", "7.15", "Operating, or financing / investing respectively, applied consistently.",
     "Interest paid: operating. Interest received: investing."),
    ("Dividends paid", "7.16", "Financing or operating, applied consistently.", "Financing."),
    ("Current / non-current distinction", "4.4", "Current and non-current classification, or liquidity order if more relevant.",
     "Current / non-current classification."),
    ("Presentation currency", "30.17-30.19", "Any currency; translate using closing and transaction-date (or average) rates.",
     "Presentation currency per the Entity sheet; average rates for income, expenses and cash flows."),
]


def _q(text, n=240):
    """Excel string literal(s) for a long text: formulas may not contain a literal longer than 255 characters."""
    text = text.replace('"', '""')
    return "&".join('"' + text[i:i + n] + '"' for i in range(0, len(text), n))


def build_policies(book, ws):
    E = book.ent
    sheet_title(ws, "Accounting policies - selection, application and disclosure decision",
                "Select the policy where the Standard gives a choice (blue). A policy is disclosed in the notes when the related amount is "
                "material, or when you override on grounds of nature (paragraph 8.5). The notes read column I.")
    widths(ws, [4, 34, 13, 44, 38, 16, 11, 12, 90, 90])
    header_row(ws, 5, ["#", "Policy area", "Reference", "Options permitted by the Standard", "Policy selected and applied",
                       "Related amount", "Material by size?", "Override: disclose? (Yes / No)", "Policy text carried to the notes",
                       "Full policy wording (edit here)"], height=42)
    book.policy_rows = {}
    for i, (head, text) in enumerate(POLICIES):
        r = FIRST + i
        ref, options, selected, choices, amount, always = POLICY_META[head]
        book.policy_rows[head] = r
        put(ws, f"A{r}", i + 1, align=CENTER)
        put(ws, f"B{r}", head, F_BOLD, align=WRAP)
        put(ws, f"C{r}", ref, align=WRAP)
        put(ws, f"D{r}", options, align=WRAP)
        put(ws, f"E{r}", selected, F_INPUT, align=WRAP, fill=FILL_INPUT if choices else None)
        if choices:
            dv = DataValidation(type="list", formula1='"' + ",".join(c.replace(",", ";") for c in choices) + '"', allow_blank=False)
            dv.add(f"E{r}")
            ws.add_data_validation(dv)
        if amount[0] == "NAT":
            f = "ABS(" + book.to_pres(nat(CY, [amount[1]])) + ")"
        elif (amount[0], amount[1]) in book.rows:
            f = f"ABS({book.cell(amount[0], amount[1], CY, 'P')})"
        else:
            f = "0"
        put(ws, f"F{r}", "=" + f, F_LINK, NUM)
        put(ws, f"G{r}", f'=IF(F{r}>={book.mat["perf"]},"Yes","No")', align=CENTER)
        put(ws, f"H{r}", "Yes" if always else None, F_INPUT, align=CENTER)
        put(ws, f"I{r}", f'=IF(OR(H{r}="Yes",AND(H{r}<>"No",G{r}="Yes")),J{r},'
                         f'"Accounting policy information for this area is not material and is not disclosed.")', align=WRAP)
        # wording follows the selection where the Standard offers a choice
        if head == "Inventories":
            a, b_ = text.split("[weighted average cost / first-in, first-out]")
            wording = f"={_q(a)}&E{r}&{_q(b_)}"
        elif head == "Property, plant and equipment":
            cost_text = text.replace(", except land, which is carried at a revalued amount, being its fair value at the date of revaluation. "
                                     "Revaluation increases are recognised in other comprehensive income and accumulated in the revaluation "
                                     "reserve", "")
            cut = text.index("Depreciation is charged")
            wording = (f'=IF(LEFT(E{r},4)="Cost",{_q(cost_text[:cost_text.index("Depreciation is charged")])},'
                       f'{_q(text[:cut])})&{_q(text[cut:])}')
        else:
            wording = text
        put(ws, f"J{r}", wording, F_INPUT, align=WRAP)
        ws.row_dimensions[r].height = 15 * max(3, -(-len(text) // 95))
    last = FIRST + len(POLICIES) - 1
    dv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
    dv.add(f"H{FIRST}:H{last}")
    ws.add_data_validation(dv)
    ws.conditional_formatting.add(f"G{FIRST}:G{last}", CellIsRule(operator="equal", formula=['"Yes"'], fill=AMBER))
    r = last + 2
    put(ws, f"B{r}", "Presentation choices applied in this workbook", F_SUB)
    r += 1
    header_row(ws, r, ["#", "Area", "Reference", "Options permitted by the Standard", "Applied"], height=20)
    for i, (area, ref, options, applied) in enumerate(PRESENTATION_CHOICES, start=1):
        r += 1
        put(ws, f"A{r}", i, align=CENTER)
        put(ws, f"B{r}", area, F_BOLD, align=WRAP)
        put(ws, f"C{r}", ref)
        put(ws, f"D{r}", options, align=WRAP)
        put(ws, f"E{r}", applied, F_INPUT, align=WRAP)
        ws.row_dimensions[r].height = 15 * max(2, -(-max(len(options), len(applied)) // 42))
    put(ws, f"B{r + 2}", "Consistency (paragraph 10.7): the same policies are applied to similar transactions and to both years presented. "
                         "A voluntary change needs the disclosures of paragraph 10.14.", F_GREY)
    ws.freeze_panes = "C6"


# ----------------------------------------------------------------------------------------------- findings
SEVERITY = {"FAIL": "High", "WARN": "Medium", "INFO": "Low"}
ACTION = {
    "Trial balance": "Correct the trial balance at source and re-run.",
    "General ledger": "Investigate the journals listed and correct or document.",
    "GL to TB": "Reconcile the ledger to the trial balance by account.",
    "Journal review": "Inspect supporting documents for the journals listed.",
    "Continuity": "Agree opening balances to the prior-year signed financial statements.",
    "Mapping": "Confirm or change the AFS code on the Mapping sheet.",
}


def build_findings(book, ws, checks):
    E = book.ent
    sheet_title(ws, "Findings and action tracker",
                "Exceptions raised by the source-data checks, plus standing completion steps. Assign an owner and target date and update the status (blue).")
    widths(ws, [9, 13, 16, 58, 11, 60, 16, 14, 13])
    header_row(ws, 5, ["ID", "Source", "Area", "Finding / action required", "Severity", "Detail", "Owner", "Target date", "Status"],
               height=20)
    rows = [(c["id"], "DQ_Log", c["area"], c["check"] + " - exception",
             SEVERITY[c["result"]], c["detail"], ACTION.get(c["area"], "Investigate."))
            for c in checks if c["result"] in SEVERITY]
    rows += [
        ("FS-01", "Notes", "Disclosure", "Complete the entity-specific narrative in the yellow cells of the Notes", "Medium",
         "Related parties, contingencies, commitments, events after the reporting period, terms of borrowings.", ""),
        ("FS-02", "Entity", "Translation", "Agree the exchange rates to the official published rates", "Medium",
         "Closing and average rates for both years, and historical rates for opening equity.", ""),
        ("FS-03", "Materiality", "Judgement", "Document the benchmark rationale and items material by nature", "Low", "Materiality sheet, sections A and C.", ""),
        ("FS-04", "Checklist", "Disclosure", "Sign off the disclosure checklist against the Standard", "Low", "Mark each row Complete or N/A.", ""),
        ("FS-05", "Directors_Report", "Governance", "Complete and approve the report of the Board of Directors", "Low",
         "Directors, interests, auditors and the date of approval.", ""),
    ]
    order = {"High": 0, "Medium": 1, "Low": 2}
    rows.sort(key=lambda x: order[x[4]])
    r = 5
    for n, (cid, src, area, finding, sev, detail, action) in enumerate(rows):
        r += 1
        put(ws, f"A{r}", cid)
        put(ws, f"B{r}", src)
        put(ws, f"C{r}", area)
        put(ws, f"D{r}", finding, align=WRAP)
        put(ws, f"E{r}", sev, F_BOLD, align=CENTER)
        put(ws, f"F{r}", (detail + (" " + action if action else "")).strip(), F_GREY, align=WRAP)
        put(ws, f"G{r}", "[owner]", F_INPUT)
        put(ws, f"H{r}", f"={E['reporting_date']}+{30 + 15 * order[sev]}", F_INPUT, DATE)
        put(ws, f"I{r}", "Open", F_INPUT, align=CENTER)
        ws.row_dimensions[r].height = 15 * max(1, min(5, -(-len(detail + action) // 62), 5))
    last = max(r, 6)
    dv = DataValidation(type="list", formula1='"Open,In progress,Closed"', allow_blank=False)
    dv.add(f"I6:I{last + 40}")
    ws.add_data_validation(dv)
    for sev, colour in (("High", RED), ("Medium", AMBER)):
        ws.conditional_formatting.add(f"E6:E{last}", CellIsRule(operator="equal", formula=[f'"{sev}"'], fill=colour))
    ws.conditional_formatting.add(f"I6:I{last}", CellIsRule(operator="equal", formula=['"Closed"'], fill=GREEN))
    ws.freeze_panes = "A6"
    ws.auto_filter.ref = f"A5:I{last}"
    book.findings_last = last


# ----------------------------------------------------------------------------------------------- directors' report
def build_directors_report(book, ws):
    E = book.ent
    P = lambda sheet, key, fy=CY: book.cell(sheet, key, fy, "P")
    F = lambda sheet, key, fy=CY: book.cell(sheet, key, fy)
    widths(ws, [4, 58, 18, 18, 16, 16])
    ws.sheet_view.showGridLines = False
    put(ws, "B1", f"={E['entity_name']}", F_TITLE)
    put(ws, "B2", "Report of the Board of Directors", F_SUB)
    put(ws, "B3", f'="For the year ended "&TEXT({E["reporting_date"]},"d mmmm yyyy")')
    put(ws, "B4", "Blue text and yellow cells are to be completed or confirmed by the directors.", F_GREY)
    state = {"r": 6}

    def head(text):
        put(ws, f"B{state['r']}", text, F_BOLD)
        state["r"] += 1

    def para(text, inp=None):
        r = state["r"]
        ws.merge_cells(f"B{r}:F{r}")
        placeholder = "[" in text and not text.startswith("=")
        put(ws, f"B{r}", text, F_INPUT if (placeholder or inp) else F_NORMAL, align=WRAP, fill=FILL_INPUT if placeholder else None)
        ws.row_dimensions[r].height = 14.5 * max(1, -(-len(text) // 118)) + 3
        state["r"] += 2

    para(f'="The Board of Directors (the \'Directors\') presents its report together with the financial statements of "&{E["entity_name"]}'
         f'&" (the \'Company\') for the year ended "&TEXT({E["reporting_date"]},"d mmmm yyyy")&"."')
    head("The Company")
    para(f'="The Company is a "&{E["legal_form"]}&" incorporated in "&{E["country"]}&" under registration number "&{E["registration_no"]}'
         f'&". Its registered office is at "&{E["registered_office"]}&"."')
    head("Principal activities")
    para(f'="The principal activities of the Company are "&{E["principal_activities"]}&". There were no significant changes in the nature '
         f'of these activities during the year."')
    head("Financial results")
    r = state["r"]
    for col, f, font in (("C", f"={E['pres_label']}", F_BOLD), ("D", f"={E['pres_label']}", F_BOLD),
                         ("E", f"={E['currency']}", F_GREY), ("F", f"={E['currency']}", F_GREY)):
        put(ws, f"{col}{r}", f, font, align=RIGHT)
    for col, fy, font in (("C", CY, F_BOLD), ("D", PY, F_BOLD), ("E", CY, F_GREY), ("F", PY, F_GREY)):
        put(ws, f"{col}{r + 1}", book.year(fy), font, "0", RIGHT, B_BOTTOM)
    put(ws, f"B{r + 1}", None, border=B_BOTTOM)
    r += 2
    for cap, sheet, key, bold in (("Revenue", "SoCI", "REV", False), ("Profit before tax", "SoCI", "PBT", False),
                                  ("Income tax expense", "SoCI", "TAX", False), ("Profit for the year", "SoCI", "PROFIT", True),
                                  ("Total comprehensive income for the year", "SoCI", "TCI", True),
                                  ("Total assets", "SoFP", "TA", False), ("Total equity", "SoFP", "T_EQ", False)):
        put(ws, f"B{r}", cap, F_BOLD if bold else F_NORMAL)
        for col, f in (("C", P(sheet, key)), ("D", P(sheet, key, PY)), ("E", F(sheet, key)), ("F", F(sheet, key, PY))):
            put(ws, f"{col}{r}", "=" + f, F_BOLD if bold else F_NORMAL, NUM, border=B_TOP if bold else None)
        r += 1
    state["r"] = r + 1
    head("Dividends")
    div = lambda fy: f"({lead('close', fy, ['RE.DIV'])}+{glm(fy, ['RE'], ['DIVIDEND'])})"
    para(f'="Dividends of "&TEXT({div(CY)},"#,##0")&" "&{E["currency"]}&" were declared during the year ("&TEXT(YEAR({E["reporting_date"]})-1,"0")'
         f'&": "&TEXT({div(PY)},"#,##0")&" "&{E["currency"]}&"). Dividends proposed after the reporting date and not recognised amount to "'
         f'&TEXT({E["dividend_proposed"]},"#,##0")&" "&{E["currency"]}&"."')
    head("Share capital")
    para(f'="At the reporting date the Company had "&TEXT({E["shares_issued_cy"]},"#,##0")&" ordinary shares issued and fully paid ("'
         f'&TEXT(YEAR({E["reporting_date"]})-1,"0")&": "&TEXT({E["shares_issued_py"]},"#,##0")&") with a par value of "'
         f'&TEXT({E["par_value"]},"#,##0.00")&" "&{E["currency"]}&" each. "&TEXT({E["shares_issued_cy"]}-{E["shares_issued_py"]},"#,##0")'
         f'&" shares were issued during the year."')
    head("Reserves and provisions")
    para("There were no material movements to or from reserves and provisions during the year other than those disclosed in the financial statements.", True)
    head("Bad and doubtful receivables")
    para("Before the financial statements were prepared, the Directors took reasonable steps to ascertain that action had been taken in "
         "relation to the writing off of bad receivables and the making of an allowance for doubtful receivables, and satisfied themselves "
         "that all known bad receivables had been written off and an adequate allowance had been made. At the date of this report the "
         "Directors are not aware of any circumstances which would render the amount written off or the allowance inadequate to any material extent.", True)
    head("Assets")
    para("Before the financial statements were prepared, the Directors took reasonable steps to ensure that any assets which were unlikely "
         "to be realised in the ordinary course of business at their carrying amounts had been written down to an amount which they might "
         "be expected to realise. At the date of this report the Directors are not aware of any circumstances which would render the values "
         "attributed to the assets in the financial statements misleading.", True)
    head("Contingent and other liabilities")
    para("At the date of this report there is: (a) no charge on the assets of the Company which has arisen since the end of the year to "
         "secure the liabilities of any other person, and (b) no contingent liability which has arisen since the end of the year, other than "
         "in the ordinary course of business. [Confirm, or describe the exceptions.]")
    head("Going concern")
    para("The Directors have assessed the Company's ability to continue as a going concern, taking into account all available information "
         "about the future covering at least twelve months from the reporting date, and have a reasonable expectation that the Company has "
         "adequate resources to continue in operation. The financial statements are therefore prepared on the going concern basis.", True)
    head("Events after the end of the reporting period")
    para("[State that no significant events occurred after the reporting date requiring adjustment or disclosure, or describe them.]")
    head("The Board of Directors")
    para(f'="The Directors who served during the year and to the date of this report are: "&{E["directors"]}&"."')
    head("Directors' interests and benefits")
    para("[State the directors' interests in the shares of the Company. State that no arrangement existed during or at the end of the year "
         "whose object was to enable the Directors to acquire benefits by means of the acquisition of shares or debentures, and that no "
         "Director received or became entitled to a benefit by reason of a contract with the Company other than as disclosed in the notes.]")
    head("Auditors")
    para(f'="The financial statements have been audited by "&{E["auditors"]}&", who have expressed their willingness to continue in office."')
    head("Directors' responsibility in respect of the financial statements")
    para("The Directors are responsible for ensuring that the financial statements present fairly, in all material respects, the financial "
         "position of the Company at the reporting date and its financial performance and cash flows for the year then ended. In preparing "
         "them the Directors are required to: (i) adopt appropriate accounting policies, supported by reasonable and prudent judgements and "
         "estimates, and apply them consistently; (ii) comply with the IFRS for SMEs Accounting Standard or, if there has been any departure "
         "in the interest of fair presentation, ensure that it is disclosed, explained and quantified; (iii) maintain adequate accounting "
         "records and an effective system of internal control; (iv) prepare the financial statements on the going concern basis unless it "
         "is inappropriate to assume that the Company will continue in business; and (v) effectively control and direct the Company in all "
         "material decisions affecting its operations and performance. The Directors confirm that they have complied with these requirements.", True)
    head("Approval of the financial statements")
    para(f'="The accompanying financial statements, which present fairly, in all material respects, the financial position of the Company as '
         f'at "&TEXT({E["reporting_date"]},"d mmmm yyyy")&" and its financial performance and cash flows for the year then ended in accordance '
         f'with the IFRS for SMEs Accounting Standard, were approved by "&{E["authorised_by"]}&" on "&TEXT({E["authorised_date"]},"d mmmm yyyy")&"."')
    r = state["r"] + 1
    put(ws, f"B{r}", "Signed in accordance with a resolution of the Board of Directors:")
    put(ws, f"B{r + 4}", "______________________________")
    put(ws, f"C{r + 4}", "______________________________")
    put(ws, f"B{r + 5}", "[Name], Chairman", F_INPUT, fill=FILL_INPUT)
    put(ws, f"C{r + 5}", "[Name], Director", F_INPUT, fill=FILL_INPUT)
    put(ws, f"B{r + 6}", f'={E["country"]}&", "&TEXT({E["authorised_date"]},"d mmmm yyyy")')
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


# ----------------------------------------------------------------------------------------------- dashboard
def _style_axis(chart, y_fmt="#,##0"):
    chart.roundedCorners = False
    if chart.legend is not None:
        chart.legend.overlay = False
        chart.legend.position = "b"
    if hasattr(chart, "x_axis"):
        chart.x_axis.delete = False
        chart.y_axis.delete = False
        chart.y_axis.number_format = y_fmt
        chart.y_axis.majorGridlines.spPr = _line(GRID)
        chart.x_axis.tickLblPos = "low"
        chart.x_axis.spPr = _line("C3C2B7")


def _line(colour):
    from openpyxl.chart.shapes import GraphicalProperties
    gp = GraphicalProperties()
    gp.line.solidFill = colour
    gp.line.width = 6350
    return gp


def _colour(series, hexcolour, line=False):
    if line:
        series.graphicalProperties.line.solidFill = hexcolour
        series.graphicalProperties.line.width = 25400
        series.smooth = False
    else:
        series.graphicalProperties.solidFill = hexcolour
        series.graphicalProperties.line.solidFill = "FFFFFF"


def build_dash_data(book, ws):
    """Formula tables that feed the dashboard charts (also the table view of every chart)."""
    E = book.ent
    P = lambda sheet, key, fy=CY: book.cell(sheet, key, fy, "P")
    sheet_title(ws, "Dashboard data - tables behind the dashboard charts", "All formulas; amounts in the presentation currency.")
    widths(ws, [44, 18, 18, 4, 44, 18, 18])
    d = {}
    # 1 asset composition
    put(ws, "A4", "1. Asset composition (current year)", F_BOLD)
    r = 5
    assets = [(c, cap) for c, cap, st, sec in FACES if sec in ("NCA", "CA") and c in book.active_faces]
    for code, cap in assets:
        put(ws, f"A{r}", cap + (" (non-current)" if code == "OFA_NC" else ""))
        put(ws, f"B{r}", "=" + P("SoFP", code), fmt=NUM)
        r += 1
    d["assets"] = (5, r - 1)
    # 2 performance
    r += 1
    put(ws, f"A{r}", "2. Performance", F_BOLD)
    put(ws, f"B{r}", book.year(CY), F_BOLD, "0")
    put(ws, f"C{r}", book.year(PY), F_BOLD, "0")
    d["perf_head"] = r
    for cap, key in (("Revenue", "REV"), ("Gross profit", "GP"), ("Operating profit", "OP"), ("Profit for the year", "PROFIT")):
        r += 1
        put(ws, f"A{r}", cap)
        put(ws, f"B{r}", "=" + P("SoCI", key), fmt=NUM)
        put(ws, f"C{r}", "=" + P("SoCI", key, PY), fmt=NUM)
    d["perf"] = (d["perf_head"] + 1, r)
    # 3 cash flows
    r += 2
    put(ws, f"A{r}", "3. Cash flows by activity", F_BOLD)
    put(ws, f"B{r}", book.year(CY), F_BOLD, "0")
    put(ws, f"C{r}", book.year(PY), F_BOLD, "0")
    d["cf_head"] = r
    for cap, key in (("Operating", "OPER"), ("Investing", "INVEST"), ("Financing", "FIN")):
        r += 1
        put(ws, f"A{r}", cap)
        put(ws, f"B{r}", "=" + P("SoCF", key), fmt=NUM)
        put(ws, f"C{r}", "=" + P("SoCF", key, PY), fmt=NUM)
    d["cf"] = (d["cf_head"] + 1, r)
    # 4 findings by severity
    r += 2
    put(ws, f"A{r}", "4. Open findings by severity", F_BOLD)
    d["sev_head"] = r
    for sev in ("High", "Medium", "Low"):
        r += 1
        put(ws, f"A{r}", sev)
        put(ws, f"B{r}", f'=COUNTIFS(Findings!$E$6:$E${book.findings_last},"{sev}",Findings!$I$6:$I${book.findings_last},"<>Closed")')
    d["sev"] = (d["sev_head"] + 1, r)
    # 5 expenses by nature (right-hand block)
    put(ws, "E4", "5. Expenses by nature (current year)", F_BOLD)
    q = 5
    for code, cap in NATURES:
        if code != "OTHER" and code not in book.active_natures:
            continue
        key = "N_" + code
        put(ws, f"E{q}", cap.replace(" of property, plant and equipment", "").replace(" recognised as an expense", ""))
        put(ws, f"F{q}", "=-" + P("SoCI_Nature", key), fmt=NUM)
        q += 1
    d["exp"] = (5, q - 1)
    # 6 monthly revenue
    q += 1
    put(ws, f"E{q}", "6. Monthly revenue per the general ledger", F_BOLD)
    put(ws, f"F{q}", book.year(CY), F_BOLD, "0")
    put(ws, f"G{q}", book.year(PY), F_BOLD, "0")
    d["rev_head"] = q
    rev = ["REV.GOODS", "REV.SERV"]
    for m in range(12):
        q += 1
        put(ws, f"E{q}", f'=TEXT(EDATE({E["reporting_date"]},{m - 12})+1,"mmm")')
        for col, fy, back in (("F", CY, 12), ("G", PY, 24)):
            start = f"EDATE({E['reporting_date']},{m - back})"
            end = f"EDATE({E['reporting_date']},{m - back + 1})"
            terms = "+".join(f'SUMIFS(GL!$J:$J,GL!$K:$K,"{l}",GL!$A:$A,"{fy}",GL!$B:$B,">"&{start},GL!$B:$B,"<="&{end})' for l in rev)
            put(ws, f"{col}{q}", f"=-({terms}){book.conv('A', fy)}", fmt=NUM)
    d["rev"] = (d["rev_head"] + 1, q)
    # 7 ageing of receivables
    q += 2
    put(ws, f"E{q}", "7. Age of trade receivables (gross, current year)", F_BOLD)
    d["age_head"] = q
    if ("Notes", "AGE_ND") in book.rows:
        nd = book.rows[("Notes", "AGE_ND")]
        for i, cap in enumerate(["Not yet due", "Up to 1 month", "1 to 3 months", "3 to 12 months", "Over 1 year"]):
            q += 1
            put(ws, f"E{q}", cap)
            put(ws, f"F{q}", f"=Notes!C{nd + i}", fmt=NUM)
    d["age"] = (d["age_head"] + 1, q)
    book.dash = d


def build_dashboard(book, ws):
    E, d = book.ent, book.dash
    P = lambda sheet, key, fy=CY: book.cell(sheet, key, fy, "P")
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 85
    widths(ws, [27, 2] + [13.5] * 12 + [2])
    navy = PatternFill("solid", fgColor=NAVY)
    white = PatternFill("solid", fgColor="FFFFFF")
    page = PatternFill("solid", fgColor="F4F6FA")
    for r in range(1, 76):
        for c in range(1, 16):
            ws.cell(row=r, column=c).fill = navy if (c == 1 or r <= 3) else page
    # top band
    ws.row_dimensions[1].height, ws.row_dimensions[2].height, ws.row_dimensions[3].height = 8, 30, 20
    put(ws, "A2", "  FINANCIAL REPORTING", Font(name=FONT, size=11, bold=True, color="FFFFFF"), align=Alignment(vertical="bottom"))
    put(ws, "A3", "  & ASSURANCE DASHBOARD", Font(name=FONT, size=9, bold=True, color="9EC5F4"))
    put(ws, "C2", f"={E['entity_name']}", Font(name=FONT, size=16, bold=True, color="FFFFFF"))
    put(ws, "C3", f'="Year ended "&TEXT({E["reporting_date"]},"d mmmm yyyy")&"   |   IFRS for SMEs Accounting Standard (third edition)   |   Functional currency "'
                  f'&{E["currency"]}&"   |   Presentation currency "&{E["pres_label"]}', Font(name=FONT, size=9, color="D9E1F2"))
    put(ws, "M2", f"=Checks!C3", Font(name=FONT, size=10, bold=True, color="FFFFFF"), align=Alignment(horizontal="right"))
    # sidebar navigation
    nav = [("Dashboard", "Dashboard"), ("Entity & rates", "Entity"), ("General ledger", "GL"), ("Trial balance", "TB"),
           ("Adjusting entries", "AJE"), ("Adjusted trial balance", "ATB"), ("Mapping", "Mapping"), ("Lead schedule", "Lead"),
           ("Data quality log", "DQ_Log"), ("Live checks", "Checks"), ("Findings & actions", "Findings"),
           ("Materiality", "Materiality"), ("Accounting policies", "Policies"), ("Directors' report", "Directors_Report"),
           ("Financial position", "SoFP"), ("Comprehensive income", "SoCI"), ("Changes in equity", "SoCE"),
           ("Cash flows", "SoCF"), ("Notes", "Notes"), ("Comparative analysis", "Analysis"), ("Disclosure checklist", "Checklist"),
           ("Compliance checklist", "Compliance_Summary")]
    for i, (label, sheet) in enumerate(nav):
        c = ws.cell(row=5 + i * 2, column=1, value=f'=HYPERLINK("#\'{sheet}\'!A1","   {label}")')
        c.font = Font(name=FONT, size=10, bold=(i == 0), color="FFFFFF")
        if i == 0:
            c.fill = PatternFill("solid", fgColor="2A78D6")

    # KPI tiles
    thin = Side(style="thin", color=GRID)

    def tile(col, row, label, formula, fmt=NUM, sub=None, colour=INK):
        c1, c2 = chr(ord("C") + 2 * col), chr(ord("C") + 2 * col + 1)
        for r in range(row, row + 4):
            for cc in (c1, c2):
                ws[f"{cc}{r}"].fill = white
                ws[f"{cc}{r}"].border = Border(top=thin if r == row else None, bottom=thin if r == row + 3 else None,
                                               left=thin if cc == c1 else None, right=thin if cc == c2 else None)
        put(ws, f"{c1}{row}", " " + label, Font(name=FONT, size=9, color="52514E"))
        ws.merge_cells(f"{c1}{row + 1}:{c2}{row + 2}")
        put(ws, f"{c1}{row + 1}", "=" + formula, Font(name=FONT, size=18, bold=True, color=colour), fmt,
            Alignment(horizontal="center", vertical="center"))
        if sub:
            put(ws, f"{c1}{row + 3}", sub, Font(name=FONT, size=8, color=MUTED))

    lab = f'=" "&{E["pres_label"]}'
    yoy = lambda sheet, key: (f'=" "&IF({P(sheet, key, PY)}=0,"n/a",TEXT({P(sheet, key)}/{P(sheet, key, PY)}-1,"+0.0%;-0.0%")'
                              f'&" vs prior year")')
    tile(0, 5, "Revenue", P("SoCI", "REV"), sub=yoy("SoCI", "REV"))
    tile(1, 5, "Profit for the year", P("SoCI", "PROFIT"), sub=yoy("SoCI", "PROFIT"))
    tile(2, 5, "Total assets", P("SoFP", "TA"), sub=yoy("SoFP", "TA"))
    tile(3, 5, "Total equity", P("SoFP", "T_EQ"), sub=yoy("SoFP", "T_EQ"))
    tile(4, 5, "Cash and cash equivalents", P("SoCF", "CLOSE"), sub=yoy("SoCF", "CLOSE"))
    tile(5, 5, "Overall materiality", book.mat["overall"][0:] , sub=lab)
    cl = book.checks_last
    tile(0, 10, "Accounts in the trial balance", f"COUNTA(TB!A{FIRST}:A{book.tb_last})", "#,##0", " mapped to AFS components", "2A78D6")
    tile(1, 10, "GL lines tested", f"COUNTA(GL!C{FIRST}:C{book.gl_last})", "#,##0", " both years, after cleaning", "2A78D6")
    tile(2, 10, "Adjusting entry lines", f"COUNTA(AJE!B{FIRST}:B{FIRST + 59})", "#,##0", " carried to the ATB", "2A78D6")
    tile(3, 10, "Live checks passed", f'COUNTIF(Checks!D6:D{cl},"PASS")/COUNTA(Checks!D6:D{cl})', "0%",
         f'=" "&COUNTIF(Checks!D6:D{cl},"FAIL")&" failing of "&COUNTA(Checks!D6:D{cl})', "006300")
    tile(4, 10, "Open findings", f'COUNTIF(Findings!I6:I{book.findings_last},"<>Closed")', "#,##0",
         f'=" "&COUNTIFS(Findings!E6:E{book.findings_last},"High",Findings!I6:I{book.findings_last},"<>Closed")&" high severity"', "D03B3B")
    ct = book.compliance_total
    tile(5, 10, "Compliance checklist complete", f'IFERROR({ct}+0,0)', "0%",
         f'=" "&{ct.replace("K", "J")}&" of "&{ct.replace("K", "F")}&" applicable items open"', "EB6834")

    # charts
    def place(chart, col, row, title, w=10.2, h=7.6):
        chart.title = title
        chart.title.overlay = False
        chart.width, chart.height = w, h
        ws.add_chart(chart, f"{col}{row}")

    def cats(head, rng, col="A"):
        return Reference(book.wb["Dash_Data"], min_col=ord(col) - 64, min_row=rng[0], max_row=rng[1])

    dd = book.wb["Dash_Data"]
    # 1 doughnut - asset composition
    ch = DoughnutChart()
    ch.add_data(Reference(dd, min_col=2, min_row=d["assets"][0], max_row=d["assets"][1]), titles_from_data=False)
    ch.set_categories(cats(None, d["assets"]))
    ch.holeSize = 60
    for i in range(d["assets"][1] - d["assets"][0] + 1):
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = SERIES[i % len(SERIES)]
        pt.graphicalProperties.line.solidFill = "FFFFFF"
        ch.series[0].dPt.append(pt)
    ch.dataLabels = DataLabelList()
    ch.dataLabels.showPercent, ch.dataLabels.showVal, ch.dataLabels.showSerName = True, False, False
    ch.dataLabels.showCatName = ch.dataLabels.showLeaderLines = False
    _style_axis(ch)
    place(ch, "C", 15, "1. Asset composition")
    # 2 performance, two years
    ch = BarChart()
    ch.type, ch.grouping, ch.gapWidth, ch.overlap = "col", "clustered", 80, -10
    ch.add_data(Reference(dd, min_col=2, max_col=3, min_row=d["perf_head"], max_row=d["perf"][1]), titles_from_data=True)
    ch.set_categories(cats(None, d["perf"]))
    for s_, c_ in zip(ch.series, SERIES):
        _colour(s_, c_)
    _style_axis(ch)
    place(ch, "G", 15, "2. Performance - two-year comparison")
    # 3 cash flows
    ch = BarChart()
    ch.type, ch.grouping, ch.gapWidth, ch.overlap = "col", "clustered", 80, -10
    ch.add_data(Reference(dd, min_col=2, max_col=3, min_row=d["cf_head"], max_row=d["cf"][1]), titles_from_data=True)
    ch.set_categories(cats(None, d["cf"]))
    for s_, c_ in zip(ch.series, SERIES):
        _colour(s_, c_)
    _style_axis(ch)
    place(ch, "K", 15, "3. Cash flows by activity")
    # 4 expenses by nature (single series: no legend)
    ch = BarChart()
    ch.type, ch.gapWidth = "bar", 60
    ch.add_data(Reference(dd, min_col=6, min_row=d["exp"][0], max_row=d["exp"][1]), titles_from_data=False)
    ch.set_categories(cats(None, d["exp"], "E"))
    _colour(ch.series[0], SERIES[0])
    ch.legend = None
    _style_axis(ch)
    ch.x_axis.scaling.orientation = "maxMin"
    place(ch, "C", 32, "4. Expenses by nature - current year")
    # 5 monthly revenue
    ch = LineChart()
    ch.add_data(Reference(dd, min_col=6, max_col=7, min_row=d["rev_head"], max_row=d["rev"][1]), titles_from_data=True)
    ch.set_categories(cats(None, d["rev"], "E"))
    for s_, c_ in zip(ch.series, SERIES):
        _colour(s_, c_, line=True)
    _style_axis(ch)
    place(ch, "G", 32, "5. Monthly revenue - two years")
    # 6 receivables ageing
    ch = BarChart()
    ch.type, ch.gapWidth = "col", 60
    ch.add_data(Reference(dd, min_col=6, min_row=d["age"][0], max_row=d["age"][1]), titles_from_data=False)
    ch.set_categories(cats(None, d["age"], "E"))
    _colour(ch.series[0], SERIES[0])
    ch.legend = None
    ch.dataLabels = DataLabelList()
    ch.dataLabels.showVal = True
    ch.dataLabels.showSerName = ch.dataLabels.showCatName = ch.dataLabels.showLegendKey = False
    _style_axis(ch)
    place(ch, "K", 32, "6. Age of trade receivables")

    # tables
    def table_head(col, row, labels, spans):
        c = ord(col)
        for lab, span in zip(labels, spans):
            cell = f"{chr(c)}{row}"
            put(ws, cell, lab, Font(name=FONT, size=9, bold=True, color="FFFFFF"), fill=navy,
                align=Alignment(horizontal="left" if c == ord(col) else "center", vertical="center"))
            for k in range(1, span):
                ws[f"{chr(c + k)}{row}"].fill = navy
            c += span

    put(ws, "C49", "7. Key ratios", Font(name=FONT, size=11, bold=True, color=INK))
    table_head("C", 50, ["Ratio", f"=Analysis!B5", f"=Analysis!C5", "Change"], [3, 1, 1, 1])
    for col in "FG":
        ws[f"{col}50"].number_format = "0"
    for i, (row, fmt) in enumerate(book.ratio_rows[:11]):
        r = 51 + i
        for col in "CDEFGH":
            ws[f"{col}{r}"].fill = white
        put(ws, f"C{r}", f"=Analysis!A{row}", Font(name=FONT, size=9), fill=white)
        put(ws, f"F{r}", f"=Analysis!B{row}", Font(name=FONT, size=9, bold=True), fmt, CENTER, fill=white)
        put(ws, f"G{r}", f"=Analysis!C{row}", Font(name=FONT, size=9), fmt, CENTER, fill=white)
        put(ws, f"H{r}", f"=Analysis!D{row}", Font(name=FONT, size=9), fmt, CENTER, fill=white)

    put(ws, "J49", "8. Priority actions", Font(name=FONT, size=11, bold=True, color=INK))
    table_head("J", 50, ["ID", "Finding / action", "Severity", "Status"], [1, 2, 1, 1])
    for i in range(11):
        r, fr = 51 + i, 6 + i
        for col in "JKLMN":
            ws[f"{col}{r}"].fill = white
        if fr > book.findings_last:
            continue
        put(ws, f"J{r}", f"=Findings!A{fr}", Font(name=FONT, size=9), fill=white)
        put(ws, f"K{r}", f"=Findings!D{fr}", Font(name=FONT, size=9), fill=white)
        put(ws, f"M{r}", f"=Findings!E{fr}", Font(name=FONT, size=9, bold=True), align=CENTER, fill=white)
        put(ws, f"N{r}", f"=Findings!I{fr}", Font(name=FONT, size=9), align=CENTER, fill=white)
    for sev, colour in (("High", RED), ("Medium", AMBER)):
        ws.conditional_formatting.add("M51:M61", CellIsRule(operator="equal", formula=[f'"{sev}"'], fill=colour))
    ws.conditional_formatting.add("N51:N61", CellIsRule(operator="equal", formula=['"Closed"'], fill=GREEN))
    put(ws, "C64", "Every figure is a formula on the statements, checks and working papers; the tables behind the charts are on the "
                   "Dash_Data sheet. Use the panel on the left to open a working paper.", Font(name=FONT, size=8, italic=True, color=MUTED))
