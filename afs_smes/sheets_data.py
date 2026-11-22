"""Input and working-paper sheets: Guide, Entity, DQ_Log, Mapping, Lead, TB, AJE, GL, Lists."""
from datetime import date, datetime

from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

from .taxonomy import FACE, LEAVES, MVT_TYPES, NATURES
from .xl import (B_BOTTOM, CENTER, DATE, F_BOLD, F_GREY, F_INPUT, F_LINK, F_NORMAL, F_SUB, FILL_BAND,
                 FILL_INPUT, NUM2, WRAP, header_row, put, sheet_title, widths)

GREEN = PatternFill("solid", bgColor="C6EFCE")
RED = PatternFill("solid", bgColor="FFC7CE")
AMBER = PatternFill("solid", bgColor="FFEB9C")
FIRST = 6  # first data row on every working-paper sheet (header on row 5)


def status_formatting(ws, rng):
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"PASS"'], fill=GREEN))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"OK"'], fill=GREEN))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"FAIL"'], fill=RED))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"UNMAPPED"'], fill=RED))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"WARN"'], fill=AMBER))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"REVIEW"'], fill=AMBER))


def build_guide(book, ws):
    sheet_title(ws, "IFRS for SMEs (third edition) - annual financial statements workbook")
    widths(ws, [4, 26, 100])
    r = 3
    put(ws, f"B{r}", "Workflow", F_SUB)
    steps = [
        ("1  Read", "GL and TB exports are read as delivered (any column order, title rows, text numbers, mixed dates)."),
        ("2  Clean", "Every clean-up action is listed in DQ_Log (file, source row, issue, action). Nothing is changed silently."),
        ("3  Check", "DQ_Log: completeness, accuracy and reliability tests run on the source data. "
                     "Checks: live reconciliations that recalculate whenever this workbook changes."),
        ("4  Map", "Mapping: each ledger account is assigned to an AFS component (and expenses to a nature). "
                   "Review the blue cells; rows marked REVIEW were mapped by fallback."),
        ("5  Adjust", "AJE: post closing/audit adjustments and reclassifications by account. They flow into TB column O."),
        ("6  Draft", "SoFP, SoCI, SoCE, SoCF and Notes are formulas on Lead/TB/GL. Nothing is typed into the statements."),
        ("7  Judge", "Materiality: thresholds and line-by-line assessment. Policies: select each accounting policy; material policies flow into Note 3."),
        ("8  Review", "Dashboard, Findings (action tracker), Analysis (two-year comparison) and Checklist (disclosures by paragraph)."),
    ]
    for k, v in steps:
        r += 1
        put(ws, f"B{r}", k, F_BOLD)
        put(ws, f"C{r}", v, F_NORMAL, align=WRAP)
        ws.row_dimensions[r].height = 28
    r += 2
    put(ws, f"B{r}", "Colour legend", F_SUB)
    legend = [
        ("Blue text", "Input or mapping you may edit (Entity, Mapping columns D and F, AJE, GL movement type, note narratives).", F_INPUT, None),
        ("Yellow fill", "Entity-specific text or figure that must be completed or confirmed before issue.", F_NORMAL, FILL_INPUT),
        ("Black text", "Formula - do not overwrite.", F_NORMAL, None),
        ("Green text", "Formula that links to another sheet.", F_LINK, None),
    ]
    for k, v, f, fill in legend:
        r += 1
        put(ws, f"B{r}", k, f, fill=fill)
        put(ws, f"C{r}", v, F_NORMAL, align=WRAP)
    r += 2
    put(ws, f"B{r}", "Conventions", F_SUB)
    notes = [
        "Ledger data (TB, GL, Lead, AJE) is signed: debit positive, credit negative. Statements show natural signs; deductions in brackets.",
        "The trial balances are pre-closing: income, expenses and dividends declared still sit in their own accounts. "
        "Retained earnings on the SoFP = ledger retained earnings + profit for the year - dividends declared.",
        "Roll-forward notes (PPE, intangibles, borrowings, provisions, equity) and several cash flow lines are driven by the "
        "movement type in GL column L. It is pre-filled from the other accounts in each journal and can be overridden line by line.",
        "Rows for line items with no balance are hidden, not deleted: unhide them if an account is remapped. "
        "If you add accounts or journals, re-run the tool so every sheet is rebuilt.",
        "Statements are displayed to the nearest currency unit; underlying figures keep their cents, so a total may differ by 1 from the sum of displayed lines.",
        "Paragraph references are to the IFRS for SMEs Accounting Standard, third edition (issued February 2025, effective for annual "
        "periods beginning on or after 1 January 2027, earlier application permitted - paragraph A1). Accounting-policy wording is a "
        "drafting template: tailor it to the entity's facts before issue.",
    ]
    for t in notes:
        r += 1
        put(ws, f"B{r}", "-", F_NORMAL, align=CENTER)
        ws.merge_cells(f"B{r}:B{r}")
        put(ws, f"C{r}", t, F_NORMAL, align=WRAP)
        ws.row_dimensions[r].height = 15 * (1 + len(t) // 105)
    ws.sheet_view.showGridLines = False


ENTITY_FIELDS = [
    ("entity_name", "Entity name", "[Entity name]", None),
    ("legal_form", "Legal form", "[legal form]", None),
    ("country", "Country of incorporation", "[Country]", None),
    ("registered_office", "Registered office / principal place of business", "[Address]", None),
    ("principal_activities", "Principal activities", "[principal activities]", None),
    ("reporting_date", "Reporting date (end of current year)", None, DATE),
    ("prior_date", "End of comparative year", None, DATE),
    ("currency", "Functional currency", "USD", None),
    ("pres_currency", "Presentation currency", None, None),
    ("pres_scale", "Presentation currency amounts are shown in units of", 1, "#,##0"),
    ("pres_label", "Presentation currency label used in headings", None, None),
    ("pres_reason", "Reason for using a presentation currency different from the functional currency",
     "[state the reason, e.g. a requirement of local law or regulation]", None),
    ("fx_close_cy", "Exchange rate - closing, current year end (presentation currency per 1 unit of functional currency)", 1, "#,##0.00"),
    ("fx_avg_cy", "Exchange rate - average for the current year", 1, "#,##0.00"),
    ("fx_close_py", "Exchange rate - closing, prior year end", 1, "#,##0.00"),
    ("fx_avg_py", "Exchange rate - average for the prior year", 1, "#,##0.00"),
    ("fx_close_ppy", "Exchange rate - closing, at the beginning of the prior year", 1, "#,##0.00"),
    ("fx_hist_capital", "Historical rate - share capital and share premium at the beginning of the prior year", 1, "#,##0.00"),
    ("fx_hist_re", "Historical (blended) rate - retained earnings and reserves at the beginning of the prior year", 1, "#,##0.00"),
    ("registration_no", "Company registration number", "[registration number]", None),
    ("directors", "Directors who served during the year and to the date of the report", "[names and positions of directors]", None),
    ("auditors", "Auditors", "[name of audit firm]", None),
    ("tax_rate", "Applicable income tax rate", 0.2, "0.0%"),
    ("authorised_date", "Date financial statements authorised for issue", None, DATE),
    ("authorised_by", "Authorised for issue by", "[the Board of Directors]", None),
    ("dividend_proposed", "Dividend proposed/declared after the reporting date, not recognised (total)", 0, "#,##0"),
    ("shares_authorised", "Number of shares authorised", 0, "#,##0"),
    ("shares_issued_cy", "Number of shares issued and fully paid - current year", 0, "#,##0"),
    ("shares_issued_py", "Number of shares issued and fully paid - prior year", 0, "#,##0"),
    ("par_value", "Par value per share", 0, "#,##0.00"),
    ("employees_cy", "Average number of employees - current year", 0, "#,##0"),
    ("employees_py", "Average number of employees - prior year", 0, "#,##0"),
    ("tolerance", "Tolerance for reconciliation checks (currency units)", 0.5, "0.00"),
    ("materiality", "Analytical review: minimum change to flag (currency units)", 5000, "#,##0"),
    ("threshold", "Analytical review: minimum % change to flag", 0.10, "0.0%"),
]


def build_entity(book, ws, cfg):
    sheet_title(ws, "Entity information and parameters", "Blue cells are inputs; they feed statement headings, notes and checks.")
    widths(ws, [4, 58, 60])
    header_row(ws, 4, ["", "Item", "Value"], height=20)
    for i, (key, label, default, fmt) in enumerate(ENTITY_FIELDS):
        r = 5 + i
        put(ws, f"B{r}", label)
        val = cfg.get(key, default)
        if key in ("reporting_date", "authorised_date") and isinstance(val, str):
            val = datetime.strptime(val, "%Y-%m-%d").date()
        if key == "prior_date":
            val = f"=EDATE({book.ent['reporting_date']},-12)"
        if key == "pres_currency" and val is None:
            val = cfg.get("currency", "USD")
        if key == "pres_label":
            val = (f'={book.ent["pres_currency"]}&IF({book.ent["pres_scale"]}=1000,"\'000",'
                   f'IF({book.ent["pres_scale"]}=1000000," million",""))')
        if key == "authorised_date" and val is None:
            val = "[date]"
        c = put(ws, f"C{r}", val, F_NORMAL if key in ("prior_date", "pres_label") else F_INPUT, fmt,
                align=Alignment(horizontal="left", wrap_text=True, vertical="top"))
        if isinstance(val, str) and val.startswith("["):
            c.fill = FILL_INPUT
        book.ent[key] = f"Entity!$C${r}"
    ws.sheet_view.showGridLines = False


def build_dq_log(book, ws, checks, cleaning, sources):
    sheet_title(ws, "Data quality log - source data checks and cleaning actions",
                f"Produced by the engine when the workbook was built ({datetime.now():%d %b %Y %H:%M}). "
                "Static evidence: re-run the tool after correcting the source files.")
    widths(ws, [10, 16, 70, 9, 11, 90])
    put(ws, "A3", "Source files: " + "; ".join(sources), F_GREY)
    header_row(ws, 5, ["ID", "Area", "Check", "Result", "Exceptions", "Detail"], height=20)
    r = 5
    for c in checks:
        r += 1
        for j, k in enumerate(["id", "area", "check", "result", "exceptions", "detail"], start=1):
            cell = ws.cell(row=r, column=j, value=c[k])
            cell.font, cell.alignment = F_NORMAL, WRAP
        ws.cell(row=r, column=4).alignment = CENTER
        ws.row_dimensions[r].height = 15 * max(1, min(6, 1 + len(str(c["detail"])) // 95))
    status_formatting(ws, f"D6:D{r}")
    ws.conditional_formatting.add(f"D6:D{r}", CellIsRule(operator="equal", formula=['"INFO"'], fill=FILL_BAND))
    r += 2
    put(ws, f"A{r}", "Cleaning log - every change made to the data as delivered", F_SUB)
    r += 1
    header_row(ws, r, ["", "File", "Issue found", "Row", "", "Action taken"], height=20)
    for c in cleaning:
        r += 1
        put(ws, f"B{r}", c["file"], align=WRAP)
        put(ws, f"C{r}", c["issue"], align=WRAP)
        put(ws, f"D{r}", str(c["row"]), align=CENTER)
        put(ws, f"F{r}", c["action"], align=WRAP)
        ws.row_dimensions[r].height = 15 * max(1, 1 + len(c["issue"]) // 75)
    ws.freeze_panes = "A6"


def build_lead(book, ws):
    sheet_title(ws, "Lead schedule - balances by AFS component (debit +, credit -)",
                "One row per AFS component. Mapping column D must use a code from column A.")
    widths(ws, [15, 58, 13, 50, 9, 16, 16, 16, 16, 16, 16])
    header_row(ws, 5, ["AFS code", "AFS component", "Face line", "Line item on the face of the statement",
                       "Statement", "CY closing (final)", "PY closing", "CY opening", "PY opening", "CY movement",
                       "PY movement"])
    for i, (code, caption, fc, _) in enumerate(LEAVES):
        r = FIRST + i
        put(ws, f"A{r}", code)
        put(ws, f"B{r}", caption)
        put(ws, f"C{r}", fc)
        put(ws, f"D{r}", FACE[fc][1])
        put(ws, f"E{r}", FACE[fc][2])
        for col, tbcol in (("F", "O"), ("G", "I"), ("H", "J"), ("I", "F")):
            put(ws, f"{col}{r}", f"=SUMIFS(ATB!${tbcol}:${tbcol},ATB!$D:$D,$A{r})", F_LINK, NUM2)
        put(ws, f"J{r}", f"=F{r}-H{r}", fmt=NUM2)
        put(ws, f"K{r}", f"=G{r}-I{r}", fmt=NUM2)
    last = FIRST + len(LEAVES) - 1
    put(ws, f"B{last + 1}", "Total (must be zero)", F_BOLD)
    for col in "FGHI":
        put(ws, f"{col}{last + 1}", f"=SUM({col}{FIRST}:{col}{last})", F_BOLD, NUM2, border=B_BOTTOM)
    ws.freeze_panes = "C6"
    book.lead_last = last


def build_lists(book, ws):
    sheet_title(ws, "Lists - codes used in drop-downs")
    widths(ws, [12, 60, 4, 12, 60])
    header_row(ws, 5, ["Nature", "Expense nature (Mapping column F)", "", "Movement", "GL movement type (GL column L)"], height=20)
    for i, (c, t) in enumerate(NATURES):
        put(ws, f"A{FIRST + i}", c)
        put(ws, f"B{FIRST + i}", t)
    for i, (c, t) in enumerate(MVT_TYPES):
        put(ws, f"D{FIRST + i}", c)
        put(ws, f"E{FIRST + i}", t)


def build_mapping(book, ws, mapping):
    sheet_title(ws, "Mapping of ledger accounts to AFS components",
                "Edit the blue cells (AFS code, nature). Class: A asset, L liability, E equity, I income, C cost of sales, X expense, T tax.")
    widths(ws, [11, 46, 7, 15, 50, 10, 46, 10, 34, 16, 16, 12])
    header_row(ws, 5, ["Account", "Account name", "Class", "AFS code", "AFS component", "Expense nature",
                       "Line item on the face", "Statement", "Mapping basis", "CY balance (final)", "PY balance",
                       "Status"])
    n = len(mapping)
    for i, m in enumerate(mapping.itertuples()):
        r = FIRST + i
        put(ws, f"A{r}", m.account)
        put(ws, f"B{r}", m.account_name)
        put(ws, f"C{r}", m.cls, align=CENTER)
        put(ws, f"D{r}", m.leaf, F_INPUT)
        put(ws, f"F{r}", m.nature or None, F_INPUT, align=CENTER)
        for col, src in (("E", "B"), ("G", "D"), ("H", "E")):
            put(ws, f"{col}{r}", f'=IFERROR(INDEX(Lead!${src}:${src},MATCH($D{r},Lead!$A:$A,0)),"")', F_LINK)
        put(ws, f"I{r}", m.basis, F_GREY)
        put(ws, f"J{r}", f"=SUMIFS(ATB!$O:$O,ATB!$A:$A,$A{r})", F_LINK, NUM2)
        put(ws, f"K{r}", f"=SUMIFS(ATB!$I:$I,ATB!$A:$A,$A{r})", F_LINK, NUM2)
        put(ws, f"L{r}", f'=IF(ISNUMBER(MATCH($D{r},Lead!$A:$A,0)),IF(ISNUMBER(SEARCH("review",$I{r})),"REVIEW","OK"),"UNMAPPED")',
            align=CENTER)
    last = FIRST + n - 1
    dv = DataValidation(type="list", formula1=f"Lead!$A${FIRST}:$A${book.lead_last}", allow_blank=False,
                        showErrorMessage=True, errorTitle="AFS code", error="Choose a code from the Lead sheet.")
    dv.add(f"D{FIRST}:D{last}")
    dv2 = DataValidation(type="list", formula1=f"Lists!$A${FIRST}:$A${FIRST + len(NATURES) - 1}", allow_blank=True)
    dv2.add(f"F{FIRST}:F{last}")
    ws.add_data_validation(dv)
    ws.add_data_validation(dv2)
    status_formatting(ws, f"L{FIRST}:L{last}")
    ws.freeze_panes = "C6"
    ws.auto_filter.ref = f"A5:L{last}"
    book.map_last = last


def build_atb(book, ws, mapping):
    sheet_title(ws, "Adjusted trial balance (ATB) - mapped, in functional currency (debit +, credit -)",
                "Columns F-M are linked to the unadjusted TB sheet. O = unadjusted closing + adjustments from AJE. P-U are control columns and must be zero.")
    widths(ws, [11, 44, 7, 13, 9] + [15] * 10 + [11, 11, 15, 11, 15, 11])
    header_row(ws, 5, ["Account", "Account name", "Class", "AFS code", "Nature", "PY opening", "PY debits", "PY credits",
                       "PY closing", "CY opening", "CY debits", "CY credits", "CY closing (unadjusted)",
                       "Adjustments (AJE)", "CY closing (adjusted)", "CY cast check", "PY cast check",
                       "CY GL movement", "CY TB v GL", "PY GL movement", "PY TB v GL"], height=42)
    for i, m in enumerate(mapping.itertuples()):
        r = FIRST + i
        put(ws, f"A{r}", m.account)
        put(ws, f"B{r}", m.account_name)
        put(ws, f"C{r}", m.cls, align=CENTER)
        put(ws, f"D{r}", f'=IFERROR(INDEX(Mapping!$D:$D,MATCH($A{r},Mapping!$A:$A,0)),"UNMAPPED")', F_LINK)
        put(ws, f"E{r}", f'=IFERROR(INDEX(Mapping!$F:$F,MATCH($A{r},Mapping!$A:$A,0))&"","")', F_LINK, align=CENTER)
        for col, f in zip("FGHIJKLM", [f"=TB!C{r}-TB!D{r}", f"=TB!E{r}", f"=TB!F{r}", f"=TB!G{r}-TB!H{r}",
                                       f"=TB!I{r}-TB!J{r}", f"=TB!K{r}", f"=TB!L{r}", f"=TB!M{r}-TB!N{r}"]):
            put(ws, f"{col}{r}", f, F_LINK, NUM2)
        put(ws, f"N{r}", f"=SUMIFS(AJE!$E:$E,AJE!$B:$B,$A{r})-SUMIFS(AJE!$F:$F,AJE!$B:$B,$A{r})", F_LINK, NUM2)
        put(ws, f"O{r}", f"=M{r}+N{r}", F_BOLD, NUM2)
        put(ws, f"P{r}", f"=ROUND(J{r}+K{r}-L{r}-M{r},2)", fmt=NUM2)
        put(ws, f"Q{r}", f"=ROUND(F{r}+G{r}-H{r}-I{r},2)", fmt=NUM2)
        put(ws, f"R{r}", f'=SUMIFS(GL!$J:$J,GL!$D:$D,$A{r},GL!$A:$A,"CY")', F_LINK, NUM2)
        put(ws, f"S{r}", f"=ROUND(K{r}-L{r}-R{r},2)", fmt=NUM2)
        put(ws, f"T{r}", f'=SUMIFS(GL!$J:$J,GL!$D:$D,$A{r},GL!$A:$A,"PY")', F_LINK, NUM2)
        put(ws, f"U{r}", f"=ROUND(G{r}-H{r}-T{r},2)", fmt=NUM2)
    last = FIRST + len(mapping) - 1
    put(ws, f"B{last + 1}", "Total", F_BOLD)
    for col in "FGHIJKLMNOPQRSTU":
        put(ws, f"{col}{last + 1}", f"=SUM({col}{FIRST}:{col}{last})", F_BOLD, NUM2, border=B_BOTTOM)
    ws.freeze_panes = "C6"
    ws.auto_filter.ref = f"A5:U{last}"
    book.tb_last = last


def build_tb(book, ws, mapping, tb_cy, tb_py):
    """Unadjusted trial balance as cleaned from the source files, in debit / credit columns."""
    sheet_title(ws, "Trial balance (TB) - unadjusted, as extracted and cleaned, in functional currency",
                "Source balances in debit and credit columns for both years. Adjustments are posted on AJE and carried to the ATB sheet.")
    widths(ws, [11, 46] + [15] * 12)
    for col, lab in (("C", "Prior year"), ("I", "Current year")):
        put(ws, f"{col}4", lab, F_BOLD)
    header_row(ws, 5, ["Account", "Account name"] + ["Opening Dr", "Opening Cr", "Movement Dr", "Movement Cr", "Closing Dr",
                                                    "Closing Cr"] * 2)
    cy = tb_cy.groupby("account")[["opening", "debit", "credit", "closing"]].sum()
    py = tb_py.groupby("account")[["opening", "debit", "credit", "closing"]].sum() if tb_py is not None else None
    for i, m in enumerate(mapping.itertuples()):
        r = FIRST + i
        put(ws, f"A{r}", m.account)
        put(ws, f"B{r}", m.account_name)
        for cols, frame in (("CDEFGH", py), ("IJKLMN", cy)):
            o, d, c, cl = (float(v) for v in (frame.loc[m.account].tolist()
                                              if frame is not None and m.account in frame.index else [0, 0, 0, 0]))
            vals = [max(o, 0), max(-o, 0), d, c, max(cl, 0), max(-cl, 0)]
            for col, v in zip(cols, vals):
                put(ws, f"{col}{r}", round(v, 2), F_INPUT, NUM2)
    last = FIRST + len(mapping) - 1
    put(ws, f"B{last + 1}", "Total", F_BOLD)
    for col in "CDEFGHIJKLMN":
        put(ws, f"{col}{last + 1}", f"=SUM({col}{FIRST}:{col}{last})", F_BOLD, NUM2, border=B_BOTTOM)
    put(ws, f"B{last + 2}", "Difference (debits less credits) - must be zero", F_GREY)
    for a_, b_ in (("C", "D"), ("E", "F"), ("G", "H"), ("I", "J"), ("K", "L"), ("M", "N")):
        put(ws, f"{b_}{last + 2}", f"={a_}{last + 1}-{b_}{last + 1}", F_GREY, NUM2)
    ws.freeze_panes = "C6"
    ws.auto_filter.ref = f"A5:N{last}"


def build_aje(book, ws, adjustments=(), rows=60):
    sheet_title(ws, "Adjusting and reclassification journal entries (current year)",
                "Enter the account code, description and amount (functional currency) in the blue area, one line per account. "
                "Each journal must balance; adjustments flow to ATB column N and from there to every statement and note.")
    widths(ws, [9, 12, 46, 60, 16, 16])
    put(ws, "D3", "Total debits / credits (difference must be zero)", F_BOLD)
    put(ws, "E3", f"=SUM(E{FIRST}:E{FIRST + rows - 1})", F_BOLD, NUM2)
    put(ws, "F3", f"=SUM(F{FIRST}:F{FIRST + rows - 1})", F_BOLD, NUM2)
    header_row(ws, 5, ["AJE no.", "Account", "Account name", "Description", "Debit", "Credit"], height=20)
    for r in range(FIRST, FIRST + rows):
        for col in "ABDEF":
            c = ws[f"{col}{r}"]
            c.font = F_INPUT
            if col in "EF":
                c.number_format = NUM2
            if col == "B":
                c.number_format = "@"
        put(ws, f"C{r}", f'=IF($B{r}="","",IFERROR(INDEX(Mapping!$B:$B,MATCH($B{r}&"",Mapping!$A:$A,0)),"ACCOUNT NOT IN TB"))', F_LINK)
    for i, a in enumerate(adjustments):
        r = FIRST + i
        ws[f"A{r}"].value, ws[f"B{r}"].value, ws[f"D{r}"].value = a.get("no"), str(a["account"]), a.get("description", "")
        ws[f"E{r}"].value, ws[f"F{r}"].value = a.get("debit") or None, a.get("credit") or None
    ws.freeze_panes = "A6"


def build_gl(book, ws, gl_frames):
    sheet_title(ws, "General ledger - cleaned journal lines for both years",
                "Column L (movement type) is pre-filled from the other accounts in each journal and drives roll-forward notes and the cash flow statement. Override in blue where needed.")
    widths(ws, [6, 13, 13, 10, 40, 48, 10, 15, 15, 15, 14, 12])
    header_row(ws, 5, ["FY", "Date", "Journal", "Account", "Account name", "Description", "Reference", "Debit",
                       "Credit", "Net (Dr +)", "AFS code", "Movement type"])
    r = FIRST - 1
    for fy, gl in gl_frames:
        if gl is None:
            continue
        for g in gl.itertuples():
            r += 1
            ws.cell(row=r, column=1, value=fy).font = F_NORMAL
            c = ws.cell(row=r, column=2, value=g.date if isinstance(g.date, date) else None)
            c.font, c.number_format = F_NORMAL, DATE
            for j, v in enumerate([g.journal, g.account, g.account_name, g.description, g.reference], start=3):
                ws.cell(row=r, column=j, value=v).font = F_NORMAL
            for j, v in ((8, g.debit), (9, g.credit)):
                c = ws.cell(row=r, column=j, value=v)
                c.font, c.number_format = F_INPUT, NUM2
            c = ws.cell(row=r, column=10, value=f"=H{r}-I{r}")
            c.font, c.number_format = F_NORMAL, NUM2
            c = ws.cell(row=r, column=11, value=f'=IFERROR(INDEX(Mapping!$D:$D,MATCH($D{r},Mapping!$A:$A,0)),"UNMAPPED")')
            c.font = F_LINK
            ws.cell(row=r, column=12, value=g.mvt or None).font = F_INPUT
    last = max(r, FIRST)
    dv = DataValidation(type="list", formula1=f"Lists!$D${FIRST}:$D${FIRST + len(MVT_TYPES) - 1}", allow_blank=True)
    dv.add(f"L{FIRST}:L{last}")
    ws.add_data_validation(dv)
    ws.freeze_panes = "A6"
    ws.auto_filter.ref = f"A5:L{last}"
    book.gl_last = last
