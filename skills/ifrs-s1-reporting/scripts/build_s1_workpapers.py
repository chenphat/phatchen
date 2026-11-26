#!/usr/bin/env python3
"""Build the IFRS S1 working-paper workbook (S1-00 to S1-14).

Usage:
    python build_s1_workpapers.py --entity "Mekong Garment Co." --year-end 2026-12-31 \
        --currency "KHR million" --out S1_Working_Papers.xlsx [--no-examples]

Every sheet has: a header block (entity, period, WP ref, IFRS anchor, preparer/reviewer sign-off),
a legend, yellow input cells, and one illustrative example row (grey italic) that users overwrite.
Formulas are live; run a recalculation (e.g. LibreOffice or the xlsx skill's recalc.py) before
sharing so cached values exist.
"""
import argparse
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

FONT = "Arial"
F_BASE = Font(name=FONT, size=10)
F_BOLD = Font(name=FONT, size=10, bold=True)
F_HEAD = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_TITLE = Font(name=FONT, size=14, bold=True, color="1F3864")
F_EX = Font(name=FONT, size=10, italic=True, color="7F7F7F")
F_NOTE = Font(name=FONT, size=9, italic=True, color="404040")
FILL_HEAD = PatternFill("solid", fgColor="1F3864")
FILL_INPUT = PatternFill("solid", fgColor="FFF2CC")
FILL_SUB = PatternFill("solid", fgColor="D9E1F2")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")

EXAMPLES = True


def header(wb, ref, title, anchor, args, purpose):
    ws = wb.create_sheet(ref)
    ws.sheet_view.showGridLines = False
    ws["A1"] = f"{ref}  {title}"
    ws["A1"].font = F_TITLE
    rows = [
        ("Entity", args.entity), ("Period ended", args.year_end), ("Presentation unit", args.currency),
        ("IFRS anchor", anchor), ("Purpose", purpose),
    ]
    for i, (k, v) in enumerate(rows, start=2):
        ws.cell(i, 1, k).font = F_BOLD
        c = ws.cell(i, 2, v)
        c.font = F_BASE
        c.alignment = Alignment(wrap_text=False)
    # sign-off block
    ws["F2"], ws["F3"], ws["F4"] = "Prepared by / date", "Reviewed by / date", "Approved by / date"
    for r in (2, 3, 4):
        ws.cell(r, 6).font = F_BOLD
        for col in (7, 8):
            ws.cell(r, col).fill = FILL_INPUT
            ws.cell(r, col).border = BORDER
    ws["A7"] = ("Legend: yellow cells = inputs; white cells with formulas = calculated (do not overwrite); "
                "grey italic row = illustrative example — overwrite or delete.")
    ws["A7"].font = F_NOTE
    return ws


def table(ws, start_row, headers, rows, widths=None, input_cols=None, example_rows=1, height=None):
    """Write a table. rows: list of lists (values or formulas with {r} placeholder for row number)."""
    for j, h in enumerate(headers, start=1):
        c = ws.cell(start_row, j, h)
        c.font, c.fill, c.alignment, c.border = F_HEAD, FILL_HEAD, Alignment(wrap_text=True, vertical="center"), BORDER
    ws.row_dimensions[start_row].height = 30
    input_cols = set(input_cols or [])
    for i, row in enumerate(rows):
        r = start_row + 1 + i
        for j, v in enumerate(row, start=1):
            if isinstance(v, str) and "{r}" in v:
                v = v.replace("{r}", str(r))
            c = ws.cell(r, j, v)
            c.border, c.alignment = BORDER, WRAP
            is_example = EXAMPLES and i < example_rows and example_rows > 0
            c.font = F_EX if is_example and not (isinstance(v, str) and v.startswith("=")) else F_BASE
            if j in input_cols and not (isinstance(v, str) and str(v).startswith("=")):
                c.fill = FILL_INPUT
        if height:
            ws.row_dimensions[r].height = height
    if widths:
        for j, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(j)].width = w
    return start_row + len(rows)


def blank_rows(n, ncols, formulas=None):
    formulas = formulas or {}
    return [[formulas.get(j, None) for j in range(1, ncols + 1)] for _ in range(n)]


def add_list_validation(ws, rng, options):
    dv = DataValidation(type="list", formula1='"' + ",".join(options) + '"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(rng)


# --------------------------------------------------------------------------------------------
def s1_00(wb, args):
    ws = header(wb, "S1-00", "Index, engagement details and timetable", "S1.64; S1.E4", args,
                "Control sheet: WP status, owners, timetable aligned to the financial close.")
    wps = [
        ("S1-01", "Readiness and gap analysis (incl. SERC Prakas 034 map)", "S1.72, E3–E6"),
        ("S1-02", "Reporting entity, boundary, period and time horizons", "S1.20, 30, 65–66, B38"),
        ("S1-03", "SRRO register and assessment", "S1.30, B1–B12"),
        ("S1-04", "Material information assessment", "S1.17–19, B13–B37"),
        ("S1-05", "Governance", "S1.26–27"),
        ("S1-06", "Strategy: five zones, financial effects, resilience", "S1.28–42"),
        ("S1-07", "Risk management", "S1.43–44"),
        ("S1-08", "Metrics (KPI cards) and targets register", "S1.45–53"),
        ("S1-09", "Connectivity to the financial statements", "S1.21–24, B39–B44"),
        ("S1-10", "Judgements, measurement uncertainty and errors", "S1.74–86"),
        ("S1-11", "IFRS S1 disclosure checklist", "Whole standard"),
        ("S1-12", "ICSR controls matrix", "COSO ICSR; S1.77–82"),
        ("S1-13", "Assurance evidence index", "ISSA 5000"),
        ("S1-14", "Review notes, exception log, ICSR maturity scorecard", "ISA 265 concepts"),
        ("GHG/SA/CR", "Climate working papers (ifrs-s2-reporting workbook)", "IFRS S2"),
    ]
    rows = [[w, t, a, None, None, None, None, None] for w, t, a in wps]
    end = table(ws, 9, ["WP", "Title", "IFRS anchor", "Owner", "Preparer", "Reviewer", "Status", "Comment"],
                rows, [12, 52, 22, 16, 14, 14, 14, 30], input_cols=[4, 5, 6, 7, 8], example_rows=0)
    add_list_validation(ws, f"G10:G{end}", ["Not started", "In progress", "For review", "Reviewed", "Final"])
    for i, (w, _, _) in enumerate(wps):
        if w in wb.sheetnames or w.startswith("S1-"):
            ws.cell(10 + i, 1).hyperlink = f"#'{w}'!A1"
    r = end + 2
    ws.cell(r, 1, "Completion").font = F_BOLD
    ws.cell(r, 2, f'=COUNTIF(G10:G{end},"Final")/COUNTA(A10:A{end})').number_format = "0%"
    r += 2
    ws.cell(r, 1, "Timetable (align to the financial close; S1.64 same-time publication)").font = F_BOLD
    tt = [
        ["Kick-off and scoping agreed with CFO, audit committee, assurer", "T-16 weeks", None, None, None],
        ["SRRO and materiality assessment approved", "T-12 weeks", None, None, None],
        ["Data collection cut-off (first 9–11 months), then year-end true-up", "T-6 / T+2 weeks", None, None, None],
        ["Draft disclosures and connectivity review with FS team", "T+4 weeks", None, None, None],
        ["Internal review and exception closure", "T+6 weeks", None, None, None],
        ["Assurance fieldwork (ISSA 5000)", "T+6–9 weeks", None, None, None],
        ["Joint sign-off with financial statements; publication", "FS date", None, None, None],
    ]
    table(ws, r + 1, ["Milestone", "Target (T = year end)", "Planned date", "Actual date", "Owner"], tt,
          input_cols=[3, 4, 5], example_rows=0)


GAP_ROWS = [
    ("Statement of compliance / basis of preparation", "S1.72, E3–E6", "—"),
    ("Reporting entity same as FS; value chain defined", "S1.20, B2–B5", "1"),
    ("SRROs identified incl. opportunities", "S1.30(a)", "3"),
    ("Time horizons defined and linked to planning", "S1.30(b)–(c)", "3"),
    ("Business model and value chain effects; concentration", "S1.32", "1"),
    ("Strategy response, progress, trade-offs", "S1.33", "5"),
    ("Current financial effects", "S1.34–35", "—"),
    ("Anticipated financial effects", "S1.34–40", "—"),
    ("Resilience assessment", "S1.41–42", "—"),
    ("Board oversight, skills, frequency, remuneration link", "S1.27(a)", "6"),
    ("Management role, controls and procedures", "S1.27(b)", "6"),
    ("Risk management processes and ERM integration", "S1.44", "4"),
    ("Metrics required by IFRS SDS and SASB industry metrics", "S1.46–48", "7"),
    ("Entity-developed metric definitions", "S1.49–50", "7"),
    ("Targets with base period and milestones", "S1.51", "8"),
    ("Connected information to FS; consistent assumptions", "S1.21–23", "—"),
    ("Sources of guidance disclosed (SASB industry)", "S1.54–59", "—"),
    ("Location, cross-referencing", "S1.60–63, B45–B47", "—"),
    ("Same timing and period as FS", "S1.64–66", "—"),
    ("Comparatives", "S1.70–71, B49–B54", "—"),
    ("Judgements", "S1.74–76", "—"),
    ("Measurement uncertainty", "S1.77–82", "—"),
    ("Errors and restatements", "S1.83–86", "—"),
    ("Climate disclosures (IFRS S2) — see S2 workbook", "IFRS S2", "25–31"),
    ("Social metrics (training, injuries, gender…) — only if material to an SRRO", "S1.46(b)", "9–18"),
    ("Environmental metrics (water, waste, energy) — only if material", "S1.46(b)", "19–24, 27–29"),
]


def s1_01(wb, args):
    ws = header(wb, "S1-01", "Readiness and gap analysis", "S1.72; S1.E3–E6; SERC Prakas 034", args,
                "Compare current reporting to IFRS S1 requirements; plan remediation. Gap: Y = met, P = partial, N = not met.")
    rows = []
    for i, (req, para, prakas) in enumerate(GAP_ROWS):
        if i == 0 and EXAMPLES:
            rows.append([req, para, prakas, "Report says 'aligned with GRI and Prakas'; no ISSB statement", "N", "High",
                         "Decide year-1 reliefs (E3, E5); draft compliance statement", "CFO", "2027-01-31", "Open"])
        else:
            rows.append([req, para, prakas, None, None, None, None, None, None, None])
    end = table(ws, 9, ["Requirement", "IFRS para", "Prakas 034 item", "Current disclosure / source", "Gap (Y/P/N)",
                        "Severity", "Remediation action", "Owner", "Due", "Status"], rows,
                [42, 16, 12, 40, 10, 10, 40, 12, 12, 10], input_cols=[4, 5, 6, 7, 8, 9, 10])
    add_list_validation(ws, f"E10:E{end}", ["Y", "P", "N", "NA"])
    add_list_validation(ws, f"F10:F{end}", ["High", "Medium", "Low"])
    add_list_validation(ws, f"J10:J{end}", ["Open", "In progress", "Closed"])
    r = end + 2
    for k, (lab, crit) in enumerate([("Met (Y)", "Y"), ("Partial (P)", "P"), ("Not met (N)", "N")]):
        ws.cell(r + k, 1, lab).font = F_BOLD
        ws.cell(r + k, 2, f'=COUNTIF(E10:E{end},"{crit}")')
    ws.cell(r + 3, 1, "Readiness % (Y + 0.5×P) / assessed").font = F_BOLD
    c = ws.cell(r + 3, 2, f"=IFERROR((B{r}+0.5*B{r+1})/(B{r}+B{r+1}+B{r+2}),0)")
    c.number_format = "0%"
    ws.cell(r + 4, 1, "High-severity open gaps").font = F_BOLD
    ws.cell(r + 4, 2, f'=COUNTIFS(F10:F{end},"High",J10:J{end},"<>Closed")')
    ws.conditional_formatting.add(f"E10:E{end}", CellIsRule(operator="equal", formula=['"N"'], fill=PatternFill("solid", fgColor="F8CBAD")))
    ws.conditional_formatting.add(f"E10:E{end}", CellIsRule(operator="equal", formula=['"P"'], fill=PatternFill("solid", fgColor="FFE699")))
    ws.conditional_formatting.add(f"E10:E{end}", CellIsRule(operator="equal", formula=['"Y"'], fill=PatternFill("solid", fgColor="C6EFCE")))


def s1_02(wb, args):
    ws = header(wb, "S1-02", "Reporting entity, boundary, period and time horizons memo",
                "S1.20, S1.30(b)–(c), S1.65–66, S1.B2–B5, B38", args,
                "Document boundary judgements so sustainability information covers the same entity and period as the FS.")
    qs = [
        ["Reporting entity: list of consolidated subsidiaries agreed to FS consolidation schedule", "S1.20, B38",
         "Parent + 2 subsidiaries per FS note 1; all 100% included", "FS note 1; group structure chart"],
        ["Joint operations / JVs / associates: treatment and rationale", "S1.20 (SRG 3.2.4)", None, None],
        ["Leased assets (lessee/lessor): treatment", "S1.20 (SRG 3.2.3)", None, None],
        ["Upstream value chain: key suppliers, resources, geographies", "S1.B2–B5, App. A", None, None],
        ["Downstream value chain: customers, distributors, financed activities", "S1.B2–B5", None, None],
        ["Investments (incl. under fiduciary duty) — not excluded under ISSB", "S1.B5", None, None],
        ["Reporting period = FS period (incl. 52/53-week)", "S1.64–65", None, None],
        ["Change in reporting period? Disclose period, reason, non-comparability", "S1.66", None, None],
        ["Acquisitions/disposals: included from/to consolidation date; disaggregation needed?", "S1.B26–B27; TIG Jun 2024", None, None],
        ["Presentation currency consistent with FS", "S1.24", None, None],
    ]
    end = table(ws, 9, ["Boundary question", "Para", "Conclusion", "Evidence / WP ref"], qs, [60, 22, 50, 30],
                input_cols=[3, 4], height=30)
    r = end + 2
    ws.cell(r, 1, "Time horizons (entity-specific; link to planning and capital allocation)").font = F_BOLD
    th = [["Short term", "0–1 years", "Annual budget", "Working capital, licences"],
          ["Medium term", "1–5 years", "3-year business plan; loan tenors", "Capex programme, buyer contracts"],
          ["Long term", ">5 years", "Asset lives; NDC 2035 / LTS 2050", "Plant life, land concessions"]]
    table(ws, r + 1, ["Horizon", "Period", "Linked planning cycle", "Decisions informed"], th, input_cols=[2, 3, 4],
          example_rows=3)


SRRO_EX = [
    ["R01", "Buyer decarbonisation requirements (EU/US brands)", "Risk", "Climate – transition", "External",
     "Downstream", "Medium", "Cash flows; access to finance", 4, 4],
    ["R03", "Grid outages from hydropower shortfall in drought", "Risk", "Climate – physical", "Dependency",
     "Own operations", "Short", "Cash flows", 4, 3],
    ["O01", "Rooftop solar PPA lowers energy cost", "Opportunity", "Climate – transition", "Dependency",
     "Own operations", "Medium", "Cash flows", 4, 2],
]


def s1_03(wb, args):
    ws = header(wb, "S1-03", "SRRO register and assessment", "S1.30; S1.54–58; S1.B1–B12", args,
                "Identify and assess sustainability-related risks and opportunities by likelihood × magnitude per time horizon. "
                "Impact materiality (GRI) is recorded separately and does not drive the ISSB conclusion.")
    headers = ["ID", "SRRO description", "Risk / Opportunity", "Category", "Source (dependency / impact / external)",
               "Value-chain location", "Horizon", "Channel (cash flow / access to finance / cost of capital)",
               "Likelihood 1–5", "Magnitude 1–5", "Score L×M", "Band", "SASB / S2 topic ref",
               "Impact materiality (GRI – memo)", "Response (incl. 'monitor only')", "Owner", "Evidence"]
    score = "=IF(AND(ISNUMBER(I{r}),ISNUMBER(J{r})),I{r}*J{r},\"\")"
    band = '=IF(K{r}="","",IF(K{r}>=12,"Include – priority",IF(K{r}>=6,"Include – monitor","Exclude – document")))'
    rows = []
    for i in range(25):
        base = SRRO_EX[i] if (EXAMPLES and i < len(SRRO_EX)) else [None] * 10
        extra = ["Apparel & footwear: Raw materials sourcing", "Medium", "2 MW solar PPA; supplier data programme", "CFO", "MA-02"] if (EXAMPLES and i < len(SRRO_EX)) else [None] * 5
        rows.append(base + [score, band] + extra)
    end = table(ws, 9, headers, rows, [7, 40, 12, 18, 18, 16, 10, 24, 10, 10, 9, 18, 24, 16, 34, 12, 12],
                input_cols=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 13, 14, 15, 16, 17], example_rows=len(SRRO_EX))
    add_list_validation(ws, f"C10:C{end}", ["Risk", "Opportunity"])
    add_list_validation(ws, f"G10:G{end}", ["Short", "Medium", "Long"])
    dv = DataValidation(type="whole", operator="between", formula1="1", formula2="5", allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"I10:J{end}")
    ws.conditional_formatting.add(f"K10:K{end}", CellIsRule(operator="greaterThanOrEqual", formula=["12"], fill=PatternFill("solid", fgColor="F8CBAD")))
    ws.conditional_formatting.add(f"K10:K{end}", CellIsRule(operator="between", formula=["6", "11"], fill=PatternFill("solid", fgColor="FFE699")))
    # Heat map
    r = end + 2
    ws.cell(r, 1, "Heat map — count of SRROs (rows = likelihood, columns = magnitude)").font = F_BOLD
    ws.cell(r + 1, 1, "L \\ M").font = F_BOLD
    for m in range(1, 6):
        ws.cell(r + 1, 1 + m, m).font = F_BOLD
    for li, l in enumerate(range(5, 0, -1)):
        ws.cell(r + 2 + li, 1, l).font = F_BOLD
        for m in range(1, 6):
            c = ws.cell(r + 2 + li, 1 + m, f"=COUNTIFS($I$10:$I${end},{l},$J$10:$J${end},{m})")
            c.border = BORDER
            c.alignment = Alignment(horizontal="center")
            prod = l * m
            c.fill = PatternFill("solid", fgColor="F8CBAD" if prod >= 12 else ("FFE699" if prod >= 6 else "C6EFCE"))
    r2 = r + 8
    notes = ["Scales (calibrate to FS materiality and risk appetite; record the basis):",
             "Likelihood: 1 remote <5%; 2 unlikely 5–25%; 3 possible 25–50%; 4 likely 50–80%; 5 almost certain >80%.",
             "Magnitude: 1 <0.5% revenue/equity; 2 0.5–1%; 3 1–3% or changes cost of funds; 4 3–5% or affects market/licence/financing access; 5 >5% or threatens business model.",
             "Consider completed mitigation only; a risk may be 'reasonably expected' from users' view even if management disagrees (SRG 4.6).",
             "Excluded SRROs: reconsider aggregate exposure of similar small risks (S1.B23). Sources: SASB topics 'shall refer to and consider' (S1.55(a))."]
    for k, n in enumerate(notes):
        ws.cell(r2 + k, 1, n).font = F_NOTE


def s1_04(wb, args):
    ws = header(wb, "S1-04", "Material information assessment", "S1.17–19; S1.B13–B37", args,
                "For each included SRRO decide which information is material: would omitting, misstating or obscuring it "
                "reasonably influence primary users' decisions?")
    headers = ["SRRO ID", "Pillar", "Candidate information item", "Source (IFRS S2 / SASB / entity)",
               "Quantitative factors", "Qualitative factors", "Aggregation (B23) / obscuring (B26–B27)",
               "Material? (Y/N)", "Rationale", "Disclosure location / note ref"]
    rows = [["R01", "Metrics & targets", "Scope 3 Cat 1 purchased fabric emissions", "IFRS S2.29(a)(vi)",
             "~75% of inventory", "Buyer contracts require supplier data", "—", "Y",
             "Key to buyer retention and green-loan terms", "Climate note §4"]] if EXAMPLES else []
    rows += blank_rows(30, 10)
    end = table(ws, 9, headers, rows, [9, 16, 36, 20, 20, 24, 22, 10, 34, 20],
                input_cols=list(range(1, 11)))
    add_list_validation(ws, f"B10:B{end}", ["Governance", "Strategy", "Risk management", "Metrics & targets"])
    add_list_validation(ws, f"H10:H{end}", ["Y", "N"])


def checklist_sheet(wb, args, ref, title, anchor, items):
    ws = header(wb, ref, title, anchor, args, "Record response, evidence and status for each requirement.")
    rows = [[a, b, c, None, None, None] for a, b, c in items]
    end = table(ws, 9, ["#", "Requirement", "Para", "Entity response / draft wording", "Evidence / WP ref", "Status"],
                rows, [6, 60, 14, 50, 22, 10], input_cols=[4, 5, 6], example_rows=0, height=30)
    add_list_validation(ws, f"F10:F{end}", ["Y", "P", "N", "NA", "R"])
    return ws, end


GOV = [("G1", "Body(ies)/individual(s) responsible for oversight of SRROs; reflected in terms of reference, mandates, policies", "27(a)(i)"),
       ("G2", "How the body ensures appropriate skills and competencies", "27(a)(ii)"),
       ("G3", "How and how often the body is informed about SRROs", "27(a)(iii)"),
       ("G4", "How SRROs are considered in strategy, major transactions and risk management, incl. trade-offs", "27(a)(iv)"),
       ("G5", "Oversight of target-setting and progress; metrics in remuneration policies", "27(a)(v)"),
       ("G6", "Management's role: delegated position/committee and how overseen", "27(b)(i)"),
       ("G7", "Controls and procedures supporting oversight; integration with other functions", "27(b)(ii)"),
       ("G8", "Integrated oversight disclosed without unnecessary duplication", "B42(b)")]

RISK = [("R1", "Inputs and parameters (data sources, scope of operations)", "44(a)(i)"),
        ("R2", "Whether and how scenario analysis informs identification", "44(a)(ii)"),
        ("R3", "How nature, likelihood and magnitude are assessed", "44(a)(iii)"),
        ("R4", "Prioritisation relative to other risks", "44(a)(iv)"),
        ("R5", "Monitoring", "44(a)(v)"),
        ("R6", "Changes compared with prior period", "44(a)(vi)"),
        ("R7", "Processes for opportunities", "44(b)"),
        ("R8", "Integration into overall risk management (ERM)", "44(c)"),
        ("R9", "Greenwashing and fraud-risk controls over disclosures (good practice; COSO/ACFE)", "—")]


def s1_06(wb, args):
    ws = header(wb, "S1-06", "Strategy: five zones, financial effects and resilience", "S1.28–42", args,
                "Build the strategy disclosure; quantify current and anticipated financial effects; test resilience.")
    zones = [["1 Why — purpose and non-negotiables", "S1.32 (context)", None, None],
             ["2 Where — markets, SRROs by horizon, concentration", "S1.30–32", None, None],
             ["3 What — priorities, funding, trade-offs", "S1.33(a),(c); S1.34–35", None, None],
             ["4 Who — owners (links to governance)", "S1.27(b); S1.33(a)", None, None],
             ["5 When — dates, progress tracking", "S1.33(b); S1.51", None, None],
             ["Added — Resilience: does strategy survive the bad scenario?", "S1.41–42", None, None]]
    end = table(ws, 9, ["Zone", "IFRS S1", "Entity disclosure points", "Evidence (WP)"], zones, [44, 22, 60, 24],
                input_cols=[3, 4], example_rows=0, height=40)
    r = end + 2
    ws.cell(r, 1, f"Financial effects ({args.currency})").font = F_BOLD
    fh = ["SRRO ID", "FS line item / note", "Current-period effect", "Significant risk of material adjustment next year? (Y/N)",
          "Anticipated effect – short", "Anticipated – medium", "Anticipated – long", "Quantitative / qualitative / range",
          "If not quantitative: reason (S1.38 not separately identifiable / high uncertainty; S1.39 skills)", "Assumptions consistent with FS? (S1.23)"]
    rows = [["R03", "Cost of sales – diesel for backup generation (Note 21)", 120, "N", 150, 250, None, "Range",
             "Long term: uncertainty too high (S1.38)", "Y – same FX and price deck"]] if EXAMPLES else []
    rows += blank_rows(15, 10)
    e2 = table(ws, r + 1, fh, rows, None, input_cols=list(range(1, 11)))
    for rr in range(r + 2, e2 + 1):
        for cc in (3, 5, 6, 7):
            ws.cell(rr, cc).number_format = "#,##0;(#,##0);-"
    ws.cell(e2 + 1, 2, "Total quantified").font = F_BOLD
    for cc in (3, 5, 6, 7):
        col = get_column_letter(cc)
        ws.cell(e2 + 1, cc, f"=SUM({col}{r+2}:{col}{e2})").number_format = "#,##0;(#,##0);-"
    r3 = e2 + 3
    ws.cell(r3, 1, "Resilience assessment (S1.41–42)").font = F_BOLD
    res = [["Approach (qualitative / quantitative; scenario analysis for climate — see SA1–SA9)", None],
           ["Time horizon(s) and date of assessment", None],
           ["Results and implications for strategy and business model", None],
           ["Capacity to adjust (financial headroom, asset redeployment)", None]]
    table(ws, r3 + 1, ["Item", "Conclusion / reference"], res, input_cols=[2], example_rows=0, height=30)


def s1_08(wb, args):
    ws = header(wb, "S1-08", "Metrics (KPI cards) and targets register", "S1.45–53; S1.B52–B53", args,
                "Define each metric with nine attributes and record targets with all S1.51 elements.")
    kh = ["KPI ID", "Metric (definition)", "Unit", "Reporting period", "Boundary", "Owner", "Source system",
          "Method / factor", "Frequency", "Evidence ref", "Source of metric (IFRS SDS / SASB / entity)", "SRRO ID",
          "Current year", "Prior year", "Change %", "Data quality score (EDQA /15)"]
    chg = '=IF(AND(ISNUMBER(M{r}),ISNUMBER(N{r}),N{r}<>0),M{r}/N{r}-1,"")'
    rows = [["DA21", "Lost-time injury frequency rate per 1,000,000 hours worked", "rate", "FY", "All sites (= FS)",
             "HR Director", "Accident register; payroll hours", "LTIs × 1,000,000 / hours", "Monthly",
             "SOC-LTI-2026-001", "SASB (entity-adapted)", "R09", 2.1, 2.6, chg, 14]] if EXAMPLES else []
    rows += [[None] * 14 + [chg, None] for _ in range(25)]
    end = table(ws, 9, kh, rows, [8, 36, 9, 10, 16, 14, 18, 22, 10, 16, 18, 8, 11, 11, 9, 12],
                input_cols=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 16])
    for rr in range(10, end + 1):
        ws.cell(rr, 15).number_format = "0.0%"
    r = end + 2
    ws.cell(r, 1, "Targets register (S1.51)").font = F_BOLD
    th = ["Target ID", "Metric", "Objective", "Part of entity", "Period", "Base period", "Base value",
          "Milestones / interim", "Target value", "Current value", "Progress to target", "Revisions & explanation",
          "Ambition or result? (label)"]
    prog = '=IF(AND(ISNUMBER(G{r}),ISNUMBER(I{r}),ISNUMBER(J{r}),I{r}<>G{r}),(J{r}-G{r})/(I{r}-G{r}),"")'
    rows = [["T01", "LTIFR", "Reduce injuries", "All factories", "2026–2028", "FY2025", 2.6, "2.0 by FY2027", 1.5, 2.1,
             prog, "None", "Ambition (target); current value = result"]] if EXAMPLES else []
    rows += [[None] * 10 + [prog, None, None] for _ in range(12)]
    e2 = table(ws, r + 1, th, rows, None, input_cols=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 13])
    for rr in range(r + 2, e2 + 1):
        ws.cell(rr, 11).number_format = "0%"


def s1_09(wb, args):
    ws = header(wb, "S1-09", "Connectivity to the financial statements", "S1.21–24; S1.B39–B44", args,
                "Show how each SRRO links to FS items, judgements and assumptions; confirm consistency or explain differences.")
    h = ["SRRO ID", "FS line item / note", "IFRS / CIFRS standard", "FS judgement or assumption", "Assumption used in sustainability disclosure",
         "Consistent? (Y/N)", "Explanation of difference (S1.23)", "Cross-reference wording"]
    rows = [["R11", "Bearer plants (Note 12)", "IAS 16 / IAS 36", "Drought frequency in value-in-use cash flows",
             "Drought scenario 3°C (SA4)", "N", "FS uses 5-year budget; scenario covers 2040", "See Note 12 and Climate §3"]] if EXAMPLES else []
    rows += blank_rows(20, 8)
    end = table(ws, 9, h, rows, [9, 26, 16, 30, 30, 10, 34, 28], input_cols=list(range(1, 9)))
    add_list_validation(ws, f"F10:F{end}", ["Y", "N"])
    r = end + 2
    prompts = ["Prompts: impairment (IAS 36) • useful lives/residual values (IAS 16/38) • provisions and contingent liabilities (IAS 37)",
               "• biological assets (IAS 41) • ECL and SICR (IFRS 9) • fair value inputs (IFRS 13) • going concern (IAS 1) • deferred tax recoverability (IAS 12)",
               "• green bonds/sustainability-linked loans (IFRS 9 classification) • emission allowances/carbon credits (accounting policy) • segment information (IFRS 8)."]
    for k, p in enumerate(prompts):
        ws.cell(r + k, 1, p).font = F_NOTE


def s1_10(wb, args):
    ws = header(wb, "S1-10", "Judgements, measurement uncertainty and errors log", "S1.74–86; S1.B55–B59", args,
                "Log significant judgements, estimates with high uncertainty, and prior-period errors.")
    h = ["Ref", "Type (Judgement / Uncertainty / Error)", "Area / disclosure", "Description", "Para",
         "Assumptions / sources of uncertainty", "Sensitivity or range", "Disclosure required? (Y/N)", "Disclosure wording / ref", "Approved by"]
    rows = [["J01", "Uncertainty", "Scope 3 Cat 1", "Spend-based factor for purchased fabric", "S1.77–82",
             "Generic EXIOBASE factor; supplier data for 20% of spend", "±30%", "Y", "Climate note §4.3", "CFO"]] if EXAMPLES else []
    rows += blank_rows(20, 10)
    end = table(ws, 9, h, rows, [7, 16, 18, 36, 10, 34, 14, 10, 24, 12], input_cols=list(range(1, 11)))
    add_list_validation(ws, f"B10:B{end}", ["Judgement", "Uncertainty", "Error"])


CHECK = [
    ("CF1", "Fair presentation; additional information where needed", "11–16"),
    ("CF2", "Material information; not obscured", "17–19, B13–B37"),
    ("CF3", "Reporting entity same as FS", "20"),
    ("CF4", "Connected information across pillars and to FS", "21–22, B39–B44"),
    ("CF5", "Consistent data and assumptions with FS; differences explained", "23"),
    ("CF6", "Currency consistent with FS", "24"),
] + [(a, b, "S1." + c) for a, b, c in GOV] + [
    ("S1", "SRROs described", "30(a)"), ("S2", "Time horizons of effects", "30(b)"),
    ("S3", "Definition of short/medium/long term and link to planning", "30(c)"),
    ("S4", "Effects on business model and value chain", "32(a)"), ("S5", "Concentration in business model/value chain", "32(b)"),
    ("S6", "Response in strategy and decision-making", "33(a)"), ("S7", "Progress against prior plans", "33(b)"),
    ("S8", "Trade-offs", "33(c)"), ("S9", "Current financial effects", "34–35"),
    ("S10", "Significant risk of material adjustment within next annual period", "34–35"),
    ("S11", "Anticipated financial position, performance and cash flows", "34–35"),
    ("S12", "Approach: reasonable & supportable info; commensurate with skills", "37"),
    ("S13", "Quantitative omitted only if not separately identifiable / high uncertainty", "38"),
    ("S14", "Anticipated quantitative omitted for lack of skills/resources", "39"),
    ("S15", "Explanation and qualitative info when quantitative not provided", "40"),
    ("S16", "Resilience assessment", "41–42"),
] + [(a, b, "S1." + c) for a, b, c in RISK[:8]] + [
    ("M1", "Metrics required by applicable IFRS SDS", "46(a)"), ("M2", "Metrics used to measure/monitor SRROs and targets", "46(b)"),
    ("M3", "Industry metrics (SASB)", "47–48"), ("M4", "Source and definition of other/entity-developed metrics", "49–50"),
    ("M5", "Targets: metric, objective, scope, period, base, milestones, performance, revisions", "51"),
    ("M6", "Consistent metric definitions; redefinitions explained", "50, B52"),
    ("GR1", "Sources of guidance considered (SASB 'refer to and consider')", "54–58"),
    ("GR2", "Sources of guidance disclosed incl. SASB industry", "59"),
    ("GR3", "Location in general purpose financial reports; clearly identifiable", "60–62"),
    ("GR4", "Cross-references meet conditions", "63, B45–B47"),
    ("GR5", "Same timing and period as FS", "64–65"), ("GR6", "Change in period disclosures", "66"),
    ("GR7", "Events after reporting period", "67–68"), ("GR8", "Comparative information", "70–71, B49–B54"),
    ("GR9", "Explicit and unreserved statement of compliance (only if fully compliant)", "72"),
    ("GR10", "Information prohibited by law", "73"), ("GR11", "Commercially sensitive opportunity info", "B34–B37"),
    ("J1", "Significant judgements", "74–76"), ("J2", "Measurement uncertainty", "77–82"), ("J3", "Errors and restatement", "83–86"),
    ("T1", "Transition reliefs used and disclosed (comparatives, timing, climate-first)", "E3–E6"),
]


def s1_11(wb, args):
    ws, end = checklist_sheet(wb, args, "S1-11", "IFRS S1 disclosure checklist", "IFRS S1 (June 2023)", CHECK)
    r = end + 2
    for k, s in enumerate(["Y", "P", "N", "NA", "R"]):
        ws.cell(r + k, 2, f"Count {s}").font = F_BOLD
        ws.cell(r + k, 3, f'=COUNTIF(F10:F{end},"{s}")')
    ws.cell(r + 5, 2, "Can an explicit and unreserved compliance statement be made? (no N or P)").font = F_BOLD
    ws.cell(r + 5, 3, f'=IF(COUNTA(F10:F{end})<ROWS(F10:F{end}),"Incomplete",IF(C{r+1}+C{r+2}=0,"Yes","No"))')


def s1_12(wb, args):
    ws = header(wb, "S1-12", "ICSR controls matrix", "COSO 2013 / COSO ICSR (2023); S1.77–82", args,
                "Document key controls over each disclosed metric: maker-checker-approver, frequency, evidence, test result.")
    h = ["Control ID", "KPI / disclosure", "Risk addressed (C/A/CO/V/P)", "Control description", "Type (prevent/detect; manual/IT)",
         "COSO component", "Maker", "Checker", "Approver", "Frequency", "Evidence", "Design OK?", "Operating test result", "Deficiency (none/design/operating; significant?)"]
    rows = [["C-ELEC-01", "Scope 2 electricity kWh", "C, A, CO", "Monthly reconciliation of EDC bills to meter readings and GL electricity expense",
             "Detect; manual", "Control activities", "Facilities officer", "Finance analyst", "CFO", "Monthly",
             "ENV-ELEC-2026-003", "Y", "1 exception: Dec bill accrued late", "Operating – not significant"]] if EXAMPLES else []
    rows += blank_rows(25, 14)
    end = table(ws, 9, h, rows, [11, 20, 12, 40, 14, 16, 14, 14, 12, 10, 18, 9, 24, 22], input_cols=list(range(1, 15)))
    add_list_validation(ws, f"F10:F{end}", ["Control environment", "Risk assessment", "Control activities", "Information & communication", "Monitoring"])
    ws.cell(end + 2, 1, "Assertions: C = completeness, A = accuracy, CO = cut-off, V = valuation/measurement, P = presentation/classification. "
                        "Segregation: maker ≠ checker ≠ approver.").font = F_NOTE


def s1_13(wb, args):
    ws = header(wb, "S1-13", "Assurance evidence index", "ISSA 5000", args,
                "Cross-reference every disclosed figure and key statement to its working paper and source evidence.")
    h = ["Disclosure ref (page/para)", "Figure / statement", "Value", "Unit", "IFRS para", "WP ref", "Source evidence",
         "Reconciled to source / GL?", "Reviewed by", "In assurance scope? (Limited/Reasonable/No)", "Open issue"]
    rows = [["Climate note §4, table 1", "Scope 2 location-based emissions", 70.5, "tCO2e", "S2.29(a)(i),(v)", "GHG-05",
             "EDC invoices Jan–Dec; ENV-ELEC-2026-001", "Y", "Reviewer A", "Limited", "None"]] if EXAMPLES else []
    rows += blank_rows(40, 11)
    end = table(ws, 9, h, rows, [20, 34, 12, 9, 16, 10, 34, 12, 12, 14, 18], input_cols=list(range(1, 12)))
    add_list_validation(ws, f"J10:J{end}", ["Limited", "Reasonable", "No"])
    r = end + 2
    ws.cell(r, 1, "Figures indexed").font = F_BOLD
    ws.cell(r, 2, f"=COUNTA(B10:B{end})")
    ws.cell(r + 1, 1, "Not reconciled").font = F_BOLD
    ws.cell(r + 1, 2, f'=COUNTIF(H10:H{end},"N")')


def s1_14(wb, args):
    ws = header(wb, "S1-14", "Review notes, exception log and ICSR maturity scorecard", "ISA 265 concepts; ISSA 5000", args,
                "Record review points and exceptions to closure; score assurance readiness.")
    h = ["No.", "WP ref", "Review point / exception", "Raised by", "Date", "Impact (metric / disclosure)",
         "Response / action", "Owner", "Cleared by", "Status"]
    rows = [[1, "GHG-05", "Grid factor copied from prior year (0.50); current factor 0.5878 not applied", "Reviewer A", "2027-02-10",
             "Scope 2 understated 10.5 tCO2e (17.6%)", "Recalculate; add factor register control", "Sust. lead", None, "Open"]] if EXAMPLES else []
    rows += blank_rows(25, 10)
    end = table(ws, 9, h, rows, [6, 10, 44, 12, 11, 28, 34, 12, 12, 10], input_cols=list(range(1, 11)))
    add_list_validation(ws, f"J10:J{end}", ["Open", "Cleared"])
    r = end + 2
    ws.cell(r, 1, "Open review points").font = F_BOLD
    ws.cell(r, 2, f'=COUNTIF(J10:J{end},"Open")')
    r += 2
    ws.cell(r, 1, "ICSR maturity scorecard (each dimension scored 0–20)").font = F_BOLD
    dims = [["1 Governance and leadership", "Board oversight of data; named executive owner; data integrity in code", 14],
            ["2 Policies and procedures", "KPI definitions; maker-checker-approver; approved factor policy", 13],
            ["3 Data collection and validation", "Registers; ledger reconciliation; cut-off", 16],
            ["4 Internal review and audit", "Independent recalculation; control testing; deficiencies closed", 10],
            ["5 External assurance and disclosure", "Evidence index; balanced disclosure; restatement policy", 11]]
    e2 = table(ws, r + 1, ["Dimension", "What good looks like", "Score (0–20)"], dims, None, input_cols=[3], example_rows=5)
    ws.cell(e2 + 1, 2, "Total / 100").font = F_BOLD
    ws.cell(e2 + 1, 3, f"=SUM(C{r+2}:C{e2})")
    ws.cell(e2 + 2, 2, "Band").font = F_BOLD
    ws.cell(e2 + 2, 3, f'=IF(C{e2+1}>=85,"Ready for reasonable assurance",IF(C{e2+1}>=70,"Ready for limited assurance",IF(C{e2+1}>=50,"Developing","Initial")))')


def main():
    global EXAMPLES
    p = argparse.ArgumentParser()
    p.add_argument("--entity", default="[Entity name]")
    p.add_argument("--year-end", default="[YYYY-MM-DD]")
    p.add_argument("--currency", default="[presentation currency, e.g. KHR million]")
    p.add_argument("--out", default="S1_Working_Papers.xlsx")
    p.add_argument("--no-examples", action="store_true", help="omit illustrative example rows")
    args = p.parse_args()
    EXAMPLES = not args.no_examples

    wb = Workbook()
    wb.remove(wb.active)
    s1_00(wb, args)
    s1_01(wb, args)
    s1_02(wb, args)
    s1_03(wb, args)
    s1_04(wb, args)
    checklist_sheet(wb, args, "S1-05", "Governance", "S1.26–27", GOV)
    s1_06(wb, args)
    checklist_sheet(wb, args, "S1-07", "Risk management", "S1.43–44", RISK)
    s1_08(wb, args)
    s1_09(wb, args)
    s1_10(wb, args)
    s1_11(wb, args)
    s1_12(wb, args)
    s1_13(wb, args)
    s1_14(wb, args)
    for ws in wb.worksheets:
        ws.freeze_panes = "A10" if ws.max_row >= 10 else None
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
    # fix hyperlinks on index now that all sheets exist
    idx = wb["S1-00"]
    for row in idx.iter_rows(min_row=10, max_row=24, max_col=1):
        c = row[0]
        if c.value in wb.sheetnames:
            c.hyperlink = f"#'{c.value}'!A1"
            c.font = Font(name=FONT, size=10, color="0563C1", underline="single")
        else:
            c.hyperlink = None
    wb.save(args.out)
    print(f"Saved {args.out} with sheets: {', '.join(wb.sheetnames)}")


if __name__ == "__main__":
    main()
