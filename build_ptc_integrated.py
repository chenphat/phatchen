"""Integrates the PTC audit working-paper file into a full audit-process, ISA / ISQM / IESBA compliance layer.

Input : the PTC working-paper workbook (path in argv[1], default ./2_PTC_Audit_Working_Papers.xlsx)
Output: PTC_Integrated_Audit_File.xlsx

All 51 original sheets are kept unchanged except for additive rows appended to Index, G2_WP_Register and
G4_ISA_Map. New H-series working papers follow the file's own conventions (rows 1-3 heading, row 5 sign-off,
blue = input, black = formula, light-yellow = key output).
"""
import sys
from copy import copy

from openpyxl import load_workbook
from openpyxl.chart import BarChart, DoughnutChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation
from datetime import date

SRC = sys.argv[1] if len(sys.argv) > 1 else "2_PTC_Audit_Working_Papers.xlsx"
OUT = "PTC_Integrated_Audit_File.xlsx"

# house style (taken from the PTC file)
SLATE, TEAL, AMBER, RUST, GREY = "22303C", "1F6F78", "E08A00", "B8442F", "7F7F7F"
KEY, BORDER_C, INPUT = "FFF2CC", "BFBFBF", "0000FF"
WHITE, BG, MUTED, INK, HEAD = "FFFFFF", "EEF1F4", "595959", "000000", "F2F2F2"
GREEN, RED, BLUE, GOLD, GOLD_L, HSEC = "2E7D32", "C62828", "1F4E79", "FCD34D", "FFFBEB", "5B3A8C"

thin = Side(style="thin", color=BORDER_C)
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
BOTTOM = Border(bottom=thin)


def font(size=10, bold=False, color=INK, italic=False, underline=None):
    return Font(name="Arial", size=size, bold=bold, color=color, italic=italic, underline=underline)


def fill(c):
    return PatternFill("solid", start_color=c, end_color=c)


def align(h="left", v="center", wrap=False, rot=0, shrink=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap, text_rotation=rot, shrink_to_fit=shrink)


def paint(ws, r1, c1, r2, c2, color):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(r, c).fill = fill(color)


def put(ws, r, c, value, fnt=None, al=None, to=None, fmt=None, bg=None, border=None):
    if to and (to[0] != r or to[1] != c):
        if bg:
            paint(ws, r, c, to[0], to[1], bg)
        ws.merge_cells(start_row=r, start_column=c, end_row=to[0], end_column=to[1])
    cell = ws.cell(r, c, value)
    cell.font = fnt or font()
    cell.alignment = al or align()
    if fmt:
        cell.number_format = fmt
    if bg:
        cell.fill = fill(bg)
    if border:
        cell.border = border
    return cell


def status_cf(ws, rng, mapping, size=10):
    for text, color in mapping.items():
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{text}"'],
                                      font=Font(name="Arial", size=size, bold=True, color=WHITE), fill=fill(color)))


wb = load_workbook(SRC)
ORIG = list(wb.sheetnames)
CLIENT_LINE = wb["G2_WP_Register"]["A3"].value
LAST = 80


def new_sheet(name, idx=None):
    ws = wb.create_sheet(name, idx) if idx is not None else wb.create_sheet(name)
    ws.sheet_properties.tabColor = HSEC
    ws.sheet_view.showGridLines = False
    return ws


def wp_header(ws, code_title, note, ncols):
    """Rows 1-5 in the file's own layout; row 5 sign-off left for the preparer and reviewer."""
    put(ws, 1, 1, "Funan Consulting — Audit Working Papers", font(9, True, GREY))
    put(ws, 2, 1, code_title, font(14, True, SLATE))
    put(ws, 3, 1, CLIENT_LINE, font(9, color=MUTED))
    put(ws, 4, 1, note, font(9, italic=True, color=MUTED), align(wrap=True, v="top"), to=(4, ncols))
    ws.row_dimensions[4].height = 30
    put(ws, 5, 1, "Prepared: ____, ____  |  Reviewed: ____, ____  |  Partner/EQR: ____, ____  |  Conclusion: ____   "
        "(H-series added by the integration — preparer and reviewer to sign; see G2 and RP-05)",
        font(10, color=INPUT))
    c = put(ws, 4, ncols + 2, "◄ H0 Dashboard", font(9, True, BLUE, underline="single"))
    c.hyperlink = "#'H0_Dashboard'!A1"


def table(ws, headers, rows, widths, hrow=6, formula_cols=(), fmts=None, wrap=True, height=None, center_cols=()):
    fmts = fmts or {}
    for j, h in enumerate(headers, 1):
        put(ws, hrow, j, h, font(10, True, WHITE), align("center", wrap=True), bg=SLATE, border=BOX)
        ws.column_dimensions[L(j)].width = widths[j - 1]
    ws.row_dimensions[hrow].height = 32
    for i, row in enumerate(rows):
        r = hrow + 1 + i
        for j, v in enumerate(row, 1):
            isf = j in formula_cols
            put(ws, r, j, v, font(10, color=INK if isf else INPUT),
                align("center" if (isf or j in center_cols) else "left", "top", wrap=wrap), fmt=fmts.get(j), border=BOX)
        if height:
            ws.row_dimensions[r].height = height
    ws.freeze_panes = ws.cell(hrow + 1, 1)
    return hrow + 1, hrow + len(rows)


# =====================================================================================  phases / process map data
PHASES = [
    ("1. Acceptance & ethics", "ISQM 1.30; ISA 210; ISA 220R.16–24; IESBA Code", ["A1_Acceptance", "A3_Independence", "A5_Team_Competence"]),
    ("2. Planning", "ISA 300; ISA 320; ISA 520", ["C0_Client_Info", "A2_Team_Budget", "A4_PBC_Tracker", "B1_Materiality", "E2_Analytics"]),
    ("3. Risk assessment", "ISA 315R; ISA 240", ["B8_Risk_Factors", "B7_Inherent_Risk", "B9_Fraud_Risk", "B2_Risk_Register",
                                               "B3_Controls_IT", "B4_Team_Discussion", "B11_RMM_Assessment"]),
    ("4. Risk response & controls", "ISA 330; ISA 240.32; ISA 530", ["B5_Risk_Response", "B10_Controls_DI", "D1_Audit_Programme",
                                                                      "D3_Sampling", "B6_Fraud_Journals"]),
    ("5. Substantive fieldwork", "ISA 500-series; ISA 540R; ISA 550; ISA 620", ["C1_TB_Lead", "C5_Mapping_FS", "C2_AJE", "C3_FS_USD_KHR",
                                                                                 "C4_Opening_Bal", "C6_SoCE_SoCF", "C7_Notes", "D2_Contracts",
                                                                                 "D4_Estimates", "D5_Expert", "D6_Confirmations",
                                                                                 "D7_Inventory_Count", "D8_Related_Parties", "D9_IPE"]),
    ("6. Completion", "ISA 450; ISA 560; ISA 570R; ISA 580; ISA 220R.29–40", ["D10_Subsequent_Events", "E1_Misstatements",
                                                                               "E10_Unadjusted_Summary", "E3_Going_Concern",
                                                                               "E7_GC_Assessment", "E6_Disclosure", "E5_Review_Log",
                                                                               "E4_Completion", "E8_Finalisation"]),
    ("7. Reporting", "ISA 700-series; ISA 260R; ISA 265", ["E9_Opinion", "H5_Findings_5Cs", "H6_5Cs_Report"]),
    ("8. Engagement quality review", "ISQM 2; ISA 220R.36", ["F1_EQR"]),
    ("9. Post-audit & monitoring", "ISA 230.14–16; ISQM 1.35–54", ["F2_Post_Audit"]),
    ("10. QA & compliance", "ISA 230; ISQM 1; ISQM 2; IESBA Code", ["G1_QA_Readiness", "G2_WP_Register", "G3_WP_Standards",
                                                                    "G4_ISA_Map", "H0_Dashboard", "H1_Process_Map",
                                                                    "H2_ISA_Compliance", "H3_ISQM_Compliance",
                                                                    "H4_IESBA_Compliance", "H7_Review_Points", "H8_Calc"]),
]
NEW = {  # H-series: content (Index) and objective (G2)
    "H0_Dashboard": ("Audit process and compliance dashboard", "Summarise file status, risks, controls, misstatements, quality gates and compliance by live link"),
    "H1_Process_Map": ("Audit process map: phase → working paper → sign-off", "Show every working paper in its audit phase with objective, sign-off and conclusion from G2"),
    "H2_ISA_Compliance": ("ISA compliance matrix (ISA 200–810)", "Confirm each ISA is complied with, not applicable (with reason) or open, with evidence"),
    "H3_ISQM_Compliance": ("ISQM 1, ISQM 2 and ISA 220 (Revised) compliance matrix", "Map firm- and engagement-level quality management requirements to evidence"),
    "H4_IESBA_Compliance": ("IESBA Code compliance matrix (ethics and independence)", "Confirm fundamental principles and Part 4A independence requirements are evidenced"),
    "H5_Findings_5Cs": ("Management letter points in 5Cs format (ISA 265)", "Document each deficiency as Criteria, Condition, Cause, Consequence and Corrective action"),
    "H6_5Cs_Report": ("Printable 5Cs finding report", "Present a single management letter point in 5Cs format for communication"),
    "H7_Review_Points": ("Compliance review points raised on this file", "Record gaps found in the compliance review and track clearance"),
    "H8_Calc": ("Calculations behind the dashboard", "Supporting calculations for H0 (no judgement)"),
}
for p in PHASES:
    for wp in p[2]:
        assert wp in ORIG or wp in NEW, wp
assert sorted(w for p in PHASES for w in p[2] if w in ORIG) == sorted(s for s in ORIG if s != "Index"), "phase map incomplete"

# create sheets (dashboard first, the rest at the end)
dash = new_sheet("H0_Dashboard", 0)
pm = new_sheet("H1_Process_Map")
isa = new_sheet("H2_ISA_Compliance")
isqm = new_sheet("H3_ISQM_Compliance")
ies = new_sheet("H4_IESBA_Compliance")
fnd = new_sheet("H5_Findings_5Cs")
rep = new_sheet("H6_5Cs_Report")
rp = new_sheet("H7_Review_Points")
calc = new_sheet("H8_Calc")

# =====================================================================================  additive rows on Index, G2, G4
def append_like(ws, src_row, values, link_col=None):
    r = ws.max_row + 1
    for j, v in enumerate(values, 1):
        s = ws.cell(src_row, j)
        c = ws.cell(r, j, v)
        c.font, c.fill, c.border, c.alignment, c.number_format = copy(s.font), copy(s.fill), copy(s.border), copy(s.alignment), s.number_format
    if link_col:
        ws.cell(r, link_col).hyperlink = f"#'{values[link_col - 1]}'!A1"
    return r


ix = wb["Index"]
for name, (content, _) in NEW.items():
    append_like(ix, 56, [name, content, "Integration — process, compliance and reporting"], link_col=1)
g2 = wb["G2_WP_Register"]
for name, (_, obj) in NEW.items():
    append_like(g2, 55, [name, obj, None, None, None, None, None, None, None])
g4 = wb["G4_ISA_Map"]
append_like(g4, 43, ["Detailed compliance matrices", "ISA 200–810, ISQM 1, ISQM 2, ISA 220 (Revised), IESBA Code", "Yes",
                     "Requirement-level evidence, status and review points", "H2, H3, H4, H7"])

# =====================================================================================  H1 Process map
wp_header(pm, "H1 — Audit process map: phase → working paper → sign-off (ISA 230.9; ISA 220 (Revised).29–34)",
          "Every working paper in the file placed in its audit phase. Objective, sign-offs and conclusion are live links to G2; "
          "sign-off status is a formula. Click a working-paper name to open it.", 11)
rows = []
r = 7
for ph, stds, wps in PHASES:
    for k, wp in enumerate(wps, 1):
        G = lambda col: f"INDEX(G2_WP_Register!${col}$7:${col}${LAST},MATCH($C{r},G2_WP_Register!$A$7:$A${LAST},0))"
        title = wb[wp]["A2"].value if wp in ORIG else None
        std = title[title.find("(") + 1: title.rfind(")")] if title and "(" in title else stds
        rows.append([ph, k, wp,
                     f'=IFERROR(INDEX(Index!$B$7:$B${LAST},MATCH($C{r},Index!$A$7:$A${LAST},0)),"")',
                     f'=IFERROR({G("B")},"Not in register")', std,
                     f'=IFERROR(IF({G("C")}="","",{G("C")}&", "&{G("D")}),"")',
                     f'=IFERROR(IF({G("E")}="","",{G("E")}&", "&{G("F")}),"")',
                     f'=IFERROR(IF({G("G")}="","",{G("G")}&", "&{G("H")}),"")',
                     f'=IFERROR(IF({G("I")}="","",{G("I")}),"")',
                     f'=IFERROR(IF({G("C")}="","Awaiting preparer sign-off",IF(OR({G("E")}="",{G("E")}="—"),'
                     f'IF({G("G")}="EQR","Signed off (EQR)","Awaiting review"),"Signed off")),"Not in register")'])
        r += 1
PM_FIRST, PM_LAST = table(pm, ["Audit phase", "Step", "Working paper", "Content", "Objective (G2)", "Standards (per WP heading)",
                               "Prepared", "Reviewed", "Partner / EQR", "Conclusion (G2)", "Sign-off status"],
                          rows, [24, 6, 22, 34, 40, 34, 18, 18, 18, 36, 20], formula_cols=(4, 5, 7, 8, 9, 10, 11), center_cols=(2,))
for rr in range(PM_FIRST, PM_LAST + 1):
    c = pm.cell(rr, 3)
    c.hyperlink = f"#'{c.value}'!A1"
    c.font = font(10, color="0563C1", underline="single")
    pm.row_dimensions[rr].height = 40
status_cf(pm, f"K{PM_FIRST}:K{PM_LAST}", {"Signed off": GREEN, "Signed off (EQR)": GREEN, "Awaiting review": AMBER,
                                          "Awaiting preparer sign-off": RED, "Not in register": RED})
PMR = lambda col: f"H1_Process_Map!${col}${PM_FIRST}:${col}${PM_LAST}"


def signoff(cell):
    return (f'=IF(OR({cell}="—",{cell}=""),"—",IFERROR(INDEX({PMR("K")},MATCH({cell},{PMR("C")},0)),"WP not found"))')


ASSESS = ["Adequate", "Partly", "Outside workbook", "Not documented", "Firm level", "N/A"]
COMP_COL = {"Complied": GREEN, "Partly complied": AMBER, "Confirm on file": BLUE, "Gap": RED,
            "Awaiting sign-off": GREY, "Firm system (outside file)": TEAL, "Not applicable": "BDBDBD"}


def comp_status(a, w, s):
    return (f'=IF({a}="N/A","Not applicable",IF({a}="Firm level","Firm system (outside file)",IF({a}="Outside workbook",'
            f'"Confirm on file",IF(AND({w}<>"—",LEFT({s},6)<>"Signed"),"Awaiting sign-off",IF({a}="Adequate","Complied",'
            f'IF({a}="Partly","Partly complied","Gap"))))))')


def assess_dv(ws, col, r1, r2):
    dv = DataValidation(type="list", formula1='"' + ",".join(ASSESS) + '"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"{col}{r1}:{col}{r2}")


# =====================================================================================  H2 ISA compliance
ISA = [
    ("ISA 200", "Overall objectives; professional scepticism and judgement", "Comply with relevant ethical requirements; plan and perform the audit with professional scepticism; exercise professional judgement throughout", "Yes", "B4_Team_Discussion", "E5_Review_Log; E9_Opinion", "Adequate", "", ""),
    ("ISA 210", "Agreeing the terms of audit engagements", "Preconditions: acceptable framework and management's acknowledgement of responsibilities (210.6); terms agreed in an engagement letter (210.9–12)", "Yes", "A1_Acceptance", "Engagement letter signed 9 Oct 2026 — Letters file B", "Outside workbook", "RP-04", "Preconditions documented in A1 #6–8; signed letter is not in this workbook"),
    ("ISA 220 (Revised)", "Quality management for an audit", "Partner responsibility; ethics and independence; acceptance; resources; direction, supervision and review; consultation; EQR; documentation (220R.13–41) — see H3", "Yes", "A5_Team_Competence", "A3_Independence; E5_Review_Log; E4_Completion; F1_EQR", "Partly", "RP-02; RP-06", "Detail in H3"),
    ("ISA 230", "Audit documentation", "Form, content and extent (230.8–11); preparer and reviewer identified (230.9); file assembly within 60 days, no deletion afterwards (230.14–16)", "Yes", "G1_QA_Readiness", "G2_WP_Register; G3_WP_Standards; F2_Post_Audit", "Adequate", "RP-05; RP-11", "Original WPs signed except G2 itself; H-series awaits sign-off"),
    ("ISA 240", "Fraud (version effective before 15 Dec 2026)", "Team discussion; fraud risk factors; presumed revenue risk; management override — journals, estimate bias, unusual transactions; documentation", "Yes", "B9_Fraud_Risk", "B4_Team_Discussion; B6_Fraud_Journals; E4 #11, #27", "Adequate", "", "ISA 240 (Revised) applies from FY2027 (F2 #12)"),
    ("ISA 250 (Revised)", "Laws and regulations", "Understand the legal framework; procedures on laws with direct effect; identified or suspected non-compliance; reporting responsibilities", "Yes", "D1_Audit_Programme", "Programme section 17 (W1); E4 #10, #26", "Adequate", "", "Open tax audit disclosed as contingent liability"),
    ("ISA 260 (Revised)", "Communication with those charged with governance", "Planned scope and timing; significant findings; independence (where required); written communications", "Yes", "E4_Completion", "Planning letter 20 Nov 2026; clearance meeting 12 Mar 2027 — Letters file C", "Outside workbook", "RP-04", ""),
    ("ISA 265", "Communicating deficiencies in internal control", "Determine significant deficiencies; communicate them in writing to TCWG on a timely basis; communicate other deficiencies to management (265.8–11)", "Yes", "B10_Controls_DI", "H5_Findings_5Cs; management letter 29 Mar 2027 — Letters file F", "Partly", "RP-01; RP-09", "Timing of written communication and completeness of points"),
    ("ISA 300", "Planning an audit", "Overall strategy and audit plan; changes during the audit; initial-audit activities incl. predecessor communication (300.7–13)", "Yes", "A2_Team_Budget", "Strategy and plan memorandum — file 5; E5_Review_Log s3", "Outside workbook", "RP-04", "Strategy memo not in this workbook"),
    ("ISA 315 (Revised 2019)", "Identifying and assessing the risks of material misstatement", "Understanding of entity, framework and system of internal control incl. IT; inherent risk factors; significant risks; D&I of controls; stand-back; documentation (315R.38)", "Yes", "B2_Risk_Register", "B3, B7, B8, B10, B11, C0", "Adequate", "", ""),
    ("ISA 320", "Materiality", "Overall, performance and specific materiality; revision as the audit progresses; documentation (320.10–14)", "Yes", "B1_Materiality", "E8_Finalisation", "Adequate", "", "500 / 300 / trivial 25 / related parties 50"),
    ("ISA 330", "Responses to assessed risks", "Overall responses; further procedures responsive to each assessed risk; tests of controls; substantive procedures for material classes; linkage and results documented (330.28–30)", "Yes", "B5_Risk_Response", "D1_Audit_Programme; B10_Controls_DI", "Partly", "RP-03", "B5 concludes on 10 of the 15 risks in B2"),
    ("ISA 402", "Service organisations", "Not applicable — no service organisation used", "No", "—", "Strategy memo s3.2", "N/A", "", ""),
    ("ISA 450", "Evaluation of misstatements", "Clearly trivial threshold; accumulate; request correction; evaluate uncorrected; written representation; documentation (450.15)", "Yes", "E1_Misstatements", "E10_Unadjusted_Summary; C2_AJE", "Adequate", "", "Uncorrected 275 = 55% of materiality"),
    ("ISA 500", "Audit evidence", "Sufficient appropriate evidence; information produced by the entity tested for accuracy and completeness (500.9)", "Yes", "D9_IPE", "D1_Audit_Programme", "Adequate", "", ""),
    ("ISA 501", "Inventory; litigation and claims", "Attend physical counts (501.4); enquire of lawyers about litigation and claims (501.9–10)", "Yes", "D7_Inventory_Count", "D6_Confirmations; E4 #10, #25", "Adequate", "", "79% of inventory attended; rest by alternative procedures"),
    ("ISA 505", "External confirmations", "Control over requests; non-responses; exceptions; alternative procedures", "Yes", "D6_Confirmations", "", "Adequate", "", ""),
    ("ISA 510", "Initial audit engagements — opening balances", "Read prior FS and predecessor's report; evidence that opening balances are free of material misstatement; consistent policies", "Yes", "C4_Opening_Bal", "A1 #14", "Adequate", "", ""),
    ("ISA 520", "Analytical procedures", "Substantive analytical procedures: expectation, threshold, investigation; final analytical review", "Yes", "E2_Analytics", "", "Adequate", "", ""),
    ("ISA 530", "Audit sampling", "Sample design, size and selection; projection and evaluation of misstatements", "Yes", "D3_Sampling", "E1_Misstatements U3", "Adequate", "", ""),
    ("ISA 540 (Revised)", "Accounting estimates", "Understanding; inherent and control risk; responses incl. retrospective review; indicators of bias; disclosures; documentation (540R.39)", "Yes", "D4_Estimates", "D2_Contracts; D5_Expert", "Adequate", "", ""),
    ("ISA 550", "Related parties", "Identify related parties and transactions; respond to risks; arm's-length assertions; disclosure; documentation (550.28)", "Yes", "D8_Related_Parties", "B2 R09; B10 C15", "Adequate", "", ""),
    ("ISA 560", "Subsequent events", "Procedures covering the period to the date of the auditor's report", "Yes", "D10_Subsequent_Events", "E4 #7; E8 section 3", "Adequate", "", ""),
    ("ISA 570 (Revised)", "Going concern (version effective before 15 Dec 2026)", "Evaluate management's assessment (at least 12 months); events beyond; material uncertainty conclusion; reporting", "Yes", "E7_GC_Assessment", "E3_Going_Concern", "Adequate", "", "ISA 570 (Revised 2024) applies from FY2027"),
    ("ISA 580", "Written representations", "Representations dated as at the report date and in the required form (580.14)", "Yes", "E4_Completion", "Representation letter 22 Mar 2027 — Letters file D; E8 section 3", "Outside workbook", "RP-04", ""),
    ("ISA 600 (Revised)", "Group audits", "Not applicable — single entity", "No", "—", "", "N/A", "", ""),
    ("ISA 610 (Revised 2013)", "Using the work of internal auditors", "Not applicable — no internal audit function", "No", "—", "B3 control environment", "N/A", "", ""),
    ("ISA 620", "Using the work of an auditor's expert", "Need for an expert; competence, capabilities and objectivity; agreed scope; adequacy of the work (620.9–13)", "Yes", "D5_Expert", "D2_Contracts", "Adequate", "", "Quantity surveyor"),
    ("ISA 700 (Revised)", "Forming an opinion and reporting", "Evaluate the FS and form the opinion (700.10–15); report content; dated no earlier than evidence, approval and EQR (700.49)", "Yes", "E9_Opinion", "Auditor's report 22 Mar 2027 — Letters file E; E8 section 3", "Adequate", "", "Report wording held in Letters file"),
    ("ISA 701", "Key audit matters", "Not applicable — not a listed entity and not requested (E9 row 21)", "No", "—", "E9_Opinion", "N/A", "", ""),
    ("ISA 705 (Revised)", "Modifications to the opinion", "Considered — not required after corrections (E9 decision rules)", "Considered", "E9_Opinion", "E1_Misstatements", "Adequate", "", ""),
    ("ISA 706 (Revised)", "Emphasis of Matter / Other Matter paragraphs", "Other Matter paragraph — predecessor auditor; Emphasis of Matter considered and not required (E5 s2)", "Yes", "E9_Opinion", "E5_Review_Log", "Adequate", "", ""),
    ("ISA 710", "Comparative information", "Corresponding figures audited by predecessor — Other Matter paragraph (710.13)", "Yes", "C4_Opening_Bal", "E9_Opinion row 31", "Adequate", "", ""),
    ("ISA 720 (Revised)", "Other information", "Read and consider the Directors' report; final version obtained before the report date", "Yes", "E4_Completion", "A4 PBC row 36", "Adequate", "", ""),
    ("ISA 800 / 805 / 810 (Revised)", "Special purpose frameworks; single statements; summary FS", "Not applicable — general purpose financial statements", "No", "—", "", "N/A", "", ""),
]
wp_header(isa, "H2 — ISA compliance matrix: every ISA, evidence and status (ISA 200.18; ISA 230; G4)",
          "Builds on G4 at requirement level. Column G (blue) is the reviewer's evidence assessment; H pulls the primary WP's sign-off "
          "from H1; I is the compliance status formula. 'Outside workbook' = evidence referenced to the Letters / strategy files, "
          "which must be confirmed on the full file.", 11)
rows = []
for i, x in enumerate(ISA):
    r = 7 + i
    rows.append([x[0], x[1], x[2], x[3], x[4], x[5], x[6], signoff(f"E{r}"), comp_status(f"G{r}", f"E{r}", f"H{r}"), x[7], x[8]])
I_FIRST, I_LAST = table(isa, ["Standard", "Subject", "Key requirements (documentation focus)", "Applies?", "Primary WP",
                              "Other evidence", "Evidence assessment", "Primary WP sign-off", "Compliance status",
                              "Review point", "Comment"],
                        rows, [20, 28, 52, 10, 20, 34, 15, 18, 18, 11, 34], formula_cols=(8, 9), center_cols=(4, 7, 10), height=52)
assess_dv(isa, "G", I_FIRST, I_LAST)
status_cf(isa, f"I{I_FIRST}:I{I_LAST}", COMP_COL)
for rr in range(I_FIRST, I_LAST + 1):
    if isa.cell(rr, 5).value not in ("—", None):
        isa.cell(rr, 5).hyperlink = f"#'{isa.cell(rr, 5).value}'!A1"

# =====================================================================================  H3 ISQM
Q = [
    ("ISQM 1.28", "Governance and leadership — culture, leadership responsibility and accountability for quality", "Firm", "—", "Firm quality manual; leadership evaluations", "Firm level", "", ""),
    ("ISQM 1.25–27", "Firm's risk assessment process — quality objectives, quality risks and responses", "Firm", "—", "Firm risk-assessment documentation", "Firm level", "", ""),
    ("ISQM 1.29", "Relevant ethical requirements — independence policies, confirmations, breaches", "Firm + engagement", "A3_Independence", "A1 #9–11; E4 #18", "Partly", "RP-02", "Independence before acceptance not cross-referenced"),
    ("ISQM 1.30", "Acceptance and continuance — integrity, competence, resources, ethics, timetable", "Engagement", "A1_Acceptance", "F2 #6 (FY2027 continuance)", "Adequate", "", "Accepted with conditions 2 Oct 2026"),
    ("ISQM 1.31", "Engagement performance — direction, supervision, review, consultation, differences of opinion, retention (31(f))", "Engagement", "E5_Review_Log", "E4; G1; F2 #3", "Adequate", "", ""),
    ("ISQM 1.32", "Resources — human (competence, time), technological, intellectual, service providers", "Firm + engagement", "A5_Team_Competence", "A2_Team_Budget; D5_Expert; IT specialist", "Adequate", "", ""),
    ("ISQM 1.33", "Information and communication — within the firm, with the team and with external parties", "Firm + engagement", "E5_Review_Log", "E4 #15 (TCWG); firm communications", "Adequate", "", ""),
    ("ISQM 1.34", "Specified responses — independence confirmations, complaints, TCWG communication, EQR criteria (34(f))", "Firm + engagement", "A1_Acceptance", "A1 #18 (EQR required); F1_EQR", "Adequate", "", ""),
    ("ISQM 1.35–47", "Monitoring and remediation — inspection of completed engagements, root cause, remediation", "Firm", "F2_Post_Audit", "F2 #8–9", "Adequate", "", "Selected for 2027 cyclical inspection"),
    ("ISQM 1.53–54", "Annual evaluation of the system of quality management by leadership", "Firm", "F2_Post_Audit", "F2 #10", "Firm level", "", ""),
    ("ISQM 1.57–60", "Documentation of the system of quality management", "Firm", "—", "Firm QM documentation", "Firm level", "", ""),
    ("ISQM 2.17–23", "EQ reviewer: appointment and eligibility — competence, authority, objectivity, cooling-off", "Engagement", "F1_EQR", "A1 #18", "Adequate", "", "Partner from another office"),
    ("ISQM 2.24–27", "Performance of the EQR — timing, procedures, significant judgements, notification before the report date", "Engagement", "F1_EQR", "E8 section 3; E5", "Partly", "RP-06", "EQR completed before the partner's final review"),
    ("ISQM 2.28–30", "Documentation of the EQR", "Engagement", "F1_EQR", "", "Adequate", "", ""),
    ("ISA 220R.13–15", "Leadership — partner takes overall responsibility; sufficient and appropriate involvement", "Engagement", "E5_Review_Log", "F1 #8 (partner c.5% of hours)", "Adequate", "", ""),
    ("ISA 220R.16–21", "Relevant ethical requirements including independence", "Engagement", "A3_Independence", "A1; E4 #18", "Partly", "RP-02", ""),
    ("ISA 220R.22–24", "Acceptance and continuance", "Engagement", "A1_Acceptance", "", "Adequate", "", ""),
    ("ISA 220R.25–28", "Engagement resources — competence, time, specialists and experts", "Engagement", "A5_Team_Competence", "A2_Team_Budget", "Adequate", "", ""),
    ("ISA 220R.29–34", "Direction, supervision and review; partner review of significant matters", "Engagement", "E5_Review_Log", "E4 #1–2", "Adequate", "", "152 review notes raised and cleared"),
    ("ISA 220R.35", "Consultation on difficult or contentious matters", "Engagement", "E5_Review_Log", "E5 s2 (4 consultations)", "Adequate", "", ""),
    ("ISA 220R.36", "Engagement quality review completed before the report is dated", "Engagement", "F1_EQR", "E8 section 3", "Partly", "RP-06", ""),
    ("ISA 220R.37–38", "Differences of opinion", "Engagement", "E5_Review_Log", "—", "Partly", "RP-08", "No explicit statement that none arose"),
    ("ISA 220R.39–40", "Monitoring and remediation; stand-back on overall responsibility", "Engagement", "E5_Review_Log", "E5 s4 stand-back 19 Mar 2027", "Adequate", "", ""),
    ("ISA 220R.41", "Engagement documentation", "Engagement", "G1_QA_Readiness", "G2_WP_Register", "Adequate", "RP-11", ""),
]
wp_header(isqm, "H3 — Quality management compliance: ISQM 1, ISQM 2 and ISA 220 (Revised)",
          "Firm-level ISQM 1 components are evidenced in the firm's system of quality management, not in this engagement file "
          "('Firm level'); where the engagement touches a component the evidence WP is given. Column F (blue) is the reviewer's assessment.", 10)
rows = []
for i, x in enumerate(Q):
    r = 7 + i
    rows.append([x[0], x[1], x[2], x[3], x[4], x[5], signoff(f"D{r}"), comp_status(f"F{r}", f"D{r}", f"G{r}"), x[6], x[7]])
Q_FIRST, Q_LAST = table(isqm, ["Standard / paragraph", "Requirement", "Level", "Engagement evidence (WP)", "Other evidence",
                               "Evidence assessment", "WP sign-off", "Compliance status", "Review point", "Comment"],
                        rows, [18, 56, 16, 22, 34, 15, 18, 22, 11, 34], formula_cols=(7, 8), center_cols=(3, 6, 9), height=40)
assess_dv(isqm, "F", Q_FIRST, Q_LAST)
status_cf(isqm, f"H{Q_FIRST}:H{Q_LAST}", COMP_COL)

# =====================================================================================  H4 IESBA
E = [
    ("Part 1, s110–114", "Fundamental principles: integrity, objectivity, professional competence and due care, confidentiality, professional behaviour", "A1_Acceptance", "Client integrity assessed (A1 #2); team competence (A5); confidentiality under firm policy", "Adequate", "", ""),
    ("Part 1, s120", "Conceptual framework — identify, evaluate and address threats to compliance", "A1_Acceptance", "Threats from fees, non-assurance services, change of auditor and long association considered at acceptance", "Adequate", "", ""),
    ("Part 3, s310", "Conflicts of interest", "A1_Acceptance", "Firm conflict check not documented in A1", "Partly", "RP-07", ""),
    ("Part 3, s320", "Professional appointments — change of auditor; communication with predecessor", "A1_Acceptance", "Predecessor contacted with consent 21 Sep 2026; reply 28 Sep 2026 (A1 #4–5)", "Adequate", "", ""),
    ("Part 4A, s410", "Fees — level, dependence, contingent fees, overdue fees", "A1_Acceptance", "Fee below 5% of firm revenue; not contingent (A1 #11); invoiced 24 Mar 2027, none overdue (E4 #32)", "Adequate", "", ""),
    ("Part 4A, s420", "Gifts and hospitality", "A3_Independence", "Declared 'None' by every team member", "Adequate", "", ""),
    ("Part 4A, s430", "Actual or threatened litigation", "A1_Acceptance", "Not documented", "Partly", "RP-07", ""),
    ("Part 3, s360", "Responding to non-compliance with laws and regulations (NOCLAR)", "E4_Completion", "None identified; further reporting considered (E4 #26); open tax audit disclosed", "Adequate", "", ""),
    ("Part 4A, s400", "Independence for the engagement period and the period covered by the FS; documentation", "A3_Independence", "Engagement declarations dated 5 Oct 2026 onward; acceptance approved 2 Oct 2026", "Partly", "RP-02", ""),
    ("Part 4A, s510–511", "Financial interests; loans and guarantees", "A3_Independence", "None declared", "Adequate", "", ""),
    ("Part 4A, s520–521", "Business, family and personal relationships", "A3_Independence", "None declared", "Adequate", "", ""),
    ("Part 4A, s522–525", "Recent service with the client; director or officer; employment with the client; temporary personnel", "A3_Independence", "Not covered by the A3 declaration columns", "Partly", "RP-07", ""),
    ("Part 4A, s540", "Long association of personnel", "A3_Independence", "First year for all members (A3 column F); EQR never served the client", "Adequate", "", ""),
    ("Part 4A, s600–610", "Non-assurance services (accounting, valuation, tax, IT, legal, recruiting, corporate finance)", "A1_Acceptance", "No NAS; request to draft FS declined (A1 #10); tax compliance by another firm", "Adequate", "", ""),
    ("Part 4A, s400 (breaches)", "Breaches of independence identified and addressed", "A3_Independence", "No breaches reported at planning or completion", "Adequate", "", ""),
    ("Part 4A", "Independence reconfirmed at completion", "A3_Independence", "All members reconfirmed by 19 Mar 2027 (A3 column H; E4 #18)", "Adequate", "", ""),
    ("ISA 260R / Part 4A", "Independence communicated to those charged with governance where required", "E4_Completion", "Covered in TCWG communications — Letters file C", "Outside workbook", "RP-04", ""),
    ("Forward-looking", "IESBA Code revisions effective for later periods (e.g. using the work of an external expert) — check the Handbook and plan for FY2027", "F2_Post_Audit", "F2 #12 plans for ISA 240 (Revised) and ISA 570 (Revised 2024) only", "Partly", "RP-10", "Confirm effective dates in the IESBA Handbook"),
]
wp_header(ies, "H4 — IESBA Code compliance: fundamental principles, Part 3 and Part 4A independence",
          "Maps the IESBA Code (including the International Independence Standards) to evidence in A1, A3 and E4. Column E (blue) "
          "is the reviewer's assessment; F pulls the WP sign-off; G is the compliance status formula.", 9)
rows = []
for i, x in enumerate(E):
    r = 7 + i
    rows.append([x[0], x[1], x[2], x[3], x[4], signoff(f"C{r}"), comp_status(f"E{r}", f"C{r}", f"F{r}"), x[5], x[6]])
E_FIRST, E_LAST = table(ies, ["Code section", "Requirement", "Evidence (WP)", "Engagement response", "Evidence assessment",
                              "WP sign-off", "Compliance status", "Review point", "Comment"],
                        rows, [20, 52, 20, 52, 15, 18, 20, 11, 30], formula_cols=(6, 7), center_cols=(5, 8), height=40)
assess_dv(ies, "E", E_FIRST, E_LAST)
status_cf(ies, f"G{E_FIRST}:G{E_LAST}", COMP_COL)
for ws_, col, a, b in ((isqm, 4, Q_FIRST, Q_LAST), (ies, 3, E_FIRST, E_LAST)):
    for rr in range(a, b + 1):
        v = ws_.cell(rr, col).value
        if v and v != "—":
            ws_.cell(rr, col).hyperlink = f"#'{v}'!A1"

# =====================================================================================  H5 Findings 5Cs
F = [  # id, key control, other controls, risk, title, criteria, condition, cause, consequence, corrective, response, owner, communicated
    ("ML 1", "C18", "C13, C19", "R02", "QuickBooks access & audit trail",
     "Each user should have an individual login matched to role; leavers removed promptly; closed periods locked; manual journals approved by someone other than the preparer before posting (B10 C13, C18, C19).",
     "Two shared 'Admin' logins are in use; a leaver's account is still active; there is no periodic access review; no closing-date password is set, so posted entries can be edited or deleted; 2 of 5 manual journals were posted before CFO approval.",
     "QuickBooks was installed without security configuration, nobody owns user access, and the system does not enforce journal approval — the CFO can post his own journals.",
     "Unauthorised or back-dated changes to the accounting records could go undetected, increasing the risk of management override and fraud (R02). The audit could not rely on system controls and extended journal testing to the full audit log.",
     "Issue individual user IDs and remove shared admin logins; disable leavers on their last day; review access quarterly; set the closing-date password after each month-end; require documented approval of every manual journal by someone other than the preparer.",
     "Agreed. Individual logins and closing-date password from April 2027; quarterly access review by the Finance Manager. (illustrative)", "CFO", "Written — management letter 29 Mar 2027"),
    ("ML 2", "C2", "", "R01", "Contract cost-to-complete review",
     "Monthly project cost reviews should challenge cost-to-complete estimates against independent data, with defined thresholds and evidenced sign-off (B10 C2; IFRS for SMEs Section 23).",
     "Three months of review packs had no sign-off or queries; two contract budgets were unchanged since tender. Corrections of 850 (AJE 1, C-04) and 310 (AJE 2, C-06 onerous contract) were needed.",
     "The review has no defined precision or escalation threshold, and the CFO relies on project managers' figures without independent data.",
     "Revenue, contract assets and provisions may be materially misstated; corrected misstatements of 1,160 in profit before tax arose in this area.",
     "Set review thresholds; require a quarterly QS re-forecast of costs to complete with variance explanations; evidence the CFO's challenge and sign-off; introduce a process to identify loss-making contracts.",
     "Agreed. Quarterly QS re-forecast and loss-contract review from Q2 2027. (illustrative)", "CFO / Head of QS", "Written — management letter 29 Mar 2027"),
    ("ML 3", "C4", "C14", "R01, R06, R10", "Spreadsheet & FS close controls",
     "Revenue journals and the draft FS, including the KHR translation, should be prepared from checked data and reviewed against a checklist before approval (B10 C4, C14).",
     "2 of 5 revenue journals were approved after posting and the stage-of-completion spreadsheet is not checked; the CFO's review of the draft FS did not detect an error in translating retained earnings into KHR.",
     "Key spreadsheets have no version control, cell protection or independent check, and there is no financial close or disclosure checklist.",
     "Errors in revenue and in the KHR presentation financial statements may not be detected before issue.",
     "Protect and version-control key spreadsheets and reconcile them to QuickBooks monthly; approve journals before posting; adopt a close and disclosure checklist that includes a review of the translation.",
     "Agreed. Close checklist adopted for FY2027 interim. (illustrative)", "Finance Manager", "Written — management letter 29 Mar 2027"),
    ("ML 4", "C15", "", "R09", "Related party approval",
     "Related party transactions should be recorded in a register and approved by directors without an interest (B10 C15; IFRS for SMEs Section 33).",
     "There is no related party register; equipment-leasing transactions with a shareholder-director's company were approved by that director.",
     "The company has no policy for identifying related parties or for approving their transactions.",
     "Transactions may not be at arm's length or fully disclosed, misleading users of the financial statements (R09).",
     "Keep a register updated from annual declarations by directors and key management; require approval by non-interested directors; benchmark terms against market rates.",
     "Agreed. Register and annual declarations from June 2027. (illustrative)", "Board", "Written — management letter 29 Mar 2027"),
    ("ML 5", "C10", "", "R15", "Cash payment of site labour",
     "Engaging, recording and paying casual labour should be segregated, and each worker should sign for payment (B10 C10).",
     "The site manager engages, records and pays casual labour in cash; 1 of 3 pay sheets inspected had missing worker signatures.",
     "Site procedures are designed for speed, and finance performs no independent check.",
     "Payments to fictitious workers or overpayments could occur and go undetected (R15 fraud consideration).",
     "Pay casual labour by bank or mobile transfer where possible; have a finance officer witness cash payouts; reconcile pay sheets to attendance records weekly.",
     "Agreed in part — mobile payment pilot on two sites. (illustrative)", "HR Manager / Site Managers", "Written — management letter 29 Mar 2027"),
    ("ML 6", "C20", "", "—", "Data backup & recovery",
     "Backups should be stored off-site and restoration tested periodically (B10 C20).",
     "Daily backups are kept only on a local NAS on the same premises; restoration has never been tested.",
     "Backup arrangements are not specified in the IT vendor's service agreement.",
     "Accounting records could be lost through fire, theft or ransomware, disrupting reporting and breaking the audit trail.",
     "Keep an encrypted off-site or cloud copy; test restoration at least twice a year and document the results.",
     "Agreed. Cloud backup contracted; first restore test June 2027. (illustrative)", "IT vendor / Finance Manager", "Written — management letter 29 Mar 2027"),
    ("V 1", "C6", "", "R07", "Three-way match",
     "Payments should be released only when the purchase order, goods received note and invoice agree (B10 C6).",
     "1 of 5 payments tested was made on the invoice and order only, without a goods received note.",
     "The match is manual and QuickBooks does not enforce it.",
     "Payment may be made for goods not received.",
     "Require Accounts Payable to attach a signed goods received note before payment; consider system-enforced matching.",
     "Raised verbally with the CFO (B10 note). (illustrative)", "Accounts Payable", "Verbal — CFO; completion memorandum"),
    ("V 2", "C8", "", "R07", "Subcontractor accrual completeness",
     "The year-end accrual for uncertified subcontractor work should be complete (B10 C8).",
     "There is no completeness check over certificates received after the cut-off date; AJE 3 (280) for unrecorded certificates was needed.",
     "The accrual relies only on certificates received by the cut-off date, with no follow-up.",
     "Cost of construction and payables may be understated.",
     "Obtain site engineers' estimates of uncertified work at year end and follow up certificates received after the cut-off date.",
     "Raised verbally with the CFO (B10 note). (illustrative)", "Finance Manager / QS", "Verbal — CFO; completion memorandum"),
]
wp_header(fnd, "H5 — Management letter points in 5Cs format (ISA 265.9–11; ISA 260 (Revised))",
          "Each deficiency from B10 written as Criteria – Condition – Cause – Consequence – Corrective action. ISA 265.11 requires the "
          "written communication to describe each significant deficiency (Condition) and explain its potential effects (Consequence). "
          "Classification and the B10 reporting reference are live links. Management responses and target dates are illustrative.", 19)
rows = []
for i, x in enumerate(F):
    r = 7 + i
    rows.append([x[0], x[1], x[2], x[3], x[4],
                 f'=IFERROR(INDEX(B10_Controls_DI!$M$7:$M$26,MATCH(B{r},B10_Controls_DI!$B$7:$B$26,0)),"Not in B10")',
                 f'=IFERROR(INDEX(B10_Controls_DI!$N$7:$N$26,MATCH(B{r},B10_Controls_DI!$B$7:$B$26,0)),"")',
                 x[5], x[6], x[7], x[8], x[9], x[10], x[11], x[12], date(2027, 6, 30), "Follow up at FY2027 interim (F2 #7)", "Open",
                 f'=IF(A{r}="","",IF(COUNTA(H{r}:L{r})=5,"Complete","Missing "&(5-COUNTA(H{r}:L{r}))&" C"))'])
F_FIRST, F_LAST = table(fnd, ["Point", "Key control (B10)", "Other controls", "Risk ref", "Title", "ISA 265 classification (B10)",
                              "Reported as (B10)", "1. CRITERIA\nWhat should happen", "2. CONDITION\nWhat actually happened",
                              "3. CAUSE\nWhy it happened", "4. CONSEQUENCE\nImpact / risk", "5. CORRECTIVE ACTION\nRecommendation",
                              "Management response", "Owner", "Communicated", "Target date", "Follow-up", "Status", "5Cs check"],
                        rows, [7, 9, 10, 9, 26, 15, 10, 34, 36, 30, 32, 36, 28, 15, 18, 11, 18, 9, 11],
                        formula_cols=(6, 7, 19), fmts={16: "dd-mmm-yyyy"}, center_cols=(1, 2, 3, 4, 16, 18), height=110)
for j in range(8, 13):
    fnd.cell(6, j).fill = fill("8A5A00")
    for rr in range(F_FIRST, F_LAST + 1):
        fnd.cell(rr, j).fill = fill(GOLD_L)
fnd.row_dimensions[6].height = 44
fnd.freeze_panes = "F7"
status_cf(fnd, f"F{F_FIRST}:F{F_LAST}", {"Significant deficiency": RED, "Deficiency": AMBER})
status_cf(fnd, f"S{F_FIRST}:S{F_LAST}", {"Complete": GREEN})
status_cf(fnd, f"R{F_FIRST}:R{F_LAST}", {"Open": AMBER, "Closed": GREEN})
fnd.conditional_formatting.add(f"S{F_FIRST}:S{F_LAST}", FormulaRule(formula=[f'LEFT($S{F_FIRST},7)="Missing"'],
                               font=Font(name="Arial", size=10, bold=True, color=WHITE), fill=fill(RED)))
dv = DataValidation(type="list", formula1='"Open,In progress,Closed"', allow_blank=True)
fnd.add_data_validation(dv); dv.add(f"R{F_FIRST}:R{LAST}")
FR = lambda col: f"H5_Findings_5Cs!${col}${F_FIRST}:${col}${LAST}"

# =====================================================================================  H6 5Cs report
ws = rep
ws.column_dimensions["A"].width = 2
for col in "BCDEF":
    ws.column_dimensions[col].width = 26
put(ws, 1, 2, "Funan Consulting — Audit Working Papers", font(9, True, GREY))
put(ws, 2, 2, "H6 — Management letter point: the 5Cs (ISA 265; ISA 260 (Revised))", font(14, True, SLATE))
put(ws, 3, 2, CLIENT_LINE, font(9, color=MUTED))
c = put(ws, 4, 2, "◄ H0 Dashboard", font(9, True, BLUE, underline="single")); c.hyperlink = "#'H0_Dashboard'!A1"
c = put(ws, 4, 6, "H5 register ►", font(9, True, BLUE, underline="single"), align("right")); c.hyperlink = "#'H5_Findings_5Cs'!A1"
put(ws, 6, 2, "▼ Select point", font(10, True, SLATE), align("right"))
put(ws, 6, 3, "ML 1", font(12, True, INPUT), align("center"), bg="FFF7CC", border=BOX)
dv = DataValidation(type="list", formula1=f"=H5_Findings_5Cs!$A${F_FIRST}:$A${F_LAST}")
ws.add_data_validation(dv); dv.add("C6")
put(ws, 6, 4, "Choose ML 1–6 (written) or V 1–2 (verbal).", font(8, italic=True, color=MUTED), to=(6, 6))
for j, (nm, ds) in enumerate([("Criteria", "What should happen (policy, procedure, law, standard)."), ("Condition", "What actually happened."),
                              ("Cause", "Why the deviation occurred."), ("Consequence (Effect)", "The impact or risk resulting from the issue."),
                              ("Corrective Action", "What should be done (recommendation).")]):
    c_ = 2 + j
    put(ws, 8, c_, j + 1, font(11, True, WHITE), align("center"), bg=SLATE)
    put(ws, 9, c_, nm, font(10, True, SLATE), align("center", wrap=True), bg=GOLD)
    put(ws, 10, c_, ds, font(8), align("center", "top", wrap=True), bg=GOLD_L)
ws.row_dimensions[10].height = 30
LK = lambda col: f"INDEX(H5_Findings_5Cs!${col}${F_FIRST}:${col}${F_LAST},MATCH($C$6,H5_Findings_5Cs!$A${F_FIRST}:$A${F_LAST},0))"
MISSING = "[Not yet documented — complete before issuing]"
put(ws, 12, 2, f'=IFERROR($C$6&":  "&{LK("E")},"Select a valid point")', font(13, True, WHITE), bg=SLATE, to=(12, 6))
ws.row_dimensions[12].height = 26
for j, (lab, col) in enumerate([("ISA 265 classification", "F"), ("Key control (B10)", "B"), ("Risk ref (B2)", "D"),
                                ("Communicated", "O"), ("Status", "R")]):
    put(ws, 13, 2 + j, lab, font(8, True, MUTED), align("center"), bg=HEAD)
    put(ws, 14, 2 + j, f'=IFERROR({LK(col)},"")', font(10, True), align("center", wrap=True), border=BOX)
ws.row_dimensions[14].height = 30
status_cf(ws, "B14", {"Significant deficiency": RED, "Deficiency": AMBER})
for i, (lab, sub, col) in enumerate([("1.  CRITERIA", "What should happen", "H"), ("2.  CONDITION", "What actually happened", "I"),
                                     ("3.  CAUSE", "Why it happened", "J"), ("4.  CONSEQUENCE", "Effect — impact or risk", "K"),
                                     ("5.  CORRECTIVE ACTION", "Recommendation", "L")]):
    r = 16 + i
    put(ws, r, 2, f"{lab}\n{sub}", font(10, True, SLATE), align(wrap=True), bg=GOLD)
    put(ws, r, 3, f'=IFERROR(IF({LK(col)}="","{MISSING}",{LK(col)}),"")', font(10), align(wrap=True), to=(r, 6))
    for cc_ in range(3, 7):
        ws.cell(r, cc_).border = BOTTOM
    ws.row_dimensions[r].height = 62
put(ws, 22, 2, "Management response", font(9, True, WHITE), bg=SLATE)
put(ws, 22, 3, f'=IFERROR({LK("M")},"")', font(10, italic=True), align(wrap=True), to=(22, 6))
ws.row_dimensions[22].height = 34
for j, (lab, col, fm) in enumerate([("Owner", "N", None), ("Target date", "P", "dd-mmm-yyyy"), ("Follow-up", "Q", None),
                                    ("Other controls", "C", None), ("5Cs check", "S", None)]):
    put(ws, 24, 2 + j, lab, font(8, True, MUTED), align("center"), bg=HEAD)
    put(ws, 25, 2 + j, f'=IFERROR({LK(col)},"")', font(10, True), align("center", wrap=True), fmt=fm, border=BOX)
ws.row_dimensions[25].height = 30
status_cf(ws, "F25", {"Complete": GREEN})
put(ws, 27, 2, "ISA 265.11: the written communication of significant deficiencies includes a description of each deficiency and an "
    "explanation of its potential effects, with enough context for TCWG and management to understand it.",
    font(8, italic=True, color=MUTED), align(wrap=True), to=(28, 6))
ws.page_setup.orientation = "portrait"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_area = "A1:G28"

# =====================================================================================  H7 review points
RP = [
    ("RP-01", "E4_Completion; G1_QA_Readiness", "ISA 265.9; ISA 260 (Revised)", "High",
     "Significant deficiencies must be communicated in writing to those charged with governance on a timely basis, and checklists should reflect what has actually been done.",
     "E4 step 16 and G1 item 27 record significant deficiencies as communicated in writing (signed off 18–22 Mar 2027), but the management letter is dated 29 Mar 2027 (A2 milestones; F2 #4).",
     "The completion checklist may overstate what was done before the report was signed, and an inspector would treat this as a documentation failure.",
     "File the written communication given at the 12 Mar 2027 clearance meeting, or correct E4/G1 to show the actual date and add a note on the timing."),
    ("RP-02", "A1_Acceptance; A3_Independence", "ISQM 1.29–30; ISA 220R.16–17; IESBA Part 4A s400", "Medium",
     "Independence must be evaluated before accepting the engagement, and that evaluation must be evidenced.",
     "Acceptance was approved on 2 Oct 2026 and A1 #9 concludes that declarations were signed, but the engagement-specific declarations in A3 are dated 5 Oct 2026 or later.",
     "The file does not show the independence evidence the acceptance decision relied on.",
     "Cross-reference A1 #9 to the firm's independence and conflict check and the annual declarations held before 2 Oct 2026, or re-date the acceptance conclusion."),
    ("RP-03", "B5_Risk_Response; E8_Finalisation", "ISA 330.28(b); ISA 315R", "Medium",
     "Every assessed risk should be linked to the further procedures performed and their results.",
     "B2 lists 15 risks; B5 concludes on 10 (R01–R09, R11). R10, R12, R13, R14 and R15 are responded to in D1 but have no conclusion in B5, and E8 row 19 hard-codes 10.",
     "The linkage from risk to response to conclusion is incomplete for five risks.",
     "Add R10 and R12–R15 to B5 with their D1 references and conclusions, and change the E8 check to count the risks in B2."),
    ("RP-04", "G4_ISA_Map (Letters B–F, file 5)", "ISA 230.8; ISA 210; ISA 260R; ISA 580; ISA 700; ISA 265", "Medium",
     "The audit file must contain the engagement letter, strategy memorandum, TCWG communications, written representations, auditor's report and management letter.",
     "These documents are referenced as 'Letters B–F' and 'file 5' but are not in this workbook.",
     "A reviewer using only this workbook cannot confirm compliance with ISAs 210, 260, 265, 300, 580 and 700.",
     "Attach the documents or index them by file location in G2 so they can be inspected with the workbook."),
    ("RP-05", "H0–H8 (this integration)", "ISA 230.9", "Medium",
     "Each working paper must show who prepared and reviewed it and when.",
     "The nine H-series sheets added by this integration are in G2 with no preparer or reviewer.",
     "Until they are signed, they are not part of the reviewed audit evidence.",
     "The audit manager should review the H-series, complete row 5 and G2, and amend any evidence assessments they disagree with."),
    ("RP-06", "F1_EQR; E9_Opinion; E8_Finalisation", "ISQM 2.24–27; ISA 220R.36", "Medium",
     "The EQ reviewer must review the FS and proposed report and complete the review before the report is dated.",
     "The EQR was completed on 16 Mar 2027, before the partner's final review (19 Mar) and the final opinion WP (E9 prepared 19 Mar; reviewed as a draft on 16 Mar).",
     "If anything significant changed after 16 Mar, the EQR did not cover it.",
     "Document that no significant changes arose after the EQR, or have the EQ reviewer confirm the final FS and report before 22 Mar 2027."),
    ("RP-07", "A1_Acceptance; A3_Independence", "IESBA s310, s430, s522–525", "Low",
     "Threats from conflicts of interest, litigation, recent service, director or officer roles and employment with the client should be considered.",
     "The A3 declaration columns cover financial interests, relationships, non-assurance services, gifts and tenure only; A1 does not record a conflict-of-interest or litigation check.",
     "Independence evidence is incomplete for these sections of the Code.",
     "Extend the A3 template (or reference the firm's declaration form) and add a conflict and litigation check to A1."),
    ("RP-08", "E5_Review_Log", "ISA 220R.37–38", "Low",
     "Differences of opinion within the team, with the EQ reviewer or with consulted parties must be dealt with and documented.",
     "E5 records consultations and reviewer queries but does not state whether any differences of opinion arose.",
     "An inspector cannot tell whether the requirement was considered.",
     "Add a statement to E5 or the completion memorandum that no differences of opinion arose (or how they were resolved)."),
    ("RP-09", "B3_Controls_IT; B10_Controls_DI", "ISA 265.8–10", "Low",
     "Each deficiency identified should be evaluated and, where it merits attention, communicated.",
     "B3 marks the control environment (dominant shareholder-director, no audit committee) as an 'ISA 265 point', but it is not among ML 1–6 or the verbal points in B10.",
     "A deficiency may have been identified but not communicated.",
     "Evaluate it in B10 and either add it to the management letter (H5) or record why it was not communicated."),
    ("RP-10", "F2_Post_Audit", "IESBA Code; ISQM 1.32", "Low",
     "The firm should plan for changes in standards and the Code that affect the next audit.",
     "F2 #12 plans for ISA 240 (Revised) and ISA 570 (Revised 2024) only; it does not mention IESBA Code revisions effective for later periods.",
     "Ethics changes for FY2027 may be missed.",
     "Check the IESBA Handbook for revisions effective for FY2027 (e.g. using the work of an external expert) and add them to F2 #12."),
    ("RP-11", "G2_WP_Register", "ISA 230.9", "Low",
     "Every working paper, including the register itself, should show who prepared and reviewed it and when.",
     "G2 has no row-5 sign-off line and does not list itself, so H1 shows it as 'Not in register'. All other original WPs are signed.",
     "The register that evidences sign-off is itself not evidenced as reviewed.",
     "Add a row-5 sign-off to G2 and a G2 row for the register, then clear this point."),
]
wp_header(rp, "H7 — Compliance review points raised on this file (ISA 220 (Revised).29–34; ISQM 1.31)",
          "Points found when the file was mapped to the ISAs, ISQM 1 / ISQM 2 and the IESBA Code (H2–H4). Each point gives the "
          "criteria, condition, consequence and recommendation. Status, cleared-by and date are inputs for the audit manager.", 12)
SHORT = {"RP-01": "Management letter dated after checklist shows it sent", "RP-02": "Independence evidence post-dates acceptance",
         "RP-03": "B5 concludes on 10 of 15 assessed risks", "RP-04": "Letters B–F and strategy memo not in workbook",
         "RP-05": "H-series working papers await sign-off", "RP-06": "EQR completed before partner's final review",
         "RP-07": "Independence declaration omits s310, s430, s522–525", "RP-08": "No statement on differences of opinion",
         "RP-09": "Control-environment deficiency not communicated", "RP-10": "IESBA changes for FY2027 not planned",
         "RP-11": "G2 register has no sign-off of its own"}
rows = [[x[0], SHORT[x[0]], x[1], x[2], x[3], x[4], x[5], x[6], x[7], "AM" if x[3] != "High" else "EP", "Open", None, None] for x in RP]
R_FIRST, R_LAST = table(rp, ["Ref", "Issue", "Working paper(s)", "Standard", "Severity", "Criteria", "Condition (what we found)", "Consequence",
                             "Recommendation", "Owner", "Status", "Cleared by", "Date cleared"],
                        rows, [8, 26, 22, 22, 10, 36, 44, 30, 40, 8, 10, 10, 12], center_cols=(1, 5, 10, 11, 12, 13), height=88,
                        fmts={13: "dd-mmm-yyyy"})
status_cf(rp, f"E{R_FIRST}:E{R_LAST}", {"High": RED, "Medium": AMBER, "Low": GREEN})
status_cf(rp, f"K{R_FIRST}:K{R_LAST}", {"Open": AMBER, "Cleared": GREEN})
dv = DataValidation(type="list", formula1='"Open,Cleared"', allow_blank=True); rp.add_data_validation(dv); dv.add(f"K{R_FIRST}:K{LAST}")
dv = DataValidation(type="list", formula1='"High,Medium,Low"', allow_blank=True); rp.add_data_validation(dv); dv.add(f"E{R_FIRST}:E{LAST}")
RR = lambda col: f"H7_Review_Points!${col}${R_FIRST}:${col}${LAST}"

# =====================================================================================  H8 Calc
put(calc, 1, 1, "Funan Consulting — Audit Working Papers", font(9, True, GREY))
put(calc, 2, 1, "H8 — Calculations behind H0 (no judgement; do not edit)", font(14, True, SLATE))
for col, w in zip("ABCDEFGH", [34, 14, 14, 14, 14, 14, 14, 14]):
    calc.column_dimensions[col].width = w


def chead(r, labels):
    for j, h in enumerate(labels, 1):
        put(calc, r, j, h, font(9, True, WHITE), align("center", wrap=True), bg=SLATE)


chead(4, ["Audit phase", "Working papers", "Signed off", "% signed off"])
for i, (ph, _, _) in enumerate(PHASES):
    r = 5 + i
    put(calc, r, 1, ph, font(9))
    put(calc, r, 2, f'=COUNTIF({PMR("A")},A{r})', font(9))
    put(calc, r, 3, f'=COUNTIFS({PMR("A")},A{r},{PMR("K")},"Signed*")', font(9))
    put(calc, r, 4, f"=IF(B{r}=0,0,C{r}/B{r})", font(9), fmt="0%")
PH_TOT = 5 + len(PHASES)
put(calc, PH_TOT, 1, "Total", font(9, True))
put(calc, PH_TOT, 2, f"=SUM(B5:B{PH_TOT-1})", font(9, True)); put(calc, PH_TOT, 3, f"=SUM(C5:C{PH_TOT-1})", font(9, True))
put(calc, PH_TOT, 4, f"=IF(B{PH_TOT}=0,0,C{PH_TOT}/B{PH_TOT})", font(9, True), fmt="0%")

CST = ["Complied", "Partly complied", "Confirm on file", "Gap", "Awaiting sign-off"]
r0 = PH_TOT + 2
chead(r0, ["Framework"] + CST + ["Applicable", "% complied"])
FW = [("ISA", f"H2_ISA_Compliance!$I${I_FIRST}:$I${I_LAST}"), ("ISQM / ISA 220R", f"H3_ISQM_Compliance!$H${Q_FIRST}:$H${Q_LAST}"),
      ("IESBA Code", f"H4_IESBA_Compliance!$G${E_FIRST}:$G${E_LAST}")]
for i, (nm, rng) in enumerate(FW):
    r = r0 + 1 + i
    put(calc, r, 1, nm, font(9))
    for j, s in enumerate(CST):
        put(calc, r, 2 + j, f'=COUNTIF({rng},"{s}")', font(9))
    put(calc, r, 7, f"=SUM(B{r}:F{r})", font(9))
    put(calc, r, 8, f"=IF(G{r}=0,0,B{r}/G{r})", font(9), fmt="0%")
CMP_FIRST = r0 + 1
CMP_TOT = r0 + 4
put(calc, CMP_TOT, 1, "All frameworks", font(9, True))
for j in range(2, 8):
    put(calc, CMP_TOT, j, f"=SUM({L(j)}{CMP_FIRST}:{L(j)}{CMP_TOT-1})", font(9, True))
put(calc, CMP_TOT, 8, f"=IF(G{CMP_TOT}=0,0,B{CMP_TOT}/G{CMP_TOT})", font(9, True), fmt="0%")
NA_ROW = CMP_TOT + 1
put(calc, NA_ROW, 1, "Not applicable / firm system (excluded)", font(9))
put(calc, NA_ROW, 2, '=' + "+".join(f'COUNTIF({rng},"Not applicable")+COUNTIF({rng},"Firm system (outside file)")' for _, rng in FW), font(9))

r1 = NA_ROW + 2
chead(r1, ["Risk spectrum (B2)", "Count"])
for i, s in enumerate(["Significant", "Higher", "Lower"]):
    put(calc, r1 + 1 + i, 1, s, font(9)); put(calc, r1 + 1 + i, 2, f'=COUNTIF(B2_Risk_Register!$I$7:$I$21,A{r1+1+i})', font(9))
RISK = r1
r2 = r1 + 5
chead(r2, ["Controls D&I (B10)", "Count"])
for i, (s, f_) in enumerate([("Implemented", "=B10_Controls_DI!C30"), ("Partly implemented", "=B10_Controls_DI!C31"),
                             ("Not implemented / not effective", "=B10_Controls_DI!C32")]):
    put(calc, r2 + 1 + i, 1, s, font(9)); put(calc, r2 + 1 + i, 2, f_, font(9))
CTRL = r2
r3 = r2 + 5
chead(r3, ["Hours by phase (A2)", "Hours"])
for i, col in enumerate("CDEFGH"):
    put(calc, r3 + 1 + i, 1, f"=A2_Team_Budget!{col}6", font(9)); put(calc, r3 + 1 + i, 2, f"=A2_Team_Budget!{col}15", font(9))
HRS = r3
r4 = r3 + 8
chead(r4, ["KPI", "Value"])
KP = {}
kpis = [
    ("WPs in register", f'=COUNTA({PMR("C")})', None),
    ("WPs signed off", f'=COUNTIF({PMR("K")},"Signed*")', None),
    ("WPs awaiting sign-off", f'=COUNTIF({PMR("K")},"Awaiting*")', None),
    ("Risks identified", "=COUNTA(B2_Risk_Register!$A$7:$A$21)", None),
    ("Significant risks", "=E8_Finalisation!B18", None),
    ("Fraud risks (Yes)", '=COUNTIF(B2_Risk_Register!$J$7:$J$21,"Yes*")', None),
    ("Controls evaluated", "=B10_Controls_DI!C29", None),
    ("Controls implemented", "=B10_Controls_DI!C30", None),
    ("Significant deficiencies", "=B10_Controls_DI!C33", None),
    ("Other deficiencies", "=B10_Controls_DI!C34", None),
    ("Controls relied on (operating effectively)", "=E8_Finalisation!B34", None),
    ("Open review points", f'=COUNTIF({RR("K")},"Open")', None),
    ("Open review points - High", f'=COUNTIFS({RR("K")},"Open",{RR("E")},"High")', None),
    ("Open review points - Medium", f'=COUNTIFS({RR("K")},"Open",{RR("E")},"Medium")', None),
    ("Open review points - Low", f'=COUNTIFS({RR("K")},"Open",{RR("E")},"Low")', None),
    ("5Cs points", f'=COUNTA({FR("A")})', None),
    ("5Cs points complete", f'=COUNTIF({FR("S")},"Complete")', None),
    ("Significant deficiency points", f'=COUNTIF({FR("F")},"Significant deficiency")', None),
    ("Team members reconfirmed independence", "=COUNTA(A3_Independence!$H$7:$H$16)", None),
    ("Team members declared", "=COUNTA(A3_Independence!$A$7:$A$16)", None),
]
for i, (k, f_, fm) in enumerate(kpis):
    r = r4 + 1 + i
    put(calc, r, 1, k, font(9)); put(calc, r, 2, f_, font(9, True), fmt=fm)
    KP[k] = f"H8_Calc!$B${r}"

# =====================================================================================  H0 Dashboard
C0, NC, ROWS = 4, 36, 64
cc = lambda i: C0 + i
dash.sheet_view.showRowColHeaders = False
dash.sheet_view.zoomScale = 85
dash.column_dimensions["A"].width = 1.5
dash.column_dimensions["B"].width = 27
dash.column_dimensions["C"].width = 2
for i in range(NC):
    dash.column_dimensions[L(cc(i))].width = 5
dash.column_dimensions[L(cc(NC))].width = 1.5
for r in range(1, ROWS + 1):
    dash.row_dimensions[r].height = 16.5
paint(dash, 1, 1, ROWS, 2, SLATE)
paint(dash, 1, 3, ROWS, cc(NC), BG)
paint(dash, 1, 1, 4, cc(NC), SLATE)
put(dash, 2, 2, "PTC AUDIT FILE", font(15, True, WHITE), align("left", "bottom"))
put(dash, 3, 2, "PROCESS & COMPLIANCE", font(10, True, "F5C26B"), align("left", "top"))


def top(c1, c2, label, value, fmt=None):
    put(dash, 2, cc(c1), label, font(8, color="C9D1D9"), align("left", "bottom"), to=(2, cc(c2)))
    put(dash, 3, cc(c1), value, font(10, True), align("left", shrink=True), to=(3, cc(c2)), fmt=fmt, bg=WHITE)


top(0, 7, "Client", "Prasat Thmey Construction Co., Ltd. (PTC)")
top(9, 11, "Year end", "31 Dec 2026")
top(13, 17, "Framework", "IFRS for SMEs (3rd ed.)")
top(19, 25, "Opinion (E9)", "=E9_Opinion!B40")
top(27, 34, "Ready to sign? (E8)", "=E8_Finalisation!B48")

nav = [("H0_Dashboard", "Dashboard"), ("Index", "File index"), ("H1_Process_Map", "Audit process map"),
       ("H2_ISA_Compliance", "ISA compliance"), ("H3_ISQM_Compliance", "ISQM / ISA 220R"), ("H4_IESBA_Compliance", "IESBA Code"),
       ("H5_Findings_5Cs", "5Cs management letter"), ("H6_5Cs_Report", "5Cs report (print)"), ("H7_Review_Points", "Review points"),
       ("A1_Acceptance", "A  Acceptance & team"), ("B2_Risk_Register", "B  Risk assessment"), ("C1_TB_Lead", "C  TB, FS & notes"),
       ("D1_Audit_Programme", "D  Programme & testing"), ("E8_Finalisation", "E  Completion"), ("E9_Opinion", "E9 Opinion"),
       ("F1_EQR", "F  EQR & post-audit"), ("G1_QA_Readiness", "G  QA & standards")]
for i, (sh, lab) in enumerate(nav):
    r = 6 + 2 * i
    c = put(dash, r, 2, "  ▸  " + lab, font(10, True, WHITE if i == 0 else "D0D7DE"), bg=TEAL if i == 0 else SLATE)
    c.hyperlink = f"#'{sh}'!A1"
    dash.row_dimensions[r].height = 20
put(dash, 41, 2, "Live links to the working papers. Blue cells on H2–H4 and H7 are reviewer judgements. "
    "Fictional training client — illustrative data.", font(8, italic=True, color="AEB7C2"), align(wrap=True, v="top"), to=(45, 2))

cards = [
    ("WORKING PAPERS", f"={KP['WPs signed off']}", "0", TEAL, f'="signed off of "&{KP["WPs in register"]}&" · "&{KP["WPs awaiting sign-off"]}&" awaiting"'),
    ("SIGNIFICANT RISKS", f"={KP['Significant risks']}", "0", RUST, f'="of "&{KP["Risks identified"]}&" risks · "&{KP["Fraud risks (Yes)"]}&" fraud (ISA 240)"'),
    ("CONTROLS IMPLEMENTED", f"={KP['Controls implemented']}", "0", BLUE, f'="of "&{KP["Controls evaluated"]}&" · "&{KP["Significant deficiencies"]}&" significant deficiencies"'),
    ("UNCORRECTED (USD'000)", "=E1_Misstatements!D16", "#,##0", AMBER, '=TEXT(E1_Misstatements!D19,"0%")&" of materiality "&B1_Materiality!C13'),
    ("COMPLIANCE", f"=H8_Calc!$H${CMP_TOT}", "0%", GREEN, f'="complied · "&H8_Calc!$C${CMP_TOT}&" partly · "&H8_Calc!$D${CMP_TOT}&" confirm on file"'),
    ("OPEN REVIEW POINTS", f"={KP['Open review points']}", "0", RED, f'="High "&{KP["Open review points - High"]}&" · Medium "&{KP["Open review points - Medium"]}&" · Low "&{KP["Open review points - Low"]}'),
]
for i, (lab, f_, fm, col, sub) in enumerate(cards):
    c1, c2 = cc(i * 6), cc(i * 6 + 4)
    paint(dash, 6, c1, 9, c2, WHITE)
    paint(dash, 6, c1, 9, c1, col)
    dash.merge_cells(start_row=6, start_column=c1, end_row=9, end_column=c1)
    put(dash, 6, c1 + 1, lab, font(8, True, MUTED), align("left", "bottom"), to=(6, c2))
    put(dash, 7, c1 + 1, f_, font(22, True, col), align("left"), to=(8, c2), fmt=fm)
    put(dash, 9, c1 + 1, sub, font(8, color=MUTED), align("left", "top", shrink=True), to=(9, c2))


def panel(top_, c1, ncols, nrows, title):
    paint(dash, top_, cc(c1), top_ + nrows, cc(c1 + ncols - 1), WHITE)
    put(dash, top_, cc(c1), title, font(10, True, SLATE), to=(top_, cc(c1 + ncols - 1)))
    dash.row_dimensions[top_].height = 21


def clean(ch, w, h):
    ch.width, ch.height = w, h
    ch.roundedCorners = False
    ch.graphical_properties = GraphicalProperties(ln=LineProperties(noFill=True))
    if ch.legend is not None:
        ch.legend.overlay = False


def line_item(r, c1, c2, c3, label, f_, fm=None, color=INK):
    put(dash, r, cc(c1), label, font(8, color=MUTED), al=align(shrink=True), to=(r, cc(c2)))
    put(dash, r, cc(c2 + 1), f_, font(10, True, color), align("right", shrink=True), to=(r, cc(c3)), fmt=fm)
    for c in range(cc(c1), cc(c3) + 1):
        dash.cell(r, c).border = BOTTOM


T1, T2, T3, BODY = 11, 29, 47, 16

# P1 process by phase
panel(T1, 0, 8, BODY, " 1. Audit process — sign-off by phase")
put(dash, T1 + 1, cc(0), "Phase", font(8, True, MUTED), to=(T1 + 1, cc(3)), bg=HEAD)
for j, h in enumerate(["WPs", "Signed", "% signed"]):
    put(dash, T1 + 1, cc(4 + j) if j < 2 else cc(6), h, font(8, True, MUTED), align("center"),
        to=(T1 + 1, cc(4 + j) if j < 2 else cc(7)), bg=HEAD)
for i in range(len(PHASES)):
    r = T1 + 2 + i
    put(dash, r, cc(0), f"=H8_Calc!A{5+i}", font(8, True), al=align(shrink=True), to=(r, cc(3)))
    put(dash, r, cc(4), f"=H8_Calc!B{5+i}", font(9), align("center"))
    put(dash, r, cc(5), f"=H8_Calc!C{5+i}", font(9), align("center"))
    put(dash, r, cc(6), f"=H8_Calc!D{5+i}", font(9, True), align("center"), to=(r, cc(7)), fmt="0%")
    for c in range(cc(0), cc(8)):
        dash.cell(r, c).border = BOTTOM
dash.conditional_formatting.add(f"{L(cc(6))}{T1+2}:{L(cc(6))}{T1+11}",
                                DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="9FD3C7"))
r = T1 + 12
put(dash, r, cc(0), "Total", font(8, True), to=(r, cc(3)), bg=HEAD)
put(dash, r, cc(4), f"=H8_Calc!B{PH_TOT}", font(9, True), align("center"), bg=HEAD)
put(dash, r, cc(5), f"=H8_Calc!C{PH_TOT}", font(9, True), align("center"), bg=HEAD)
put(dash, r, cc(6), f"=H8_Calc!D{PH_TOT}", font(9, True), align("center"), to=(r, cc(7)), fmt="0%", bg=HEAD)
put(dash, T1 + 14, cc(0), "Live from G2 via H1. H-series added by this integration awaits sign-off (RP-05).",
    font(7, italic=True, color=MUTED), align(wrap=True, v="top"), to=(T1 + 15, cc(7)))

# P2 compliance chart
panel(T1, 9, 8, BODY, " 2. Compliance by framework (H2–H4)")
ch = BarChart(); ch.type = "bar"; ch.grouping = "stacked"; ch.overlap = 100
data = Reference(calc, min_col=2, max_col=6, min_row=r0, max_row=r0 + 3)
ch.add_data(data, titles_from_data=True)
ch.set_categories(Reference(calc, min_col=1, min_row=r0 + 1, max_row=r0 + 3))
for s, col in zip(ch.series, [GREEN, AMBER, BLUE, RED, GREY]):
    s.graphicalProperties.solidFill = col
    s.graphicalProperties.line.solidFill = col
ch.dataLabels = DataLabelList(); ch.dataLabels.showVal = True; ch.dataLabels.numFmt = "0;;;"
ch.dataLabels.showSerName = ch.dataLabels.showCatName = ch.dataLabels.showLegendKey = False
ch.x_axis.delete = False; ch.y_axis.delete = False
ch.x_axis.scaling.orientation = "maxMin"
ch.y_axis.majorGridlines = None
ch.y_axis.number_format = "0"
ch.legend.position = "b"
ch.gapWidth = 50
clean(ch, 8.2, 7.4)
dash.add_chart(ch, f"{L(cc(9))}{T1+1}")
line_item(T1 + 15, 9, 13, 16, "Complied of applicable requirements", f"=H8_Calc!$H${CMP_TOT}", "0%", GREEN)
line_item(T1 + 16, 9, 13, 16, "Not applicable / firm-system items (excluded)", f"=H8_Calc!$B${NA_ROW}", "0")

# P3 risk spectrum
panel(T1, 18, 8, BODY, " 3. Risks of material misstatement (B2)")
ch = DoughnutChart()
ch.add_data(Reference(calc, min_col=2, min_row=RISK, max_row=RISK + 3), titles_from_data=True)
ch.set_categories(Reference(calc, min_col=1, min_row=RISK + 1, max_row=RISK + 3))
for i, col in enumerate([RUST, AMBER, "9FD3C7"]):
    pt = DataPoint(idx=i); pt.graphicalProperties.solidFill = col; pt.graphicalProperties.line.solidFill = WHITE
    ch.series[0].dPt.append(pt)
ch.holeSize = 55
ch.dataLabels = DataLabelList(); ch.dataLabels.showVal = True
ch.dataLabels.showSerName = ch.dataLabels.showCatName = ch.dataLabels.showLegendKey = ch.dataLabels.showPercent = False
ch.legend.position = "b"
clean(ch, 8.2, 6.6)
dash.add_chart(ch, f"{L(cc(18))}{T1+1}")
line_item(T1 + 14, 18, 22, 25, "Significant risks (each has a concluded response)", f"={KP['Significant risks']}", "0", RUST)
line_item(T1 + 15, 18, 22, 25, "Fraud risks flagged 'Yes' (ISA 240)", f"={KP['Fraud risks (Yes)']}", "0")
line_item(T1 + 16, 18, 22, 25, "Risk responses concluded in B5 (RP-03)", "=E8_Finalisation!B19&\" of \"&" + KP["Risks identified"], None, AMBER)

# P4 controls
panel(T1, 27, 8, BODY, " 4. Internal control — design & implementation (B10)")
ch = DoughnutChart()
ch.add_data(Reference(calc, min_col=2, min_row=CTRL, max_row=CTRL + 3), titles_from_data=True)
ch.set_categories(Reference(calc, min_col=1, min_row=CTRL + 1, max_row=CTRL + 3))
for i, col in enumerate([GREEN, AMBER, RED]):
    pt = DataPoint(idx=i); pt.graphicalProperties.solidFill = col; pt.graphicalProperties.line.solidFill = WHITE
    ch.series[0].dPt.append(pt)
ch.holeSize = 55
ch.dataLabels = DataLabelList(); ch.dataLabels.showVal = True
ch.dataLabels.showSerName = ch.dataLabels.showCatName = ch.dataLabels.showLegendKey = ch.dataLabels.showPercent = False
ch.legend.position = "b"
clean(ch, 8.2, 6.6)
dash.add_chart(ch, f"{L(cc(27))}{T1+1}")
line_item(T1 + 14, 27, 31, 34, "Significant deficiencies (ISA 265)", f"={KP['Significant deficiencies']}", "0", RED)
line_item(T1 + 15, 27, 31, 34, "Other deficiencies", f"={KP['Other deficiencies']}", "0", AMBER)
line_item(T1 + 16, 27, 31, 34, "Controls relied on — operating effectively", f"={KP['Controls relied on (operating effectively)']}", "0", GREEN)

# P5 materiality & misstatements
panel(T2, 0, 8, BODY, " 5. Materiality & misstatements (B1, E1)")
items = [("Overall materiality", "=B1_Materiality!C13", "#,##0", SLATE), ("Performance materiality", "=B1_Materiality!C17", "#,##0", SLATE),
         ("Clearly trivial threshold", "=B1_Materiality!C18", "#,##0", SLATE),
         ("Specific materiality — related parties", "=B1_Materiality!C19", "#,##0", SLATE),
         ("Corrected misstatements (PBT)", "=E1_Misstatements!D10", "#,##0", GREEN),
         ("Uncorrected misstatements (PBT)", "=E1_Misstatements!D16", "#,##0", RED),
         ("Uncorrected / materiality", "=E1_Misstatements!D19", "0%", RED)]
for i, (lab, f_, fm, col) in enumerate(items):
    line_item(T2 + 1 + 2 * i, 0, 4, 7, lab, f_, fm, col)
    dash.row_dimensions[T2 + 1 + 2 * i].height = 18
dash.conditional_formatting.add(f"{L(cc(5))}{T2+13}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="F4B6B6"))
put(dash, T2 + 15, cc(0), "=E1_Misstatements!D21", font(8, True, GREEN), align(wrap=True, v="top"), to=(T2 + 16, cc(7)))

# P6 quality gates
panel(T2, 9, 8, BODY, " 6. Completion & quality gates (E8, G1, F1)")
gates = [("Open items on E8 finalisation", "=E8_Finalisation!B47", "0"),
         ("Completion checklist done (E4)", '=COUNTIF(E4_Completion!D7:D39,"Yes")&" / "&COUNTA(E4_Completion!A7:A39)', None),
         ("QA self-check met (G1)", '=COUNTIF(G1_QA_Readiness!E7:E43,"Met")&" / "&COUNTA(G1_QA_Readiness!A7:A43)', None),
         ("EQR steps done (F1)", '=COUNTIF(F1_EQR!D7:D18,"Yes")&" / "&COUNTA(F1_EQR!A7:A18)', None),
         ("Review notes cleared (E5)", '=E5_Review_Log!D12&" / "&E5_Review_Log!C12', None),
         ("Independence reconfirmed (A3)", f'={KP["Team members reconfirmed independence"]}&" / "&{KP["Team members declared"]}', None),
         ("PBC items: late / outstanding (A4)", '=A4_PBC_Tracker!D39&" / "&A4_PBC_Tracker!E39', None),
         ("File assembly (ISA 230.14)", "=E8_Finalisation!F45", None)]
for i, (lab, f_, fm) in enumerate(gates):
    line_item(T2 + 1 + 2 * i, 9, 13, 16, lab, f_, fm, SLATE)
    dash.row_dimensions[T2 + 1 + 2 * i].height = 18

# P7 hours
panel(T2, 18, 8, BODY, " 7. Resources — budgeted hours by phase (A2)")
ch = BarChart(); ch.type = "col"
ch.add_data(Reference(calc, min_col=2, min_row=HRS, max_row=HRS + 6), titles_from_data=True)
ch.set_categories(Reference(calc, min_col=1, min_row=HRS + 1, max_row=HRS + 6))
ch.series[0].graphicalProperties.solidFill = TEAL
ch.dataLabels = DataLabelList(); ch.dataLabels.showVal = True
ch.dataLabels.showSerName = ch.dataLabels.showCatName = ch.dataLabels.showLegendKey = False
ch.legend = None
ch.x_axis.delete = False; ch.y_axis.delete = False
ch.y_axis.majorGridlines = None
ch.y_axis.number_format = "0"
ch.gapWidth = 40
clean(ch, 8.2, 6.6)
dash.add_chart(ch, f"{L(cc(18))}{T2+1}")
line_item(T2 + 14, 18, 22, 25, "Total budgeted hours", "=A2_Team_Budget!I15", "#,##0", SLATE)
line_item(T2 + 15, 18, 22, 25, "Fee recovery on standard cost", "=A2_Team_Budget!J16", "0%", AMBER)
line_item(T2 + 16, 18, 22, 25, "Partner + manager share of hours", "=A2_Team_Budget!B17", "0%", SLATE)

# P8 5Cs
panel(T2, 27, 8, BODY, " 8. Management letter — 5Cs points (H5)")
put(dash, T2 + 1, cc(27), "Point", font(8, True, MUTED), align("center"), bg=HEAD)
put(dash, T2 + 1, cc(28), "Title", font(8, True, MUTED), to=(T2 + 1, cc(31)), bg=HEAD)
put(dash, T2 + 1, cc(32), "ISA 265", font(8, True, MUTED), align("center"), to=(T2 + 1, cc(33)), bg=HEAD)
put(dash, T2 + 1, cc(34), "5Cs", font(8, True, MUTED), align("center"), bg=HEAD)
for k in range(len(F)):
    r = T2 + 2 + k
    src = F_FIRST + k
    put(dash, r, cc(27), f"=H5_Findings_5Cs!A{src}", font(8, True), align("center"))
    put(dash, r, cc(28), f"=H5_Findings_5Cs!E{src}", font(8), al=align(shrink=True), to=(r, cc(31)))
    put(dash, r, cc(32), f'=IF(H5_Findings_5Cs!F{src}="Significant deficiency","Significant","Deficiency")', font(8, True),
        align("center"), to=(r, cc(33)))
    put(dash, r, cc(34), f'=IF(H5_Findings_5Cs!S{src}="Complete","✓","!")', font(9, True, GREEN), align("center"))
    for c in range(cc(27), cc(35)):
        dash.cell(r, c).border = BOTTOM
status_cf(dash, f"{L(cc(32))}{T2+2}:{L(cc(32))}{T2+9}", {"Significant": RED, "Deficiency": AMBER}, size=8)
line_item(T2 + 12, 27, 31, 34, "Points written in full 5Cs", f'={KP["5Cs points complete"]}&" of "&{KP["5Cs points"]}', None, GREEN)
line_item(T2 + 13, 27, 31, 34, "Written (ML 1–6) / verbal (V 1–2)",
          f'=COUNTIF({FR("O")},"Written*")&" / "&COUNTIF({FR("O")},"Verbal*")', None, SLATE)
put(dash, T2 + 15, cc(27), "Open H6 to print any point for the management letter.", font(7, italic=True, color=MUTED), to=(T2 + 15, cc(34)))

# P9 review points
panel(T3, 0, 17, 15, " 9. Compliance review points (H7)")
hdr = [("Ref", 0, 1), ("Issue", 2, 8), ("Standard", 9, 13), ("Severity", 14, 15), ("Status", 16, 16)]
for h, a, b in hdr:
    put(dash, T3 + 1, cc(a), h, font(8, True, MUTED), align("center" if a not in (2, 9) else "left"), to=(T3 + 1, cc(b)), bg=HEAD)
for k in range(len(RP)):
    r = T3 + 2 + k
    dash.row_dimensions[r].height = 21
    src = R_FIRST + k
    for (h, a, b), col in zip(hdr, "ABDEK"):
        put(dash, r, cc(a), f"=H7_Review_Points!{col}{src}", font(7 if col == "D" else 8, col == "A"),
            align("left" if col in "BD" else "center", wrap=True, shrink=col == "B"), to=(r, cc(b)))
    for c in range(cc(0), cc(17)):
        dash.cell(r, c).border = BOTTOM
status_cf(dash, f"{L(cc(14))}{T3+2}:{L(cc(14))}{T3+12}", {"High": RED, "Medium": AMBER, "Low": GREEN}, size=8)
status_cf(dash, f"{L(cc(16))}{T3+2}:{L(cc(16))}{T3+12}", {"Open": AMBER, "Cleared": GREEN}, size=8)

# P10 timeline
panel(T3, 18, 17, 15, " 10. Engagement timetable (A2 milestones)")
put(dash, T3 + 1, cc(18), "Milestone", font(8, True, MUTED), to=(T3 + 1, cc(27)), bg=HEAD)
put(dash, T3 + 1, cc(28), "Date", font(8, True, MUTED), align("center"), to=(T3 + 1, cc(31)), bg=HEAD)
put(dash, T3 + 1, cc(32), "Standard", font(8, True, MUTED), align("center"), to=(T3 + 1, cc(34)), bg=HEAD)
for k, src in enumerate([22, 23, 26, 27, 29, 31, 33, 35, 36, 37]):
    r = T3 + 2 + k
    put(dash, r, cc(18), f"=A2_Team_Budget!A{src}", font(8), al=align(shrink=True), to=(r, cc(27)))
    put(dash, r, cc(28), f"=A2_Team_Budget!B{src}", font(8, True), align("center", shrink=True), to=(r, cc(31)))
    put(dash, r, cc(32), f"=A2_Team_Budget!C{src}", font(8, color=MUTED), align("center", shrink=True), to=(r, cc(34)))
    for c in range(cc(18), cc(35)):
        dash.cell(r, c).border = BOTTOM

dash.page_setup.orientation = "landscape"
dash.page_setup.fitToWidth = 1; dash.page_setup.fitToHeight = 1
dash.sheet_properties.pageSetUpPr.fitToPage = True
dash.print_area = f"A1:{L(cc(NC))}{ROWS}"
wb.active = 0
for ws_ in wb.worksheets:
    ws_.sheet_view.tabSelected = ws_.title == "H0_Dashboard"

wb.save(OUT)
print("saved", OUT, "sheets:", len(wb.sheetnames))
