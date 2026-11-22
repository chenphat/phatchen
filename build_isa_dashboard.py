"""Builds ISA_Audit_Dashboard.xlsx - a formula-driven ISA audit process dashboard.

All figures are illustrative teaching data. Run:  python build_isa_dashboard.py
"""
from datetime import date

from openpyxl import Workbook
from openpyxl.chart import BarChart, DoughnutChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule, IconSetRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

OUT = "ISA_Audit_Dashboard.xlsx"
LAST = 300  # data ranges run to this row so new rows are picked up

NAVY, NAVY2, BLUE, BG, WHITE = "0B1F4B", "132C63", "2563EB", "EEF2F8", "FFFFFF"
GREEN, AMBER, RED, PURPLE, TEAL, GREY = "16A34A", "F59E0B", "DC2626", "7C3AED", "0D9488", "94A3B8"
INK, MUTED, LINE, HEAD = "0F172A", "64748B", "E2E8F0", "F1F5F9"
INPUT_BLUE = "0000FF"


def font(size=10, bold=False, color=INK, italic=False, underline=None):
    return Font(name="Arial", size=size, bold=bold, color=color, italic=italic, underline=underline)


def fill(color):
    return PatternFill("solid", start_color=color, end_color=color)


def align(h="left", v="center", wrap=False, rot=0):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap, text_rotation=rot)


THIN = Side(style="thin", color=LINE)
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BOTTOM = Border(bottom=THIN)


def paint(ws, r1, c1, r2, c2, color, border=None):
    f = fill(color)
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(r, c)
            cell.fill = f
            if border:
                cell.border = border


def put(ws, r, c, value, fnt=None, al=None, to=None, fmt=None, bg=None):
    """Write value at (r,c); optionally merge through to=(r2,c2)."""
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
    return cell


wb = Workbook()
dash = wb.active
dash.title = "Dashboard"
guide = wb.create_sheet("Guide")
st = wb.create_sheet("Settings")
mat = wb.create_sheet("Materiality")
prog = wb.create_sheet("Audit Programme")
risk = wb.create_sheet("Risk Assessment")
ctl = wb.create_sheet("Control Testing")
sub = wb.create_sheet("Substantive Testing")
pbc = wb.create_sheet("PBC Tracker")
mis = wb.create_sheet("Misstatements")
fnd = wb.create_sheet("Findings")
isa = wb.create_sheet("ISA Map")
calc = wb.create_sheet("Calc")

PHASES = ["Pre-engagement", "Planning", "Risk Assessment", "Control Testing",
          "Substantive Testing", "Completion", "Reporting"]
STATUSES = ["Not Started", "In Progress", "Under Review", "Completed"]
TEAM = ["Partner", "Manager", "Senior", "Associate 1", "Associate 2", "IT Specialist"]
SEVERITY = ["Critical", "High", "Medium", "Low"]
FSTATUS = ["Open", "In Progress", "Closed"]
CYCLES = ["Revenue", "Purchases", "Inventory", "Payroll", "Fixed Assets", "Treasury",
          "Financial Close", "IT General"]
LEVELS = ["High", "Medium", "Low"]
YESNO = ["Yes", "No"]
DESIGN = ["Effective", "Partially Effective", "Ineffective"]
EVIDENCE = ["Strong", "Adequate", "Weak"]
MTYPE = ["Factual", "Judgmental", "Projected"]
BENCH = ["Profit before tax", "Revenue", "Total assets", "Equity"]

# ---------------------------------------------------------------- Settings
st.sheet_view.showGridLines = False
put(st, 1, 1, "SETTINGS", font(14, True, WHITE), bg=NAVY, to=(1, 16))
st.row_dimensions[1].height = 26
put(st, 2, 1, "Blue text = input you may change. Lists below feed the drop-downs on every sheet.",
    font(9, italic=True, color=MUTED), to=(2, 16))
settings = [
    ("Client name", "ABC Manufacturing Co., Ltd. (illustrative)", None),
    ("Financial year end", date(2026, 6, 30), "dd-mmm-yyyy"),
    ("Reporting (as-at) date", date(2026, 11, 22), "dd-mmm-yyyy"),
    ("Audit start date", date(2026, 7, 1), "dd-mmm-yyyy"),
    ("Planned auditor's report date", date(2026, 12, 18), "dd-mmm-yyyy"),
    ("Tolerable deviation rate (controls)", 0.05, "0.0%"),
    ("Financial reporting framework", "IFRS", None),
]
for i, (k, v, f) in enumerate(settings):
    r = 3 + i
    put(st, r, 1, k, font(10, True))
    put(st, r, 2, v, font(10, color=INPUT_BLUE), fmt=f)
    st.cell(r, 1).border = st.cell(r, 2).border = BOX
put(st, 11, 1, "Notes", font(10, True))
notes = [
    "Reporting date drives every overdue / ageing test. Type =TODAY() in B5 for live tracking.",
    "Tolerable deviation rate: assumption used on 'Control Testing' to rate operating effectiveness.",
    "All client data in this workbook is illustrative sample data - overwrite it with your own.",
]
for i, n in enumerate(notes):
    put(st, 12 + i, 1, n, font(9, color=MUTED), al=align(wrap=True), to=(12 + i, 2))
    st.row_dimensions[12 + i].height = 26
lists = [("Audit phases", PHASES), ("Procedure status", STATUSES), ("Audit team", TEAM),
         ("Finding severity", SEVERITY), ("Finding status", FSTATUS), ("Cycles", CYCLES),
         ("Risk level", LEVELS), ("Yes / No", YESNO), ("Design rating", DESIGN),
         ("Evidence quality", EVIDENCE), ("Misstatement type", MTYPE),
         ("Phase filter", ["All"] + PHASES), ("Team filter", ["All"] + TEAM)]
LISTREF = {}
for j, (name, items) in enumerate(lists):
    c = 4 + j
    put(st, 3, c, name, font(9, True, WHITE), align("center", wrap=True), bg=NAVY2)
    for i, it in enumerate(items):
        put(st, 4 + i, c, it, font(9, color=INPUT_BLUE)).border = BOX
    LISTREF[name] = f"Settings!${L(c)}$4:${L(c)}${3 + len(items)}"
    st.column_dimensions[L(c)].width = 18
st.column_dimensions["A"].width = 34
st.column_dimensions["B"].width = 40
st.column_dimensions["C"].width = 3
st.row_dimensions[3].height = 24
RD = "Settings!$B$5"


def add_dv(ws, listname, col, r1=6, r2=LAST):
    dv = DataValidation(type="list", formula1="=" + LISTREF[listname], allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"{col}{r1}:{col}{r2}")


def sheet_header(ws, title, subtitle, ncols):
    ws.sheet_view.showGridLines = False
    put(ws, 1, 1, title, font(14, True, WHITE), bg=NAVY, to=(1, ncols))
    ws.row_dimensions[1].height = 26
    put(ws, 2, 1, subtitle, font(9, italic=True, color=MUTED), al=align(wrap=True), to=(2, ncols))
    ws.row_dimensions[2].height = 26
    put(ws, 3, 1, "◄ Back to Dashboard", font(9, True, BLUE, underline="single"))
    ws.cell(3, 1).hyperlink = "#'Dashboard'!A1"


def table(ws, headers, rows, widths, formats=None, formula_cols=(), hrow=5, wrap_cols=()):
    """headers: list; rows: list of lists (strings starting '=' are formulas)."""
    formats = formats or {}
    for j, h in enumerate(headers, 1):
        put(ws, hrow, j, h, font(9, True, WHITE), align("center", wrap=True), bg=NAVY2).border = BOX
        ws.column_dimensions[L(j)].width = widths[j - 1]
    ws.row_dimensions[hrow].height = 30
    for i, row in enumerate(rows):
        r = hrow + 1 + i
        for j, v in enumerate(row, 1):
            is_f = j in formula_cols
            cell = put(ws, r, j, v, font(9, color=INK if is_f else INPUT_BLUE),
                       align("left" if isinstance(v, str) and not is_f else "center",
                             wrap=j in wrap_cols), fmt=formats.get(j))
            if isinstance(v, str) and is_f:
                cell.alignment = align("center")
            cell.border = BOX
            if is_f:
                cell.fill = fill(HEAD)
    ws.freeze_panes = ws.cell(hrow + 1, 2)
    ws.auto_filter.ref = f"A{hrow}:{L(len(headers))}{hrow + len(rows)}"


def status_cf(ws, rng, mapping):
    for text, color in mapping.items():
        ws.conditional_formatting.add(
            rng, CellIsRule(operator="equal", formula=[f'"{text}"'],
                            font=Font(name="Arial", size=9, bold=True, color=WHITE), fill=fill(color)))


D = date

# ---------------------------------------------------------------- Materiality (ISA 320)
sheet_header(mat, "MATERIALITY  -  ISA 320",
             "Overall materiality, performance materiality and the clearly-trivial threshold (ISA 450). "
             "Blue = inputs; grey cells are formulas.", 5)
put(mat, 5, 1, "1. Draft financial statement data (USD) - illustrative, entered by user", font(10, True))
fin = [("Profit before tax", 4150000), ("Revenue", 48500000), ("Total assets", 39800000), ("Equity", 21300000)]
for j, h in enumerate(["Benchmark", "Amount (USD)", "% applied", "Indicative materiality (USD)", "Common range (guidance)"], 1):
    put(mat, 6, j, h, font(9, True, WHITE), align("center", wrap=True), bg=NAVY2).border = BOX
ranges = ["5% - 10% (profit-oriented entities)", "0.5% - 1%", "1% - 2%", "2% - 5%"]
pcts = [0.05, 0.005, 0.01, 0.02]
for i, (k, v) in enumerate(fin):
    r = 7 + i
    put(mat, r, 1, k, font(10)).border = BOX
    put(mat, r, 2, v, font(10, color=INPUT_BLUE), align("right"), fmt="#,##0").border = BOX
    put(mat, r, 3, pcts[i], font(10, color=INPUT_BLUE), align("center"), fmt="0.0%").border = BOX
    put(mat, r, 4, f"=B{r}*C{r}", font(10), align("right"), fmt="#,##0", bg=HEAD).border = BOX
    put(mat, r, 5, ranges[i], font(9, color=MUTED)).border = BOX
put(mat, 12, 1, "2. Materiality determination", font(10, True))
rows = [
    ("Selected benchmark", "Profit before tax", None, True, "Choose from the list - stable, profit-oriented entity, users focus on earnings."),
    ("Benchmark amount (USD)", "=INDEX($B$7:$B$10,MATCH($B$13,$A$7:$A$10,0))", "#,##0", False, ""),
    ("Percentage applied", "=INDEX($C$7:$C$10,MATCH($B$13,$A$7:$A$10,0))", "0.0%", False, ""),
    ("Overall materiality (USD)", "=ROUNDDOWN(B14*B15,-3)", "#,##0", False, "Rounded down to the nearest 1,000 (ISA 320.10)."),
    ("Performance materiality %", 0.75, "0%", True, "Assumption: 75% - moderate risk, recurring engagement (typical range 50%-75%)."),
    ("Performance materiality (USD)", "=B16*B17", "#,##0", False, "ISA 320.11 - used for scoping and sample sizes."),
    ("Clearly trivial threshold %", 0.05, "0%", True, "Assumption: 5% of overall materiality (ISA 450.A2)."),
    ("Clearly trivial threshold (USD)", "=B16*B19", "#,##0", False, "Misstatements below this are not accumulated."),
]
for i, (k, v, f, inp, note) in enumerate(rows):
    r = 13 + i
    put(mat, r, 1, k, font(10, True)).border = BOX
    c = put(mat, r, 2, v, font(10, bold=not inp, color=INPUT_BLUE if inp else INK), align("right"), fmt=f,
            bg=None if inp else HEAD)
    c.border = BOX
    put(mat, r, 3, note, font(9, italic=True, color=MUTED), to=(r, 5))
dv = DataValidation(type="list", formula1="=$A$7:$A$10")
mat.add_data_validation(dv)
dv.add("B13")
for col, w in zip("ABCDE", [34, 22, 12, 26, 44]):
    mat.column_dimensions[col].width = w
OM, PM, TRIV = "Materiality!$B$16", "Materiality!$B$18", "Materiality!$B$20"

# ---------------------------------------------------------------- Audit Programme
P = [  # phase idx, ISA, procedure, owner, start, due, status, completed, budget, actual, wp
    (0, "ISA 210", "Evaluate client continuance and integrity of management", "Partner", D(2026,7,1), D(2026,7,6), "Completed", D(2026,7,5), 4, 4, "A1"),
    (0, "ISA 220", "Confirm independence and ethical requirements for the engagement team", "Manager", D(2026,7,1), D(2026,7,8), "Completed", D(2026,7,8), 3, 3, "A2"),
    (0, "ISA 210", "Agree terms of engagement and obtain signed engagement letter", "Partner", D(2026,7,6), D(2026,7,13), "Completed", D(2026,7,10), 3, 2, "A3"),
    (0, "ISA 220", "Assign engagement team; assess competence, time and resources", "Partner", D(2026,7,8), D(2026,7,15), "Completed", D(2026,7,15), 2, 2, "A4"),
    (0, "ISA 300", "Review prior-year working papers and points carried forward", "Senior", D(2026,7,13), D(2026,7,20), "Completed", D(2026,7,21), 6, 7, "A5"),
    (1, "ISA 300", "Develop overall audit strategy and timetable", "Manager", D(2026,7,27), D(2026,8,5), "Completed", D(2026,8,4), 8, 9, "B1"),
    (1, "ISA 315", "Understand the entity, its environment and the reporting framework", "Senior", D(2026,8,3), D(2026,8,14), "Completed", D(2026,8,14), 16, 18, "B2"),
    (1, "ISA 320", "Determine overall and performance materiality", "Manager", D(2026,8,10), D(2026,8,14), "Completed", D(2026,8,13), 4, 4, "B3"),
    (1, "ISA 520", "Perform preliminary analytical procedures", "Associate 1", D(2026,8,10), D(2026,8,19), "Completed", D(2026,8,19), 10, 11, "B4"),
    (1, "ISA 250", "Understand the legal and regulatory framework", "Senior", D(2026,8,12), D(2026,8,21), "Completed", D(2026,8,20), 5, 5, "B5"),
    (1, "ISA 550", "Identify related parties and related party transactions", "Associate 2", D(2026,8,12), D(2026,8,21), "Completed", D(2026,8,24), 6, 7, "B6"),
    (1, "ISA 300", "Prepare detailed audit plan and hold team planning meeting", "Manager", D(2026,8,17), D(2026,8,26), "Completed", D(2026,8,26), 6, 6, "B7"),
    (1, "ISA 260", "Communicate planned scope and timing to those charged with governance", "Partner", D(2026,8,24), D(2026,8,31), "Completed", D(2026,8,31), 3, 3, "B8"),
    (2, "ISA 315", "Walkthroughs - revenue and receivables process", "Associate 1", D(2026,8,31), D(2026,9,8), "Completed", D(2026,9,8), 12, 13, "C1"),
    (2, "ISA 315", "Walkthroughs - purchases, payables and payroll processes", "Associate 2", D(2026,8,31), D(2026,9,8), "Completed", D(2026,9,9), 12, 12, "C2"),
    (2, "ISA 315", "Walkthroughs - inventory and product costing process", "Senior", D(2026,8,31), D(2026,9,10), "Completed", D(2026,9,10), 14, 16, "C3"),
    (2, "ISA 315", "Understand the IT environment and general IT controls", "IT Specialist", D(2026,9,1), D(2026,9,11), "Completed", D(2026,9,11), 12, 12, "C4"),
    (2, "ISA 240", "Team fraud discussion, management inquiries and fraud risk factors", "Manager", D(2026,9,7), D(2026,9,14), "Completed", D(2026,9,14), 6, 6, "C5"),
    (2, "ISA 315", "Identify and assess risks of material misstatement (FS and assertion level)", "Manager", D(2026,9,10), D(2026,9,18), "Completed", D(2026,9,18), 10, 12, "C6"),
    (2, "ISA 540", "Understand accounting estimates (ECL, NRV, provisions)", "Senior", D(2026,9,10), D(2026,9,18), "Completed", D(2026,9,17), 6, 6, "C7"),
    (2, "ISA 330", "Design responses to assessed risks; finalise audit programme", "Manager", D(2026,9,16), D(2026,9,23), "Completed", D(2026,9,23), 8, 8, "C8"),
    (3, "ISA 330", "Test operating effectiveness - revenue controls", "Associate 1", D(2026,9,24), D(2026,10,5), "Completed", D(2026,10,5), 16, 17, "D1"),
    (3, "ISA 330", "Test operating effectiveness - purchases and payables controls", "Associate 2", D(2026,9,24), D(2026,10,5), "Completed", D(2026,10,6), 16, 16, "D2"),
    (3, "ISA 330", "Test operating effectiveness - inventory and costing controls", "Senior", D(2026,9,28), D(2026,10,9), "Completed", D(2026,10,9), 18, 21, "D3"),
    (3, "ISA 330", "Test operating effectiveness - payroll controls", "Associate 2", D(2026,10,5), D(2026,10,12), "Completed", D(2026,10,12), 10, 10, "D4"),
    (3, "ISA 330", "Test IT general controls (access, change management, backup)", "IT Specialist", D(2026,9,28), D(2026,11,13), "In Progress", None, 20, 16, "D5"),
    (3, "ISA 330", "Test operating effectiveness - treasury and financial close controls", "Associate 1", D(2026,10,12), D(2026,10,21), "Completed", D(2026,10,22), 12, 12, "D6"),
    (3, "ISA 265", "Evaluate control deficiencies and revise risk assessment", "Manager", D(2026,10,19), D(2026,10,28), "Completed", D(2026,10,28), 6, 7, "D7"),
    (4, "ISA 505", "Send and follow up bank and receivable confirmations", "Associate 1", D(2026,10,19), D(2026,11,13), "In Progress", None, 14, 12, "E1"),
    (4, "ISA 501", "Inventory existence - reconcile year-end count attended to final listing", "Senior", D(2026,10,19), D(2026,10,30), "Completed", D(2026,10,30), 16, 18, "E2"),
    (4, "ISA 500", "Inventory valuation - standard cost, overhead absorption and NRV", "Senior", D(2026,10,26), D(2026,11,25), "Under Review", None, 20, 22, "E3"),
    (4, "ISA 530", "Test of details - revenue occurrence and cut-off", "Associate 1", D(2026,10,26), D(2026,11,10), "Completed", D(2026,11,10), 18, 19, "E4"),
    (4, "ISA 500", "Test of details - purchases, payables and unrecorded liabilities", "Associate 2", D(2026,10,26), D(2026,11,10), "Completed", D(2026,11,12), 16, 17, "E5"),
    (4, "ISA 520", "Substantive analytical procedures - payroll and operating expenses", "Associate 2", D(2026,11,2), D(2026,11,13), "Completed", D(2026,11,13), 10, 9, "E6"),
    (4, "ISA 500", "PPE - additions, disposals, depreciation and impairment indicators", "Associate 1", D(2026,11,9), D(2026,11,20), "In Progress", None, 12, 8, "E7"),
    (4, "ISA 540", "Audit accounting estimates - ECL allowance and provisions", "Senior", D(2026,11,9), D(2026,11,27), "In Progress", None, 14, 6, "E8"),
    (4, "ISA 240", "Test journal entries and management override of controls", "Manager", D(2026,11,9), D(2026,11,27), "In Progress", None, 12, 5, "E9"),
    (4, "ISA 550", "Test related party transactions and disclosures", "Associate 2", D(2026,11,16), D(2026,11,30), "In Progress", None, 8, 2, "E10"),
    (4, "ISA 500", "Taxation, borrowings and equity", "Associate 1", D(2026,11,16), D(2026,12,2), "Not Started", None, 10, 0, "E11"),
    (5, "ISA 560", "Subsequent events review", "Senior", D(2026,11,30), D(2026,12,9), "Not Started", None, 6, 0, "F1"),
    (5, "ISA 570", "Evaluate management's going concern assessment", "Manager", D(2026,11,30), D(2026,12,9), "Not Started", None, 6, 0, "F2"),
    (5, "ISA 450", "Evaluate misstatements; summary of uncorrected misstatements", "Manager", D(2026,12,1), D(2026,12,10), "Not Started", None, 5, 0, "F3"),
    (5, "ISA 520", "Final analytical review of the financial statements", "Senior", D(2026,12,3), D(2026,12,10), "Not Started", None, 5, 0, "F4"),
    (5, "ISA 580", "Obtain written representations from management", "Manager", D(2026,12,7), D(2026,12,14), "Not Started", None, 2, 0, "F5"),
    (5, "ISA 220", "Engagement partner review and engagement quality review", "Partner", D(2026,12,7), D(2026,12,15), "Not Started", None, 10, 0, "F6"),
    (6, "ISA 700", "Form the opinion and draft the auditor's report", "Partner", D(2026,12,10), D(2026,12,16), "Not Started", None, 6, 0, "G1"),
    (6, "ISA 701 / 705 / 706", "Determine key audit matters, modifications and emphasis of matter", "Partner", D(2026,12,10), D(2026,12,16), "Not Started", None, 4, 0, "G2"),
    (6, "ISA 720", "Read other information in the annual report", "Manager", D(2026,12,10), D(2026,12,16), "Not Started", None, 3, 0, "G3"),
    (6, "ISA 260 / 265", "Report to those charged with governance; issue management letter", "Manager", D(2026,12,14), D(2026,12,18), "Not Started", None, 6, 0, "G4"),
    (6, "ISA 700", "Sign and issue the auditor's report", "Partner", D(2026,12,18), D(2026,12,18), "Not Started", None, 1, 0, "G5"),
    (6, "ISA 230", "Assemble the final audit file (within 60 days of report date)", "Senior", D(2026,12,18), D(2027,2,16), "Not Started", None, 6, 0, "G6"),
]
sheet_header(prog, "AUDIT PROGRAMME  -  ISA 300 / 330",
             "One row per audit procedure across the seven phases. Enter blue columns; 'Overdue?' and 'Days overdue' are formulas "
             "(copy them down when adding rows). Example rows show the expected format.", 14)
rows = []
for i, p in enumerate(P):
    r = 6 + i
    rows.append([f"AP-{i+1:03d}", PHASES[p[0]], p[1], p[2], p[3], p[4], p[5], p[6], p[7], p[8], p[9], p[10],
                 f'=IF(A{r}="","",IF(AND(H{r}<>"Completed",G{r}<{RD}),"Yes","No"))',
                 f'=IF(M{r}="Yes",{RD}-G{r},0)'])
table(prog, ["ID", "Audit phase", "ISA ref.", "Audit procedure", "Owner", "Planned start", "Due date", "Status",
             "Completed date", "Budget hrs", "Actual hrs", "WP ref.", "Overdue?", "Days overdue"],
      rows, [9, 19, 16, 62, 13, 12, 12, 13, 13, 9, 9, 8, 10, 9],
      {6: "dd-mmm-yy", 7: "dd-mmm-yy", 9: "dd-mmm-yy", 14: "0"}, formula_cols=(13, 14))
add_dv(prog, "Audit phases", "B"); add_dv(prog, "Audit team", "E"); add_dv(prog, "Procedure status", "H")
STAT_COL = {"Completed": GREEN, "In Progress": BLUE, "Under Review": PURPLE, "Not Started": GREY}
status_cf(prog, f"H6:H{LAST}", STAT_COL)
status_cf(prog, f"M6:M{LAST}", {"Yes": RED})
PR = lambda col: f"'Audit Programme'!${col}$6:${col}${LAST}"

# ---------------------------------------------------------------- Risk Assessment (ISA 315)
R = [  # cycle, fs area, risk, assertion, L, M, control risk, significant, fraud, response, owner, status
    ("Revenue", "Revenue", "Revenue recognised before control transfers to the customer (presumed fraud risk)", "Occurrence / Cut-off", 4, 5, "Medium", "Yes", "Yes", "Extended cut-off testing +/- 10 days; test of details on year-end sales", "Associate 1", "Completed"),
    ("Financial Close", "Journal entries", "Management override of controls through manual journals", "All assertions", 3, 5, "High", "Yes", "Yes", "Journal entry testing; review estimates for bias; unusual transactions", "Manager", "In Progress"),
    ("Inventory", "Inventory", "Overhead absorption rates misapplied; cost variances not allocated to closing inventory", "Valuation", 4, 4, "High", "Yes", "No", "Recompute absorption; test variance allocation; compare standard to actual cost", "Senior", "In Progress"),
    ("Inventory", "Inventory", "Slow-moving and obsolete stock not written down to net realisable value", "Valuation", 4, 3, "Medium", "No", "No", "NRV testing against post-year-end selling prices; ageing analysis", "Senior", "In Progress"),
    ("Inventory", "Work in progress", "Stage of completion of WIP estimated incorrectly", "Valuation", 3, 3, "Medium", "No", "No", "Inspect production records; recalculate conversion cost in WIP", "Senior", "In Progress"),
    ("Inventory", "Inventory", "Quantities misstated - count errors and goods held by third parties", "Existence", 2, 4, "Low", "No", "No", "Count attendance, test counts, third-party confirmations", "Senior", "Completed"),
    ("Revenue", "Trade receivables", "Expected credit loss allowance understated on overdue balances", "Valuation", 3, 4, "Medium", "Yes", "No", "Test ECL model inputs, ageing accuracy and subsequent receipts", "Senior", "In Progress"),
    ("Revenue", "Trade receivables", "Fictitious or disputed receivable balances", "Existence", 2, 4, "Medium", "No", "No", "External confirmations; alternative procedures for non-replies", "Associate 1", "In Progress"),
    ("Purchases", "Trade payables", "Unrecorded liabilities at year end", "Completeness", 3, 3, "Medium", "No", "No", "Search for unrecorded liabilities; supplier statement reconciliations", "Associate 2", "Completed"),
    ("Purchases", "Cost of sales", "Purchases recorded in the wrong period", "Cut-off", 2, 3, "Low", "No", "No", "Cut-off testing of goods received notes around year end", "Associate 2", "Completed"),
    ("Payroll", "Payroll expense", "Ghost employees or unauthorised pay rate changes", "Occurrence", 2, 3, "Low", "No", "Yes", "Analytical procedures; test starters, leavers and rate changes", "Associate 2", "Completed"),
    ("Fixed Assets", "PPE", "Repairs capitalised; inappropriate useful lives", "Classification / Valuation", 2, 3, "Medium", "No", "No", "Vouch additions; reperform depreciation", "Associate 1", "In Progress"),
    ("Fixed Assets", "PPE", "Impairment indicators for idle production line not assessed", "Valuation", 2, 4, "Medium", "No", "No", "Review impairment assessment; inspect idle assets", "Senior", "In Progress"),
    ("Treasury", "Borrowings", "Loan covenant breach not identified; current / non-current misclassified", "Presentation", 2, 5, "Medium", "No", "No", "Recalculate covenants; inspect loan agreements; bank confirmation", "Associate 1", "Not Started"),
    ("Treasury", "Cash and bank", "Misappropriation of cash; unreconciled bank items", "Existence", 1, 4, "Low", "No", "Yes", "Bank confirmations; test year-end reconciliations", "Associate 1", "In Progress"),
    ("Financial Close", "Related parties", "Undisclosed related party transactions", "Completeness / Disclosure", 3, 4, "High", "Yes", "No", "Inspect minutes and registers; test identified transactions (ISA 550)", "Associate 2", "In Progress"),
    ("IT General", "All areas", "Inappropriate user access to the ERP affects reliability of system reports", "All assertions", 3, 4, "High", "No", "No", "ITGC testing; extend substantive work where ITGCs fail", "IT Specialist", "In Progress"),
    ("Financial Close", "Taxation", "Errors in current and deferred tax computation", "Accuracy", 2, 2, "Medium", "No", "No", "Reperform tax computation; reconcile effective tax rate", "Associate 1", "Not Started"),
    ("Financial Close", "Going concern", "Material uncertainty over going concern not disclosed", "Disclosure", 1, 5, "Low", "No", "No", "Review cash-flow forecast and financing facilities (ISA 570)", "Manager", "Not Started"),
    ("Financial Close", "Provisions", "Provision for pending legal claim omitted", "Completeness", 2, 3, "Medium", "No", "No", "Lawyer's confirmation; review legal expense accounts", "Senior", "In Progress"),
]
sheet_header(risk, "RISK ASSESSMENT  -  ISA 315 (Revised 2019) / ISA 240",
             "Likelihood and magnitude are scored 1-5. Score, inherent risk and RoMM are formulas: score >= 15 High, >= 8 Medium, "
             "otherwise Low; RoMM combines inherent and control risk. Planned responses follow ISA 330.", 16)
rows = []
lv = '{"Low","Medium","High"}'
for i, x in enumerate(R):
    r = 6 + i
    rows.append([f"R-{i+1:02d}", x[0], x[1], x[2], x[3], x[4], x[5],
                 f'=IF(A{r}="","",F{r}*G{r})',
                 f'=IF(A{r}="","",IF(H{r}>=15,"High",IF(H{r}>=8,"Medium","Low")))',
                 x[6],
                 f'=IF(A{r}="","",IF(MATCH(I{r},{lv},0)*MATCH(J{r},{lv},0)>=6,"High",IF(MATCH(I{r},{lv},0)*MATCH(J{r},{lv},0)>=3,"Medium","Low")))',
                 x[7], x[8], x[9], x[10], x[11]])
table(risk, ["Risk ID", "Cycle", "FS area", "Risk of material misstatement", "Assertion", "Likelihood (1-5)",
             "Magnitude (1-5)", "Inherent score", "Inherent risk", "Control risk", "RoMM", "Significant risk?",
             "Fraud risk? (ISA 240)", "Planned response (ISA 330)", "Owner", "Response status"],
      rows, [8, 15, 16, 58, 20, 10, 10, 9, 10, 10, 10, 11, 11, 58, 13, 13],
      formula_cols=(8, 9, 11))
add_dv(risk, "Cycles", "B"); add_dv(risk, "Risk level", "J"); add_dv(risk, "Yes / No", "L")
add_dv(risk, "Yes / No", "M"); add_dv(risk, "Audit team", "O"); add_dv(risk, "Procedure status", "P")
dv = DataValidation(type="whole", operator="between", formula1="1", formula2="5")
risk.add_data_validation(dv); dv.add(f"F6:G{LAST}")
LVL_COL = {"High": RED, "Medium": AMBER, "Low": GREEN}
for rng in (f"I6:K{LAST}",):
    status_cf(risk, rng, LVL_COL)
status_cf(risk, f"P6:P{LAST}", STAT_COL)
RR = lambda col: f"'Risk Assessment'!${col}$6:${col}${LAST}"

# ---------------------------------------------------------------- Control Testing (ISA 330)
C = [  # cycle, control, frequency, type, design, sample, tested, exceptions, evidence
    ("Revenue", "Credit limit approval before order acceptance", "Per transaction", "Preventive", "Effective", 25, 25, 0, "Strong"),
    ("Revenue", "Sales invoice matched to delivery note and customer order", "Per transaction", "Preventive", "Effective", 25, 25, 1, "Adequate"),
    ("Revenue", "Monthly receivables ageing review and follow-up", "Monthly", "Detective", "Effective", 3, 3, 0, "Adequate"),
    ("Purchases", "Purchase order approval per delegation of authority", "Per transaction", "Preventive", "Effective", 25, 25, 0, "Strong"),
    ("Purchases", "Three-way match (PO, GRN, invoice) before payment", "Per transaction", "Preventive", "Effective", 25, 25, 2, "Adequate"),
    ("Purchases", "Monthly supplier statement reconciliations", "Monthly", "Detective", "Partially Effective", 3, 3, 1, "Weak"),
    ("Inventory", "Cycle counts and investigation of variances", "Monthly", "Detective", "Effective", 3, 3, 0, "Strong"),
    ("Inventory", "Annual review and approval of standard costs and overhead rates", "Annual", "Preventive", "Partially Effective", 1, 1, 1, "Weak"),
    ("Inventory", "Monthly review of production cost variances", "Monthly", "Detective", "Effective", 3, 3, 0, "Adequate"),
    ("Payroll", "HR approval of new hires and pay rate changes", "Per transaction", "Preventive", "Effective", 15, 15, 0, "Strong"),
    ("Payroll", "Monthly payroll reconciliation to the general ledger", "Monthly", "Detective", "Effective", 3, 3, 0, "Adequate"),
    ("Fixed Assets", "Capital expenditure approval and capitalisation review", "Per transaction", "Preventive", "Effective", 15, 12, 0, "Adequate"),
    ("Treasury", "Monthly bank reconciliations prepared and reviewed", "Monthly", "Detective", "Effective", 3, 3, 0, "Strong"),
    ("Treasury", "Dual authorisation of payments", "Per transaction", "Preventive", "Effective", 25, 25, 0, "Strong"),
    ("Financial Close", "Review and approval of manual journal entries", "Per transaction", "Preventive", "Partially Effective", 25, 25, 3, "Adequate"),
    ("Financial Close", "Monthly management accounts review against budget", "Monthly", "Detective", "Effective", 3, 3, 0, "Adequate"),
    ("IT General", "User access provisioning and periodic access review", "Quarterly", "Preventive", "Ineffective", 4, 4, 2, "Weak"),
    ("IT General", "Change management - approval and testing before release", "Per change", "Preventive", "Effective", 15, 9, 0, "Adequate"),
    ("IT General", "Daily backup and periodic restore testing", "Daily", "Detective", "Effective", 25, 15, 0, "Adequate"),
]
sheet_header(ctl, "CONTROL TESTING  -  ISA 330 (tests of controls) / ISA 530",
             "Enter design rating, sample size, items tested and exceptions. Operating effectiveness compares the deviation rate with the "
             "tolerable rate in Settings (0 deviations = Effective; up to tolerable = Partially Effective; above = Ineffective).", 14)
rows = []
for i, x in enumerate(C):
    r = 6 + i
    rows.append([f"C-{i+1:02d}", x[0], x[1], x[2], x[3], x[4], x[5], x[6], x[7],
                 f'=IF(G{r}=0,0,H{r}/G{r})', f'=IF(H{r}=0,0,I{r}/H{r})',
                 f'=IF(A{r}="","",IF(H{r}=0,"Not Tested",IF(K{r}=0,"Effective",IF(K{r}<=Settings!$B$8,"Partially Effective","Ineffective"))))',
                 x[8],
                 f'=IF(A{r}="","",IF(OR(F{r}="Ineffective",L{r}="Ineffective"),"No reliance - substantive approach",IF(J{r}<1,"Testing in progress",IF(AND(F{r}="Effective",L{r}="Effective"),"Rely on control","Limited reliance"))))'])
table(ctl, ["Control ID", "Cycle", "Control description", "Frequency", "Type", "Design effectiveness", "Sample size",
            "Items tested", "Exceptions", "Sample completion", "Deviation rate", "Operating effectiveness",
            "Evidence quality", "Reliance conclusion"],
      rows, [9, 15, 56, 14, 11, 16, 8, 8, 10, 11, 10, 17, 11, 30],
      {10: "0%", 11: "0.0%"}, formula_cols=(10, 11, 12, 14))
add_dv(ctl, "Cycles", "B"); add_dv(ctl, "Design rating", "F"); add_dv(ctl, "Evidence quality", "M")
EFF_COL = {"Effective": GREEN, "Partially Effective": AMBER, "Ineffective": RED, "Not Tested": GREY}
status_cf(ctl, f"F6:F{LAST}", EFF_COL); status_cf(ctl, f"L6:L{LAST}", EFF_COL)
status_cf(ctl, f"M6:M{LAST}", {"Strong": GREEN, "Adequate": AMBER, "Weak": RED})
CR = lambda col: f"'Control Testing'!${col}$6:${col}${LAST}"

# ---------------------------------------------------------------- Substantive Testing (ISA 330/500)
S = [  # area, cycle, balance, assertions, procedure, isa, planned, tested, exceptions, wp
    ("Revenue", "Revenue", 48500000, "Occurrence, Cut-off, Accuracy", "Test of details + cut-off testing", "ISA 500 / 530", 40, 40, 1, "E4"),
    ("Trade receivables", "Revenue", 7900000, "Existence, Valuation", "External confirmations; ECL testing", "ISA 505 / 540", 30, 22, 0, "E1"),
    ("Inventory", "Inventory", 9600000, "Existence, Valuation", "Count reconciliation; cost and NRV testing", "ISA 501 / 500", 45, 38, 2, "E2-E3"),
    ("Cost of sales", "Purchases", 36200000, "Occurrence, Cut-off, Accuracy", "Test of details; variance analysis", "ISA 500 / 520", 35, 35, 1, "E5"),
    ("Trade payables", "Purchases", 5400000, "Completeness", "Unrecorded liabilities search; supplier statements", "ISA 500", 30, 30, 1, "E5"),
    ("Payroll expense", "Payroll", 6300000, "Occurrence, Accuracy", "Substantive analytical procedures", "ISA 520", 12, 12, 0, "E6"),
    ("Property, plant and equipment", "Fixed Assets", 17800000, "Existence, Valuation", "Vouch additions; reperform depreciation", "ISA 500", 25, 14, 0, "E7"),
    ("Cash and bank", "Treasury", 2250000, "Existence", "Bank confirmations; reconciliations", "ISA 505", 8, 6, 0, "E1"),
    ("Borrowings", "Treasury", 9500000, "Completeness, Presentation", "Confirmations; covenant recalculation", "ISA 505 / 500", 6, 0, 0, "E11"),
    ("Provisions", "Financial Close", 640000, "Completeness, Valuation", "Lawyer's letter; estimate testing", "ISA 501 / 540", 5, 2, 0, "E8"),
    ("Taxation", "Financial Close", 910000, "Accuracy", "Reperform tax computation", "ISA 500", 4, 0, 0, "E11"),
    ("Related party transactions", "Financial Close", 1150000, "Completeness, Disclosure", "Test identified transactions and disclosures", "ISA 550", 10, 3, 0, "E10"),
    ("Equity", "Financial Close", 21300000, "Presentation", "Agree movements to resolutions", "ISA 500", 3, 0, 0, "E11"),
    ("Prepayments", "Purchases", 120000, "Existence", "Analytical review only", "ISA 520", 0, 0, 0, "E5"),
]
sheet_header(sub, "SUBSTANTIVE TESTING  -  ISA 330 / 500 / 505 / 520 / 530",
             "Lead schedule of financial statement areas. Scoping compares the book balance with performance materiality (ISA 320). "
             "Completion and status are formulas driven by sample progress.", 13)
rows = []
for i, x in enumerate(S):
    r = 6 + i
    rows.append([x[0], x[1], x[2],
                 f'=IF(A{r}="","",IF(C{r}>{PM},"In scope - test of details","Below PM - analytical only"))',
                 x[3], x[4], x[5], x[6], x[7], x[8],
                 f'=IF(H{r}=0,0,I{r}/H{r})',
                 f'=IF(A{r}="","",IF(H{r}=0,"Not Applicable",IF(I{r}=0,"Not Started",IF(I{r}<H{r},"In Progress","Completed"))))',
                 x[9]])
table(sub, ["FS area", "Cycle", "Book balance (USD)", "Scoping vs performance materiality", "Key assertions",
            "Primary substantive procedure", "ISA ref.", "Sample planned", "Sample tested", "Exceptions",
            "Completion", "Status", "WP ref."],
      rows, [28, 15, 16, 28, 28, 42, 14, 9, 9, 10, 11, 13, 8],
      {3: "#,##0", 11: "0%"}, formula_cols=(4, 11, 12))
add_dv(sub, "Cycles", "B")
status_cf(sub, f"L6:L{LAST}", {**STAT_COL, "Not Applicable": GREY})
SR = lambda col: f"'Substantive Testing'!${col}$6:${col}${LAST}"

# ---------------------------------------------------------------- PBC tracker (ISA 500)
E = [
    ("Signed trial balance and general ledger at 30 June 2026", "Financial Close", D(2026,8,3), D(2026,8,10), D(2026,8,10)),
    ("Draft financial statements and disclosure notes", "Financial Close", D(2026,10,12), D(2026,10,30), D(2026,11,4)),
    ("Sales ledger, invoice listing and June / July delivery notes", "Revenue", D(2026,10,19), D(2026,10,26), D(2026,10,26)),
    ("Receivables ageing and ECL calculation", "Revenue", D(2026,10,19), D(2026,10,30), D(2026,11,2)),
    ("Customer confirmation replies (second requests)", "Revenue", D(2026,11,2), D(2026,11,16), None),
    ("Year-end inventory count sheets and final inventory listing", "Inventory", D(2026,10,12), D(2026,10,19), D(2026,10,19)),
    ("Standard cost cards, bill of materials and overhead absorption workings", "Inventory", D(2026,10,19), D(2026,10,30), D(2026,11,9)),
    ("Production variance reports and year-end variance allocation", "Inventory", D(2026,10,26), D(2026,11,13), None),
    ("Supplier statements and payables ledger", "Purchases", D(2026,10,19), D(2026,10,26), D(2026,10,27)),
    ("Payroll registers and headcount reconciliation", "Payroll", D(2026,10,26), D(2026,11,2), D(2026,11,2)),
    ("Fixed asset register with additions and disposals support", "Fixed Assets", D(2026,11,2), D(2026,11,13), D(2026,11,16)),
    ("Bank statements, reconciliations and confirmation authorisations", "Treasury", D(2026,10,19), D(2026,10,26), D(2026,10,23)),
    ("Loan agreements and covenant compliance calculations", "Treasury", D(2026,11,9), D(2026,11,25), None),
    ("ERP user access listing and change logs", "IT General", D(2026,9,21), D(2026,10,5), D(2026,10,16)),
    ("Related party register and board minutes", "Financial Close", D(2026,11,9), D(2026,11,20), None),
    ("Tax computation and correspondence with the tax authority", "Financial Close", D(2026,11,16), D(2026,11,30), None),
    ("Cash-flow forecast for going concern assessment", "Financial Close", D(2026,11,23), D(2026,12,4), None),
    ("Lawyer's confirmation letter", "Financial Close", D(2026,11,9), D(2026,11,27), None),
]
sheet_header(pbc, "PBC / EVIDENCE TRACKER  -  ISA 500",
             "Documents 'provided by client'. Leave 'Received date' blank until the item arrives; status and days overdue are formulas.", 8)
rows = []
for i, x in enumerate(E):
    r = 6 + i
    rows.append([f"PBC-{i+1:02d}", x[0], x[1], x[2], x[3], x[4],
                 f'=IF(A{r}="","",IF(F{r}<>"","Received",IF(E{r}<{RD},"Overdue","Outstanding")))',
                 f'=IF(G{r}="Overdue",{RD}-E{r},0)'])
table(pbc, ["Ref.", "Document / evidence requested", "Cycle", "Requested date", "Due date", "Received date",
            "Status", "Days overdue"], rows, [9, 64, 16, 13, 13, 13, 13, 10],
      {4: "dd-mmm-yy", 5: "dd-mmm-yy", 6: "dd-mmm-yy", 8: "0"}, formula_cols=(7, 8))
add_dv(pbc, "Cycles", "C")
status_cf(pbc, f"G6:G{LAST}", {"Received": GREEN, "Outstanding": AMBER, "Overdue": RED})

# ---------------------------------------------------------------- Misstatements (ISA 450)
M = [
    ("Revenue", "Cut-off: July 2026 sales recorded in June 2026", "Factual", 86000, "Yes"),
    ("Inventory", "Production overhead over-absorbed into closing finished goods", "Factual", 64500, "No"),
    ("Inventory", "NRV shortfall on obsolete items not provided", "Judgmental", 38000, "No"),
    ("Trade payables", "Unrecorded accrued liabilities at year end", "Factual", 27400, "Yes"),
    ("Cost of sales", "Projected error from purchases sample", "Projected", 21000, "No"),
    ("Trade receivables", "ECL allowance understated on balances over 120 days", "Judgmental", 45000, "No"),
    ("Property, plant and equipment", "Repairs and maintenance capitalised", "Factual", 12800, "Yes"),
    ("Payroll expense", "Year-end bonus accrual understated", "Factual", 8900, "No"),
]
sheet_header(mis, "SUMMARY OF MISSTATEMENTS  -  ISA 450",
             "Amount = effect on profit before tax (positive = profit overstated). Items below the clearly-trivial threshold are not "
             "accumulated. The summary on the right compares uncorrected misstatements with materiality.", 11)
rows = []
for i, x in enumerate(M):
    r = 6 + i
    rows.append([f"M-{i+1:02d}", x[0], x[1], x[2], x[3], x[4],
                 f'=IF(A{r}="","",IF(ABS(E{r})<{TRIV},"Clearly trivial","Accumulate"))',
                 f'=IF(AND(F{r}="No",G{r}="Accumulate"),E{r},0)'])
table(mis, ["ID", "FS area", "Description of misstatement", "Type (ISA 450.A6)", "Amount (USD)", "Corrected by management?",
            "Treatment", "Uncorrected amount (USD)"], rows, [8, 28, 54, 14, 14, 13, 15, 16],
      {5: "#,##0", 8: "#,##0;(#,##0);-"}, formula_cols=(7, 8))
add_dv(mis, "Misstatement type", "D"); add_dv(mis, "Yes / No", "F")
mis.column_dimensions["I"].width = 3
mis.column_dimensions["J"].width = 40
mis.column_dimensions["K"].width = 18
put(mis, 5, 10, "Evaluation (ISA 450.11)", font(9, True, WHITE), align("center"), bg=NAVY2, to=(5, 11))
summ = [
    ("Total misstatements identified", f"=SUM(E6:E{LAST})", "#,##0"),
    ("Corrected by management", f'=SUMIF(F6:F{LAST},"Yes",E6:E{LAST})', "#,##0"),
    ("Uncorrected (accumulated)", f"=SUM(H6:H{LAST})", "#,##0"),
    ("Overall materiality", f"={OM}", "#,##0"),
    ("Performance materiality", f"={PM}", "#,##0"),
    ("Clearly trivial threshold", f"={TRIV}", "#,##0"),
    ("Uncorrected as % of overall materiality", "=IF(K9=0,0,K8/K9)", "0%"),
]
for i, (k, v, f) in enumerate(summ):
    r = 6 + i
    put(mis, r, 10, k, font(9, True)).border = BOX
    put(mis, r, 11, v, font(9, True), align("right"), fmt=f, bg=HEAD).border = BOX
put(mis, 13, 10, "Conclusion", font(9, True))
put(mis, 14, 10,
    '=IF(K8>=K9,"Uncorrected misstatements exceed overall materiality - request correction; consider a modified opinion (ISA 705).",'
    'IF(K8>=K10,"Uncorrected misstatements exceed performance materiality - request correction and extend procedures (ISA 450).",'
    '"Uncorrected misstatements are below performance materiality - evaluate qualitative factors (ISA 450)."))',
    font(9, True, RED), align(wrap=True, v="top"), to=(17, 11))

# ---------------------------------------------------------------- Findings (ISA 265)
F = [  # cycle, finding, severity, action, owner, raised, target, status, closed
    ("IT General", "User access reviews not performed; leavers still active in the ERP", "Critical", "Disable leaver accounts; implement quarterly access review", "IT Manager", D(2026,9,11), D(2026,10,30), "In Progress", None),
    ("Inventory", "Standard costs and overhead absorption rates not reviewed or approved for FY2026", "High", "Cost accountant to update rates; CFO to approve annually", "Cost Accountant", D(2026,10,9), D(2026,11,30), "In Progress", None),
    ("Financial Close", "Manual journals posted without independent approval", "High", "Enforce maker-checker workflow in the ERP", "Financial Controller", D(2026,10,22), D(2026,11,15), "Open", None),
    ("Purchases", "Three-way match overridden for urgent supplier payments", "High", "Restrict override rights; monthly exception report to CFO", "Procurement Manager", D(2026,10,6), D(2026,11,20), "In Progress", None),
    ("Purchases", "Supplier statement reconciliations not performed consistently", "Medium", "Reconcile top 20 suppliers monthly", "Financial Controller", D(2026,10,6), D(2026,12,5), "Open", None),
    ("Revenue", "Delivery notes missing customer acknowledgement", "Medium", "Require signed proof of delivery before invoicing", "Sales Manager", D(2026,10,5), D(2026,11,10), "Closed", D(2026,11,6)),
    ("Inventory", "Obsolete stock provision methodology is undocumented", "Medium", "Document and approve provisioning policy", "Cost Accountant", D(2026,10,30), D(2026,12,10), "Open", None),
    ("Fixed Assets", "Asset register not reconciled to the general ledger monthly", "Low", "Introduce monthly reconciliation", "Financial Controller", D(2026,11,12), D(2026,12,15), "Open", None),
    ("Payroll", "Leaver records updated late in the payroll system", "Low", "HR to notify payroll within two working days", "HR Manager", D(2026,10,12), D(2026,11,5), "Closed", D(2026,11,3)),
    ("Treasury", "Petty cash counts not documented", "Low", "Monthly surprise counts signed by reviewer", "Financial Controller", D(2026,10,22), D(2026,11,30), "In Progress", None),
    ("Revenue", "ECL model lacks forward-looking information", "Medium", "Update model with macro-economic overlay", "CFO", D(2026,11,18), D(2026,12,12), "Open", None),
    ("Financial Close", "Related party register is incomplete", "High", "Annual declarations from directors and key management", "CFO", D(2026,11,19), D(2026,12,8), "Open", None),
    ("IT General", "Program changes released to production without documented testing", "Medium", "Mandatory test sign-off before release", "IT Manager", D(2026,9,11), D(2026,10,31), "In Progress", None),
    ("Inventory", "Cycle count variances adjusted without an investigation threshold", "Low", "Set threshold and approval for count adjustments", "Warehouse Manager", D(2026,10,9), D(2026,11,13), "Closed", D(2026,11,13)),
    ("IT General", "No periodic backup restore test", "Medium", "Perform and document semi-annual restore test", "IT Manager", D(2026,9,11), D(2026,11,27), "Open", None),
    ("Inventory", "Year-end cost variances not allocated between inventory and cost of sales", "Critical", "Allocate variances at each period end; management review", "Cost Accountant", D(2026,11,13), D(2026,12,4), "Open", None),
]
sheet_header(fnd, "AUDIT FINDINGS & CONTROL DEFICIENCIES  -  ISA 265 / ISA 260",
             "Enter status as Open, In Progress or Closed. 'Display status' turns to Overdue automatically once the target date passes. "
             "Critical and High findings are classified as significant deficiencies to be reported in writing to those charged with governance.", 15)
rows = []
for i, x in enumerate(F):
    r = 6 + i
    rows.append([f"F-{i+1:02d}", x[0], x[1], x[2], x[3], x[4], x[5], x[6], x[7], x[8],
                 f'=IF(A{r}="","",IF(I{r}="Closed","Closed",IF(H{r}<{RD},"Overdue",I{r})))',
                 f'=IF(OR(A{r}="",I{r}="Closed"),"",{RD}-G{r})',
                 f'=IF(L{r}="","",IF(L{r}<=30,"0-30 days",IF(L{r}<=60,"31-60 days",IF(L{r}<=90,"61-90 days","90+ days"))))',
                 f'=IF(A{r}="","",IF(OR(D{r}="Critical",D{r}="High"),"Significant deficiency","Other deficiency"))',
                 f'=IF(OR(A{r}="",I{r}="Closed"),"",H{r}+ROW()/100000)'])
table(fnd, ["Finding ID", "Cycle", "Finding / deficiency", "Severity", "Agreed management action", "Client owner",
            "Raised date", "Target date", "Status (input)", "Closed date", "Display status", "Age (days)",
            "Ageing bucket", "ISA 265 classification", "Sort key (helper)"],
      rows, [9, 15, 56, 10, 46, 19, 11, 11, 12, 11, 12, 8, 12, 21, 12],
      {7: "dd-mmm-yy", 8: "dd-mmm-yy", 10: "dd-mmm-yy", 12: "0", 15: "0.00000"}, formula_cols=(11, 12, 13, 14, 15))
add_dv(fnd, "Cycles", "B"); add_dv(fnd, "Finding severity", "D"); add_dv(fnd, "Finding status", "I")
SEV_COL = {"Critical": RED, "High": "EA580C", "Medium": AMBER, "Low": GREEN}
status_cf(fnd, f"D6:D{LAST}", SEV_COL)
FS_COL = {"Open": AMBER, "In Progress": BLUE, "Overdue": RED, "Closed": GREEN}
status_cf(fnd, f"K6:K{LAST}", FS_COL)
FR = lambda col: f"Findings!${col}$6:${col}${LAST}"

# ---------------------------------------------------------------- ISA Map
sheet_header(isa, "ISA PROCESS MAP",
             "The audit process by phase, the ISAs that govern each step, the key output and where it is tracked in this workbook.", 5)
I = [
    ("Pre-engagement", "ISA 200", "Overall objectives of the independent auditor", "Professional scepticism and judgment applied throughout", "All sheets"),
    ("Pre-engagement", "ISA 210", "Agreeing the terms of audit engagements", "Signed engagement letter", "Audit Programme"),
    ("Pre-engagement", "ISA 220", "Quality management for an audit of financial statements", "Independence confirmation; team assignment", "Audit Programme"),
    ("Planning", "ISA 300", "Planning an audit of financial statements", "Overall audit strategy and audit plan", "Audit Programme"),
    ("Planning", "ISA 320", "Materiality in planning and performing an audit", "Overall and performance materiality", "Materiality"),
    ("Planning", "ISA 250", "Consideration of laws and regulations", "Understanding of legal framework", "Audit Programme"),
    ("Planning", "ISA 260", "Communication with those charged with governance", "Planned scope and timing communicated", "Audit Programme"),
    ("Risk Assessment", "ISA 315", "Identifying and assessing the risks of material misstatement", "Risk register; significant risks", "Risk Assessment"),
    ("Risk Assessment", "ISA 240", "The auditor's responsibilities relating to fraud", "Fraud risk factors; presumed risks", "Risk Assessment"),
    ("Risk Assessment", "ISA 550", "Related parties", "Related party listing", "Risk Assessment"),
    ("Risk Assessment", "ISA 540", "Auditing accounting estimates and related disclosures", "Understanding of estimates", "Risk Assessment"),
    ("Control Testing", "ISA 330", "The auditor's responses to assessed risks", "Tests of controls; reliance conclusion", "Control Testing"),
    ("Control Testing", "ISA 265", "Communicating deficiencies in internal control", "Deficiency log; management letter points", "Findings"),
    ("Substantive Testing", "ISA 500", "Audit evidence", "Sufficient appropriate evidence per assertion", "Substantive Testing / PBC Tracker"),
    ("Substantive Testing", "ISA 501", "Audit evidence - specific considerations (inventory, litigation)", "Count attendance; lawyer's letter", "Substantive Testing"),
    ("Substantive Testing", "ISA 505", "External confirmations", "Bank and receivable confirmations", "Substantive Testing"),
    ("Substantive Testing", "ISA 520", "Analytical procedures", "Substantive and final analytical review", "Substantive Testing"),
    ("Substantive Testing", "ISA 530", "Audit sampling", "Sample sizes and projected misstatements", "Control / Substantive Testing"),
    ("Completion", "ISA 450", "Evaluation of misstatements identified during the audit", "Summary of uncorrected misstatements", "Misstatements"),
    ("Completion", "ISA 560", "Subsequent events", "Subsequent events review", "Audit Programme"),
    ("Completion", "ISA 570", "Going concern", "Going concern conclusion", "Audit Programme"),
    ("Completion", "ISA 580", "Written representations", "Signed representation letter", "Audit Programme"),
    ("Reporting", "ISA 700", "Forming an opinion and reporting on financial statements", "Auditor's report", "Audit Programme"),
    ("Reporting", "ISA 701", "Communicating key audit matters", "KAM section (listed entities)", "Audit Programme"),
    ("Reporting", "ISA 705", "Modifications to the opinion", "Qualified / adverse / disclaimer where required", "Misstatements"),
    ("Reporting", "ISA 706", "Emphasis of matter and other matter paragraphs", "Additional paragraphs where required", "Audit Programme"),
    ("Reporting", "ISA 720", "The auditor's responsibilities relating to other information", "Other information read and considered", "Audit Programme"),
    ("Reporting", "ISA 230", "Audit documentation", "Final audit file assembled within 60 days", "Audit Programme"),
]
table(isa, ["Audit phase", "ISA", "Title", "Key output / working paper", "Tracked on sheet"],
      [list(x) for x in I], [20, 11, 58, 46, 30], formula_cols=(1, 2, 3, 4, 5))
for r in range(6, 6 + len(I)):
    for c in range(1, 6):
        isa.cell(r, c).alignment = align("left")
        isa.cell(r, c).fill = fill(WHITE)

# ---------------------------------------------------------------- Dashboard geometry
C0 = 4                       # first content column (D)
NC = 36                      # content columns
cc = lambda i: C0 + i        # content col index -> sheet col
ROWS = 62
dash.sheet_view.showGridLines = False
dash.sheet_view.showRowColHeaders = False
dash.sheet_view.zoomScale = 85
dash.column_dimensions["A"].width = 1.5
dash.column_dimensions["B"].width = 25
dash.column_dimensions["C"].width = 2
for i in range(NC):
    dash.column_dimensions[L(cc(i))].width = 5
dash.column_dimensions[L(cc(NC))].width = 1.5
for r in range(1, ROWS + 1):
    dash.row_dimensions[r].height = 16.5
paint(dash, 1, 1, ROWS, 2, NAVY)
paint(dash, 1, 3, ROWS, cc(NC), BG)
paint(dash, 1, 1, 4, cc(NC), NAVY)

# header bar
put(dash, 2, 2, "ISA AUDIT", font(15, True, WHITE), align("left", "bottom"))
put(dash, 3, 2, "PROCESS DASHBOARD", font(11, True, "93C5FD"), align("left", "top"))


def top_field(c1, c2, label, value, fmt=None, is_filter=False):
    put(dash, 2, cc(c1), label, font(8, color="BFDBFE"), align("left", "bottom"), to=(2, cc(c2)))
    cell = put(dash, 3, cc(c1), value, font(10, True, INPUT_BLUE if is_filter else INK), align("left"),
               to=(3, cc(c2)), fmt=fmt, bg="FFF7CC" if is_filter else WHITE)
    return f"${L(cc(c1))}$3"


top_field(0, 8, "Client", "=Settings!B3")
top_field(10, 12, "Financial year end", "=Settings!B4", "dd-mmm-yyyy")
top_field(14, 16, "Reporting date", "=Settings!B5", "dd-mmm-yyyy")
top_field(18, 20, "Planned report date", "=Settings!B7", "dd-mmm-yyyy")
PH_SEL = top_field(23, 27, "▼ FILTER: Audit phase", "All", is_filter=True)
OW_SEL = top_field(29, 33, "▼ FILTER: Team member", "All", is_filter=True)
dv = DataValidation(type="list", formula1="=" + LISTREF["Phase filter"]); dash.add_data_validation(dv); dv.add(PH_SEL.replace("$", ""))
dv = DataValidation(type="list", formula1="=" + LISTREF["Team filter"]); dash.add_data_validation(dv); dv.add(OW_SEL.replace("$", ""))

# sidebar navigation
nav = [("Dashboard", "Dashboard"), ("Guide", "Guide & legend"), ("ISA Map", "ISA process map"), ("Settings", "Settings"),
       ("Materiality", "Materiality  (ISA 320)"), ("Audit Programme", "Audit programme  (ISA 300)"),
       ("Risk Assessment", "Risk assessment  (ISA 315)"), ("Control Testing", "Control testing  (ISA 330)"),
       ("Substantive Testing", "Substantive testing  (ISA 500)"), ("PBC Tracker", "Evidence tracker  (PBC)"),
       ("Misstatements", "Misstatements  (ISA 450)"), ("Findings", "Findings  (ISA 265)"), ("Calc", "Calculations")]
for i, (sheet, label) in enumerate(nav):
    r = 6 + 2 * i
    active = i == 0
    c = put(dash, r, 2, "  ▸  " + label, font(10, True, WHITE if active else "CBD5E1"), align("left"),
            bg=BLUE if active else NAVY)
    c.hyperlink = f"#'{sheet}'!A1"
    c.font = font(10, True, WHITE if active else "CBD5E1")
    dash.row_dimensions[r].height = 22
put(dash, 34, 2, "AUDIT PROCESS FLOW", font(8, True, "93C5FD"))
for i, p in enumerate(PHASES):
    put(dash, 35 + 2 * i, 2, f"  {i+1}. {p}", font(9, color="E2E8F0"))
    if i < len(PHASES) - 1:
        put(dash, 36 + 2 * i, 2, "      ↓", font(8, color="60A5FA"))
put(dash, 51, 2, "Yellow cells in the top bar are filters (phase and team member). They drive the KPI cards, "
    "panels 1-2, panel 7 and the trend chart.", font(8, italic=True, color="94A3B8"), align(wrap=True, v="top"), to=(56, 2))
put(dash, 58, 2, "Illustrative sample data for teaching - replace on the input sheets.",
    font(8, italic=True, color="94A3B8"), align(wrap=True, v="top"), to=(60, 2))

# ---------------------------------------------------------------- Calc sheet
calc.sheet_view.showGridLines = False
put(calc, 1, 1, "CALCULATIONS  (feeds the dashboard - do not edit)", font(12, True, WHITE), bg=NAVY, to=(1, 8))
calc.column_dimensions["A"].width = 30
for col in "BCDEFGH":
    calc.column_dimensions[col].width = 16


def chead(r, labels):
    for j, h in enumerate(labels, 1):
        put(calc, r, j, h, font(9, True, WHITE), align("center"), bg=NAVY2)


put(calc, 3, 1, "Phase criterion", font(9, True)); put(calc, 3, 2, f'=IF(Dashboard!{PH_SEL}="All","*",Dashboard!{PH_SEL})', font(9))
put(calc, 4, 1, "Team member criterion", font(9, True)); put(calc, 4, 2, f'=IF(Dashboard!{OW_SEL}="All","*",Dashboard!{OW_SEL})', font(9))
put(calc, 5, 1, "Reporting date", font(9, True)); put(calc, 5, 2, f"={RD}", font(9), fmt="dd-mmm-yyyy")
PC, OC = "$B$3", "$B$4"
base = f"{PR('B')},{PC},{PR('E')},{OC}"

chead(8, ["Procedure status", "Count"])
ST_ORDER = ["Completed", "Under Review", "In Progress", "Not Started"]
for i, s in enumerate(ST_ORDER):
    put(calc, 9 + i, 1, s, font(9)); put(calc, 9 + i, 2, f'=COUNTIFS({base},{PR("H")},A{9+i})', font(9))

chead(15, ["Audit phase", "Not Started", "In Progress", "Under Review", "Completed", "Total", "% complete"])
for i, p in enumerate(PHASES):
    r = 16 + i
    put(calc, r, 1, p, font(9))
    for j, s in enumerate(STATUSES):
        put(calc, r, 2 + j, f'=COUNTIFS({PR("B")},$A{r},{base},{PR("H")},{L(2+j)}$15)', font(9))
    put(calc, r, 6, f"=SUM(B{r}:E{r})", font(9))
    put(calc, r, 7, f"=IF(F{r}=0,0,E{r}/F{r})", font(9), fmt="0%")

chead(25, ["Open findings by severity", "Count"])
for i, s in enumerate(SEVERITY):
    put(calc, 26 + i, 1, s, font(9))
    put(calc, 26 + i, 2, f'=COUNTIFS({FR("D")},A{26+i},{FR("K")},"<>Closed")', font(9))

chead(32, ["Open findings - age", "Count"])
for i, s in enumerate(["0-30 days", "31-60 days", "61-90 days", "90+ days"]):
    put(calc, 33 + i, 1, s, font(9)); put(calc, 33 + i, 2, f'=COUNTIF({FR("M")},A{33+i})', font(9))

chead(39, ["Team member", "Budget hrs", "Actual hrs", "Remaining hrs", "% used"])
for i, t in enumerate(TEAM):
    r = 40 + i
    put(calc, r, 1, t, font(9))
    put(calc, r, 2, f'=SUMIFS({PR("J")},{PR("E")},$A{r},{PR("B")},{PC},{PR("E")},{OC})', font(9))
    put(calc, r, 3, f'=SUMIFS({PR("K")},{PR("E")},$A{r},{PR("B")},{PC},{PR("E")},{OC})', font(9))
    put(calc, r, 4, f"=B{r}-C{r}", font(9))
    put(calc, r, 5, f"=IF(B{r}=0,0,C{r}/B{r})", font(9), fmt="0%")
put(calc, 46, 1, "Total", font(9, True))
for col in "BCD":
    put(calc, 46, "ABCDE".index(col) + 1, f"=SUM({col}40:{col}45)", font(9, True))
put(calc, 46, 5, "=IF(B46=0,0,C46/B46)", font(9, True), fmt="0%")

chead(48, ["Month end", "Label", "Actual cumulative %", "Planned cumulative %"])
NM = 8
for i in range(NM):
    r = 49 + i
    put(calc, r, 1, f"=EOMONTH(Settings!$B$6,{i})", font(9), fmt="dd-mmm-yyyy")
    put(calc, r, 2, f'=TEXT(A{r},"mmm yy")', font(9))
    put(calc, r, 3, f'=IF($K$3=0,0,IF(EOMONTH(A{r},-1)>=$B$5,"",COUNTIFS({base},{PR("I")},"<="&MIN(A{r},$B$5))/$K$3))', font(9), fmt="0%")
    put(calc, r, 4, f'=IF($K$3=0,0,COUNTIFS({base},{PR("G")},"<="&A{r})/$K$3)', font(9), fmt="0%")

chead(59, ["Controls by cycle", "Controls", "Design effective %", "Operating effective %", "Sample completion %"])
for i, cy in enumerate(CYCLES):
    r = 60 + i
    put(calc, r, 1, cy, font(9))
    put(calc, r, 2, f'=COUNTIF({CR("B")},$A{r})', font(9))
    put(calc, r, 3, f'=IF(B{r}=0,0,COUNTIFS({CR("B")},$A{r},{CR("F")},"Effective")/B{r})', font(9), fmt="0%")
    put(calc, r, 4, f'=IF(B{r}=0,0,COUNTIFS({CR("B")},$A{r},{CR("L")},"Effective")/B{r})', font(9), fmt="0%")
    put(calc, r, 5, f'=IF(SUMIF({CR("B")},$A{r},{CR("G")})=0,0,SUMIF({CR("B")},$A{r},{CR("H")})/SUMIF({CR("B")},$A{r},{CR("G")}))', font(9), fmt="0%")

kp = [
    ("Total procedures (filtered)", "=SUM(B9:B12)", None),
    ("Completed", "=B9", None),
    ("% complete", "=IF(K3=0,0,K4/K3)", "0%"),
    ("Overdue procedures", f'=COUNTIFS({base},{PR("M")},"Yes")', None),
    ("Open findings", "=SUM(B26:B29)", None),
    ("Total findings raised", f'=COUNTA({FR("A")})', None),
    ("Overdue finding actions", f'=COUNTIF({FR("K")},"Overdue")', None),
    ("Significant risks", f'=COUNTIF({RR("L")},"Yes")', None),
    ("Total risks identified", f'=COUNTA({RR("A")})', None),
    ("Fraud risks (ISA 240)", f'=COUNTIF({RR("M")},"Yes")', None),
    ("PBC items not yet received", "=COUNTIF('PBC Tracker'!$G$6:$G$300,\"Outstanding\")+COUNTIF('PBC Tracker'!$G$6:$G$300,\"Overdue\")", None),
    ("PBC items overdue", "=COUNTIF('PBC Tracker'!$G$6:$G$300,\"Overdue\")", None),
    ("High RoMM risks", f'=COUNTIF({RR("K")},"High")', None),
    ("In progress + under review", "=B10+B11", None),
    ("Controls tested (sample completion)", f'=IF(SUM({CR("G")})=0,0,SUM({CR("H")})/SUM({CR("G")}))', "0%"),
]
put(calc, 2, 10, "KPI", font(9, True, WHITE), bg=NAVY2); put(calc, 2, 11, "Value", font(9, True, WHITE), bg=NAVY2)
KP = {}
for i, (k, v, f) in enumerate(kp):
    put(calc, 3 + i, 10, k, font(9)); put(calc, 3 + i, 11, v, font(9, True), fmt=f)
    KP[k] = f"Calc!$K${3+i}"
calc.column_dimensions["J"].width = 34

# ---------------------------------------------------------------- KPI cards
cards = [
    ("AUDIT PROCEDURES", f"={KP['Total procedures (filtered)']}", "0", BLUE,
     f'="In progress / review: "&{KP["In progress + under review"]}'),
    ("COMPLETED", f"={KP['Completed']}", "0", GREEN,
     f'=TEXT({KP["% complete"]},"0%")&" of programme"'),
    ("HOURS USED", "=Calc!$E$46", "0%", PURPLE,
     '=TEXT(Calc!$C$46,"#,##0")&" of "&TEXT(Calc!$B$46,"#,##0")&" budget hrs"'),
    ("OPEN FINDINGS", f"={KP['Open findings']}", "0", AMBER,
     f'="of "&{KP["Total findings raised"]}&" raised (ISA 265)"'),
    ("OVERDUE ITEMS", f"={KP['Overdue procedures']}+{KP['Overdue finding actions']}+{KP['PBC items overdue']}", "0", RED,
     f'="Proc. "&{KP["Overdue procedures"]}&" · Findings "&{KP["Overdue finding actions"]}&" · PBC "&{KP["PBC items overdue"]}'),
    ("SIGNIFICANT RISKS", f"={KP['Significant risks']}", "0", TEAL,
     f'="of "&{KP["Total risks identified"]}&" risks · "&{KP["Fraud risks (ISA 240)"]}&" fraud (ISA 240)"'),
]
for i, (label, formula, fmt, color, subt) in enumerate(cards):
    c1, c2 = cc(i * 6), cc(i * 6 + 4)
    paint(dash, 6, c1, 9, c2, WHITE)
    paint(dash, 6, c1, 9, c1, color)
    dash.merge_cells(start_row=6, start_column=c1, end_row=9, end_column=c1)
    put(dash, 6, c1 + 1, label, font(8, True, MUTED), align("left", "bottom"), to=(6, c2))
    put(dash, 7, c1 + 1, formula, font(22, True, color), align("left"), to=(8, c2), fmt=fmt)
    put(dash, 9, c1 + 1, subt, font(8, color=MUTED), align("left", "top"), to=(9, c2))
dash.row_dimensions[6].height = 18
dash.row_dimensions[9].height = 18


def panel(top, c1, ncols, nrows, title):
    paint(dash, top, cc(c1), top + nrows, cc(c1 + ncols - 1), WHITE)
    put(dash, top, cc(c1), title, font(10, True, NAVY), align("left"), to=(top, cc(c1 + ncols - 1)))
    dash.row_dimensions[top].height = 21


def clean(ch, w, h):
    ch.width, ch.height = w, h
    ch.roundedCorners = False
    ch.graphical_properties = GraphicalProperties(ln=LineProperties(noFill=True))
    if ch.legend is not None:
        ch.legend.overlay = False


T1, T2, T3 = 11, 29, 47
BODY = 16
CW, CH = 8.2, 8.6   # chart cm inside an 8-column panel

# --- Panel 1: procedure status doughnut
panel(T1, 0, 8, BODY, " 1. Audit Progress by Status")
ch = DoughnutChart()
ch.add_data(Reference(calc, min_col=2, min_row=8, max_row=12), titles_from_data=True)
ch.set_categories(Reference(calc, min_col=1, min_row=9, max_row=12))
ch.holeSize = 58
for i, col in enumerate([GREEN, PURPLE, BLUE, GREY]):
    pt = DataPoint(idx=i); pt.graphicalProperties.solidFill = col
    pt.graphicalProperties.line.solidFill = WHITE
    ch.series[0].dPt.append(pt)
ch.dataLabels = DataLabelList(); ch.dataLabels.showVal = True
ch.dataLabels.showSerName = ch.dataLabels.showCatName = ch.dataLabels.showLegendKey = ch.dataLabels.showPercent = False
ch.legend.position = "b"
clean(ch, CW, CH)
dash.add_chart(ch, f"{L(cc(0))}{T1+1}")

# --- Panel 2: ISA phase status matrix
panel(T1, 9, 8, BODY, " 2. Audit Phase Status Matrix (ISA process)")
r = T1 + 1
dash.row_dimensions[r].height = 24
put(dash, r, cc(9), "Phase", font(8, True, MUTED), align("left"), to=(r, cc(11)), bg=HEAD)
for j, h in enumerate(["Not started", "In progress", "Under review", "Done", "% done"]):
    put(dash, r, cc(12 + j), h, font(7, True, MUTED), align("center", wrap=True), bg=HEAD)
for i, p in enumerate(PHASES):
    r1 = T1 + 2 + 2 * i
    put(dash, r1, cc(9), f"=Calc!A{16+i}", font(8, True), align("left", wrap=True), to=(r1 + 1, cc(11)), bg=WHITE)
    for j, col in enumerate("BCDEG"):
        put(dash, r1, cc(12 + j), f"=Calc!{col}{16+i}", font(9, j >= 3), align("center"),
            to=(r1 + 1, cc(12 + j)), fmt="0%" if col == "G" else "0;;", bg=WHITE)
    for c in range(cc(9), cc(17)):
        dash.cell(r1 + 1, c).border = BOTTOM
rng = f"{L(cc(16))}{T1+2}:{L(cc(16))}{T1+15}"
dash.conditional_formatting.add(rng, DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="86EFAC"))
r = T1 + 16
put(dash, r, cc(9), "Total", font(8, True), to=(r, cc(11)), bg=HEAD)
for j, col in enumerate("BCDE"):
    put(dash, r, cc(12 + j), f"=SUM(Calc!{col}16:{col}22)", font(9, True), align("center"), bg=HEAD)
put(dash, r, cc(16), f"={KP['% complete']}", font(9, True), align("center"), fmt="0%", bg=HEAD)

# --- Panel 3: control testing matrix
panel(T1, 18, 8, BODY, " 3. Control Testing Evidence Matrix (ISA 330)")
r = T1 + 1
put(dash, r, cc(18), "Cycle", font(8, True, MUTED), align("left"), to=(r, cc(19)), bg=HEAD)
for j, h in enumerate(["Design effective", "Operating effective", "Sample completion"]):
    put(dash, r, cc(20 + 2 * j), h, font(7, True, MUTED), align("center", wrap=True), to=(r, cc(21 + 2 * j)), bg=HEAD)
for i, cy in enumerate(CYCLES):
    r = T1 + 2 + i
    dash.row_dimensions[r].height = 21
    put(dash, r, cc(18), f"=Calc!A{60+i}", font(8, True), Alignment(horizontal="left", vertical="center", shrink_to_fit=True), to=(r, cc(19)), bg=WHITE)
    for j, col in enumerate("CDE"):
        put(dash, r, cc(20 + 2 * j), f"=Calc!{col}{60+i}", font(9), align("center"), to=(r, cc(21 + 2 * j)), fmt="0%", bg=WHITE)
    for c in range(cc(18), cc(26)):
        dash.cell(r, c).border = BOTTOM
rng = f"{L(cc(20))}{T1+2}:{L(cc(25))}{T1+9}"
dash.conditional_formatting.add(rng, IconSetRule("3Symbols", "num", [0, 0.5, 0.8]))
r = T1 + 11
put(dash, r, cc(18), "Controls tested (items)", font(8, color=MUTED), to=(r, cc(22)))
put(dash, r, cc(23), f"={KP['Controls tested (sample completion)']}", font(10, True, TEAL), align("right"), to=(r, cc(25)), fmt="0%")
r += 1
put(dash, r, cc(18), "Controls with no reliance", font(8, color=MUTED), to=(r, cc(22)))
put(dash, r, cc(23), f'=COUNTIF({CR("N")},"No reliance*")&" of "&COUNTA({CR("A")})', font(10, True, RED), align("right"), to=(r, cc(25)))
r += 1
put(dash, r, cc(18), "PBC items not yet received", font(8, color=MUTED), to=(r, cc(22)))
put(dash, r, cc(23), f"={KP['PBC items not yet received']}", font(10, True, AMBER), align("right"), to=(r, cc(25)))
put(dash, T1 + 15, cc(18), "● ≥80%   ● 50-79%   ● <50% of controls / sample items in the cycle",
    font(7, italic=True, color=MUTED), to=(T1 + 15, cc(25)))

# --- Panel 4: findings by severity
panel(T1, 27, 8, BODY, " 4. Open Findings by Severity (ISA 265)")
ch = BarChart(); ch.type = "bar"
ch.add_data(Reference(calc, min_col=2, min_row=25, max_row=29), titles_from_data=True)
ch.set_categories(Reference(calc, min_col=1, min_row=26, max_row=29))
for i, col in enumerate([RED, "EA580C", AMBER, GREEN]):
    pt = DataPoint(idx=i); pt.graphicalProperties.solidFill = col
    ch.series[0].dPt.append(pt)
ch.dataLabels = DataLabelList(); ch.dataLabels.showVal = True
ch.dataLabels.showSerName = ch.dataLabels.showCatName = ch.dataLabels.showLegendKey = False
ch.legend = None
ch.gapWidth = 60
ch.x_axis.scaling.orientation = "maxMin"
ch.y_axis.crosses = "max"
ch.x_axis.delete = False; ch.y_axis.delete = False
ch.y_axis.majorGridlines = None
ch.y_axis.scaling.min = 0
ch.y_axis.majorUnit = 1
ch.y_axis.number_format = "0"
clean(ch, CW, CH)
dash.add_chart(ch, f"{L(cc(27))}{T1+1}")

# --- Panel 5: finding ageing
panel(T2, 0, 8, BODY, " 5. Management Action Ageing (open findings)")
ch = BarChart(); ch.type = "col"
ch.add_data(Reference(calc, min_col=2, min_row=32, max_row=36), titles_from_data=True)
ch.set_categories(Reference(calc, min_col=1, min_row=33, max_row=36))
for i, col in enumerate([GREEN, AMBER, "EA580C", RED]):
    pt = DataPoint(idx=i); pt.graphicalProperties.solidFill = col
    ch.series[0].dPt.append(pt)
ch.dataLabels = DataLabelList(); ch.dataLabels.showVal = True
ch.dataLabels.showSerName = ch.dataLabels.showCatName = ch.dataLabels.showLegendKey = False
ch.legend = None
ch.gapWidth = 50
ch.x_axis.delete = False; ch.y_axis.delete = False
ch.y_axis.majorGridlines = None
ch.y_axis.scaling.min = 0
ch.y_axis.majorUnit = 2
ch.y_axis.number_format = "0"
clean(ch, CW, CH)
dash.add_chart(ch, f"{L(cc(0))}{T2+1}")

# --- Panel 6: risk heat map
panel(T2, 9, 8, BODY, " 6. Risk Heat Map (ISA 315) - number of risks")
put(dash, T2 + 1, cc(9), "Likelihood", font(8, True, MUTED), align("center", rot=90), to=(T2 + 10, cc(9)), bg=WHITE)
for i, lik in enumerate([5, 4, 3, 2, 1]):
    r1 = T2 + 1 + 2 * i
    put(dash, r1, cc(10), lik, font(8, True, MUTED), align("center"), to=(r1 + 1, cc(10)), bg=WHITE)
    for m in range(1, 6):
        score = lik * m
        color = "F87171" if score >= 15 else ("FCD34D" if score >= 8 else "86EFAC")
        c = put(dash, r1, cc(10 + m), f'=COUNTIFS({RR("F")},{lik},{RR("G")},{m})', font(11, True), align("center"),
                to=(r1 + 1, cc(10 + m)), fmt="0;;", bg=color)
        for rr in (r1, r1 + 1):
            dash.cell(rr, cc(10 + m)).border = Border(left=Side(style="thin", color=WHITE), right=Side(style="thin", color=WHITE),
                                                      top=Side(style="thin", color=WHITE), bottom=Side(style="thin", color=WHITE))
for m in range(1, 6):
    put(dash, T2 + 11, cc(10 + m), m, font(8, True, MUTED), align("center"))
put(dash, T2 + 12, cc(11), "Magnitude of misstatement", font(8, True, MUTED), align("center"), to=(T2 + 12, cc(15)))
for j, (lab, key, col) in enumerate([("Significant", "Significant risks", RED), ("Fraud risks", "Fraud risks (ISA 240)", PURPLE),
                                     ("High RoMM", "High RoMM risks", "EA580C"), ("Total risks", "Total risks identified", NAVY)]):
    put(dash, T2 + 14, cc(9 + 2 * j), f"={KP[key]}", font(12, True, col), align("center"), to=(T2 + 14, cc(10 + 2 * j)))
    put(dash, T2 + 15, cc(9 + 2 * j), lab, font(7, color=MUTED), align("center"), to=(T2 + 15, cc(10 + 2 * j)))

# --- Panel 7: resource allocation
panel(T2, 18, 8, BODY, " 7. Audit Resource Allocation (hours)")
r = T2 + 1
put(dash, r, cc(18), "Team member", font(8, True, MUTED), to=(r, cc(19)), bg=HEAD)
for j, h in enumerate(["Budget", "Actual", "% used"]):
    put(dash, r, cc(20 + 2 * j), h, font(8, True, MUTED), align("center"), to=(r, cc(21 + 2 * j)), bg=HEAD)
for i in range(len(TEAM) + 1):
    r1 = T2 + 2 + 2 * i
    src = 40 + i
    bold = i == len(TEAM)
    bg = HEAD if bold else WHITE
    put(dash, r1, cc(18), f"=Calc!A{src}", font(8, True), to=(r1 + 1, cc(19)), bg=bg)
    for j, col in enumerate("BCE"):
        put(dash, r1, cc(20 + 2 * j), f"=Calc!{col}{src}", font(9, bold), align("center"), to=(r1 + 1, cc(21 + 2 * j)),
            fmt="0%" if col == "E" else "#,##0", bg=bg)
    for c in range(cc(18), cc(26)):
        dash.cell(r1 + 1, c).border = BOTTOM
rng = f"{L(cc(24))}{T2+2}:{L(cc(24))}{T2+13}"
dash.conditional_formatting.add(rng, DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1.2, color="93C5FD"))
put(dash, T2 + 16, cc(18), "% used above 100% = budget overrun on work assigned.", font(7, italic=True, color=MUTED), to=(T2 + 16, cc(25)))

# --- Panel 8: materiality & misstatements
panel(T2, 27, 8, BODY, " 8. Materiality & Misstatements (ISA 320 / 450)")
items = [("Overall materiality", f"={OM}", "#,##0", NAVY), ("Performance materiality", f"={PM}", "#,##0", NAVY),
         ("Clearly trivial threshold", f"={TRIV}", "#,##0", NAVY),
         ("Misstatements identified", "=Misstatements!$K$6", "#,##0", INK),
         ("Corrected by management", "=Misstatements!$K$7", "#,##0", GREEN),
         ("Uncorrected (accumulated)", "=Misstatements!$K$8", "#,##0", RED),
         ("Uncorrected / materiality", "=Misstatements!$K$12", "0%", RED)]
for i, (lab, f, fmt, col) in enumerate(items):
    r1 = T2 + 1 + 2 * i
    put(dash, r1, cc(27), lab, font(8, color=MUTED), to=(r1 + 1, cc(31)), bg=WHITE)
    put(dash, r1, cc(32), f, font(10, True, col), align("right"), to=(r1 + 1, cc(34)), fmt=fmt, bg=WHITE)
    for c in range(cc(27), cc(35)):
        dash.cell(r1 + 1, c).border = BOTTOM
rng = f"{L(cc(32))}{T2+13}"
dash.conditional_formatting.add(rng, DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="FCA5A5"))
put(dash, T2 + 15, cc(27), "=Misstatements!$J$14", font(7, True, RED), align(wrap=True, v="top"), to=(T2 + 16, cc(34)), bg=WHITE)
dash.row_dimensions[T2 + 15].height = 18
dash.row_dimensions[T2 + 16].height = 18

# --- Panel 9: priority actions
NB = 13
panel(T3, 0, 17, NB, " 9. Priority Audit Actions (open findings by target date)")
r = T3 + 1
spans = [("Finding", 0, 1, "A", None), ("Area", 2, 4, "B", None), ("Finding / deficiency", 5, 9, "C", None),
         ("Client owner", 10, 12, "F", None), ("Target date", 13, 14, "H", "dd-mmm-yy"), ("Status", 15, 16, "K", None)]
for h, a, b, _, _ in spans:
    put(dash, r, cc(a), h, font(8, True, MUTED), align("left" if a in (5,) else "center"), to=(r, cc(b)), bg=HEAD)
for k in range(1, 7):
    r1 = T3 + 2 * k
    for h, a, b, col, fmt in spans:
        f = (f'=IFERROR(INDEX({FR(col)},MATCH(SMALL({FR("O")},{k}),{FR("O")},0)),"")')
        put(dash, r1, cc(a), f, font(8, h == "Finding"), align("left" if col == "C" else "center", wrap=True),
            to=(r1 + 1, cc(b)), fmt=fmt, bg=WHITE)
    for c in range(cc(0), cc(17)):
        dash.cell(r1 + 1, c).border = BOTTOM
rng = f"{L(cc(15))}{T3+2}:{L(cc(16))}{T3+13}"
status_cf(dash, rng, FS_COL)

# --- Panel 10: progress trend
panel(T3, 18, 17, NB, " 10. Audit Completion Trend - actual vs planned (cumulative % of procedures)")
bar = BarChart(); bar.type = "col"
bar.add_data(Reference(calc, min_col=3, min_row=48, max_row=48 + NM), titles_from_data=True)
bar.set_categories(Reference(calc, min_col=2, min_row=49, max_row=48 + NM))
bar.series[0].graphicalProperties.solidFill = TEAL
bar.gapWidth = 60
line = LineChart()
line.add_data(Reference(calc, min_col=4, min_row=48, max_row=48 + NM), titles_from_data=True)
line.series[0].graphicalProperties.line.solidFill = NAVY
line.series[0].graphicalProperties.line.width = 22000
line.series[0].marker.symbol = "circle"
line.series[0].marker.size = 5
line.series[0].marker.graphicalProperties.solidFill = NAVY
line.series[0].marker.graphicalProperties.line.solidFill = NAVY
bar.y_axis.number_format = "0%"
bar.y_axis.scaling.min = 0; bar.y_axis.scaling.max = 1
bar.y_axis.majorUnit = 0.25
bar.x_axis.delete = False; bar.y_axis.delete = False
bar.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill=LINE))
bar.legend.position = "b"
bar += line
line.series[0].smooth = False
clean(bar, 16.8, 6.9)
dash.add_chart(bar, f"{L(cc(18))}{T3+1}")

dash.page_setup.orientation = "landscape"
dash.page_setup.fitToWidth = 1
dash.page_setup.fitToHeight = 1
dash.sheet_properties.pageSetUpPr.fitToPage = True
dash.print_area = f"A1:{L(cc(NC))}{ROWS}"

# ---------------------------------------------------------------- Guide
guide.sheet_view.showGridLines = False
put(guide, 1, 1, "GUIDE & LEGEND", font(14, True, WHITE), bg=NAVY, to=(1, 3))
guide.row_dimensions[1].height = 26
put(guide, 2, 1, "◄ Back to Dashboard", font(9, True, BLUE, underline="single")).hyperlink = "#'Dashboard'!A1"
guide.column_dimensions["A"].width = 30
guide.column_dimensions["B"].width = 95
guide.column_dimensions["C"].width = 22
g = [
    ("PURPOSE", None),
    ("What this is", "A formula-driven dashboard that tracks a financial statement audit through the seven phases of the ISA audit process, from engagement acceptance to the auditor's report."),
    ("Sample data", "Every figure is illustrative teaching data for a fictitious manufacturing client. Overwrite the input sheets with a real engagement."),
    ("HOW TO USE", None),
    ("1. Settings", "Set the client, year end and reporting date. The reporting date drives all overdue and ageing tests (type =TODAY() for live tracking)."),
    ("2. Materiality", "Enter draft financial data, choose the benchmark and percentages. Overall, performance and trivial thresholds calculate automatically."),
    ("3. Input sheets", "Update Audit Programme, Risk Assessment, Control Testing, Substantive Testing, PBC Tracker, Misstatements and Findings."),
    ("4. Adding rows", "Type directly below the last row (ranges read to row 300) and copy the grey formula cells down from the row above."),
    ("5. Dashboard filters", "The two yellow cells in the top bar filter by audit phase and team member. They affect the KPI cards for procedures, panels 1, 2, 7 and 10."),
    ("COLOUR LEGEND", None),
    ("Blue text", "Input - type or select from the drop-down."),
    ("Black text on grey", "Formula - do not overwrite."),
    ("Yellow cell (Dashboard)", "Filter selector."),
    ("Calc sheet", "Helper tables behind the charts - do not edit."),
    ("KEY RULES BUILT INTO THE FORMULAS", None),
    ("Inherent risk (ISA 315)", "Likelihood x magnitude (1-5 each): score >= 15 High, >= 8 Medium, otherwise Low."),
    ("RoMM", "Inherent risk x control risk (Low=1, Medium=2, High=3): product >= 6 High, >= 3 Medium, otherwise Low."),
    ("Operating effectiveness", "Deviation rate = exceptions / items tested. Zero = Effective; up to the tolerable rate in Settings = Partially Effective; above = Ineffective."),
    ("Scoping (ISA 320)", "FS areas with a balance above performance materiality are in scope for tests of details."),
    ("Misstatements (ISA 450)", "Items below the clearly-trivial threshold are not accumulated. Uncorrected items are compared with performance and overall materiality."),
    ("Findings (ISA 265)", "Critical and High findings are treated as significant deficiencies. A finding shows Overdue when not closed and the target date is before the reporting date."),
    ("Thresholds are assumptions", "The scoring cut-offs, 75% performance materiality, 5% trivial threshold and 5% tolerable deviation rate are common practice assumptions, not ISA requirements. Adjust to firm methodology."),
]
r = 4
for k, v in g:
    if v is None:
        put(guide, r, 1, k, font(10, True, WHITE), bg=NAVY2, to=(r, 2))
    else:
        put(guide, r, 1, k, font(10, True), align(v="top"))
        put(guide, r, 2, v, font(10), align(wrap=True, v="top"))
        guide.row_dimensions[r].height = 30 if len(v) > 100 else 18
    r += 1

for ws in wb.worksheets:
    ws.sheet_properties.tabColor = BLUE if ws.title == "Dashboard" else (GREY if ws.title in ("Calc", "Guide", "ISA Map", "Settings") else NAVY)

wb.save(OUT)
print("saved", OUT)
