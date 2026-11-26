import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.chart import BarChart, DoughnutChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.comments import Comment

OUT = "/home/user/phatchen/Audit_Working_Papers_ISA_ISQM_IESBA.xlsx"
D = dt.date

# ---------- styles ----------
F = "Arial"
DK, MD, LT, PALE, BG = "1B5E20", "2E7D32", "43A047", "E8F5E9", "F3F6F3"
INP = "FFF9C4"
def font(sz=10, b=False, c="000000", i=False): return Font(name=F, size=sz, bold=b, color=c, italic=i)
def fill(c): return PatternFill("solid", start_color=c, end_color=c)
thin = Side(style="thin", color="BDBDBD")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
gs = Side(style="medium", color=LT)
GBOX = Border(left=gs, right=gs, top=gs, bottom=gs)
WRAP = Alignment(wrap_text=True, vertical="top")
CEN = Alignment(horizontal="center", vertical="center", wrap_text=True)
NUM = '#,##0;(#,##0);"-"'
PCT = '0.0%;(0.0%);"-"'
DATE = 'dd-mmm-yy'

wb = Workbook()
names = {}
def dname(n, ref): names[n] = ref

order = ["Dashboard", "LayersMap", "Setup", "WP_Index", "L1_Strategy", "L2_Planning",
         "L3_Methodology", "L4_Execution", "L5_Findings", "L6_Oversight", "DashData", "Standards", "Lists"]
ws0 = wb.active; ws0.title = order[0]
S = {order[0]: ws0}
for n in order[1:]: S[n] = wb.create_sheet(n)
for ws in S.values():
    ws.sheet_view.showGridLines = False

tabcol = {"Dashboard": DK, "LayersMap": DK, "Setup": "F9A825", "WP_Index": "F9A825", "DashData": "9E9E9E",
          "Standards": "9E9E9E", "Lists": "9E9E9E"}
for n, ws in S.items(): ws.sheet_properties.tabColor = tabcol.get(n, LT)

def put(ws, ref, v, f=None, fl=None, nf=None, al=None, bd=None):
    c = ws[ref]; c.value = v
    c.font = f or font()
    if fl: c.fill = fill(fl)
    if nf: c.number_format = nf
    if al: c.alignment = al
    if bd: c.border = bd
    return c

def widths(ws, w):
    for k, v in w.items(): ws.column_dimensions[k].width = v

# Lists ------------------------------------------------------------
L = S["Lists"]
lists = {
    "YesNo": ["Yes", "No", "N/A"],
    "Status": ["Not started", "In progress", "Prepared", "Reviewed", "Closed"],
    "Layer": ["Strategy", "Planning", "Methodology", "Execution", "Findings", "Oversight"],
    "Severity": ["High", "Medium", "Low"],
    "IssueStatus": ["Open", "In progress", "Closed"],
    "MSType": ["Factual", "Judgmental", "Projected"],
    "Benchmark": ["Profit before tax", "Revenue", "Total assets", "Equity"],
    "Done": ["Done", "Pending", "N/A"],
    "Threat": ["Self-interest", "Self-review", "Advocacy", "Familiarity", "Intimidation"],
    "Freq": ["Annual", "Quarterly", "Monthly", "Weekly", "Daily", "Multiple daily"],
    "Confidence": ["95%", "90%", "80%"],
    "ConfStatus": ["Received - agreed", "Received - reconciled", "Not received - alt. procedures", "Outstanding"],
    "Approach": ["Controls reliance + substantive", "Substantive only"],
}
put(L, "A1", "Validation lists (do not edit)", font(11, True))
for j, (k, vals) in enumerate(lists.items()):
    col = chr(65 + j)
    put(L, f"{col}2", k, font(b=True))
    for i, v in enumerate(vals): put(L, f"{col}{3+i}", v)
    dname("lst_" + k, f"Lists!${col}$3:${col}${2+len(vals)}")
    L.column_dimensions[col].width = 26

def dv(ws, rng, lst):
    d = DataValidation(type="list", formula1="=lst_" + lst, allow_blank=True)
    ws.add_data_validation(d); d.add(rng)

def inp(c): c.fill = fill(INP); return c

# Setup ------------------------------------------------------------
st = S["Setup"]
widths(st, {"A": 3, "B": 38, "C": 34, "D": 60})
put(st, "B2", "ENGAGEMENT SETUP  |  Key inputs used across all working papers", font(14, True, "FFFFFF"), DK)
st.merge_cells("B2:D2"); st.row_dimensions[2].height = 26
setup = [
    ("Client name", "ABC Trading LLC", "Illustrative client - replace", "Client"),
    ("Financial year end", D(2025, 12, 31), "Period covered by the audit", "FYE"),
    ("Financial reporting framework", "IFRS Accounting Standards", "ISA 210.6(a) acceptability", "Framework"),
    ("Presentation currency / unit", "MNT '000", "All amounts in thousands", "Unit"),
    ("Entity type", "Non-PIE (private company)", "PIE status drives IESBA rotation & fee rules", "EntityType"),
    ("Audit firm", "[Audit Firm Name]", "", "Firm"),
    ("Engagement partner", "B. Bat-Erdene", "ISA 220 (Revised) - overall responsibility", "EP"),
    ("Engagement manager", "S. Oyunaa", "", "Mgr"),
    ("Engagement quality reviewer", "D. Ganbold", "ISQM 2 - required by firm policy", "EQR"),
    ("Planned audit report date", D(2026, 3, 31), "ISA 700.49 - not before sufficient evidence obtained", "RptDate"),
    ("Dashboard status (as-of) date", D(2026, 3, 20), "Drives ageing & countdown KPIs", "AsOf"),
    ("File assembly deadline", "=Set_RptDate+60", "ISA 230.A21 / ISQM 1.A83 - normally within 60 days", "AssemblyDate"),
    ("Total firm fee income (MNT '000)", 2400000, "For IESBA 410 fee-dependency test", "FirmFees"),
    ("Audit fee - this client (MNT '000)", 120000, "Engagement letter", "AuditFee"),
    ("Non-assurance fees - this client (MNT '000)", 30000, "Tax compliance (IESBA 604)", "NASFee"),
]
put(st, "B4", "Input", font(b=True, c="FFFFFF"), MD); put(st, "C4", "Value", font(b=True, c="FFFFFF"), MD)
put(st, "D4", "Note / source", font(b=True, c="FFFFFF"), MD)
for i, (lab, val, note, nm) in enumerate(setup):
    r = 5 + i
    put(st, f"B{r}", lab, bd=BOX)
    c = put(st, f"C{r}", val, f=font(c="0000FF") if not str(val).startswith("=") else font(), bd=BOX)
    if not str(val).startswith("="): inp(c)
    if isinstance(val, D) or nm == "AssemblyDate": c.number_format = DATE
    if isinstance(val, int): c.number_format = NUM
    put(st, f"D{r}", note, font(9, i=True, c="616161"), bd=BOX)
    dname("Set_" + nm, f"Setup!$C${r}")
r = 5 + len(setup) + 1
put(st, f"B{r}", "LEGEND", font(b=True))
put(st, f"B{r+1}", "Yellow cell, blue text", font(c="0000FF"), INP); put(st, f"C{r+1}", "Input - edit these")
put(st, f"B{r+2}", "Black text", font()); put(st, f"C{r+2}", "Formula - do not overwrite")
put(st, f"B{r+3}", "Green text", font(c="008000")); put(st, f"C{r+3}", "Link from another working paper")
put(st, f"B{r+5}", "All client data in this file is illustrative, for teaching. Replace with real engagement data.",
    font(9, i=True, c="C62828"))

# ---------- layer sheet helpers ----------
COLW = {"A": 6, "B": 44, "C": 18, "D": 16, "E": 16, "F": 16, "G": 15, "H": 15, "I": 15, "J": 14, "K": 18, "L": 16, "M": 14}
anchors = {}

def sheet_header(ws, title, layer_no, objective, stds):
    widths(ws, COLW)
    put(ws, "A1", f"LAYER {layer_no}  |  {title}", font(15, True, "FFFFFF"), DK)
    ws.merge_cells("A1:L1"); ws.row_dimensions[1].height = 28
    rows = [("Client", "=Set_Client"), ("Period end", "=Set_FYE"), ("Framework", "=Set_Framework"),
            ("Objective", objective), ("Standards", stds)]
    for i, (k, v) in enumerate(rows):
        r = 2 + i
        put(ws, f"A{r}", k, font(9, True), PALE); ws.merge_cells(f"A{r}:A{r}")
        c = put(ws, f"B{r}", v, font(9, c="008000" if str(v).startswith("=") else "000000"), al=WRAP)
        if k == "Period end": c.number_format = DATE; c.alignment = Alignment(horizontal="left")
        ws.merge_cells(f"B{r}:F{r}")
    # sign-off block (ISA 230.9)
    for j, (k, v) in enumerate([("Prepared by", ""), ("Date", ""), ("Reviewed by", ""), ("Date", ""),
                                ("Partner review", "")]):
        r = 2 + j
        put(ws, f"H{r}", k, font(9, True), PALE, bd=BOX)
        c = put(ws, f"I{r}", v, font(9, c="0000FF"), bd=BOX); inp(c)
        ws.merge_cells(f"I{r}:J{r}")
        if k == "Date": c.number_format = DATE
    put(ws, "K2", "ISA 230.9: record who performed\nthe work, when, and who\nreviewed it and when.", font(8, i=True, c="616161"), al=WRAP)
    ws.merge_cells("K2:L4")
    ws.row_dimensions[5].height = 30
    ws.freeze_panes = "A8"
    return 8

def section(ws, r, ref, title, std):
    put(ws, f"A{r}", ref, font(11, True, "FFFFFF"), MD)
    put(ws, f"B{r}", title, font(11, True, "FFFFFF"), MD)
    ws.merge_cells(f"B{r}:H{r}")
    put(ws, f"I{r}", std, font(9, True, "FFFFFF"), MD, al=Alignment(horizontal="right"))
    ws.merge_cells(f"I{r}:L{r}")
    ws.row_dimensions[r].height = 20
    anchors[ref] = (ws.title, r)
    return r + 1

def table_head(ws, r, heads, start="A"):
    for j, h in enumerate(heads):
        col = chr(ord(start) + j)
        put(ws, f"{col}{r}", h, font(9, True, "FFFFFF"), DK, al=CEN, bd=BOX)
    ws.row_dimensions[r].height = 30
    return r + 1

def cell(ws, ref, v, nf=None, is_input=False, link=False, center=False):
    c = put(ws, ref, v, font(9, c="0000FF" if is_input else ("008000" if link else "000000")),
            nf=nf, al=CEN if center else WRAP, bd=BOX)
    if is_input: inp(c)
    return c

def checklist(ws, r, items, wp):
    """items: (question, std ref, response, comment)"""
    r = table_head(ws, r, ["#", "Procedure / requirement", "Standard ref", "Response", "Comment / evidence obtained"])
    ws.merge_cells(f"E{r-1}:J{r-1}")
    top = r
    for i, (q, ref, resp, com) in enumerate(items, 1):
        cell(ws, f"A{r}", i, center=True); cell(ws, f"B{r}", q); cell(ws, f"C{r}", ref)
        cell(ws, f"D{r}", resp, is_input=True, center=True)
        cell(ws, f"E{r}", com, is_input=True); ws.merge_cells(f"E{r}:J{r}")
        cell(ws, f"K{r}", wp, center=True)
        ws.row_dimensions[r].height = 30
        r += 1
    dv(ws, f"D{top}:D{r-1}", "YesNo")
    put(ws, f"A{r}", "Concl.", font(9, True), PALE, bd=BOX)
    put(ws, f"B{r}", f'=IF(COUNTBLANK(D{top}:D{r-1})>0,COUNTBLANK(D{top}:D{r-1})&" item(s) pending",'
                     f'IF(COUNTIF(D{top}:D{r-1},"No")>0,COUNTIF(D{top}:D{r-1},"No")&" exception(s) - escalate to Engagement Partner",'
                     f'"All requirements addressed - no exceptions"))', font(9, True, MD), PALE, bd=BOX)
    ws.merge_cells(f"B{r}:J{r}")
    return r + 2

# =====================================================================
# LAYER 4 first concepts needed by names; but each sheet is independent via names.
# =====================================================================

# ---------------- L1 STRATEGY ----------------
ws = S["L1_Strategy"]
r = sheet_header(ws, "AUDIT STRATEGY LAYER", 1,
                 "Decide whether to accept/continue the engagement, agree terms, confirm independence and set the overall audit strategy and direction.",
                 "ISQM 1 para 30; ISA 210; ISA 220 (Revised); ISA 300; ISA 260; IESBA Code Parts 1, 3 & 4A")
r = section(ws, r, "A1", "Engagement acceptance & continuance", "ISQM 1.30 | ISA 220R.22-24")
r = checklist(ws, r, [
    ("Integrity of owners, key management and TCWG assessed (no information indicating lack of integrity)", "ISQM 1.30(a)(ii)", "Yes", "Background & media searches performed; no adverse findings"),
    ("Firm has competence, capabilities, time and resources to perform the engagement", "ISQM 1.30(a)(i); ISA 220R.25", "Yes", "Team includes IFRS & retail sector experience"),
    ("Relevant ethical requirements incl. independence can be complied with", "ISA 220R.16-21; IESBA 400", "Yes", "See A3 & A4"),
    ("Significant matters from prior engagement / previous auditor considered", "ISA 220R.23; IESBA 320", "Yes", "Predecessor communication not required (continuing)"),
    ("Financial & reputational risk of the client evaluated (fee level, going concern, litigation)", "ISQM 1.30(c)", "Yes", "Low-moderate risk rating"),
    ("Acceptance/continuance approved by Engagement Partner (and risk partner where high risk)", "ISQM 1.30; ISA 220R.22", "Yes", "Approved 10-Sep-25"),
], "A1")
r = section(ws, r, "A2", "Terms of engagement & preconditions for an audit", "ISA 210")
r = checklist(ws, r, [
    ("Financial reporting framework is acceptable", "ISA 210.6(a)", "Yes", "IFRS - general purpose framework"),
    ("Management acknowledges responsibility for preparation of FS", "ISA 210.6(b)(i)", "Yes", ""),
    ("Management acknowledges responsibility for internal control", "ISA 210.6(b)(ii)", "Yes", ""),
    ("Management will provide access to all information & unrestricted access to persons", "ISA 210.6(b)(iii)", "Yes", ""),
    ("Engagement letter covers objective, scope, responsibilities, framework, expected report form", "ISA 210.10", "Yes", "Signed by client CEO 22-Sep-25"),
    ("Recurring audit: need to revise terms considered", "ISA 210.13", "Yes", "No change in circumstances"),
], "A2")

r = section(ws, r, "A3", "Independence & ethics declarations - engagement team", "IESBA Part 4A | ISA 220R.16-21")
heads = ["#", "Team member", "Role", "Financial interest?", "Loans / guarantees?", "Business relationship?",
         "Family / personal relationship?", "Employment with client?", "Years on engagement", "Declaration date", "Assessment", "IESBA ref"]
r = table_head(ws, r, heads)
team = [("B. Bat-Erdene", "Engagement partner", "No", "No", "No", "No", "No", 4, D(2025, 9, 12)),
        ("D. Ganbold", "EQ reviewer", "No", "No", "No", "No", "No", 1, D(2025, 9, 12)),
        ("S. Oyunaa", "Manager", "No", "No", "No", "No", "No", 3, D(2025, 9, 15)),
        ("T. Enkhjin", "Senior", "No", "No", "No", "Yes", "No", 2, D(2025, 9, 15)),
        ("M. Saruul", "Staff", "No", "No", "No", "No", "No", 1, D(2025, 9, 16)),
        ("N. Temuulen", "Staff", "No", "No", "No", "No", "No", 1, None),
        ("IT auditor (firm)", "Specialist", "No", "No", "No", "No", "No", 1, D(2025, 10, 2))]
t0 = r
for i, row in enumerate(team, 1):
    cell(ws, f"A{r}", i, center=True)
    for j, v in enumerate(row):
        col = chr(66 + j)
        cell(ws, f"{col}{r}", v, nf=DATE if j == 8 else None, is_input=j >= 2, center=j >= 2)
    cell(ws, f"K{r}", f'=IF(J{r}="","Declaration outstanding",IF(COUNTIF(D{r}:H{r},"Yes")>0,"Threat - evaluate (A4)",'
                      f'IF(AND(C{r}="Engagement partner",I{r}>7),"Rotation required (IESBA 540)","Independent")))', center=True)
    cell(ws, f"L{r}", "510-524; 540", center=True)
    r += 1
dv(ws, f"D{t0}:H{r-1}", "YesNo")
dname("Ind_Assess", f"L1_Strategy!$K${t0}:$K${r-1}")
dname("Ind_Date", f"L1_Strategy!$J${t0}:$J${r-1}")
put(ws, f"B{r}", "Declarations received", font(9, True), PALE, bd=BOX)
put(ws, f"C{r}", f'=COUNT(J{t0}:J{r-1})&" of "&COUNTA(B{t0}:B{r-1})', font(9, True), PALE, bd=BOX)
put(ws, f"E{r}", "Independent (no threats)", font(9, True), PALE, bd=BOX)
put(ws, f"F{r}", f'=COUNTIF(K{t0}:K{r-1},"Independent")', font(9, True), PALE, bd=BOX)
r += 2

r = section(ws, r, "A4", "Threats & safeguards register (conceptual framework)", "IESBA 120; 300 series")
r = table_head(ws, r, ["#", "Threat identified", "Category", "Level (before)", "Safeguard / action", "", "", "Level (after)",
                       "Acceptable?", "Approved by", "Date", "IESBA ref"])
ws.merge_cells(f"E{r-1}:G{r-1}")
threats = [("Senior T. Enkhjin's sibling is accounts payable clerk at client", "Familiarity", "Medium",
            "Senior removed from payables & purchases testing; work reviewed by manager", "Low", "Yes", "B. Bat-Erdene", D(2025, 9, 18), "521"),
           ("Firm prepares corporate income tax return", "Self-review", "Medium",
            "Separate tax team; management takes all decisions; no effect on FS amounts audited", "Low", "Yes", "B. Bat-Erdene", D(2025, 9, 18), "604"),
           ("Engagement partner 4 years on engagement", "Familiarity", "Low",
            "Monitor; non-PIE - no mandatory rotation; EQR in place", "Low", "Yes", "Risk partner", D(2025, 9, 18), "540"),
           ("Audit fee overdue > 6 months at planning date", "Self-interest", "Low",
            "Fees settled 30-Sep-25 before report", "Low", "Yes", "B. Bat-Erdene", D(2025, 10, 1), "410.7")]
t0 = r
for i, row in enumerate(threats, 1):
    cell(ws, f"A{r}", i, center=True); cell(ws, f"B{r}", row[0], is_input=True)
    cell(ws, f"C{r}", row[1], is_input=True, center=True); cell(ws, f"D{r}", row[2], is_input=True, center=True)
    cell(ws, f"E{r}", row[3], is_input=True); ws.merge_cells(f"E{r}:G{r}")
    cell(ws, f"H{r}", row[4], is_input=True, center=True); cell(ws, f"I{r}", row[5], is_input=True, center=True)
    cell(ws, f"J{r}", row[6], is_input=True); cell(ws, f"K{r}", row[7], DATE, True, center=True); cell(ws, f"L{r}", row[8], center=True)
    ws.row_dimensions[r].height = 32
    r += 1
dv(ws, f"C{t0}:C{r-1}", "Threat"); dv(ws, f"D{t0}:D{r-1}", "Severity"); dv(ws, f"H{t0}:H{r-1}", "Severity"); dv(ws, f"I{t0}:I{r-1}", "YesNo")
r += 1

r = section(ws, r, "A5", "Fee dependency & non-assurance services", "IESBA 410; 600-610")
fee_rows = [("Audit fee from client", "=Set_AuditFee", NUM), ("Non-assurance fees from client", "=Set_NASFee", NUM),
            ("Total fees from client", None, NUM), ("Total firm fee income", "=Set_FirmFees", NUM),
            ("Client fees as % of firm total", None, PCT), ("NAS fees as % of audit fee", None, PCT),
            ("Fee dependency threshold (PIE, 2 consecutive years)", 0.15, PCT), ("Conclusion", None, None)]
r0 = r
for i, (lab, v, nf) in enumerate(fee_rows):
    rr = r + i
    cell(ws, f"B{rr}", lab)
    if lab.startswith("Total fees"): v = f"=C{r0}+C{r0+1}"
    if lab.startswith("Client fees"): v = f"=IFERROR(C{r0+2}/C{r0+3},0)"
    if lab.startswith("NAS fees"): v = f"=IFERROR(C{r0+1}/C{r0},0)"
    if lab == "Conclusion":
        v = (f'=IF(C{r0+4}>C{r0+6},"Fee dependency exceeds 15% - disclose to TCWG & apply pre/post-issuance review (IESBA 410.4)",'
             f'"Below 15% threshold - self-interest threat at acceptable level")')
    c = cell(ws, f"C{rr}", v, nf, is_input=(lab.startswith("Fee dependency")), link=str(v).startswith("=Set"))
    if lab == "Conclusion": ws.merge_cells(f"C{rr}:J{rr}"); c.font = font(9, True, MD)
r += len(fee_rows) + 1

r = section(ws, r, "A6", "Overall audit strategy", "ISA 300.7-8")
r = table_head(ws, r, ["#", "Strategy element", "Standard ref", "Decision / plan", "", "", "", "", "", "", "Owner"])
ws.merge_cells(f"D{r-1}:J{r-1}")
strat = [("Scope: characteristics of the engagement, framework, locations", "ISA 300.8(a)", "Single entity, HQ + 3 retail branches + 1 warehouse; IFRS"),
         ("Reporting objectives, timing and nature of communications", "ISA 300.8(b)", "Planning meeting Oct-25; interim Dec-25; final report 31-Mar-26"),
         ("Significant factors directing the team's efforts (materiality, risk areas)", "ISA 300.8(c)", "Revenue cut-off, inventory valuation, mgmt override, IT general controls"),
         ("Results of preliminary activities & prior-year experience", "ISA 300.8(d)", "PY: inventory NRV adjustment proposed; controls over sales generally effective"),
         ("Nature, timing & extent of resources (team, experts, EQR, partner involvement)", "ISA 300.8(e); ISA 220R.25", "Partner 10%, IT auditor for ITGCs, EQR required; budget per B10"),
         ("Audit approach: combined (controls reliance) for revenue & purchases; substantive for others", "ISA 330.7", "See C1 audit programme")]
for i, (a, b, c_) in enumerate(strat, 1):
    cell(ws, f"A{r}", i, center=True); cell(ws, f"B{r}", a); cell(ws, f"C{r}", b)
    cell(ws, f"D{r}", c_, is_input=True); ws.merge_cells(f"D{r}:J{r}"); cell(ws, f"K{r}", "EP / Manager", center=True)
    ws.row_dimensions[r].height = 30; r += 1
r += 1
r = section(ws, r, "A7", "Communication of planned scope & timing to TCWG", "ISA 260.15")
r = checklist(ws, r, [
    ("Planned scope and timing communicated, incl. significant risks identified", "ISA 260.15; A12-A13", "Yes", "Planning letter issued 20-Oct-25"),
    ("Auditor's responsibilities regarding FS audit communicated", "ISA 260.14", "Yes", ""),
    ("Independence statement planned for TCWG (listed entities required)", "ISA 260.17", "N/A", "Non-listed - firm policy: provide anyway"),
    ("Form, timing and expected general content of communications agreed", "ISA 260.18", "Yes", ""),
], "A7")

# ---------------- L2 PLANNING ----------------
ws = S["L2_Planning"]
r = sheet_header(ws, "AUDIT PLANNING LAYER", 2,
                 "Understand the entity and its system of internal control, set materiality and identify & assess risks of material misstatement.",
                 "ISA 300; ISA 315 (Revised 2019); ISA 320; ISA 240; ISA 250; ISA 550; ISA 570 (Revised); ISA 510")
r = section(ws, r, "B1", "Understanding the entity, its environment & framework", "ISA 315R.19")
r = checklist(ws, r, [
    ("Organisational structure, ownership, governance and business model understood", "ISA 315R.19(a)(i)", "Yes", "Family-owned; Board of 5; retail of consumer goods"),
    ("Industry, regulatory and other external factors (incl. tax, inflation, FX) understood", "ISA 315R.19(a)(ii)", "Yes", "MNT depreciation affects imported inventory cost"),
    ("Measures used to assess financial performance (KPIs, bonuses) understood", "ISA 315R.19(a)(iii)", "Yes", "Mgmt bonus linked to EBITDA - fraud risk factor"),
    ("Applicable framework, accounting policies & reasons for changes understood", "ISA 315R.19(b)", "Yes", "IFRS 16 & IFRS 15 policies unchanged"),
    ("Inherent risk factors considered (complexity, subjectivity, change, uncertainty, bias)", "ISA 315R.19(c)", "Yes", "Applied in B4 register"),
], "B1")
r = section(ws, r, "B2", "Understanding the components of the system of internal control", "ISA 315R.21-26")
r = checklist(ws, r, [
    ("Control environment evaluated", "ISA 315R.21", "Yes", "Tone at top adequate; no internal audit function"),
    ("Entity's risk assessment process evaluated", "ISA 315R.22-23", "Yes", "Informal process - deficiency noted (E2)"),
    ("Process to monitor the system of internal control evaluated", "ISA 315R.24", "Yes", ""),
    ("Information system & communication understood (incl. flows of transactions)", "ISA 315R.25", "Yes", "Walkthroughs: sales, purchases, payroll, inventory"),
    ("Control activities: controls addressing RMM identified; design & implementation evaluated", "ISA 315R.26", "Yes", "See D2"),
    ("IT environment & general IT controls (access, change, operations) evaluated", "ISA 315R.26(b)-(c)", "Yes", "IT auditor involved; user access deficiency"),
], "B2")

r = section(ws, r, "B3", "Materiality", "ISA 320.10-11 | ISA 450")
r = table_head(ws, r, ["", "Item", "Value", "Note / rationale"])
ws.merge_cells(f"D{r-1}:J{r-1}")
m0 = r
mat = [
    ("Benchmark selected", "Profit before tax", None, "Profit-oriented entity; users focus on earnings (ISA 320.A4)", True),
    ("Benchmark amount (unadjusted, from D1)", f'=IF(C{m0}="Profit before tax",TB_PBT,IF(C{m0}="Revenue",TB_Rev,IF(C{m0}="Total assets",TB_TA,TB_Eq)))', NUM, "Linked to lead schedule", False),
    ("Percentage applied", 0.05, PCT, None, True),
    ("Typical range for benchmark", f'=IF(C{m0}="Profit before tax","5% - 10%",IF(C{m0}="Revenue","0.5% - 1%",IF(C{m0}="Total assets","1% - 2%","2% - 5%")))', None, "Firm methodology guidance", False),
    ("OVERALL MATERIALITY (OM)", f"=ROUND(C{m0+1}*C{m0+2},-3)", NUM, "Rounded to nearest thousand", False),
    ("Performance materiality %", 0.75, PCT, "Moderate risk: PY misstatements found, controls generally effective", True),
    ("PERFORMANCE MATERIALITY (PM)", f"=ROUND(C{m0+4}*C{m0+5},-3)", NUM, "ISA 320.11", False),
    ("Clearly trivial %", 0.05, PCT, "Threshold for accumulating misstatements", True),
    ("CLEARLY TRIVIAL THRESHOLD (CT)", f"=ROUND(C{m0+4}*C{m0+7},-3)", NUM, "ISA 450.5 & A2", False),
    ("Specific materiality - related party transactions & directors' remuneration", f"=ROUND(C{m0+4}*0.5,-3)", NUM, "Qualitatively sensitive disclosures (ISA 320.10)", False),
]
for i, (lab, v, nf, note, isin) in enumerate(mat):
    rr = r + i
    b = cell(ws, f"B{rr}", lab)
    c = cell(ws, f"C{rr}", v, nf, is_input=isin, link=(i == 1))
    if lab.isupper() or lab.startswith(("OVERALL", "PERFORMANCE M", "CLEARLY TRIVIAL T")):
        b.font = font(9, True); c.font = font(10, True, DK); b.fill = fill(PALE); c.fill = fill(PALE)
    cell(ws, f"D{rr}", note or (f'="Guidance range: "&C{m0+3}' if i == 2 else ""), is_input=isin and i != 2)
    ws.merge_cells(f"D{rr}:J{rr}")
dv(ws, f"C{m0}", "Benchmark")
dname("Mat_OM", f"L2_Planning!$C${m0+4}"); dname("Mat_PM", f"L2_Planning!$C${m0+6}")
dname("Mat_CT", f"L2_Planning!$C${m0+8}"); dname("Mat_Bench", f"L2_Planning!$C${m0}")
r += len(mat) + 1

r = section(ws, r, "B4", "Risk assessment register - risks of material misstatement", "ISA 315R.28-37 | ISA 330")
r = table_head(ws, r, ["ID", "Risk description", "FS area", "Assertion(s)", "Likelihood (1-5)", "Magnitude (1-5)",
                       "Inherent risk score", "Spectrum of IR", "Significant risk?", "Control risk", "Planned response", "WP ref"])
risks = [("Fraud in revenue recognition - cut-off around year end (presumed)", "Revenue", "Occurrence, Cut-off", 4, 5, "Yes", "Not tested", "Extended cut-off testing 10 days either side; credit notes after YE", "D3"),
         ("Management override of controls (presumed)", "All", "All", 4, 5, "Yes", "Not tested", "JE testing, estimates bias review, unusual transactions", "D6"),
         ("Inventory NRV - slow-moving imported goods with MNT depreciation", "Inventory", "Valuation", 4, 4, "No", "Not tested", "Ageing analysis, post-YE selling prices, NRV recompute", "D3"),
         ("Inventory existence across 3 branches & warehouse", "Inventory", "Existence", 3, 4, "No", "Rely", "Attend counts at warehouse + 2 branches; test counts", "D5"),
         ("Receivables recoverability (ECL - IFRS 9)", "Trade receivables", "Valuation", 3, 3, "No", "Not tested", "Ageing, after-date cash, ECL matrix recompute", "D4"),
         ("Completeness of accruals & payables", "Payables & accruals", "Completeness", 3, 3, "No", "Rely", "Search for unrecorded liabilities", "D3"),
         ("PPE capitalisation vs repairs", "PPE", "Existence, Classification", 2, 3, "No", "Not tested", "Vouch additions > PM/2; physical inspection", "D3"),
         ("Borrowings covenant compliance & classification", "Borrowings", "Presentation", 2, 4, "No", "Not tested", "Covenant recalculation; bank confirmation", "D4"),
         ("Related party transactions not identified / disclosed", "Disclosures", "Completeness", 3, 2, "No", "Not tested", "RP inquiries, board minutes, ledger scans", "B7"),
         ("Cash and bank balances", "Cash", "Existence", 1, 2, "No", "Rely", "Bank confirmations & reconciliations", "D4")]
r0 = r
for i, row in enumerate(risks, 1):
    cell(ws, f"A{r}", f"R{i:02d}", center=True); cell(ws, f"B{r}", row[0], is_input=True)
    cell(ws, f"C{r}", row[1], is_input=True); cell(ws, f"D{r}", row[2], is_input=True)
    cell(ws, f"E{r}", row[3], is_input=True, center=True); cell(ws, f"F{r}", row[4], is_input=True, center=True)
    cell(ws, f"G{r}", f"=E{r}*F{r}", center=True)
    cell(ws, f"H{r}", f'=IF(I{r}="Yes","Significant",IF(G{r}>=12,"Higher",IF(G{r}>=6,"Moderate","Low")))', center=True)
    cell(ws, f"I{r}", row[5], is_input=True, center=True); cell(ws, f"J{r}", row[6], is_input=True, center=True)
    cell(ws, f"K{r}", row[7], is_input=True); cell(ws, f"L{r}", row[8], center=True)
    ws.row_dimensions[r].height = 36; r += 1
dv(ws, f"I{r0}:I{r-1}", "YesNo")
dname("Risk_Spectrum", f"L2_Planning!$H${r0}:$H${r-1}")
ws.conditional_formatting.add(f"H{r0}:H{r-1}", CellIsRule(operator="equal", formula=['"Significant"'], fill=fill("FFCDD2"), font=Font(name=F, bold=True, color="B71C1C")))
ws.conditional_formatting.add(f"H{r0}:H{r-1}", CellIsRule(operator="equal", formula=['"Higher"'], fill=fill("FFE0B2")))
put(ws, f"B{r}", "Scoring: Likelihood x Magnitude. >=12 Higher, 6-11 Moderate, <6 Low; 'Significant' where judged at the upper end of the spectrum (ISA 315R.12(l)) incl. presumed fraud risks (ISA 240.26-27, 31).", font(8, i=True, c="616161"), al=WRAP)
ws.merge_cells(f"B{r}:L{r}"); ws.row_dimensions[r].height = 24
r += 2

r = section(ws, r, "B5", "Fraud risk assessment", "ISA 240")
r = checklist(ws, r, [
    ("Engagement team discussion on susceptibility of FS to fraud held & documented", "ISA 240.16", "Yes", "Planning meeting 15-Oct-25, all team attended"),
    ("Inquiries of management, TCWG and others regarding fraud", "ISA 240.18-22", "Yes", "No known fraud reported"),
    ("Fraud risk factors (incentive/pressure, opportunity, rationalisation) evaluated", "ISA 240.24", "Yes", "EBITDA-linked bonus = incentive"),
    ("Presumption of fraud risk in revenue recognition evaluated (rebuttal documented if rebutted)", "ISA 240.26-27, 47", "Yes", "Not rebutted - R01"),
    ("Responses to management override: JE testing, estimates review, unusual transactions", "ISA 240.32-33", "Yes", "See D6"),
], "B5")
r = section(ws, r, "B6", "Laws & regulations", "ISA 250 (Revised)")
r = checklist(ws, r, [
    ("Understanding of legal & regulatory framework applicable to entity obtained", "ISA 250.13", "Yes", "Tax, labour, consumer protection, AML"),
    ("Evidence of compliance with laws having direct effect on FS (e.g. tax)", "ISA 250.14", "Yes", ""),
    ("Inquiries & inspection of correspondence with regulators", "ISA 250.15", "Yes", "Tax inspection 2024 - closed, no penalties"),
    ("Written representation on non-compliance obtained", "ISA 250.17", None, "At completion - F2"),
], "B6")
r = section(ws, r, "B7", "Related parties", "ISA 550")
r = checklist(ws, r, [
    ("Identity of related parties & changes from PY obtained from management", "ISA 550.13", "Yes", "List of 6 related entities"),
    ("Controls to identify, account for & approve RP transactions understood", "ISA 550.14", "Yes", ""),
    ("Records reviewed for undisclosed RP relationships / transactions", "ISA 550.15", "Yes", "Board minutes, contracts, large payments"),
    ("Significant RP transactions outside normal course of business evaluated", "ISA 550.23", "N/A", "None identified"),
    ("RP disclosures (IAS 24) adequate", "ISA 550.25", None, "Final review"),
], "B7")
r = section(ws, r, "B8", "Going concern - preliminary assessment", "ISA 570 (Revised).10-11")
r = checklist(ws, r, [
    ("Events or conditions casting significant doubt considered at planning", "ISA 570.10", "Yes", "Profitable; positive operating cash flow"),
    ("Management's preliminary assessment obtained & discussed", "ISA 570.10(a)", "Yes", "12-month cash forecast"),
    ("Borrowing facilities & covenants reviewed for maturity within 12 months", "ISA 570.A3", "Yes", "Loan matures 2028; covenants met"),
], "B8")
r = section(ws, r, "B9", "Opening balances (initial engagements)", "ISA 510")
r = checklist(ws, r, [
    ("Engagement is an initial audit engagement", "ISA 510.2", "No", "Continuing engagement - year 4"),
    ("Opening balances agreed to prior period closing balances", "ISA 510.6(a)", "Yes", "Agreed to PY signed FS"),
    ("Accounting policies consistently applied", "ISA 510.7", "Yes", ""),
], "B9")
r = section(ws, r, "B10", "Resource plan & audit timeline", "ISA 300.9 | ISA 220R.25-28")
r = table_head(ws, r, ["#", "Phase / milestone", "Start", "End", "Days", "% complete", "Status", "Responsible", "Layer"])
phases = [("Acceptance, independence & engagement letter", D(2025, 9, 1), D(2025, 9, 26), 1.0, "EP", "Strategy"),
          ("Planning & risk assessment", D(2025, 10, 1), D(2025, 11, 14), 1.0, "Manager", "Planning"),
          ("Interim fieldwork - controls testing", D(2025, 11, 17), D(2025, 12, 12), 1.0, "Senior", "Execution"),
          ("Year-end inventory count attendance", D(2025, 12, 31), D(2026, 1, 2), 1.0, "Senior", "Execution"),
          ("Final fieldwork - substantive procedures", D(2026, 2, 2), D(2026, 3, 13), 0.9, "Senior", "Execution"),
          ("Findings, misstatements & management letter", D(2026, 3, 9), D(2026, 3, 20), 0.7, "Manager", "Findings"),
          ("Completion, EQR & TCWG communication", D(2026, 3, 16), D(2026, 3, 27), 0.4, "EP / EQR", "Oversight"),
          ("Audit report signed", D(2026, 3, 31), D(2026, 3, 31), 0.0, "EP", "Oversight"),
          ("Final file assembly (within 60 days)", D(2026, 4, 1), D(2026, 5, 30), 0.0, "Manager", "Oversight")]
p0 = r
for i, (ph, s, e, pc, who, lay) in enumerate(phases, 1):
    cell(ws, f"A{r}", i, center=True); cell(ws, f"B{r}", ph)
    cell(ws, f"C{r}", s, DATE, True, center=True); cell(ws, f"D{r}", e, DATE, True, center=True)
    cell(ws, f"E{r}", f"=D{r}-C{r}+1", "0", center=True); cell(ws, f"F{r}", pc, "0%", True, center=True)
    cell(ws, f"G{r}", f'=IF(F{r}>=1,"Complete",IF(Set_AsOf>D{r},"Overdue",IF(F{r}>0,"On track","Scheduled")))', center=True)
    cell(ws, f"H{r}", who, center=True); cell(ws, f"I{r}", lay, center=True)
    r += 1
dname("TL_Phase", f"L2_Planning!$B${p0}:$B${r-1}"); dname("TL_Pct", f"L2_Planning!$F${p0}:$F${r-1}")
ws.conditional_formatting.add(f"G{p0}:G{r-1}", CellIsRule(operator="equal", formula=['"Overdue"'], fill=fill("FFCDD2")))
ws.conditional_formatting.add(f"G{p0}:G{r-1}", CellIsRule(operator="equal", formula=['"Complete"'], fill=fill("C8E6C9")))
PH = (p0, r - 1)

# ---------------- L3 METHODOLOGY ----------------
ws = S["L3_Methodology"]
r = sheet_header(ws, "AUDIT METHODOLOGY LAYER", 3,
                 "Design further audit procedures responsive to assessed risks, using standard procedures, sampling and evaluation methods.",
                 "ISA 330; ISA 500; ISA 520; ISA 530; ISA 540 (Revised); ISA 600 (Revised); ISA 610; ISA 620; ISA 402")
r = section(ws, r, "C1", "Audit programme - responses by FS area & assertion", "ISA 330.6-7, 18")
r = table_head(ws, r, ["#", "FS area", "Highest IR", "E/O", "C", "A/V", "Cut-off", "P&D", "Approach", "", "Key procedures", "WP ref"])
ws.merge_cells(f"I{r-1}:J{r-1}")
prog = [("Revenue", "Significant", "x", "x", "x", "x", "x", "Controls reliance + substantive", "ToC over sales approval; substantive analytics; cut-off", "D3"),
        ("Inventory", "Higher", "x", "x", "x", "x", "", "Controls reliance + substantive", "Count attendance; NRV testing; costing", "D5"),
        ("Trade receivables", "Moderate", "x", "", "x", "", "x", "Substantive only", "Confirmations; ECL recompute", "D4"),
        ("PPE", "Moderate", "x", "x", "x", "", "x", "Substantive only", "Vouch additions; depreciation proof-in-total", "D3"),
        ("Payables & accruals", "Moderate", "", "x", "x", "x", "x", "Controls reliance + substantive", "Unrecorded liabilities search", "D3"),
        ("Borrowings", "Moderate", "x", "x", "x", "", "x", "Substantive only", "Bank confirmation; covenant test", "D4"),
        ("Cash", "Low", "x", "x", "x", "", "", "Substantive only", "Bank confirmations & reconciliations", "D4"),
        ("Journals & estimates", "Significant", "x", "x", "x", "", "x", "Substantive only", "JE testing; retrospective review of estimates", "D6")]
c0 = r
for i, row in enumerate(prog, 1):
    cell(ws, f"A{r}", i, center=True); cell(ws, f"B{r}", row[0]); cell(ws, f"C{r}", row[1], center=True)
    for j in range(5): cell(ws, f"{chr(68+j)}{r}", row[2+j], is_input=True, center=True)
    cell(ws, f"I{r}", row[7], is_input=True); ws.merge_cells(f"I{r}:J{r}")
    cell(ws, f"K{r}", row[8], is_input=True); cell(ws, f"L{r}", row[9], center=True)
    ws.row_dimensions[r].height = 30; r += 1
dv(ws, f"I{c0}:I{r-1}", "Approach")
put(ws, f"B{r}", "E/O = Existence/Occurrence; C = Completeness; A/V = Accuracy/Valuation; P&D = Presentation & disclosure (ISA 315R.A190).", font(8, i=True, c="616161"))
r += 2

r = section(ws, r, "C2", "Sampling methodology - monetary unit sampling calculator", "ISA 530")
r = table_head(ws, r, ["", "Parameter", "Value", "Note"])
ws.merge_cells(f"D{r-1}:J{r-1}")
s0 = r
samp = [("Population (e.g. trade receivables, MNT '000)", "=INDEX(TB_CY,MATCH(\"Trade receivables\",TB_Acct,0))", NUM, "Linked to D1 lead schedule", False, True),
        ("Tolerable misstatement (= performance materiality)", "=Mat_PM", NUM, "ISA 530.A3", False, True),
        ("Expected misstatement", 5000, NUM, "Based on PY results", True, False),
        ("Desired confidence level", "95%", None, "Higher for significant risks", True, False),
        ("Reliability (confidence) factor", f'=IF(C{s0+3}="95%",3,IF(C{s0+3}="90%",2.31,1.61))', "0.00", "Poisson factors, 0 errors expected", False, False),
        ("Expansion factor", f'=IF(C{s0+3}="95%",1.6,IF(C{s0+3}="90%",1.5,1.3))', "0.00", "Approximation per audit sampling guides", False, False),
        ("SAMPLE SIZE", f"=IFERROR(ROUNDUP(C{s0}*C{s0+4}/(C{s0+1}-C{s0+2}*C{s0+5}),0),0)", "0", "n = Pop x RF / (TM - EM x EF)", False, False),
        ("SAMPLING INTERVAL", f"=IFERROR(C{s0}/C{s0+6},0)", NUM, "Select every Nth MNT; items > interval tested 100%", False, False)]
for i, (lab, v, nf, note, isin, lk) in enumerate(samp):
    rr = r + i
    b = cell(ws, f"B{rr}", lab); c = cell(ws, f"C{rr}", v, nf, isin, lk, center=True)
    if lab.isupper(): b.font = font(9, True); c.font = font(10, True, DK); b.fill = fill(PALE); c.fill = fill(PALE)
    cell(ws, f"D{rr}", note); ws.merge_cells(f"D{rr}:J{rr}")
dv(ws, f"C{s0+3}", "Confidence")
dname("Samp_Size", f"L3_Methodology!$C${s0+6}")
r += len(samp) + 1
put(ws, f"B{r}", "Tests of controls - sample size guide by control frequency (firm methodology)", font(9, True))
r += 1
r = table_head(ws, r, ["", "Control frequency", "Min sample", "Max sample"])
f0 = r
for fq, mn, mx in [("Annual", 1, 1), ("Quarterly", 2, 2), ("Monthly", 2, 5), ("Weekly", 5, 15), ("Daily", 20, 40), ("Multiple daily", 25, 60)]:
    cell(ws, f"B{r}", fq); cell(ws, f"C{r}", mn, center=True); cell(ws, f"D{r}", mx, center=True); r += 1
dname("Freq_Name", f"L3_Methodology!$B${f0}:$B${r-1}"); dname("Freq_Min", f"L3_Methodology!$C${f0}:$C${r-1}")
r += 1

r = section(ws, r, "C3", "Substantive analytical procedures & ratio analysis", "ISA 520.5-7")
r = table_head(ws, r, ["#", "Ratio / relationship", "Current year", "Prior year", "Change", "Expectation", "Threshold", "Difference vs expectation", "Investigate?", "", "Explanation obtained"])
ws.merge_cells(f"I{r-1}:J{r-1}")
def tb(a, col="TB_CY"): return f'INDEX({col},MATCH("{a}",TB_Acct,0))'
ratios = [("Gross margin %", f'=1-{tb("Cost of sales")}/-{tb("Revenue")}', f'=1-{tb("Cost of sales","TB_PY")}/-{tb("Revenue","TB_PY")}', 0.35, 0.02, PCT),
          ("Receivable days", f'={tb("Trade receivables")}/-{tb("Revenue")}*365', f'={tb("Trade receivables","TB_PY")}/-{tb("Revenue","TB_PY")}*365', 65, 7, "0.0"),
          ("Inventory days", f'={tb("Inventory")}/{tb("Cost of sales")}*365', f'={tb("Inventory","TB_PY")}/{tb("Cost of sales","TB_PY")}*365', 120, 10, "0.0"),
          ("Payable days", f'=-{tb("Trade payables")}/{tb("Cost of sales")}*365', f'=-{tb("Trade payables","TB_PY")}/{tb("Cost of sales","TB_PY")}*365', 90, 10, "0.0"),
          ("Current ratio (x)", f'=({tb("Cash")}+{tb("Trade receivables")}+{tb("Inventory")})/-({tb("Trade payables")}+{tb("Accruals")}+{tb("Income tax payable")})',
           f'=({tb("Cash","TB_PY")}+{tb("Trade receivables","TB_PY")}+{tb("Inventory","TB_PY")})/-({tb("Trade payables","TB_PY")}+{tb("Accruals","TB_PY")}+{tb("Income tax payable","TB_PY")})', 2.2, 0.3, "0.00"),
          ("Revenue growth %", f'={tb("Revenue")}/{tb("Revenue","TB_PY")}-1', None, 0.08, 0.03, PCT)]
for i, (nm, cy, py, exp, thr, nf) in enumerate(ratios, 1):
    cell(ws, f"A{r}", i, center=True); cell(ws, f"B{r}", nm)
    cell(ws, f"C{r}", cy, nf, center=True); cell(ws, f"D{r}", py if py else "-", nf, center=True)
    cell(ws, f"E{r}", f"=IFERROR(C{r}-D{r},0)" if py else "-", nf, center=True)
    cell(ws, f"F{r}", exp, nf, True, center=True); cell(ws, f"G{r}", thr, nf, True, center=True)
    cell(ws, f"H{r}", f"=C{r}-F{r}", nf, center=True)
    cell(ws, f"I{r}", f'=IF(ABS(H{r})>G{r},"Investigate","Within threshold")', center=True); ws.merge_cells(f"I{r}:J{r}")
    cell(ws, f"K{r}", "", is_input=True); ws.merge_cells(f"K{r}:L{r}")
    r += 1
r += 1
r = section(ws, r, "C4", "Auditing accounting estimates", "ISA 540 (Revised)")
r = checklist(ws, r, [
    ("Estimates identified: inventory NRV, ECL on receivables, useful lives, accruals", "ISA 540.13", "Yes", ""),
    ("Estimation uncertainty, complexity & subjectivity assessed (spectrum of IR)", "ISA 540.16", "Yes", "NRV = higher IR"),
    ("Retrospective review of PY estimates outcome performed", "ISA 540.14", "Yes", "PY NRV under-provided by 21,000"),
    ("Method, significant assumptions and data tested OR independent point estimate developed", "ISA 540.22-29", "Yes", "Auditor range for NRV: 31,000 - 45,000"),
    ("Indicators of possible management bias evaluated", "ISA 540.32", "Yes", "Consistent optimism in NRV - communicated"),
], "C4")
r = section(ws, r, "C5", "Using the work of internal auditors", "ISA 610 (Revised 2013)")
r = checklist(ws, r, [("Entity has an internal audit function", "ISA 610.13", "No", "No IA function - not applicable"),
                      ("Direct assistance from internal auditors planned", "ISA 610.26", "N/A", "")], "C5")
r = section(ws, r, "C6", "Using the work of an auditor's expert", "ISA 620")
r = checklist(ws, r, [("Need for auditor's expert determined", "ISA 620.7", "Yes", "Firm IT audit specialist for ITGCs"),
                      ("Competence, capabilities & objectivity of expert evaluated", "ISA 620.9", "Yes", "CISA qualified; firm employee"),
                      ("Nature, scope & objectives of expert's work agreed", "ISA 620.11", "Yes", "Instruction memo 1-Oct-25"),
                      ("Adequacy of expert's work evaluated", "ISA 620.12", "Yes", "ITGC report reviewed by manager")], "C6")
r = section(ws, r, "C7", "Group audits & component auditors", "ISA 600 (Revised)")
r = checklist(ws, r, [("Entity prepares group financial statements", "ISA 600.2", "No", "Single entity - not applicable")], "C7")
r = section(ws, r, "C8", "Service organisations", "ISA 402")
r = checklist(ws, r, [("Entity uses service organisations relevant to the audit (e.g. payroll bureau)", "ISA 402.9", "Yes", "Outsourced payroll"),
                      ("Type 1/Type 2 report obtained or alternative procedures performed", "ISA 402.13-17", "Yes", "Substantive testing of payroll performed"),
                      ("Complementary user entity controls identified & tested", "ISA 402.A33", "Yes", "Payroll input approval tested")], "C8")

# ---------------- L4 EXECUTION ----------------
ws = S["L4_Execution"]
r = sheet_header(ws, "AUDIT EXECUTION LAYER", 4,
                 "Perform fieldwork: lead schedules, tests of controls, substantive tests, confirmations, count attendance, JE testing and subsequent events.",
                 "ISA 230; ISA 330; ISA 500; ISA 501; ISA 505; ISA 240.33; ISA 560; ISA 710")
r = section(ws, r, "D1", "Lead schedule / trial balance (MNT '000)", "ISA 500 | ISA 230")
r = table_head(ws, r, ["Code", "Account", "FS caption", "Class", "CY unadjusted", "Corrected adj. (E1)", "CY final", "PY final",
                       "Variance", "Var %", "Explain?", "WP ref"])
tbrows = [("1010", "Cash", "Cash and cash equivalents", "Asset", 1250000, 1305000, "D4"),
          ("1100", "Trade receivables", "Trade and other receivables", "Asset", 3480000, 2950000, "D4"),
          ("1200", "Inventory", "Inventories", "Asset", 4120000, 3640000, "D5"),
          ("1500", "Property, plant & equipment", "PPE", "Asset", 9870000, 9420000, "D3"),
          ("1600", "Intangible assets", "Intangibles", "Asset", 650000, 700000, "D3"),
          ("2010", "Trade payables", "Trade and other payables", "Liability", -2960000, -2610000, "D3"),
          ("2020", "Accruals", "Trade and other payables", "Liability", -840000, -760000, "D3"),
          ("2100", "Borrowings", "Borrowings", "Liability", -4500000, -5000000, "D4"),
          ("2200", "Income tax payable", "Current tax liability", "Liability", -310000, -280000, "D3"),
          ("3000", "Share capital", "Equity", "Equity", -5000000, -5000000, "D3"),
          ("3100", "Retained earnings (opening)", "Equity", "Equity", -4365000, -3290000, "B9"),
          ("4000", "Revenue", "Revenue", "Income", -18600000, -16900000, "D3"),
          ("5000", "Cost of sales", "Cost of sales", "Expense", 12090000, 11000000, "D3"),
          ("6000", "Administrative expenses", "Administrative expenses", "Expense", 2850000, 2700000, "D3"),
          ("6100", "Selling & distribution expenses", "Selling expenses", "Expense", 1420000, 1300000, "D3"),
          ("7000", "Finance costs", "Finance costs", "Expense", 380000, 420000, "D4"),
          ("8000", "Income tax expense", "Income tax", "Tax", 465000, 405000, "D3")]
t0 = r
for code, acct, cap, cls, cy, py, wp in tbrows:
    cell(ws, f"A{r}", code, center=True); cell(ws, f"B{r}", acct); cell(ws, f"C{r}", cap); cell(ws, f"D{r}", cls, center=True)
    cell(ws, f"E{r}", cy, NUM, True)
    cell(ws, f"F{r}", f"=SUMIFS(MS_Amt,MS_Dr,B{r},MS_Corr,\"Yes\")-SUMIFS(MS_Amt,MS_Cr,B{r},MS_Corr,\"Yes\")", NUM, link=True)
    cell(ws, f"G{r}", f"=E{r}+F{r}", NUM); cell(ws, f"H{r}", py, NUM, True)
    cell(ws, f"I{r}", f"=G{r}-H{r}", NUM); cell(ws, f"J{r}", f"=IFERROR(I{r}/ABS(H{r}),0)", PCT, center=True)
    cell(ws, f"K{r}", f'=IF(AND(ABS(I{r})>Mat_PM,ABS(J{r})>0.1),"Explain","-")', center=True)
    cell(ws, f"L{r}", wp, center=True)
    r += 1
t1 = r - 1
ws.conditional_formatting.add(f"K{t0}:K{t1}", CellIsRule(operator="equal", formula=['"Explain"'], fill=fill("FFE0B2"), font=Font(name=F, bold=True, color="E65100")))
dname("TB_Acct", f"L4_Execution!$B${t0}:$B${t1}"); dname("TB_Class", f"L4_Execution!$D${t0}:$D${t1}")
dname("TB_CY", f"L4_Execution!$E${t0}:$E${t1}"); dname("TB_PY", f"L4_Execution!$H${t0}:$H${t1}")
dname("TB_Final", f"L4_Execution!$G${t0}:$G${t1}")
tot = [("Trial balance check (must be zero)", f"=SUM(E{t0}:E{t1})", f"=SUM(G{t0}:G{t1})", f"=SUM(H{t0}:H{t1})"),
       ("Profit before tax", f'=-SUMIFS(E{t0}:E{t1},D{t0}:D{t1},"Income")-SUMIFS(E{t0}:E{t1},D{t0}:D{t1},"Expense")',
        f'=-SUMIFS(G{t0}:G{t1},D{t0}:D{t1},"Income")-SUMIFS(G{t0}:G{t1},D{t0}:D{t1},"Expense")',
        f'=-SUMIFS(H{t0}:H{t1},D{t0}:D{t1},"Income")-SUMIFS(H{t0}:H{t1},D{t0}:D{t1},"Expense")'),
       ("Revenue", f'=-SUMIFS(E{t0}:E{t1},D{t0}:D{t1},"Income")', f'=-SUMIFS(G{t0}:G{t1},D{t0}:D{t1},"Income")', f'=-SUMIFS(H{t0}:H{t1},D{t0}:D{t1},"Income")'),
       ("Total assets", f'=SUMIFS(E{t0}:E{t1},D{t0}:D{t1},"Asset")', f'=SUMIFS(G{t0}:G{t1},D{t0}:D{t1},"Asset")', f'=SUMIFS(H{t0}:H{t1},D{t0}:D{t1},"Asset")'),
       ("Equity (excl. current-year profit)", f'=-SUMIFS(E{t0}:E{t1},D{t0}:D{t1},"Equity")', f'=-SUMIFS(G{t0}:G{t1},D{t0}:D{t1},"Equity")', f'=-SUMIFS(H{t0}:H{t1},D{t0}:D{t1},"Equity")')]
for i, (lab, a, b, c_) in enumerate(tot):
    put(ws, f"B{r}", lab, font(9, True), PALE, bd=BOX); ws.merge_cells(f"B{r}:D{r}")
    for col, v in (("E", a), ("G", b), ("H", c_)): put(ws, f"{col}{r}", v, font(9, True), PALE, NUM, bd=BOX)
    if i == 0:
        put(ws, f"I{r}", f'=IF(AND(E{r}=0,G{r}=0,H{r}=0),"TB balances","TB OUT OF BALANCE")', font(9, True, MD), PALE, bd=BOX)
        ws.merge_cells(f"I{r}:K{r}")
    if i == 1: dname("TB_PBT", f"L4_Execution!$E${r}"); dname("TB_PBT_Final", f"L4_Execution!$G${r}")
    if i == 2: dname("TB_Rev", f"L4_Execution!$E${r}")
    if i == 3: dname("TB_TA", f"L4_Execution!$E${r}")
    if i == 4: dname("TB_Eq", f"L4_Execution!$E${r}")
    r += 1
put(ws, f"B{r}", "Sign convention: debits positive, credits negative. 'Explain' = movement > PM and > 10% - document explanation on supporting WP.", font(8, i=True, c="616161"))
r += 2

r = section(ws, r, "D2", "Tests of operating effectiveness of controls", "ISA 330.8-17")
r = table_head(ws, r, ["ID", "Control description", "Process", "Frequency", "Min sample (C2)", "Sample tested", "Exceptions",
                       "Result", "Rely?", "Tested by", "Reference", ""])
ctrls = [("Sales orders above credit limit approved by credit manager", "Revenue", "Daily", 25, 0),
         ("Monthly bank reconciliations prepared & reviewed", "Cash", "Monthly", 3, 0),
         ("Three-way match (PO/GRN/invoice) before payment", "Purchases", "Multiple daily", 25, 1),
         ("Payroll changes authorised by HR manager", "Payroll", "Monthly", 3, 0),
         ("Quarterly inventory cycle counts & variance review", "Inventory", "Quarterly", 2, 0),
         ("User access reviews for ERP (ITGC)", "IT", "Quarterly", 2, 1),
         ("Program change approval & testing (ITGC)", "IT", "Monthly", 3, 0),
         ("Journal entries above MNT 50m approved by CFO", "Financial close", "Weekly", 10, 0)]
k0 = r
for i, (dsc, proc, fq, n, ex) in enumerate(ctrls, 1):
    cell(ws, f"A{r}", f"C{i:02d}", center=True); cell(ws, f"B{r}", dsc); cell(ws, f"C{r}", proc, center=True)
    cell(ws, f"D{r}", fq, is_input=True, center=True)
    cell(ws, f"E{r}", f"=IFERROR(INDEX(Freq_Min,MATCH(D{r},Freq_Name,0)),\"\")", link=True, center=True)
    cell(ws, f"F{r}", n, is_input=True, center=True); cell(ws, f"G{r}", ex, is_input=True, center=True)
    cell(ws, f"H{r}", f'=IF(F{r}<E{r},"Sample too small",IF(G{r}=0,"Effective","Deviation - evaluate (ISA 330.17)"))', center=True)
    cell(ws, f"I{r}", f'=IF(H{r}="Effective","Yes","No")', center=True)
    cell(ws, f"J{r}", "T. Enkhjin" if i % 2 else "M. Saruul", center=True); cell(ws, f"K{r}", f"D2.{i}", center=True)
    ws.row_dimensions[r].height = 28; r += 1
dv(ws, f"D{k0}:D{r-1}", "Freq")
dname("Ctl_Result", f"L4_Execution!$H${k0}:$H${r-1}")
ws.conditional_formatting.add(f"H{k0}:H{r-1}", FormulaRule(formula=[f'LEFT(H{k0},9)="Deviation"'], fill=fill("FFCDD2")))
r += 1

r = section(ws, r, "D3", "Substantive tests of details - summary log", "ISA 330.18-23 | ISA 500 | ISA 530.14")
r = table_head(ws, r, ["#", "Procedure", "FS area", "Population (MNT '000)", "Value tested", "Coverage %", "Exceptions found",
                       "Projected misstatement", "vs CT", "Result", "Risk ref", "Prepared by"])
subs = [("Revenue cut-off: invoices 10 days pre/post YE vs dispatch notes", "Revenue", 1850000, 620000, 45000, "R01"),
        ("Revenue occurrence: vouch sample to orders, GDNs & cash", "Revenue", 18600000, 2400000, 0, "R01"),
        ("Inventory NRV: post-YE selling prices vs cost", "Inventory", 4120000, 1900000, 38000, "R03"),
        ("Search for unrecorded liabilities (post-YE payments)", "Payables & accruals", 960000, 610000, 22000, "R06"),
        ("PPE additions vouched to invoices & title", "PPE", 1150000, 980000, 0, "R07"),
        ("Receivables ECL matrix recompute & aged balances", "Trade receivables", 3480000, 1300000, 2240, "R05"),
        ("Expense classification review (COGS vs admin)", "Expenses", 2850000, 700000, 15000, "R02")]
u0 = r
for i, (p, a, pop, tv, ex, rk) in enumerate(subs, 1):
    cell(ws, f"A{r}", i, center=True); cell(ws, f"B{r}", p); cell(ws, f"C{r}", a, center=True)
    cell(ws, f"D{r}", pop, NUM, True); cell(ws, f"E{r}", tv, NUM, True)
    cell(ws, f"F{r}", f"=IFERROR(E{r}/D{r},0)", PCT, center=True); cell(ws, f"G{r}", ex, NUM, True)
    cell(ws, f"H{r}", f"=IF(F{r}>=0.95,G{r},IFERROR(G{r}/E{r}*D{r},0))", NUM)
    cell(ws, f"I{r}", f'=IF(H{r}>Mat_CT,"Accumulate (E1)","Trivial")', center=True)
    cell(ws, f"J{r}", f'=IF(H{r}=0,"No exceptions",IF(H{r}>Mat_PM,"Exceeds PM - extend testing","Misstatement noted"))', center=True)
    cell(ws, f"K{r}", rk, center=True); cell(ws, f"L{r}", "T. Enkhjin", center=True)
    ws.row_dimensions[r].height = 30; r += 1
put(ws, f"B{r}", "Projection: exception / value tested x population, unless coverage >= 95% (ISA 530.14). Cut-off & targeted tests are not projected where 100% of the relevant population is examined.", font(8, i=True, c="616161"), al=WRAP)
ws.merge_cells(f"B{r}:L{r}"); r += 2

r = section(ws, r, "D4", "External confirmations", "ISA 505")
r = table_head(ws, r, ["#", "Confirming party", "Type", "Balance per books", "Confirmed amount", "Difference", "Status",
                       "", "Reconciled / resolved?", "Date sent", "Date received", "Ref"])
ws.merge_cells(f"G{r-1}:H{r-1}")
confs = [("Golomt Bank - current & loan accounts", "Bank", 5200000, 5200000, "Received - agreed", "Yes", D(2026, 1, 5), D(2026, 1, 19)),
         ("Khan Bank - current account", "Bank", 550000, 550000, "Received - agreed", "Yes", D(2026, 1, 5), D(2026, 1, 21)),
         ("Customer: Nomin Retail", "Receivable", 820000, 775000, "Received - reconciled", "Yes", D(2026, 1, 8), D(2026, 1, 27)),
         ("Customer: Orgil Market", "Receivable", 610000, 610000, "Received - agreed", "Yes", D(2026, 1, 8), D(2026, 1, 25)),
         ("Customer: Sansar Stores", "Receivable", 430000, None, "Not received - alt. procedures", "Yes", D(2026, 1, 8), None),
         ("Supplier: Asia Import LLC", "Payable", 690000, 712000, "Received - reconciled", "Yes", D(2026, 1, 8), D(2026, 2, 2)),
         ("Legal counsel - litigation", "Legal", 0, 0, "Outstanding", "No", D(2026, 3, 2), None)]
q0 = r
for i, (pty, typ, bk, cf, stt, rec, ds, dr) in enumerate(confs, 1):
    cell(ws, f"A{r}", i, center=True); cell(ws, f"B{r}", pty); cell(ws, f"C{r}", typ, center=True)
    cell(ws, f"D{r}", bk, NUM, True); cell(ws, f"E{r}", cf, NUM, True)
    cell(ws, f"F{r}", f'=IF(E{r}="","n/a",E{r}-D{r})', NUM)
    cell(ws, f"G{r}", stt, is_input=True, center=True); ws.merge_cells(f"G{r}:H{r}")
    cell(ws, f"I{r}", rec, is_input=True, center=True); cell(ws, f"J{r}", ds, DATE, True, center=True)
    cell(ws, f"K{r}", dr, DATE, True, center=True); cell(ws, f"L{r}", f"D4.{i}", center=True)
    r += 1
dv(ws, f"G{q0}:G{r-1}", "ConfStatus"); dv(ws, f"I{q0}:I{r-1}", "YesNo")
dname("Conf_Status", f"L4_Execution!$G${q0}:$G${r-1}")
r += 1
r = section(ws, r, "D5", "Attendance at physical inventory count", "ISA 501.4-8")
r = checklist(ws, r, [
    ("Management's count instructions evaluated", "ISA 501.4(a)(i)", "Yes", "Instructions adequate; blind counts"),
    ("Performance of management's count procedures observed", "ISA 501.4(a)(ii)", "Yes", "Warehouse + Branch 1 & 2 on 31-Dec-25"),
    ("Inventory inspected (condition, obsolescence)", "ISA 501.4(a)(iii)", "Yes", "Damaged goods noted - fed into NRV test"),
    ("Test counts performed floor-to-sheet and sheet-to-floor", "ISA 501.4(a)(iv)", "Yes", "60 items, 1 minor difference corrected"),
    ("Final inventory records reconcile to count results; cut-off documents captured", "ISA 501.4(b)", "Yes", "Last GRN/GDN numbers recorded"),
], "D5")
r = section(ws, r, "D6", "Journal entry testing (management override)", "ISA 240.33(a) | ISA 330.20")
r = checklist(ws, r, [
    ("Completeness of JE population reconciled to TB", "ISA 240.A43", "Yes", "48,320 JEs; agreed to TB movement"),
    ("Inquiries of individuals about inappropriate/unusual activity in JE process", "ISA 240.33(a)(i)", "Yes", ""),
    ("JEs & adjustments at period end selected & tested", "ISA 240.33(a)(ii)", "Yes", "All post-closing entries tested"),
    ("Risk criteria applied: weekend/holiday, round amounts, unusual users, seldom-used accounts, post-close", "ISA 240.A44-A45", "Yes", "112 entries selected via data analytics"),
    ("Significant transactions outside normal course of business evaluated for business rationale", "ISA 240.33(c)", "Yes", "None noted"),
], "D6")
r = section(ws, r, "D7", "Litigation & claims", "ISA 501.9-12")
r = checklist(ws, r, [("Inquiry of management & in-house legal; review minutes and legal expense ledger", "ISA 501.9", "Yes", "One customer claim noted"),
                      ("Direct communication with external legal counsel", "ISA 501.10", None, "Letter sent 2-Mar-26 - outstanding")], "D7")
r = section(ws, r, "D8", "Subsequent events review", "ISA 560.6-9")
r = checklist(ws, r, [("Procedures cover period from FYE to date of auditor's report", "ISA 560.6", None, "To be updated to 31-Mar-26"),
                      ("Management inquiry; minutes; latest interim accounts reviewed", "ISA 560.7", "Yes", "Jan-Feb 26 management accounts"),
                      ("Adjusting/non-adjusting events appropriately reflected (IAS 10)", "ISA 560.8", None, ""),
                      ("Written representation on subsequent events obtained", "ISA 560.9", None, "F2")], "D8")
r = section(ws, r, "D9", "Comparative information", "ISA 710")
r = checklist(ws, r, [("Comparatives agree to PY FS / restated appropriately", "ISA 710.7(a)", "Yes", ""),
                      ("Accounting policies for comparatives consistent", "ISA 710.7(b)", "Yes", "")], "D9")

# ---------------- L5 FINDINGS ----------------
ws = S["L5_Findings"]
r = sheet_header(ws, "FINDINGS & REMEDIATION LAYER", 5,
                 "Accumulate and evaluate misstatements, communicate control deficiencies, analyse root causes and track remediation to closure.",
                 "ISA 450; ISA 265; ISA 230; ISQM 1 para 42-45 (remediation)")
r = section(ws, r, "E1", "Summary of identified misstatements (SUM/SAD) - MNT '000", "ISA 450.5-12")
r = table_head(ws, r, ["Ref", "Description", "Type", "Debit account", "Credit account", "Amount", "Profit effect (+/-)",
                       "Corrected by mgmt?", "Source WP", "Above CT?", "Root cause", ""])
ms = [("Revenue recorded for goods dispatched after year end (cut-off)", "Factual", "Revenue", "Trade receivables", 45000, "No", "D3.1", "Period-end pressure to meet targets"),
      ("Inventory NRV write-down understated - slow-moving imports", "Judgmental", "Cost of sales", "Inventory", 38000, "No", "D3.3", "No formal NRV policy"),
      ("Unrecorded utility & consulting accruals", "Factual", "Administrative expenses", "Accruals", 22000, "Yes", "D3.4", "Incomplete accrual checklist"),
      ("Depreciation of warehouse misclassified to admin", "Factual", "Cost of sales", "Administrative expenses", 15000, "No", "D3.7", "Cost centre mapping error"),
      ("Projected ECL shortfall from receivables sample", "Projected", "Administrative expenses", "Trade receivables", 6000, "No", "D3.6", "ECL matrix not updated")]
e0 = r
for i, (dsc, typ, dr, cr, amt, cor, src, rc) in enumerate(ms, 1):
    cell(ws, f"A{r}", f"M{i:02d}", center=True); cell(ws, f"B{r}", dsc, is_input=True)
    cell(ws, f"C{r}", typ, is_input=True, center=True); cell(ws, f"D{r}", dr, is_input=True); cell(ws, f"E{r}", cr, is_input=True)
    cell(ws, f"F{r}", amt, NUM, True)
    cell(ws, f"G{r}", f'=-F{r}*((COUNTIFS(TB_Acct,D{r},TB_Class,"Income")+COUNTIFS(TB_Acct,D{r},TB_Class,"Expense"))>0)'
                      f'+F{r}*((COUNTIFS(TB_Acct,E{r},TB_Class,"Income")+COUNTIFS(TB_Acct,E{r},TB_Class,"Expense"))>0)', NUM)
    cell(ws, f"H{r}", cor, is_input=True, center=True); cell(ws, f"I{r}", src, center=True)
    cell(ws, f"J{r}", f'=IF(F{r}>Mat_CT,"Yes","No - trivial")', center=True)
    cell(ws, f"K{r}", rc, is_input=True); ws.merge_cells(f"K{r}:L{r}")
    ws.row_dimensions[r].height = 30; r += 1
# spare rows for additional misstatements
for i in range(3):
    cell(ws, f"A{r}", f"M{len(ms)+i+1:02d}", center=True)
    for col in "BCDEFH": cell(ws, f"{col}{r}", None, NUM if col == "F" else None, True)
    cell(ws, f"G{r}", f'=IF(F{r}="",0,-F{r}*((COUNTIFS(TB_Acct,D{r},TB_Class,"Income")+COUNTIFS(TB_Acct,D{r},TB_Class,"Expense"))>0)'
                      f'+F{r}*((COUNTIFS(TB_Acct,E{r},TB_Class,"Income")+COUNTIFS(TB_Acct,E{r},TB_Class,"Expense"))>0))', NUM)
    cell(ws, f"I{r}", ""); cell(ws, f"J{r}", ""); cell(ws, f"K{r}", None, is_input=True); ws.merge_cells(f"K{r}:L{r}")
    r += 1
e1 = r - 1
dv(ws, f"C{e0}:C{e1}", "MSType"); dv(ws, f"H{e0}:H{e1}", "YesNo")
acct_dv = DataValidation(type="list", formula1="=TB_Acct", allow_blank=True); ws.add_data_validation(acct_dv); acct_dv.add(f"D{e0}:E{e1}")
dname("MS_Amt", f"L5_Findings!$F${e0}:$F${e1}"); dname("MS_Dr", f"L5_Findings!$D${e0}:$D${e1}")
dname("MS_Cr", f"L5_Findings!$E${e0}:$E${e1}"); dname("MS_Corr", f"L5_Findings!$H${e0}:$H${e1}")
dname("MS_Profit", f"L5_Findings!$G${e0}:$G${e1}")
summ = [("Total corrected misstatements - profit effect", f'=SUMIFS(G{e0}:G{e1},H{e0}:H{e1},"Yes")', NUM),
        ("Total UNCORRECTED misstatements - profit effect", f'=SUMIFS(G{e0}:G{e1},H{e0}:H{e1},"No")', NUM),
        ("Uncorrected as % of overall materiality", None, PCT),
        ("Overall materiality (B3)", "=Mat_OM", NUM), ("Performance materiality (B3)", "=Mat_PM", NUM),
        ("Uncorrected classification misstatements (no profit effect)", f'=SUMIFS(F{e0}:F{e1},H{e0}:H{e1},"No",G{e0}:G{e1},0)', NUM),
        ("EVALUATION (ISA 450.11)", None, None)]
s0 = r + 1
for i, (lab, v, nf) in enumerate(summ):
    rr = s0 + i
    if lab.startswith("Uncorrected as"): v = f"=IFERROR(ABS(C{s0+1})/Mat_OM,0)"
    if lab.startswith("EVALUATION"):
        v = (f'=IF(ABS(C{s0+1})>=Mat_OM,"MATERIAL - request correction; otherwise modify opinion (ISA 705)",'
             f'IF(ABS(C{s0+1})>=Mat_PM,"Below OM but above PM - request correction; reassess risk of undetected misstatement (ISA 450.6)",'
             f'"Not material individually or in aggregate"))')
    put(ws, f"B{rr}", lab, font(9, True), PALE, bd=BOX)
    c = put(ws, f"C{rr}", v, font(9, True, "008000" if str(v).startswith("=Mat") else DK), PALE, nf, bd=BOX)
    if lab.startswith("EVALUATION"): ws.merge_cells(f"C{rr}:L{rr}")
dname("MS_Uncorr", f"L5_Findings!$C${s0+1}"); dname("MS_UncorrPct", f"L5_Findings!$C${s0+2}")
dname("MS_Eval", f"L5_Findings!$C${s0+6}")
r = s0 + len(summ) + 1

r = section(ws, r, "E2", "Control deficiencies evaluation", "ISA 265.7-11")
r = table_head(ws, r, ["Ref", "Deficiency", "Process", "Source", "Could result in material misstatement?", "Compensating control?",
                       "Classification", "Communicate to", "Form", "Date communicated", "Mgmt response", ""])
defs = [("No formal entity risk assessment process", "Entity level", "B2", "No", "No"),
        ("ERP user access reviews not performed in Q2 & Q3 (ITGC)", "IT", "D2 C06", "Yes", "No"),
        ("Three-way match override without approval (1 of 25)", "Purchases", "D2 C03", "No", "Yes"),
        ("No documented NRV / slow-moving inventory policy", "Inventory", "E1 M02", "Yes", "No"),
        ("Accrual checklist not updated for new contracts", "Financial close", "E1 M03", "No", "Yes")]
g0 = r
for i, (dsc, proc, src, mm, comp) in enumerate(defs, 1):
    cell(ws, f"A{r}", f"CD{i:02d}", center=True); cell(ws, f"B{r}", dsc, is_input=True); cell(ws, f"C{r}", proc, center=True)
    cell(ws, f"D{r}", src, center=True); cell(ws, f"E{r}", mm, is_input=True, center=True); cell(ws, f"F{r}", comp, is_input=True, center=True)
    cell(ws, f"G{r}", f'=IF(AND(E{r}="Yes",F{r}="No"),"Significant deficiency","Other deficiency")', center=True)
    cell(ws, f"H{r}", f'=IF(G{r}="Significant deficiency","TCWG & management","Management")', center=True)
    cell(ws, f"I{r}", f'=IF(G{r}="Significant deficiency","Written (ISA 265.9)","Written or oral (265.10)")', center=True)
    cell(ws, f"J{r}", D(2026, 3, 20) if i < 4 else None, DATE, True, center=True)
    cell(ws, f"K{r}", "Agreed" if i < 4 else "", is_input=True); ws.merge_cells(f"K{r}:L{r}")
    ws.row_dimensions[r].height = 30; r += 1
dv(ws, f"E{g0}:F{r-1}", "YesNo")
dname("CD_Class", f"L5_Findings!$G${g0}:$G${r-1}")
ws.conditional_formatting.add(f"G{g0}:G{r-1}", CellIsRule(operator="equal", formula=['"Significant deficiency"'], fill=fill("FFCDD2"), font=Font(name=F, bold=True, color="B71C1C")))
r += 1

r = section(ws, r, "E3", "Issue log & remediation tracker", "ISA 230.8-11 | ISQM 1.42")
r = table_head(ws, r, ["ID", "Issue identified", "Layer", "Category", "Severity", "Root cause", "Recommendation", "Owner",
                       "Date raised", "Due date", "Status", "Days open"])
issues = [("Revenue cut-off error (M01)", "Execution", "Misstatement", "High", "Sales staff incentivised on dispatch date", "Align invoicing to GDN; system cut-off lock", "CFO", D(2026, 2, 16), D(2026, 3, 25), "Open"),
          ("ERP access reviews missed (CD02)", "Execution", "Control deficiency", "High", "IT manager vacancy", "Quarterly access review with sign-off", "IT Manager", D(2025, 12, 5), D(2026, 3, 31), "In progress"),
          ("No NRV policy (CD04)", "Findings", "Control deficiency", "High", "Policy gap", "Adopt NRV/slow-moving policy & ageing report", "CFO", D(2026, 2, 20), D(2026, 4, 30), "Open"),
          ("Legal letter outstanding (D7)", "Execution", "Evidence gap", "Medium", "Counsel response delayed", "Follow up; consider ISA 501.11 limitation", "Manager", D(2026, 3, 2), D(2026, 3, 27), "Open"),
          ("Independence declaration outstanding - N. Temuulen", "Strategy", "Compliance", "Medium", "New joiner onboarding", "Obtain before further work; firm system reminder", "Manager", D(2026, 2, 10), D(2026, 3, 13), "Open"),
          ("Expense misclassification (M04)", "Execution", "Misstatement", "Low", "Cost centre mapping", "Update mapping table", "Chief Accountant", D(2026, 2, 25), D(2026, 3, 31), "In progress"),
          ("Accruals completeness (M03)", "Execution", "Misstatement", "Medium", "Checklist out of date", "Monthly contract review", "Chief Accountant", D(2026, 2, 18), D(2026, 3, 6), "Closed"),
          ("Informal risk assessment (CD01)", "Planning", "Control deficiency", "Low", "Owner-managed culture", "Annual risk workshop by Board", "CEO", D(2025, 11, 10), D(2026, 6, 30), "In progress"),
          ("Planning materiality documentation incomplete", "Planning", "Documentation", "Low", "Template not updated", "Use current methodology template", "Senior", D(2025, 11, 12), D(2025, 11, 20), "Closed"),
          ("Subsequent events update to report date", "Oversight", "Evidence gap", "Medium", "Timing", "Update procedures to 31-Mar-26", "Senior", D(2026, 3, 16), D(2026, 3, 31), "Open"),
          ("Sample size short for C05 cycle counts", "Methodology", "Documentation", "Low", "Template error", "Re-perform with 2 quarters", "Senior", D(2026, 1, 15), D(2026, 1, 30), "Closed")]
i0 = r
for i, row in enumerate(issues, 1):
    cell(ws, f"A{r}", f"I{i:02d}", center=True)
    for j, v in enumerate(row):
        col = chr(66 + j)
        cell(ws, f"{col}{r}", v, DATE if j in (7, 8) else None, True, center=j in (1, 2, 3, 6, 7, 8, 9))
    cell(ws, f"L{r}", f'=IF(K{r}="Closed","-",Set_AsOf-I{r})', "0", center=True)
    ws.row_dimensions[r].height = 30; r += 1
dv(ws, f"C{i0}:C{r-1}", "Layer"); dv(ws, f"E{i0}:E{r-1}", "Severity"); dv(ws, f"K{i0}:K{r-1}", "IssueStatus")
dname("Iss_Sev", f"L5_Findings!$E${i0}:$E${r-1}"); dname("Iss_Status", f"L5_Findings!$K${i0}:$K${r-1}")
dname("Iss_Due", f"L5_Findings!$J${i0}:$J${r-1}")
ws.conditional_formatting.add(f"E{i0}:E{r-1}", CellIsRule(operator="equal", formula=['"High"'], fill=fill("FFCDD2")))
ws.conditional_formatting.add(f"E{i0}:E{r-1}", CellIsRule(operator="equal", formula=['"Medium"'], fill=fill("FFE0B2")))
ws.conditional_formatting.add(f"K{i0}:K{r-1}", CellIsRule(operator="equal", formula=['"Closed"'], fill=fill("C8E6C9")))
ws.conditional_formatting.add(f"J{i0}:J{r-1}", FormulaRule(formula=[f'AND(K{i0}<>"Closed",J{i0}<Set_AsOf)'], font=Font(name=F, bold=True, color="C62828")))
put(ws, f"B{r}", "Red due date = overdue vs dashboard as-of date. Closure must be validated by the reviewer (ISA 220R.30-31) before status 'Closed'.", font(8, i=True, c="616161"))
r += 2
r = section(ws, r, "E4", "Management letter (communication of deficiencies)", "ISA 265.9-11")
r = checklist(ws, r, [
    ("Significant deficiencies communicated in writing to TCWG on a timely basis", "ISA 265.9", "Yes", "Letter dated 20-Mar-26"),
    ("Other deficiencies of sufficient importance communicated to management", "ISA 265.10(b)", "Yes", ""),
    ("Letter includes description, potential effects & explanation that audit not designed to opine on IC", "ISA 265.11", "Yes", ""),
    ("Management responses & remediation plans obtained and logged in E3", "ISQM 1.42", None, "2 responses pending"),
], "E4")

# ---------------- L6 OVERSIGHT ----------------
ws = S["L6_Oversight"]
r = sheet_header(ws, "AUDIT OVERSIGHT & COMPLETION LAYER", 6,
                 "Communicate with TCWG, obtain representations, complete engagement quality review, form the opinion and monitor quality (ISQM 1).",
                 "ISA 260; ISA 580; ISA 570; ISA 701; ISA 700/705/706; ISA 720; ISA 220 (Revised); ISQM 1; ISQM 2")
r = section(ws, r, "F1", "Communications with those charged with governance", "ISA 260 (Revised) | ISA 265")
r = table_head(ws, r, ["#", "Matter required to be communicated", "Standard ref", "Required?", "Form", "Date", "Status", "Ref"])
comms = [("Auditor's responsibilities in relation to the FS audit", "ISA 260.14", "Yes", "Written", D(2025, 10, 20), "Done"),
         ("Planned scope & timing incl. significant risks", "ISA 260.15", "Yes", "Written", D(2025, 10, 20), "Done"),
         ("Significant qualitative aspects of accounting practices", "ISA 260.16(a)", "Yes", "Meeting", None, "Pending"),
         ("Significant difficulties encountered during the audit", "ISA 260.16(b)", "Yes", "Meeting", None, "Pending"),
         ("Uncorrected misstatements & request for correction", "ISA 450.12-13", "Yes", "Written", None, "Pending"),
         ("Significant deficiencies in internal control", "ISA 265.9", "Yes", "Written", D(2026, 3, 20), "Done"),
         ("Statement on independence & threats/safeguards", "ISA 260.17; IESBA 400", "Yes", "Written", None, "Pending"),
         ("Fraud or suspected fraud identified", "ISA 240.42", "N/A", "", None, "N/A"),
         ("Non-compliance with laws & regulations", "ISA 250.23", "N/A", "", None, "N/A"),
         ("Expected modifications / EOM in auditor's report", "ISA 705.30; 706.12", "Yes", "Meeting", None, "Pending")]
h0 = r
for i, (m, ref, req, frm, dt_, stt) in enumerate(comms, 1):
    cell(ws, f"A{r}", i, center=True); cell(ws, f"B{r}", m); cell(ws, f"C{r}", ref)
    cell(ws, f"D{r}", req, is_input=True, center=True); cell(ws, f"E{r}", frm, is_input=True, center=True)
    cell(ws, f"F{r}", dt_, DATE, True, center=True); cell(ws, f"G{r}", stt, is_input=True, center=True); cell(ws, f"H{r}", f"F1.{i}", center=True)
    r += 1
dv(ws, f"D{h0}:D{r-1}", "YesNo"); dv(ws, f"G{h0}:G{r-1}", "Done")
dname("Comm_Status", f"L6_Oversight!$G${h0}:$G${r-1}")
ws.conditional_formatting.add(f"G{h0}:G{r-1}", CellIsRule(operator="equal", formula=['"Pending"'], fill=fill("FFE0B2")))
ws.conditional_formatting.add(f"G{h0}:G{r-1}", CellIsRule(operator="equal", formula=['"Done"'], fill=fill("C8E6C9")))
r += 1
r = section(ws, r, "F2", "Written representations", "ISA 580")
r = checklist(ws, r, [
    ("Management responsibility for preparation of FS per framework", "ISA 580.10", None, "Draft letter sent 18-Mar-26"),
    ("All relevant information & access provided; all transactions recorded", "ISA 580.11", None, ""),
    ("Uncorrected misstatements immaterial - schedule attached", "ISA 450.14", None, ""),
    ("Fraud, laws & regulations, related parties, subsequent events, going concern", "ISA 240.40; 250.17; 550.26; 560.9; 570.16(e)", None, ""),
    ("Dated as near as practicable to, but not after, the auditor's report date", "ISA 580.14", None, ""),
], "F2")
r = section(ws, r, "F3", "Going concern - final conclusion", "ISA 570 (Revised).16-24")
r = checklist(ws, r, [
    ("Management's assessment covers at least 12 months from FS date", "ISA 570.13", "Yes", ""),
    ("Material uncertainty related to going concern exists", "ISA 570.18", "No", "No events/conditions identified"),
    ("Disclosures adequate (incl. 'close call' disclosures)", "ISA 570.19-20", "Yes", ""),
], "F3")
r = section(ws, r, "F4", "Key audit matters (if communicated)", "ISA 701")
r = table_head(ws, r, ["#", "Matter", "Why a KAM (significant attention)", "", "", "How addressed in the audit", "", "", "", "Include?"])
ws.merge_cells(f"C{r-1}:E{r-1}"); ws.merge_cells(f"F{r-1}:I{r-1}")
for i, (m, why, how, inc) in enumerate([
        ("Valuation of inventories", "Significant judgement in NRV; MNT depreciation; PY adjustment", "Ageing analysis; post-YE prices; independent range; bias review", "Yes"),
        ("Revenue recognition - cut-off", "Presumed fraud risk; cut-off error identified", "Extended cut-off; credit note review; analytics", "Yes")], 1):
    cell(ws, f"A{r}", i, center=True); cell(ws, f"B{r}", m, is_input=True)
    cell(ws, f"C{r}", why, is_input=True); ws.merge_cells(f"C{r}:E{r}")
    cell(ws, f"F{r}", how, is_input=True); ws.merge_cells(f"F{r}:I{r}")
    cell(ws, f"J{r}", inc, is_input=True, center=True); ws.row_dimensions[r].height = 30; r += 1
put(ws, f"B{r}", "KAMs mandatory for listed entities (ISA 701.5); voluntary for this non-listed client per engagement letter.", font(8, i=True, c="616161")); r += 2

r = section(ws, r, "F5", "Engagement quality review", "ISQM 2 | ISA 220R.36")
r = checklist(ws, r, [
    ("EQ reviewer appointed meets eligibility (competence, authority, objectivity)", "ISQM 2.18", "Yes", "D. Ganbold, partner"),
    ("Cooling-off: reviewer not EP of this engagement in prior 2 years", "ISQM 2.19 (IESBA 325)", "Yes", ""),
    ("Significant judgements & conclusions reviewed (risks, materiality, misstatements, KAMs)", "ISQM 2.25", None, "Scheduled 25-Mar-26"),
    ("Independence evaluation by engagement team reviewed", "ISQM 2.25(d)", None, ""),
    ("Draft FS and auditor's report reviewed", "ISQM 2.25(e)", None, ""),
    ("EQ review completed BEFORE the report is dated", "ISA 220R.36(c); ISQM 2.27", None, ""),
], "F5")
r = section(ws, r, "F6", "Completion checklist & partner sign-off", "ISA 220R.37-40 | ISA 230")
r = checklist(ws, r, [
    ("All working papers prepared and reviewed (see WP_Index)", "ISA 220R.30-31; ISA 230.9", None, ""),
    ("Engagement partner sufficiently & appropriately involved throughout", "ISA 220R.40(a)", "Yes", ""),
    ("Consultation on difficult/contentious matters completed and implemented", "ISA 220R.35", "Yes", "Inventory NRV consulted with technical dept"),
    ("Differences of opinion resolved", "ISA 220R.37", "N/A", "None"),
    ("Final analytical review performed: FS consistent with understanding", "ISA 520.6", None, ""),
    ("Disclosure checklist (IFRS) completed", "ISA 700.13", None, ""),
], "F6")

r = section(ws, r, "F7", "Forming the opinion", "ISA 700 (Revised) | 705 (Revised) | 706 (Revised)")
r = table_head(ws, r, ["", "Decision factor", "Answer", "Basis"])
ws.merge_cells(f"D{r-1}:J{r-1}")
o0 = r
opin = [("Uncorrected misstatements material? (from E1)", '=IF(ABS(MS_Uncorr)>=Mat_OM,"Yes","No")', "ISA 450.11 / 705.6(a)", False),
        ("Inability to obtain sufficient appropriate evidence (scope limitation)?", "No", "ISA 705.6(b)", True),
        ("If modified - effects pervasive?", "No", "ISA 705.A5", True),
        ("Material uncertainty related to going concern? (F3)", "No", "ISA 570.22", True),
        ("Emphasis of matter required?", "No", "ISA 706.8", True),
        ("PROPOSED OPINION", None, "ISA 705 decision matrix", False),
        ("Additional report section", None, "", False)]
for i, (lab, v, basis, isin) in enumerate(opin):
    rr = r + i
    if lab == "PROPOSED OPINION":
        v = (f'=IF(AND(C{o0+1}="Yes",C{o0+2}="Yes"),"Disclaimer of opinion",IF(AND(C{o0}="Yes",C{o0+2}="Yes"),"Adverse opinion",'
             f'IF(OR(C{o0}="Yes",C{o0+1}="Yes"),"Qualified opinion","Unmodified opinion")))')
    if lab.startswith("Additional"):
        v = (f'=IF(C{o0+3}="Yes","Material Uncertainty Related to Going Concern section",IF(C{o0+4}="Yes","Emphasis of Matter paragraph","None"))')
    b = cell(ws, f"B{rr}", lab); c = cell(ws, f"C{rr}", v, None, isin, link=(i == 0), center=True)
    if lab.isupper(): b.font = font(9, True); c.font = font(11, True, DK); b.fill = fill(PALE); c.fill = fill(PALE)
    cell(ws, f"D{rr}", basis); ws.merge_cells(f"D{rr}:J{rr}")
for k in range(1, 5): dv(ws, f"C{o0+k}", "YesNo")
dname("Opinion", f"L6_Oversight!$C${o0+5}")
r += len(opin) + 1
r = section(ws, r, "F8", "Other information", "ISA 720 (Revised)")
r = checklist(ws, r, [("Other information (annual report) identified & obtained", "ISA 720.13", "Yes", "Directors' report"),
                      ("Read & considered for material inconsistency with FS / auditor's knowledge", "ISA 720.14-15", None, "")], "F8")
r = section(ws, r, "F9", "Assembly of final audit file", "ISA 230.14-16 | ISQM 1.A83")
r = table_head(ws, r, ["", "Item", "Value"])
put(ws, f"B{r}", "Audit report date", font(9), bd=BOX); put(ws, f"C{r}", "=Set_RptDate", font(9, c="008000"), nf=DATE, bd=BOX)
put(ws, f"B{r+1}", "Assembly deadline (report date + 60 days)", font(9), bd=BOX); put(ws, f"C{r+1}", "=Set_AssemblyDate", font(9, c="008000"), nf=DATE, bd=BOX)
put(ws, f"B{r+2}", "Actual assembly date", font(9), bd=BOX); inp(put(ws, f"C{r+2}", None, font(9, c="0000FF"), nf=DATE, bd=BOX))
put(ws, f"B{r+3}", "Status", font(9, True), PALE, bd=BOX)
put(ws, f"C{r+3}", f'=IF(C{r+2}="","Pending",IF(C{r+2}<=C{r+1},"Assembled on time","LATE - report to quality management"))', font(9, True, DK), PALE, bd=BOX)
ws.merge_cells(f"C{r+3}:F{r+3}")
put(ws, f"B{r+4}", "After assembly no documentation may be deleted or discarded before the end of the retention period (ISA 230.15) - min. 5 years (ISQM 1.A85).", font(8, i=True, c="616161"))
r += 6
r = section(ws, r, "F10", "ISQM 1 monitoring - engagement quality indicators", "ISQM 1.35-47")
r = table_head(ws, r, ["#", "Quality indicator", "Target", "Actual", "RAG", "Source"])
qa = [("Working papers reviewed before report date", 1.0, "=DD_PctComplete", PCT, "WP_Index"),
      ("Independence declarations received", 1.0, "=DD_IndPct", PCT, "A3"),
      ("TCWG communications completed", 1.0, "=DD_CommPct", PCT, "F1"),
      ("Controls tested & effective", 0.8, "=DD_CtlPct", PCT, "D2"),
      ("Confirmations received (excl. outstanding)", 0.85, "=DD_ConfPct", PCT, "D4"),
      ("Partner hours as % of total actual hours", 0.1, 0.09, PCT, "Time system"),
      ("High-severity issues open", 0, "=DD_OpenHigh", "0", "E3"),
      ("Actual hours vs budget", 1.0, "=DD_HrsPct", PCT, "WP_Index")]
a0 = r
for i, (ind, tgt, act, nf, srcx) in enumerate(qa, 1):
    cell(ws, f"A{r}", i, center=True); cell(ws, f"B{r}", ind); cell(ws, f"C{r}", tgt, nf, True, center=True)
    cell(ws, f"D{r}", act, nf, isinstance(act, float), link=str(act).startswith("="), center=True)
    if ind.startswith("High-severity"):
        cell(ws, f"E{r}", f'=IF(D{r}<=C{r},"Green",IF(D{r}<=2,"Amber","Red"))', center=True)
    elif ind.startswith("Actual hours"):
        cell(ws, f"E{r}", f'=IF(D{r}<=1,"Green",IF(D{r}<=1.1,"Amber","Red"))', center=True)
    else:
        cell(ws, f"E{r}", f'=IF(D{r}>=C{r},"Green",IF(D{r}>=C{r}*0.8,"Amber","Red"))', center=True)
    cell(ws, f"F{r}", srcx, center=True); r += 1
for col_, clr in (("Green", "C8E6C9"), ("Amber", "FFE0B2"), ("Red", "FFCDD2")):
    ws.conditional_formatting.add(f"E{a0}:E{r-1}", CellIsRule(operator="equal", formula=[f'"{col_}"'], fill=fill(clr)))
dname("QA_RAG", f"L6_Oversight!$E${a0}:$E${r-1}")

# ---------------- WP INDEX ----------------
wpi = S["WP_Index"]
widths(wpi, {"A": 7, "B": 13, "C": 46, "D": 30, "E": 14, "F": 11, "G": 14, "H": 11, "I": 13, "J": 9, "K": 9, "L": 9, "M": 12})
put(wpi, "A1", "WORKING PAPER INDEX & DOCUMENT REGISTER  |  drives the dashboard", font(14, True, "FFFFFF"), DK)
wpi.merge_cells("A1:M1"); wpi.row_dimensions[1].height = 26
put(wpi, "A2", "=Set_Client&\"  |  FYE \"&TEXT(Set_FYE,\"dd mmm yyyy\")&\"  |  ISA 230 - audit documentation; ISQM 1.31(f) - engagement documentation\"", font(9, i=True))
wpi.merge_cells("A2:M2")
put(wpi, "A3", "Update Status, sign-offs and hours (yellow). Status values: Not started / In progress / Prepared / Reviewed / Closed. 'Completed' on the dashboard = Reviewed + Closed.", font(8, i=True, c="616161"))
wpi.merge_cells("A3:M3")
heads = ["Ref", "Layer", "Working paper", "Standard reference", "Prepared by", "Prep date", "Reviewed by", "Review date",
         "Status", "Budget hrs", "Actual hrs", "Var hrs", "Go to WP"]
for j, h in enumerate(heads):
    put(wpi, f"{chr(65+j)}5", h, font(9, True, "FFFFFF"), DK, al=CEN, bd=BOX)
wpi.row_dimensions[5].height = 28
wps = [
 ("A1", "Strategy", "Engagement acceptance & continuance", "ISQM 1.30; ISA 220R.22-24", "Reviewed", 6, 6),
 ("A2", "Strategy", "Terms of engagement & preconditions", "ISA 210", "Closed", 3, 2.5),
 ("A3", "Strategy", "Independence declarations - team", "IESBA Part 4A; ISA 220R.16-21", "Prepared", 4, 5),
 ("A4", "Strategy", "Threats & safeguards register", "IESBA 120; 300", "Reviewed", 3, 3),
 ("A5", "Strategy", "Fee dependency & non-assurance services", "IESBA 410; 600", "Reviewed", 2, 1.5),
 ("A6", "Strategy", "Overall audit strategy", "ISA 300.7-8", "Reviewed", 8, 9),
 ("A7", "Strategy", "Planned scope communicated to TCWG", "ISA 260.15", "Closed", 2, 2),
 ("B1", "Planning", "Understanding the entity & environment", "ISA 315R.19", "Reviewed", 12, 14),
 ("B2", "Planning", "Understanding system of internal control", "ISA 315R.21-26", "Reviewed", 18, 21),
 ("B3", "Planning", "Materiality", "ISA 320; ISA 450", "Closed", 3, 3),
 ("B4", "Planning", "Risk assessment register", "ISA 315R.28-37", "Reviewed", 10, 11),
 ("B5", "Planning", "Fraud risk assessment", "ISA 240", "Reviewed", 6, 6),
 ("B6", "Planning", "Laws & regulations", "ISA 250", "Prepared", 4, 3),
 ("B7", "Planning", "Related parties", "ISA 550", "In progress", 5, 3),
 ("B8", "Planning", "Going concern - preliminary", "ISA 570.10", "Reviewed", 3, 3),
 ("B9", "Planning", "Opening balances", "ISA 510", "Closed", 2, 1),
 ("B10", "Planning", "Resource plan & timeline", "ISA 300.9; ISA 220R.25", "Reviewed", 3, 3),
 ("C1", "Methodology", "Audit programme by FS area & assertion", "ISA 330.6-7", "Reviewed", 8, 8),
 ("C2", "Methodology", "Sampling methodology (MUS)", "ISA 530", "Reviewed", 4, 4),
 ("C3", "Methodology", "Analytical procedures", "ISA 520", "Prepared", 8, 9),
 ("C4", "Methodology", "Accounting estimates", "ISA 540R", "Prepared", 12, 15),
 ("C5", "Methodology", "Internal audit function", "ISA 610", "Closed", 1, 0.5),
 ("C6", "Methodology", "Auditor's expert (IT)", "ISA 620", "Reviewed", 3, 3),
 ("C7", "Methodology", "Group audit considerations", "ISA 600R", "Closed", 1, 0.5),
 ("C8", "Methodology", "Service organisations (payroll)", "ISA 402", "Reviewed", 4, 4),
 ("D1", "Execution", "Lead schedule / trial balance", "ISA 500; ISA 230", "Reviewed", 6, 7),
 ("D2", "Execution", "Tests of controls", "ISA 330.8-17", "Reviewed", 40, 46),
 ("D3", "Execution", "Substantive tests of details", "ISA 330.18; ISA 500", "Prepared", 90, 98),
 ("D4", "Execution", "External confirmations", "ISA 505", "In progress", 14, 12),
 ("D5", "Execution", "Inventory count attendance", "ISA 501.4-8", "Reviewed", 16, 18),
 ("D6", "Execution", "Journal entry testing", "ISA 240.33", "Prepared", 12, 14),
 ("D7", "Execution", "Litigation & claims", "ISA 501.9-12", "In progress", 3, 1),
 ("D8", "Execution", "Subsequent events", "ISA 560", "In progress", 4, 2),
 ("D9", "Execution", "Comparative information", "ISA 710", "Reviewed", 2, 2),
 ("E1", "Findings", "Summary of misstatements", "ISA 450", "Prepared", 5, 6),
 ("E2", "Findings", "Control deficiencies evaluation", "ISA 265", "Prepared", 4, 4),
 ("E3", "Findings", "Issue log & remediation tracker", "ISA 230; ISQM 1.42", "In progress", 4, 3),
 ("E4", "Findings", "Management letter", "ISA 265.9-11", "In progress", 5, 3),
 ("F1", "Oversight", "Communications with TCWG", "ISA 260; ISA 265", "In progress", 6, 3),
 ("F2", "Oversight", "Written representations", "ISA 580", "Not started", 2, 0),
 ("F3", "Oversight", "Going concern - final conclusion", "ISA 570.16-24", "Prepared", 3, 2),
 ("F4", "Oversight", "Key audit matters", "ISA 701", "In progress", 4, 2),
 ("F5", "Oversight", "Engagement quality review", "ISQM 2; ISA 220R.36", "Not started", 10, 0),
 ("F6", "Oversight", "Completion checklist & partner sign-off", "ISA 220R.37-40", "Not started", 6, 0),
 ("F7", "Oversight", "Forming the opinion", "ISA 700R; 705R; 706R", "Not started", 4, 0),
 ("F8", "Oversight", "Other information", "ISA 720R", "Not started", 2, 0),
 ("F9", "Oversight", "Assembly of final file", "ISA 230.14; ISQM 1.A83", "Not started", 4, 0),
 ("F10", "Oversight", "ISQM 1 monitoring & quality indicators", "ISQM 1.35-47", "In progress", 3, 1)]
preparer = {"Strategy": "S. Oyunaa", "Planning": "S. Oyunaa", "Methodology": "T. Enkhjin", "Execution": "T. Enkhjin",
            "Findings": "S. Oyunaa", "Oversight": "S. Oyunaa"}
reviewer = {"Strategy": "B. Bat-Erdene", "Planning": "B. Bat-Erdene", "Methodology": "S. Oyunaa", "Execution": "S. Oyunaa",
            "Findings": "B. Bat-Erdene", "Oversight": "B. Bat-Erdene"}
r = 6
w0 = r
base = {"Strategy": D(2025, 9, 20), "Planning": D(2025, 11, 10), "Methodology": D(2025, 11, 14), "Execution": D(2026, 2, 27),
        "Findings": D(2026, 3, 16), "Oversight": D(2026, 3, 18)}
for ref, lay, title, std, stt, bh, ah in wps:
    cell(wpi, f"A{r}", ref, center=True); cell(wpi, f"B{r}", lay, center=True); cell(wpi, f"C{r}", title); cell(wpi, f"D{r}", std)
    prepd = stt in ("Prepared", "Reviewed", "Closed")
    revd = stt in ("Reviewed", "Closed")
    cell(wpi, f"E{r}", preparer[lay] if prepd else None, is_input=True, center=True)
    cell(wpi, f"F{r}", base[lay] if prepd else None, DATE, True, center=True)
    cell(wpi, f"G{r}", reviewer[lay] if revd else None, is_input=True, center=True)
    cell(wpi, f"H{r}", (base[lay] + dt.timedelta(days=3)) if revd else None, DATE, True, center=True)
    cell(wpi, f"I{r}", stt, is_input=True, center=True)
    cell(wpi, f"J{r}", bh, "0.0", True, center=True); cell(wpi, f"K{r}", ah, "0.0", True, center=True)
    cell(wpi, f"L{r}", f"=K{r}-J{r}", '0.0;(0.0);"-"', center=True)
    sh, ar = anchors[ref]
    c = cell(wpi, f"M{r}", f'=HYPERLINK("#{sh}!A{ar}","Open >")', center=True)
    c.font = Font(name=F, size=9, color="1565C0", underline="single")
    r += 1
w1 = r - 1
put(wpi, f"C{r}", "TOTAL", font(9, True), PALE, bd=BOX)
for col in "JKL": put(wpi, f"{col}{r}", f"=SUM({col}{w0}:{col}{w1})", font(9, True), PALE, "0.0", CEN, BOX)
put(wpi, f"I{r}", f'=COUNTA(I{w0}:I{w1})&" WPs"', font(9, True), PALE, al=CEN, bd=BOX)
dv(wpi, f"I{w0}:I{w1}", "Status"); dv(wpi, f"B{w0}:B{w1}", "Layer")
stcol = {"Not started": "EEEEEE", "In progress": "FFE0B2", "Prepared": "BBDEFB", "Reviewed": "C8E6C9", "Closed": "81C784"}
for k, v in stcol.items():
    wpi.conditional_formatting.add(f"I{w0}:I{w1}", CellIsRule(operator="equal", formula=[f'"{k}"'], fill=fill(v)))
wpi.conditional_formatting.add(f"L{w0}:L{w1}", CellIsRule(operator="greaterThan", formula=["0"], font=Font(name=F, color="C62828", bold=True)))
wpi.freeze_panes = "A6"
dname("WP_Layer", f"WP_Index!$B${w0}:$B${w1}"); dname("WP_Status", f"WP_Index!$I${w0}:$I${w1}")
dname("WP_Bud", f"WP_Index!$J${w0}:$J${w1}"); dname("WP_Act", f"WP_Index!$K${w0}:$K${w1}")

# ---------------- DASHDATA ----------------
dd = S["DashData"]
widths(dd, {"A": 34, "B": 13, "C": 13, "D": 13, "E": 13, "F": 13, "G": 11, "H": 12, "I": 12, "J": 12})
put(dd, "A1", "Dashboard calculation engine (formulas only - feeds charts & KPI tiles)", font(12, True))
hd = ["Layer", "Not started", "In progress", "Prepared", "Reviewed", "Closed", "Total", "% Complete", "Budget hrs", "Actual hrs"]
for j, h in enumerate(hd): put(dd, f"{chr(65+j)}3", h, font(9, True, "FFFFFF"), DK, al=CEN, bd=BOX)
for i, lay in enumerate(lists["Layer"]):
    rr = 4 + i
    put(dd, f"A{rr}", lay, bd=BOX)
    for j, s in enumerate(lists["Status"]):
        col = chr(66 + j)
        put(dd, f"{col}{rr}", f'=COUNTIFS(WP_Layer,$A{rr},WP_Status,{col}$3)', bd=BOX, al=CEN)
    put(dd, f"G{rr}", f"=SUM(B{rr}:F{rr})", bd=BOX, al=CEN)
    put(dd, f"H{rr}", f"=IFERROR((E{rr}+F{rr})/G{rr},0)", nf="0%", bd=BOX, al=CEN)
    put(dd, f"I{rr}", f"=SUMIFS(WP_Bud,WP_Layer,$A{rr})", nf="0", bd=BOX, al=CEN)
    put(dd, f"J{rr}", f"=SUMIFS(WP_Act,WP_Layer,$A{rr})", nf="0", bd=BOX, al=CEN)
put(dd, "A10", "Total", font(b=True), bd=BOX)
for col in "BCDEFGIJ": put(dd, f"{col}10", f"=SUM({col}4:{col}9)", font(b=True), bd=BOX, al=CEN)
put(dd, "H10", "=IFERROR((E10+F10)/G10,0)", font(b=True), nf="0%", bd=BOX, al=CEN)
for i, lay in enumerate(lists["Layer"]): dname(f"DD_Pct_{lay}", f"DashData!$H${4+i}")

put(dd, "A13", "Findings by severity", font(9, True, "FFFFFF"), DK); put(dd, "B13", "All", font(9, True, "FFFFFF"), DK); put(dd, "C13", "Open", font(9, True, "FFFFFF"), DK)
for i, s in enumerate(lists["Severity"]):
    rr = 14 + i; put(dd, f"A{rr}", s, bd=BOX)
    put(dd, f"B{rr}", f'=COUNTIF(Iss_Sev,A{rr})', bd=BOX, al=CEN)
    put(dd, f"C{rr}", f'=COUNTIFS(Iss_Sev,A{rr},Iss_Status,"<>Closed")', bd=BOX, al=CEN)

put(dd, "A19", "Risk spectrum (ISA 315R)", font(9, True, "FFFFFF"), DK); put(dd, "B19", "Count", font(9, True, "FFFFFF"), DK)
for i, s in enumerate(["Low", "Moderate", "Higher", "Significant"]):
    rr = 20 + i; put(dd, f"A{rr}", s, bd=BOX); put(dd, f"B{rr}", f'=COUNTIF(Risk_Spectrum,A{rr})', bd=BOX, al=CEN)

put(dd, "A26", "Misstatements vs materiality", font(9, True, "FFFFFF"), DK); put(dd, "B26", "MNT '000", font(9, True, "FFFFFF"), DK)
for i, (lab, f_) in enumerate([("Uncorrected (aggregate)", "=ABS(MS_Uncorr)"), ("Clearly trivial", "=Mat_CT"),
                               ("Performance materiality", "=Mat_PM"), ("Overall materiality", "=Mat_OM")]):
    rr = 27 + i; put(dd, f"A{rr}", lab, bd=BOX); put(dd, f"B{rr}", f_, nf=NUM, bd=BOX)

put(dd, "A33", "Timeline phase", font(9, True, "FFFFFF"), DK); put(dd, "B33", "% done", font(9, True, "FFFFFF"), DK)
for i in range(PH[1] - PH[0] + 1):
    rr = 34 + i
    put(dd, f"A{rr}", f"=INDEX(TL_Phase,{i+1})", bd=BOX); put(dd, f"B{rr}", f"=INDEX(TL_Pct,{i+1})", nf="0%", bd=BOX)
tl_end = 34 + PH[1] - PH[0]

kp = 46
put(dd, f"A{kp}", "KPI", font(9, True, "FFFFFF"), DK); put(dd, f"B{kp}", "Value", font(9, True, "FFFFFF"), DK)
kpis = [("TotalWP", "Total working papers", "=G10", "0"),
        ("PctComplete", "% WPs completed (reviewed/closed)", "=H10", "0%"),
        ("Completed", "WPs completed", "=E10+F10", "0"),
        ("SigRisks", "Significant risks", '=COUNTIF(Risk_Spectrum,"Significant")', "0"),
        ("TotalRisks", "Total risks assessed", "=COUNTA(Risk_Spectrum)", "0"),
        ("OpenIssues", "Open findings", '=COUNTIF(Iss_Status,"<>Closed")', "0"),
        ("OpenHigh", "Open high-severity findings", '=COUNTIFS(Iss_Status,"<>Closed",Iss_Sev,"High")', "0"),
        ("Overdue", "Overdue findings", '=COUNTIFS(Iss_Status,"<>Closed",Iss_Due,"<"&Set_AsOf)', "0"),
        ("HrsBud", "Budget hours", "=I10", "#,##0"),
        ("HrsAct", "Actual hours", "=J10", "#,##0"),
        ("HrsPct", "Actual / budget", "=IFERROR(J10/I10,0)", "0%"),
        ("IndRecv", "Independence declarations received", "=COUNT(Ind_Date)", "0"),
        ("IndTot", "Team members", "=ROWS(Ind_Date)", "0"),
        ("IndPct", "Declarations %", f"=IFERROR(B{kp+12}/B{kp+13},0)", "0%"),
        ("IndThreat", "Members with threats / pending", '=ROWS(Ind_Assess)-COUNTIF(Ind_Assess,"Independent")', "0"),
        ("CommPct", "TCWG communications done", '=IFERROR(COUNTIF(Comm_Status,"Done")/(ROWS(Comm_Status)-COUNTIF(Comm_Status,"N/A")),0)', "0%"),
        ("CtlPct", "Controls effective", '=IFERROR(COUNTIF(Ctl_Result,"Effective")/ROWS(Ctl_Result),0)', "0%"),
        ("ConfPct", "Confirmations received", '=IFERROR(COUNTIF(Conf_Status,"Received*")/ROWS(Conf_Status),0)', "0%"),
        ("DaysToReport", "Days to report date", "=Set_RptDate-Set_AsOf", "0"),
        ("SigDef", "Significant deficiencies", '=COUNTIF(CD_Class,"Significant deficiency")', "0"),
        ("QAGreen", "Quality indicators green", '=COUNTIF(QA_RAG,"Green")&" / "&ROWS(QA_RAG)', "@"),
        ("EQR", "EQR status", '=INDEX(WP_Status,MATCH("F5",WP_Index!$A$6:$A$200,0))', "@")]
for i, (nm, lab, f_, nf) in enumerate(kpis):
    rr = kp + 1 + i
    put(dd, f"A{rr}", lab, bd=BOX); put(dd, f"B{rr}", f_, nf=nf, bd=BOX, al=CEN)
    dname("DD_" + nm, f"DashData!$B${rr}")

# ---------------- STANDARDS ----------------
sd = S["Standards"]
widths(sd, {"A": 16, "B": 58, "C": 16, "D": 40})
put(sd, "A1", "STANDARDS LIBRARY - ISA / ISQM / IESBA mapped to audit layers", font(13, True, "FFFFFF"), DK); sd.merge_cells("A1:D1")
for j, h in enumerate(["Standard", "Title", "Layer", "Working papers"]):
    put(sd, f"{chr(65+j)}3", h, font(9, True, "FFFFFF"), MD, al=CEN, bd=BOX)
stds = [("ISQM 1", "Quality Management for Firms that Perform Audits or Reviews", "Strategy / Oversight", "A1, E3, F10"),
        ("ISQM 2", "Engagement Quality Reviews", "Oversight", "F5"),
        ("IESBA Code", "International Code of Ethics for Professional Accountants (incl. Independence Standards)", "Strategy", "A3, A4, A5"),
        ("ISA 200", "Overall Objectives of the Independent Auditor", "All", "-"),
        ("ISA 210", "Agreeing the Terms of Audit Engagements", "Strategy", "A2"),
        ("ISA 220 (Revised)", "Quality Management for an Audit of Financial Statements", "Strategy / Oversight", "A1, A3, B10, F5, F6"),
        ("ISA 230", "Audit Documentation", "All", "WP_Index, D1, E3, F9"),
        ("ISA 240", "Auditor's Responsibilities Relating to Fraud", "Planning / Execution", "B5, D6"),
        ("ISA 250 (Revised)", "Consideration of Laws and Regulations", "Planning", "B6"),
        ("ISA 260 (Revised)", "Communication with Those Charged with Governance", "Strategy / Oversight", "A7, F1"),
        ("ISA 265", "Communicating Deficiencies in Internal Control", "Findings", "E2, E4"),
        ("ISA 300", "Planning an Audit of Financial Statements", "Strategy / Planning", "A6, B10"),
        ("ISA 315 (Revised 2019)", "Identifying and Assessing the Risks of Material Misstatement", "Planning", "B1, B2, B4"),
        ("ISA 320", "Materiality in Planning and Performing an Audit", "Planning", "B3"),
        ("ISA 330", "Auditor's Responses to Assessed Risks", "Methodology / Execution", "C1, D2, D3"),
        ("ISA 402", "Audit Considerations Relating to Service Organizations", "Methodology", "C8"),
        ("ISA 450", "Evaluation of Misstatements Identified during the Audit", "Findings", "E1"),
        ("ISA 500", "Audit Evidence", "Execution", "D1, D3"),
        ("ISA 501", "Audit Evidence - Specific Considerations", "Execution", "D5, D7"),
        ("ISA 505", "External Confirmations", "Execution", "D4"),
        ("ISA 510", "Initial Audit Engagements - Opening Balances", "Planning", "B9"),
        ("ISA 520", "Analytical Procedures", "Methodology", "C3"),
        ("ISA 530", "Audit Sampling", "Methodology", "C2"),
        ("ISA 540 (Revised)", "Auditing Accounting Estimates and Related Disclosures", "Methodology", "C4"),
        ("ISA 550", "Related Parties", "Planning", "B7"),
        ("ISA 560", "Subsequent Events", "Execution", "D8"),
        ("ISA 570 (Revised)", "Going Concern", "Planning / Oversight", "B8, F3"),
        ("ISA 580", "Written Representations", "Oversight", "F2"),
        ("ISA 600 (Revised)", "Audits of Group Financial Statements", "Methodology", "C7"),
        ("ISA 610 (Revised 2013)", "Using the Work of Internal Auditors", "Methodology", "C5"),
        ("ISA 620", "Using the Work of an Auditor's Expert", "Methodology", "C6"),
        ("ISA 700 (Revised)", "Forming an Opinion and Reporting on Financial Statements", "Oversight", "F7"),
        ("ISA 701", "Communicating Key Audit Matters", "Oversight", "F4"),
        ("ISA 705 (Revised)", "Modifications to the Opinion", "Oversight", "F7"),
        ("ISA 706 (Revised)", "Emphasis of Matter & Other Matter Paragraphs", "Oversight", "F7"),
        ("ISA 710", "Comparative Information", "Execution", "D9"),
        ("ISA 720 (Revised)", "Auditor's Responsibilities Relating to Other Information", "Oversight", "F8")]
for i, row in enumerate(stds):
    for j, v in enumerate(row): put(sd, f"{chr(65+j)}{4+i}", v, font(9), PALE if i % 2 else None, bd=BOX, al=WRAP)
put(sd, f"A{5+len(stds)}", "Paragraph references follow the IAASB Handbook and IESBA Code; verify against the edition adopted locally (e.g. by the Mongolian Institute of Certified Public Accountants).", font(8, i=True, c="616161"))

# ---------------- LAYERS MAP ----------------
lm = S["LayersMap"]
for col in "ABCDEFGHIJKLMNOPQ": lm.column_dimensions[col].width = 9.5
lm.column_dimensions["A"].width = 2
for rr in range(1, 63): lm.row_dimensions[rr].height = 15
put(lm, "B1", "AUDIT LAYERS ACROSS THE ENGAGEMENT", font(16, True, "FFFFFF"), "101010"); lm.merge_cells("B1:Q2")
lm["B1"].alignment = Alignment(vertical="center", indent=1)
put(lm, "B3", "=Set_Client&\"  |  FYE \"&TEXT(Set_FYE,\"dd-mmm-yyyy\")&\"  |  Completion % = WPs Reviewed/Closed ÷ total in layer (live from WP_Index)\"",
    font(9, i=True, c="C5E1A5"), "101010"); lm.merge_cells("B3:Q3")
for rr in range(1, 63):
    for col in "BCDEFGHIJKLMNOPQ":
        if not lm[f"{col}{rr}"].fill or lm[f"{col}{rr}"].fill.fgColor.rgb in (None, "00000000"):
            lm[f"{col}{rr}"].fill = fill("101010")
boxes = {
    "Strategy": ("G5", ["Audit objective definition", "Scope alignment", "Stakeholder expectations", "Audit direction setting"],
                 ["Acceptance & continuance (A1)", "Engagement letter (A2)", "Independence (A3-A5)", "Overall strategy (A6)"], "L1_Strategy"),
    "Planning": ("B19", ["Understanding entity & IC", "Materiality setting", "Risk assessment", "Resource & timeline"],
                 ["Materiality WP (B3)", "Risk register (B4)", "Fraud brainstorm (B5)", "Resource plan (B10)"], "L2_Planning"),
    "Execution": ("L19", ["Fieldwork execution", "Control testing", "Data validation", "Process walkthrough"],
                  ["Lead schedule (D1)", "Tests of controls (D2)", "Confirmations (D4)", "JE testing (D6)"], "L4_Execution"),
    "Methodology": ("B33", ["Standard procedures", "Sampling techniques", "Evaluation methods", "Quality standards"],
                    ["Audit programme (C1)", "MUS calculator (C2)", "Analytics (C3)", "Estimates (C4)"], "L3_Methodology"),
    "Oversight": ("L33", ["TCWG interaction", "Quality assurance reviews", "Independence monitoring", "Performance evaluation"],
                  ["TCWG log (F1)", "EQ review (F5)", "Opinion matrix (F7)", "ISQM 1 metrics (F10)"], "L6_Oversight"),
    "Findings": ("G47", ["Issue identification", "Root cause analysis", "Recommendation design", "Remediation tracking"],
                 ["Misstatements SUM (E1)", "Deficiencies (E2)", "Issue log (E3)", "Management letter (E4)"], "L5_Findings"),
}
from openpyxl.utils import column_index_from_string as ci, get_column_letter as gl
for lay, (anchor, focus, ctrls_, sheet) in boxes.items():
    c0 = ci(anchor.rstrip("0123456789")); r0 = int(anchor[len(anchor.rstrip("0123456789")):])
    cA, cZ = gl(c0), gl(c0 + 5)
    title = {"Findings": "Findings & Remediation Layer"}.get(lay, f"Audit {lay} Layer")
    c = put(lm, f"{cA}{r0}", f'=HYPERLINK("#{sheet}!A1","{title}  >")', font(11, True, "FFFFFF"), "101010", al=CEN)
    lm.merge_cells(f"{cA}{r0}:{cZ}{r0+1}")
    put(lm, f"{cA}{r0+2}", "Common Focus Areas:", font(9, True, "9CCC65", i=True), "101010")
    for k, t in enumerate(focus):
        put(lm, f"{cA}{r0+3+k}", "▪ " + t, font(9, c="FFFFFF"), "101010"); lm.merge_cells(f"{cA}{r0+3+k}:{cZ}{r0+3+k}")
    put(lm, f"{cA}{r0+7}", "Key Controls / WPs:", font(9, True, "9CCC65", i=True), "101010")
    for k, t in enumerate(ctrls_):
        put(lm, f"{cA}{r0+8+k}", "▪ " + t, font(9, c="FFFFFF"), "101010"); lm.merge_cells(f"{cA}{r0+8+k}:{cZ}{r0+8+k}")
    put(lm, f"{cA}{r0+12}", "Completion", font(9, True, "9CCC65"), "101010")
    put(lm, f"{gl(c0+2)}{r0+12}", f"=DD_Pct_{lay}", font(11, True, "B2FF59"), "101010", "0%", CEN)
    put(lm, f"{gl(c0+3)}{r0+12}", f'=REPT("█",ROUND(DD_Pct_{lay}*10,0))&REPT("░",10-ROUND(DD_Pct_{lay}*10,0))', font(9, c="76FF03"), "101010")
    lm.merge_cells(f"{gl(c0+3)}{r0+12}:{cZ}{r0+12}")
    sideL = Side(style="medium", color="76FF03")
    for rr in range(r0, r0 + 13):
        for cc in range(c0, c0 + 6):
            cl = lm.cell(rr, cc)
            cl.border = Border(left=sideL if cc == c0 else None, right=sideL if cc == c0 + 5 else None,
                               top=sideL if rr == r0 else None, bottom=sideL if rr == r0 + 12 else None)
put(lm, "H26", "AUDIT", font(30, True, "B2FF59"), "101010", al=CEN); lm.merge_cells("H26:K31")
ring = Side(style="thick", color="76FF03")
for rr in range(26, 32):
    for cc in range(ci("H"), ci("K") + 1):
        lm.cell(rr, cc).border = Border(left=ring if cc == ci("H") else None, right=ring if cc == ci("K") else None,
                                        top=ring if rr == 26 else None, bottom=ring if rr == 31 else None)
put(lm, "H32", '="Overall "&TEXT(DD_PctComplete,"0%")&" complete"', font(10, True, "FFFFFF"), "101010", al=CEN); lm.merge_cells("H32:K32")
put(lm, "B61", "Click a layer title to open its working papers.  |  ISA · ISQM 1 · ISQM 2 · IESBA Code", font(9, i=True, c="9CCC65"), "101010")
lm.merge_cells("B61:Q61")

# ---------------- DASHBOARD ----------------
db = S["Dashboard"]
for col in range(1, 30): db.column_dimensions[gl(col)].width = 8.5
db.column_dimensions["A"].width = 2; db.column_dimensions["B"].width = 3; db.column_dimensions["C"].width = 16; db.column_dimensions["D"].width = 2
for rr in range(1, 80):
    db.row_dimensions[rr].height = 16
    for cc in range(1, 26): db.cell(rr, cc).fill = fill(BG)
# Banner
for cc in range(2, 26):
    for rr in (1, 2, 3): db.cell(rr, cc).fill = fill("FFFFFF")
put(db, "C1", "AW", font(16, True, "FFFFFF"), DK, al=CEN); db.merge_cells("C1:C3")
put(db, "E1", "AUDIT DASHBOARD", font(20, True, DK), "FFFFFF"); db.merge_cells("E1:L2")
put(db, "E3", "Overview of ISA / ISQM / IESBA engagement status", font(10, c="424242"), "FFFFFF"); db.merge_cells("E3:L3")
for k, (lab, f_, nf) in enumerate([("Client", "=Set_Client", None), ("Year end", "=Set_FYE", "dd mmm yyyy"), ("Status as of", "=Set_AsOf", "dd mmm yyyy")]):
    cA = gl(14 + k * 4); cZ = gl(14 + k * 4 + 3)
    put(db, f"{cA}1", lab, font(8, True, "616161"), "FFFFFF"); db.merge_cells(f"{cA}1:{cZ}1")
    put(db, f"{cA}2", f_, font(10, True, DK), "FFFFFF", nf, Alignment(horizontal="left", vertical="center"), GBOX)
    db.merge_cells(f"{cA}2:{cZ}3")
for cc in range(2, 26): db.cell(4, cc).fill = fill(DK)
db.row_dimensions[4].height = 5

# Nav panel
for rr in range(6, 47):
    for cc in (2, 3): db.cell(rr, cc).fill = fill(DK)
navs = [("Overview", "LayersMap"), ("Setup", "Setup"), ("WP Index", "WP_Index"), ("1 Strategy", "L1_Strategy"),
        ("2 Planning", "L2_Planning"), ("3 Methodology", "L3_Methodology"), ("4 Execution", "L4_Execution"),
        ("5 Findings", "L5_Findings"), ("6 Oversight", "L6_Oversight"), ("Standards", "Standards")]
for k, (lab, sh) in enumerate(navs):
    rr = 7 + k * 3
    c = put(db, f"C{rr}", f'=HYPERLINK("#{sh}!A1","{lab}")', font(10, True, DK if k == 0 else "FFFFFF"),
            "C5E1A5" if k == 0 else DK, al=Alignment(horizontal="left", vertical="center", indent=1))
    db.merge_cells(f"C{rr}:C{rr+1}")

def tile(r0, c0, w, label, valf, nf, subf, accent=LT, h=4):
    cA, cZ = gl(c0), gl(c0 + w - 1)
    for rr in range(r0, r0 + h):
        for cc in range(c0, c0 + w):
            cl = db.cell(rr, cc); cl.fill = fill("FFFFFF")
            cl.border = Border(left=gs if cc == c0 else None, right=gs if cc == c0 + w - 1 else None,
                               top=Side(style="thick", color=accent) if rr == r0 else None, bottom=gs if rr == r0 + h - 1 else None)
    put(db, f"{cA}{r0}", label, font(9, True, "424242"), "FFFFFF", al=CEN); db.merge_cells(f"{cA}{r0}:{cZ}{r0}")
    put(db, f"{cA}{r0+1}", valf, font(18, True, DK), "FFFFFF", nf, CEN); db.merge_cells(f"{cA}{r0+1}:{cZ}{r0+2}")
    put(db, f"{cA}{r0+3}", subf, font(8, c="616161"), "FFFFFF", al=CEN); db.merge_cells(f"{cA}{r0+3}:{cZ}{r0+3}")
    db[f"{cA}{r0}"].border = Border(top=Side(style="thick", color=accent), left=gs)

T1 = [("Total Working Papers", "=DD_TotalWP", "0", '="▲ "&DD_Completed&" completed"'),
      ("WP Completion", "=DD_PctComplete", "0%", '="Reviewed + closed of "&DD_TotalWP'),
      ("Overall Materiality", "=Mat_OM", '#,##0', '="PM "&TEXT(Mat_PM,"#,##0")&" | CT "&TEXT(Mat_CT,"#,##0")'),
      ("Significant Risks", "=DD_SigRisks", "0", '="of "&DD_TotalRisks&" RMMs assessed (ISA 315R)"')]
T2 = [("Open Findings", "=DD_OpenIssues", "0", '="▼ "&DD_OpenHigh&" high | "&DD_Overdue&" overdue"'),
      ("Uncorrected Misstatements", "=ABS(MS_Uncorr)", "#,##0", '=TEXT(MS_UncorrPct,"0%")&" of OM (ISA 450)"'),
      ("Hours: Actual vs Budget", "=DD_HrsPct", "0%", '=TEXT(DD_HrsAct,"#,##0")&" / "&TEXT(DD_HrsBud,"#,##0")&" hrs"'),
      ("Independence (IESBA)", '=DD_IndRecv&" / "&DD_IndTot', "@", '=DD_IndThreat&" member(s) need follow-up"')]
for k, t in enumerate(T1): tile(6, 5 + k * 5, 4, *t, accent=[LT, LT, "1565C0", "C62828"][k])
for k, t in enumerate(T2): tile(11, 5 + k * 5, 4, *t, accent=["EF6C00", "C62828", "F9A825", LT][k])
# Proposed opinion tile (wide)

# Chart data anchors
def bar_chart(title, cats, data_ref, stacked=False, horizontal=False, colors=None, titles=True, w=12, h=6.2, legend=True, pct=False):
    ch = BarChart(); ch.type = "bar" if horizontal else "col"
    ch.title = title; ch.style = 10; ch.roundedCorners = False
    if stacked: ch.grouping = "stacked"; ch.overlap = 100
    ch.add_data(data_ref, titles_from_data=titles); ch.set_categories(cats)
    ch.width, ch.height = w, h
    ch.x_axis.delete = False; ch.y_axis.delete = False
    ch.y_axis.majorGridlines = None
    if pct: ch.y_axis.number_format = "0%"; ch.y_axis.scaling.max = 1; ch.y_axis.scaling.min = 0
    if horizontal: ch.x_axis.scaling.orientation = "maxMin"
    if legend: ch.legend.position = "b"; ch.legend.overlay = False
    else: ch.legend = None
    if colors:
        for s, col in zip(ch.series, colors):
            s.graphicalProperties.solidFill = col; s.graphicalProperties.line.solidFill = col
    if not stacked:
        ch.dataLabels = DataLabelList(showVal=True, showSerName=False, showCatName=False, showLegendKey=False, showPercent=False)
        if pct: ch.dataLabels = None
    ch.title.overlay = False
    return ch

# 1 Stacked status by layer
c1 = bar_chart("Working Papers by Layer & Status", Reference(dd, min_col=1, min_row=4, max_row=9),
               Reference(dd, min_col=2, max_col=6, min_row=3, max_row=9), stacked=True, horizontal=True,
               colors=["BDBDBD", "FFB300", "42A5F5", "66BB6A", DK], w=13, h=7)
db.add_chart(c1, "E16")
# 2 doughnut severity
c2 = DoughnutChart(); c2.title = "Findings by Severity"; c2.roundedCorners = False
c2.add_data(Reference(dd, min_col=2, min_row=13, max_row=16), titles_from_data=True)
c2.set_categories(Reference(dd, min_col=1, min_row=14, max_row=16)); c2.holeSize = 55
c2.width, c2.height = 8, 7; c2.legend.position = "r"; c2.legend.overlay = False
c2.dataLabels = DataLabelList(showVal=True, showSerName=False, showCatName=False, showLegendKey=False, showPercent=False)
from openpyxl.chart.series import DataPoint
for idx, col in enumerate(["C62828", "FFB300", "66BB6A"]):
    pt = DataPoint(idx=idx); pt.graphicalProperties.solidFill = col; c2.series[0].dPt.append(pt)
db.add_chart(c2, "N16")
# 3 Risk spectrum
c3 = bar_chart("Risk Spectrum", Reference(dd, min_col=1, min_row=20, max_row=23),
               Reference(dd, min_col=2, min_row=19, max_row=23), colors=[LT], w=8, h=7, legend=False)
pts = ["81C784", "FFB300", "EF6C00", "C62828"]
for idx, col in enumerate(pts):
    pt = DataPoint(idx=idx); pt.graphicalProperties.solidFill = col; c3.series[0].dPt.append(pt)
db.add_chart(c3, "T16")
# 4 Budget vs Actual
c4 = bar_chart("Budget vs Actual Hours by Layer", Reference(dd, min_col=1, min_row=4, max_row=9),
               Reference(dd, min_col=9, max_col=10, min_row=3, max_row=9), colors=["A5D6A7", DK], w=13, h=7)
db.add_chart(c4, "E31")
# 5 Misstatements vs materiality
c5 = bar_chart("Misstatements vs Materiality", Reference(dd, min_col=1, min_row=27, max_row=30),
               Reference(dd, min_col=2, min_row=26, max_row=30), horizontal=True, colors=[LT], w=8, h=7, legend=False)
for idx, col in enumerate(["C62828", "BDBDBD", "FFB300", DK]):
    pt = DataPoint(idx=idx); pt.graphicalProperties.solidFill = col; c5.series[0].dPt.append(pt)
db.add_chart(c5, "N31")
# 6 Timeline
c6 = bar_chart("Audit Timeline - % Complete", Reference(dd, min_col=1, min_row=34, max_row=tl_end),
               Reference(dd, min_col=2, min_row=33, max_row=tl_end), horizontal=True, colors=[LT], w=8, h=7, legend=False, pct=True)
db.add_chart(c6, "T31")

# bottom KPI strip
B = [("TCWG Comms", "=DD_CommPct", "0%", '="ISA 260 matters done"'),
     ("Controls Effective", "=DD_CtlPct", "0%", '="of controls tested (ISA 330)"'),
     ("Confirmations", "=DD_ConfPct", "0%", '="received (ISA 505)"'),
     ("Signif. Deficiencies", "=DD_SigDef", "0", '="to TCWG in writing (ISA 265)"'),
     ("Days to Report", "=DD_DaysToReport", "0", '="EQR: "&DD_EQR'),
     ("Proposed Opinion", "=Opinion", "@", '="Quality KPIs green: "&DD_QAGreen')]
for k, t in enumerate(B):
    tile(46, 5 + k * 3 + (0 if k < 5 else 0), 3 if k < 5 else 6, *t, accent=LT)
db["T47"].font = font(11, True, DK)
put(db, "E51", "=\"Misstatement evaluation: \"&MS_Eval", font(9, True, "B71C1C"), BG); db.merge_cells("E51:Y51")
put(db, "E52", "All figures update live from the layer working papers. Illustrative data for teaching - replace in Setup & layer sheets.", font(8, i=True, c="757575"), BG)
db.merge_cells("E52:Y52")
db.freeze_panes = "A5"
db.print_area = "A1:Y52"

# ---------------- defined names ----------------
for n, ref in names.items():
    wb.defined_names[n] = DefinedName(n, attr_text=ref)

# print setup
for ws in S.values():
    ws.page_setup.orientation = "landscape"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
S["LayersMap"].page_setup.fitToHeight = 1
S["Dashboard"].page_setup.fitToHeight = 1
wb.active = 0
wb.save(OUT)
print("saved", OUT, len(names), "names")
