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
         "L3_Methodology", "Sampling", "L4_Execution", "IC_Assessment", "L5_Findings", "SUAD", "L6_Oversight", "DashData", "Standards", "Lists"]
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
    "ICRating": ["Effective", "Needs improvement", "Ineffective"],
    "Nature": ["Overstatement", "Understatement", "Classification", "Disclosure"],
    "CtlType": ["Preventive", "Detective"],
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

r = section(ws, r, "C2", "Sampling methodology - policy (detail on Sampling sheet S1-S5)", "ISA 530")
r = table_head(ws, r, ["#", "Test type", "Method", "Sample size basis", "", "", "Documented in"])
ws.merge_cells(f"D{r-1}:F{r-1}")
pol = [("Tests of controls - daily / multiple-daily", "Statistical attribute sampling", "n = RF(k) / tolerable deviation rate", "Sampling S2; IC4-IC5"),
       ("Tests of controls - weekly or less frequent", "Non-statistical", "Frequency table (min-max)", "Sampling S3; IC4"),
       ("Tests of details - large monetary populations", "Monetary unit sampling (MUS)", "n = Residual pop x RF / (TM - EM x EF)", "Sampling S1, S4, S5"),
       ("Individually significant / key items", "100% examination (not sampling)", "Items > sampling interval or > PM", "Sampling S1"),
       ("Journal entries", "Risk-criteria selection with CAATs (not sampling)", "All entries meeting criteria", "D6"),
       ("Small homogeneous populations", "Non-statistical random / haphazard", "Auditor judgement, documented", "Sampling S1")]
for i, row in enumerate(pol, 1):
    cell(ws, f"A{r}", i, center=True); cell(ws, f"B{r}", row[0]); cell(ws, f"C{r}", row[1])
    cell(ws, f"D{r}", row[2]); ws.merge_cells(f"D{r}:F{r}"); cell(ws, f"G{r}", row[3], center=True)
    ws.row_dimensions[r].height = 28; r += 1
cell(ws, f"B{r}", "Tolerable misstatement for tests of details"); cell(ws, f"C{r}", "=Mat_PM", NUM, link=True)
cell(ws, f"B{r+1}", "Receivables MUS sample size (S1.1)"); cell(ws, f"C{r+1}", "=INDEX(Pop_SampleSize,1)", "0", link=True)
c = cell(ws, f"B{r+2}", '=HYPERLINK("#Sampling!A1","Open Sampling working paper >")'); c.font = Font(name=F, size=9, color="1565C0", underline="single")
r += 4

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

r = section(ws, r, "D2", "Tests of operating effectiveness of controls (detail: IC_Assessment IC1-IC6)", "ISA 330.8-17")
r = table_head(ws, r, ["", "Summary", "Value"])
for i, (lab, f_) in enumerate([("Entity-level control environment (IC1)", "=IC_COSO"), ("ITGC conclusion (IC3)", "=IC_ITGC"),
                               ("Controls in RCM (IC2)", "=IC_RCMn"), ("D&I effective (IC2)", "=IC_DIeff"),
                               ("Controls tested for operating effectiveness (IC4)", "=COUNTA(ToE_ID)"),
                               ("Operating effectively", '=COUNTIF(ToE_Result,"Effective")'), ("Not effective", '=COUNTIF(ToE_Result,"Not effective")'),
                               ("Items tested / deviations", '=SUM(ToE_Tested)&" / "&SUM(ToE_Dev)'), ("Controls relied upon", "=IC_Rely")]):
    cell(ws, f"B{r}", lab); c = cell(ws, f"C{r}", f_, link=True); ws.merge_cells(f"C{r}:J{r}"); r += 1
c = cell(ws, f"B{r}", '=HYPERLINK("#IC_Assessment!A1","Open internal control assessment & testing >")'); c.font = Font(name=F, size=9, color="1565C0", underline="single")
r += 2

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
r = section(ws, r, "E1", "Summary of misstatements (full schedule on SUAD sheet)", "ISA 450.5-12")
r = table_head(ws, r, ["", "Item", "Value"])
for i, (lab, f_, nf) in enumerate([("Misstatements logged (all, incl. trivial & corrected)", "=COUNT(MS_Amt)", "0"),
        ("Corrected by management", '=COUNTIF(MS_Corr,"Yes")', "0"),
        ("Uncorrected & accumulated (above CT / projected)", '=COUNTIFS(MS_Corr,"No",MS_Acc,"Yes")', "0"),
        ("Current-year uncorrected PBT effect", "=MS_Uncorr", NUM), ("Dual-approach aggregate incl. prior-period effects", "=MS_Agg", NUM),
        ("Aggregate as % of overall materiality", "=MS_UncorrPct", PCT), ("Overall materiality", "=Mat_OM", NUM),
        ("Performance materiality", "=Mat_PM", NUM), ("Clearly trivial threshold", "=Mat_CT", NUM),
        ("Quantitative evaluation", "=MS_Eval", None), ("Qualitative evaluation", "=MS_Qual", None)]):
    cell(ws, f"B{r}", lab); c = cell(ws, f"C{r}", f_, nf, link=True)
    if nf is None: ws.merge_cells(f"C{r}:L{r}"); c.font = font(9, True, MD)
    r += 1
c = cell(ws, f"B{r}", '=HYPERLINK("#SUAD!A1","Open Summary of Unadjusted Differences >")'); c.font = Font(name=F, size=9, color="1565C0", underline="single")
r += 2

r = section(ws, r, "E2", "Control deficiencies evaluation", "ISA 265.7-11")
r = table_head(ws, r, ["Ref", "Deficiency", "Process", "Source", "Could result in material misstatement?", "Compensating control?",
                       "Classification", "Communicate to", "Form", "Date communicated", "Mgmt response", ""])
defs = [("No formal entity risk assessment process", "Entity level", "B2", "No", "No"),
        ("ERP user access reviews not performed in Q2 & Q3 (ITGC)", "IT", "IC4 C06", "Yes", "No"),
        ("Three-way match override without approval (1 of 30)", "Purchases", "IC4 C03", "No", "Yes"),
        ("No documented NRV / slow-moving inventory policy", "Inventory", "SUAD M02", "Yes", "No"),
        ("Accrual checklist not updated for new contracts", "Financial close", "SUAD M03", "No", "Yes")]
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
opin = [("Uncorrected misstatements material? (SUAD dual aggregate)", '=IF(ABS(MS_Agg)>=Mat_OM,"Yes","No")', "ISA 450.11 / 705.6(a)", False),
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

# ======================================================================
# NEW SHEETS: IC_Assessment, Sampling, SUAD  (inserted before WP INDEX)
# ======================================================================
import math, random

def extra_widths(ws, last_col, w=12):
    for k in range(ci("N"), ci(last_col) + 1):
        ws.column_dimensions[gl(k)].width = w

from openpyxl.utils import column_index_from_string as ci, get_column_letter as gl

# ---------------- SAMPLING ----------------
ws = S["Sampling"]
r = sheet_header(ws, "AUDIT SAMPLING - POPULATIONS, SAMPLE SIZES, SELECTION & EVALUATION", "3",
                 "Define each population, confirm its completeness, determine sample sizes (statistical & non-statistical), select items and evaluate results.",
                 "ISA 530 Audit Sampling; ISA 500.7-9 (reliability of information); ISA 330.10, 23; ISA 450.A9")
extra_widths(ws, "R", 13)
widths(ws, {"B": 30, "C": 26, "D": 26})

# S2 factor tables first in layout? keep logical order: S1 register needs RF table -> names resolve anyway.
r = section(ws, r, "S1", "Population register & sample size determination", "ISA 530.6-8 | ISA 500.9")
heads = ["Ref", "Test / assertion", "Population description", "Source & completeness check", "Recon. to TB?",
         "No. of items", "Population value", "Key items (100%) value", "Residual value", "Method", "Confidence",
         "Tolerable misstatement", "Expected misstatement", "Reliability factor", "Expansion factor",
         "Sample size", "Sampling interval", "Linked WP"]
r = table_head(ws, r, heads)
pops = [
    ("S1.1", "Trade receivables - existence (confirmations)", "Customer balances at 31-Dec-25", "Aged listing footed & agreed to GL",
     '=IF(ABS(G{r}-INDEX(TB_CY,MATCH("Trade receivables",TB_Acct,0)))<1,"Agreed","DIFFERENCE")', '=COUNT(Rec_Book)', '=SUM(Rec_Book)', 0, "MUS", "95%", "=Mat_PM", 5000, "S4/S5, D4"),
    ("S1.2", "Revenue - occurrence", "Sales invoices Jan-Dec 2025", "Sales register reconciled to GL revenue",
     '=IF(ABS(G{r}+INDEX(TB_CY,MATCH("Revenue",TB_Acct,0)))<1,"Agreed","DIFFERENCE")', 48120, '=-INDEX(TB_CY,MATCH("Revenue",TB_Acct,0))', 0, "MUS", "95%", "=Mat_PM", 10000, "D3"),
    ("S1.3", "PPE additions - existence & valuation", "Capital additions 2025", "Additions schedule agreed to PPE roll-forward",
     "Agreed", 64, 1150000, 640000, "MUS", "90%", "=Mat_PM", 3000, "D3"),
    ("S1.4", "Inventory - valuation (NRV)", "SKU listing at YE (stock sheets)", "Priced stock sheets agreed to GL & count",
     '=IF(ABS(G{r}-INDEX(TB_CY,MATCH("Inventory",TB_Acct,0)))<1,"Agreed","DIFFERENCE")', 3820, '=INDEX(TB_CY,MATCH("Inventory",TB_Acct,0))', 1200000, "MUS", "95%", "=Mat_PM", 8000, "D3"),
    ("S1.5", "Payables - completeness (post-YE payments)", "Bank payments 1 Jan - 28 Feb 2026", "Cash book agreed to bank statements",
     "Agreed", 420, 960000, 350000, "MUS", "90%", "=Mat_PM", 3000, "D3"),
    ("S1.6", "Payroll - occurrence", "Employees paid Dec-25", "Payroll register to GL & bureau report",
     "Agreed", 186, 210000, 0, "Non-statistical", "Moderate", "=Mat_PM", 0, "C8"),
    ("S1.7", "Journal entries - mgmt override", "All manual JEs 2025", "JE population reconciled to TB movements",
     "Agreed", 48320, 0, 0, "100% risk-criteria (not sampling)", "-", "-", "-", "D6"),
]
p0 = r
for (ref, test, desc, src, recon, nitems, pval, keyv, meth, conf, tm, em, wp) in pops:
    cell(ws, f"A{r}", ref, center=True); cell(ws, f"B{r}", test); cell(ws, f"C{r}", desc, is_input=True); cell(ws, f"D{r}", src, is_input=True)
    cell(ws, f"E{r}", recon.format(r=r) if recon.startswith("=") else recon, is_input=not recon.startswith("="), center=True)
    cell(ws, f"F{r}", nitems, "#,##0", is_input=not str(nitems).startswith("="), link=str(nitems).startswith("="), center=True)
    cell(ws, f"G{r}", pval, NUM, is_input=not str(pval).startswith("="), link=str(pval).startswith("="))
    cell(ws, f"H{r}", keyv, NUM, True)
    cell(ws, f"I{r}", f"=G{r}-H{r}", NUM)
    cell(ws, f"J{r}", meth, is_input=True, center=True); cell(ws, f"K{r}", conf, is_input=True, center=True)
    if meth == "MUS":
        cell(ws, f"L{r}", tm, NUM, link=True); cell(ws, f"M{r}", em, NUM, True)
        cell(ws, f"N{r}", f"=INDEX(RF_Tab,1,MATCH(K{r},RF_Conf,0))", "0.00", center=True)
        cell(ws, f"O{r}", f"=INDEX(EF_Tab,MATCH(K{r},RF_Conf,0))", "0.00", center=True)
        cell(ws, f"P{r}", f"=IFERROR(ROUNDUP(I{r}*N{r}/(L{r}-M{r}*O{r}),0),0)", "#,##0", center=True)
        cell(ws, f"Q{r}", f"=IFERROR(I{r}/P{r},0)", NUM)
    elif meth == "Non-statistical":
        cell(ws, f"L{r}", tm, NUM, link=True); cell(ws, f"M{r}", em, NUM, True)
        cell(ws, f"N{r}", "-", center=True); cell(ws, f"O{r}", "-", center=True)
        cell(ws, f"P{r}", 25, "#,##0", True, center=True); cell(ws, f"Q{r}", "-", center=True)
    else:
        for col in "LMNOQ": cell(ws, f"{col}{r}", "-", center=True)
        cell(ws, f"P{r}", 112, "#,##0", True, center=True)
    cell(ws, f"R{r}", wp, center=True)
    ws.row_dimensions[r].height = 32; r += 1
dv(ws, f"K{p0}:K{p0+4}", "Confidence")
dname("Pop_SampleSize", f"Sampling!$P${p0}:$P${r-1}")
ws.conditional_formatting.add(f"E{p0}:E{r-1}", CellIsRule(operator="equal", formula=['"DIFFERENCE"'], fill=fill("FFCDD2")))
put(ws, f"B{r}", "MUS sample size n = Residual population x RF / (TM - EM x EF). Key items (individually significant, or > interval) are tested 100% and excluded from the residual. "
                 "Population completeness must be tested before sampling (ISA 500.9 / ISA 530.A12).", font(8, i=True, c="616161"), al=WRAP)
ws.merge_cells(f"B{r}:R{r}"); ws.row_dimensions[r].height = 24
REC_ROW = p0
r += 2

r = section(ws, r, "S2", "Statistical factor tables & attribute sample size calculator (tests of controls)", "ISA 530.A11-A13, App. 2-3")
r = table_head(ws, r, ["", "Deviations / errors (k)", "95%", "90%", "80%"])
rf = [(0, 3.00, 2.31, 1.61), (1, 4.75, 3.89, 2.99), (2, 6.30, 5.33, 4.28), (3, 7.76, 6.69, 5.52), (4, 9.16, 8.00, 6.72), (5, 10.52, 9.28, 7.91)]
hdr_row = r - 1
# conf header as text used for MATCH
t0 = r
for k, a, b, c_ in rf:
    cell(ws, f"B{r}", k, center=True)
    for col, v in zip("CDE", (a, b, c_)): cell(ws, f"{col}{r}", v, "0.00", center=True)
    r += 1
dname("RF_Tab", f"Sampling!$C${t0}:$E${r-1}"); dname("RF_Conf", f"Sampling!$C${hdr_row}:$E${hdr_row}")
cell(ws, f"B{r}", "Expansion factor (EF)", center=True)
for col, v in zip("CDE", (1.6, 1.5, 1.3)): cell(ws, f"{col}{r}", v, "0.00", center=True)
dname("EF_Tab", f"Sampling!$C${r}:$E${r}")
r += 1
put(ws, f"B{r}", "Poisson upper-limit (reliability) factors. Confidence 95% = 5% risk of incorrect acceptance. Expansion factors approximate those in widely used audit sampling guides - align with your firm's methodology.",
    font(8, i=True, c="616161"), al=WRAP); ws.merge_cells(f"B{r}:J{r}"); ws.row_dimensions[r].height = 24
r += 2
put(ws, f"B{r}", "Attribute sample size calculator", font(10, True, DK)); r += 1
r = table_head(ws, r, ["", "Parameter", "Value", "Note"]); ws.merge_cells(f"D{r-1}:J{r-1}")
a0 = r
attr = [("Desired confidence", "95%", None, "Higher when placing high reliance on the control", True),
        ("Tolerable deviation rate (TDR)", 0.10, PCT, "5% high reliance, 10% moderate reliance", True),
        ("Deviations planned to be allowed (k)", 0, "0", "Usually 0 for key controls", True),
        ("Reliability factor for k", f"=INDEX(RF_Tab,C{a0+2}+1,MATCH(C{a0},RF_Conf,0))", "0.00", "From table above", False),
        ("SAMPLE SIZE  n = RF / TDR", f"=ROUNDUP(C{a0+3}/C{a0+1},0)", "0", "Population > 2,000: no finite correction needed", False)]
for i, (lab, v, nf, note, isin) in enumerate(attr):
    b = cell(ws, f"B{r}", lab); c = cell(ws, f"C{r}", v, nf, isin, center=True)
    if i == 4: b.font = font(9, True); c.font = font(11, True, DK); b.fill = fill(PALE); c.fill = fill(PALE)
    cell(ws, f"D{r}", note); ws.merge_cells(f"D{r}:J{r}"); r += 1
dv(ws, f"C{a0}", "Confidence")
r += 1

r = section(ws, r, "S3", "Non-statistical sample sizes for low-frequency controls", "ISA 530.A13 | firm methodology")
r = table_head(ws, r, ["", "Control frequency", "Population (occurrences / yr)", "Min sample", "Max sample"])
f0 = r
for fq, pop_, mn, mx in [("Annual", 1, 1, 1), ("Quarterly", 4, 2, 2), ("Monthly", 12, 2, 5), ("Weekly", 52, 5, 15),
                         ("Daily", 250, 20, 40), ("Multiple daily", ">250", 25, 60)]:
    cell(ws, f"B{r}", fq); cell(ws, f"C{r}", pop_, center=True); cell(ws, f"D{r}", mn, center=True); cell(ws, f"E{r}", mx, center=True); r += 1
dname("Freq_Name", f"Sampling!$B${f0}:$B${r-1}"); dname("Freq_Min", f"Sampling!$D${f0}:$D${r-1}")
put(ws, f"B{r}", "Daily / multiple-daily controls use the statistical attribute calculator (S2) in IC4; lower-frequency controls use this table.", font(8, i=True, c="616161"))
r += 2

# S4 MUS selection - receivables ledger
r = section(ws, r, "S4", "MUS sample selection - trade receivables ledger (systematic, random start)", "ISA 530.8 & A13 | ISA 505")
cfg0 = r
cfg = [("Population value (S1.1)", f"=G{REC_ROW}", NUM), ("Sample size n (S1.1)", f"=P{REC_ROW}", "0"),
       ("Sampling interval J", f"=IFERROR(C{cfg0}/C{cfg0+1},0)", NUM), ("Random start (1 to J)", 7350, NUM),
       ("Hits generated (sample units)", None, "0"), ("Distinct items selected", None, "0"), ("Key items (book >= J)", None, "0"),
       ("Coverage of population value", None, PCT)]
for i, (lab, v, nf) in enumerate(cfg):
    cell(ws, f"B{r+i}", lab); cell(ws, f"C{r+i}", v, nf, is_input=(i == 3), link=(i < 2), center=True)
dname("Rec_J", f"Sampling!$C${cfg0+2}"); dname("Rec_Start", f"Sampling!$C${cfg0+3}")
put(ws, f"E{cfg0}", "How it works: cumulate book values; an item is selected each time the cumulative total passes start + m x J. "
                    "Items >= J are always selected (key items). Enter confirmed/audited values for selected items only.",
    font(8, i=True, c="616161"), al=WRAP); ws.merge_cells(f"E{cfg0}:K{cfg0+3}")
r += len(cfg) + 1
r = table_head(ws, r, ["#", "Customer", "Book value", "Cumulative", "Hits", "Selected?", "Key item?", "Audited / confirmed value",
                       "Misstatement", "Tainting (non-key)", "Confirmation ref"])
# generate ledger
rng = random.Random(2026)
raw = [math.exp(rng.gauss(9.2, 1.0)) for _ in range(200)]
sc = 3480000 / sum(raw)
bal = [max(500, int(round(x * sc, -1))) for x in raw]
bal[-1] += 3480000 - sum(bal)
# replicate selection in python
n_py = math.ceil(3480000 * 3.0 / (70000 - 5000 * 1.6)); J_py = 3480000 / n_py; start = 7350
cum, prev, sel_idx = 0, 0, []
for i, b in enumerate(bal):
    cum += b
    if math.floor((cum - start) / J_py) - math.floor((prev - start) / J_py) > 0: sel_idx.append(i)
    prev = cum
small_sel = [i for i in sel_idx if 10000 < bal[i] < J_py * 0.8]
err = {small_sel[3]: 0.08, small_sel[11]: 0.06}
l0 = r
for i, b in enumerate(bal):
    cell(ws, f"A{r}", i + 1, center=True); cell(ws, f"B{r}", f"CUST-{i+1:03d}")
    cell(ws, f"C{r}", b, NUM, True)
    cell(ws, f"D{r}", f"=C{r}" if i == 0 else f"=D{r-1}+C{r}", NUM)
    prevc = "0" if i == 0 else f"D{r-1}"
    cell(ws, f"E{r}", f"=INT((D{r}-Rec_Start)/Rec_J)-INT(({prevc}-Rec_Start)/Rec_J)", "0", center=True)
    cell(ws, f"F{r}", f'=IF(E{r}>0,"Yes","")', center=True)
    cell(ws, f"G{r}", f'=IF(C{r}>=Rec_J,"Yes","No")', center=True)
    av = (round(b * (1 - err[i]), -1) if i in err else b) if i in sel_idx else None
    cell(ws, f"H{r}", av, NUM, True)
    cell(ws, f"I{r}", f'=IF(AND(F{r}="Yes",H{r}<>""),C{r}-H{r},0)', NUM)
    cell(ws, f"J{r}", f'=IF(AND(F{r}="Yes",G{r}="No",I{r}<>0),I{r}/C{r},"")', PCT, center=True)
    cell(ws, f"K{r}", f'=IF(F{r}="Yes","D4.R"&TEXT(A{r},"000"),"")', center=True)
    ws.row_dimensions[r].height = 13
    r += 1
l1 = r - 1
dname("Rec_Book", f"Sampling!$C${l0}:$C${l1}"); dname("Rec_Mis", f"Sampling!$I${l0}:$I${l1}")
dname("Rec_Key", f"Sampling!$G${l0}:$G${l1}"); dname("Rec_Taint", f"Sampling!$J${l0}:$J${l1}")
ws[f"C{cfg0+4}"] = f"=SUM(E{l0}:E{l1})"; ws[f"C{cfg0+5}"] = f'=COUNTIF(F{l0}:F{l1},"Yes")'
ws[f"C{cfg0+6}"] = f'=COUNTIF(G{l0}:G{l1},"Yes")'
ws[f"C{cfg0+7}"] = f'=IFERROR(SUMIFS(C{l0}:C{l1},F{l0}:F{l1},"Yes")/C{cfg0},0)'
dname("Rec_Hits", f"Sampling!$C${cfg0+4}"); dname("Rec_Items", f"Sampling!$C${cfg0+5}"); dname("Rec_Cov", f"Sampling!$C${cfg0+7}")
ws.conditional_formatting.add(f"A{l0}:K{l1}", FormulaRule(formula=[f'$F{l0}="Yes"'], fill=fill("E8F5E9")))
ws.conditional_formatting.add(f"I{l0}:I{l1}", CellIsRule(operator="notEqual", formula=["0"], fill=fill("FFCDD2"), font=Font(name=F, bold=True, color="B71C1C")))
put(ws, f"B{r}", "Ledger total", font(9, True), PALE, bd=BOX); put(ws, f"C{r}", f"=SUM(C{l0}:C{l1})", font(9, True), PALE, NUM, bd=BOX)
put(ws, f"H{r}", "Known misst.", font(9, True), PALE, bd=BOX); put(ws, f"I{r}", f"=SUM(I{l0}:I{l1})", font(9, True), PALE, NUM, bd=BOX)
r += 2

r = section(ws, r, "S5", "MUS evaluation - projection & upper misstatement limit", "ISA 530.14-15 | ISA 450.A9")
e0 = r
ev = [("Confidence level (from S1.1)", f"=K{REC_ROW}", None), ("Tolerable misstatement (S1.1)", f"=L{REC_ROW}", NUM),
      ("Sampling interval J", "=Rec_J", NUM), ("Known misstatements in sample (all)", "=SUM(Rec_Mis)", NUM),
      ("of which: key items (actual, not projected)", '=SUMIFS(Rec_Mis,Rec_Key,"Yes")', NUM),
      ("Basic precision = RF(0) x J", f"=INDEX(RF_Tab,1,MATCH(C{e0},RF_Conf,0))*C{e0+2}", NUM)]
for i, (lab, v, nf) in enumerate(ev):
    cell(ws, f"B{r+i}", lab); cell(ws, f"C{r+i}", v, nf, link=(i < 3), center=True)
r += len(ev) + 1
r = table_head(ws, r, ["Rank", "Tainting % (ranked)", "Projected misstatement (t x J)", "Incremental factor", "Allowance (proj x factor)"])
k0 = r
for k in range(1, 6):
    cell(ws, f"A{r}", k, center=True)
    cell(ws, f"B{r}", f"=IFERROR(LARGE(Rec_Taint,{k}),0)", PCT, center=True)
    cell(ws, f"C{r}", f"=B{r}*$C${e0+2}", NUM)
    cell(ws, f"D{r}", f"=INDEX(RF_Tab,{k+1},MATCH($C${e0},RF_Conf,0))-INDEX(RF_Tab,{k},MATCH($C${e0},RF_Conf,0))", "0.00", center=True)
    cell(ws, f"E{r}", f"=C{r}*D{r}", NUM)
    r += 1
fin = [("PROJECTED MISSTATEMENT (most likely)", f"=SUM(C{k0}:C{r-1})+C{e0+4}", NUM),
       ("UPPER MISSTATEMENT LIMIT (UML)", f"=C{e0+5}+SUM(E{k0}:E{r-1})+C{e0+4}", NUM),
       ("UML as % of tolerable misstatement", f"=IFERROR(C{r+1}/C{e0+1},0)", PCT),
       ("CONCLUSION", f'=IF(C{r+1}<=C{e0+1},"UML <= TM: population accepted - record projected misstatement in SUAD",'
                      f'"UML > TM: population NOT accepted - extend testing, request adjustment or modify approach (ISA 530.A23)")', None)]
for i, (lab, v, nf) in enumerate(fin):
    b = put(ws, f"B{r+i}", lab, font(9, True), PALE, bd=BOX); c = put(ws, f"C{r+i}", v, font(10, True, DK), PALE, nf, bd=BOX)
    if i == 3: ws.merge_cells(f"C{r+i}:K{r+i}")
dname("Rec_Proj", f"Sampling!$C${r}"); dname("Rec_UML", f"Sampling!$C${r+1}"); dname("Rec_UMLpct", f"Sampling!$C${r+2}")
r += len(fin) + 1
put(ws, f"B{r}", "Taintings are ranked highest first and multiplied by incremental reliability factors (cell method). Key-item misstatements are added at actual value.", font(8, i=True, c="616161"))

# ---------------- IC ASSESSMENT ----------------
ws = S["IC_Assessment"]
r = sheet_header(ws, "INTERNAL CONTROL ASSESSMENT & TESTING", "2 & 4",
                 "Evaluate the system of internal control (COSO components), design & implementation of controls, IT general controls, and test operating effectiveness of controls relied upon.",
                 "ISA 315 (Revised 2019).21-26, A90-A175; ISA 330.8-17; ISA 265; ISA 530 (attribute sampling); COSO 2013 framework")
extra_widths(ws, "Q", 13)

r = section(ws, r, "IC1", "Entity-level controls - COSO 2013 components & 17 principles", "ISA 315R.21-25 | COSO 2013")
r = table_head(ws, r, ["#", "Principle", "Component", "Assessment", "Score", "Evidence / basis", "", "", "", "", "Deficiency ref"])
ws.merge_cells(f"F{r-1}:J{r-1}")
coso = [("Demonstrates commitment to integrity & ethical values", "Control environment", "Effective", "Code of conduct signed annually"),
        ("Board exercises oversight responsibility", "Control environment", "Needs improvement", "Board meets quarterly; no audit committee"),
        ("Establishes structure, authority & responsibility", "Control environment", "Effective", "Delegation of authority matrix approved 2024"),
        ("Demonstrates commitment to competence", "Control environment", "Effective", "Qualified CFO & chief accountant"),
        ("Enforces accountability", "Control environment", "Effective", "KPIs & performance reviews"),
        ("Specifies suitable objectives", "Risk assessment", "Needs improvement", "Objectives set informally"),
        ("Identifies & analyses risk", "Risk assessment", "Ineffective", "No formal risk register"),
        ("Assesses fraud risk", "Risk assessment", "Needs improvement", "No fraud risk assessment by management"),
        ("Identifies & analyses significant change", "Risk assessment", "Needs improvement", "FX exposure not formally assessed"),
        ("Selects & develops control activities", "Control activities", "Effective", "Documented procedures for key cycles"),
        ("Selects & develops general controls over technology", "Control activities", "Needs improvement", "Access reviews lapsed (IC3)"),
        ("Deploys through policies & procedures", "Control activities", "Needs improvement", "No NRV policy"),
        ("Uses relevant, quality information", "Information & communication", "Effective", "ERP single source of truth"),
        ("Communicates internally", "Information & communication", "Effective", "Monthly management meetings"),
        ("Communicates externally", "Information & communication", "Effective", "Customer & supplier statements"),
        ("Conducts ongoing / separate evaluations", "Monitoring activities", "Needs improvement", "No internal audit function"),
        ("Evaluates & communicates deficiencies", "Monitoring activities", "Effective", "PY management letter points tracked")]
c0 = r
for i, (p, comp, rating, ev_) in enumerate(coso, 1):
    cell(ws, f"A{r}", i, center=True); cell(ws, f"B{r}", p); cell(ws, f"C{r}", comp)
    cell(ws, f"D{r}", rating, is_input=True, center=True)
    cell(ws, f"E{r}", f'=IF(D{r}="Effective",1,IF(D{r}="Needs improvement",2,IF(D{r}="Ineffective",3,"")))', "0", center=True)
    cell(ws, f"F{r}", ev_, is_input=True); ws.merge_cells(f"F{r}:J{r}")
    cell(ws, f"K{r}", f'=IF(E{r}=3,"CD01",IF(E{r}=2,"Mgmt letter","-"))', center=True)
    r += 1
c1 = r - 1
dv(ws, f"D{c0}:D{c1}", "ICRating")
for k, v in (("Effective", "C8E6C9"), ("Needs improvement", "FFE0B2"), ("Ineffective", "FFCDD2")):
    ws.conditional_formatting.add(f"D{c0}:D{c1}", CellIsRule(operator="equal", formula=[f'"{k}"'], fill=fill(v)))
r += 1
r = table_head(ws, r, ["", "COSO component", "Avg score", "Principles", "Conclusion"]); ws.merge_cells(f"E{r-1}:H{r-1}")
cs0 = r
for comp in ["Control environment", "Risk assessment", "Control activities", "Information & communication", "Monitoring activities"]:
    cell(ws, f"B{r}", comp)
    cell(ws, f"C{r}", f'=IFERROR(AVERAGEIF(C{c0}:C{c1},B{r},E{c0}:E{c1}),0)', "0.00", center=True)
    cell(ws, f"D{r}", f'=COUNTIF(C{c0}:C{c1},B{r})', "0", center=True)
    cell(ws, f"E{r}", f'=IF(C{r}<=1.4,"Present & functioning",IF(C{r}<=2.2,"Present - improvements needed","Not functioning - significant deficiency?"))')
    ws.merge_cells(f"E{r}:H{r}"); r += 1
put(ws, f"B{r}", "OVERALL ENTITY-LEVEL CONTROL ENVIRONMENT", font(9, True), PALE, bd=BOX)
put(ws, f"C{r}", f"=AVERAGE(C{cs0}:C{r-1})", font(10, True, DK), PALE, "0.00", CEN, BOX)
put(ws, f"E{r}", f'=IF(C{r}<=1.4,"Effective - supports controls reliance",IF(C{r}<=2.2,"Generally effective - reliance possible on tested process controls","Ineffective - substantive approach"))',
    font(9, True, DK), PALE, bd=BOX); ws.merge_cells(f"E{r}:J{r}")
dname("IC_COSO", f"IC_Assessment!$E${r}")
put(ws, f"B{r+1}", "Score: Effective = 1, Needs improvement = 2, Ineffective = 3.", font(8, i=True, c="616161"))
r += 3

r = section(ws, r, "IC2", "Risk & control matrix (RCM) - design & implementation", "ISA 315R.26 & A175 | ISA 330.8")
heads = ["ID", "Control description", "Process", "Risk addressed (B4)", "Assertion", "Type", "Nature", "Frequency",
         "Key control?", "Walkthrough done?", "Design effective?", "Implemented?", "D&I conclusion", "Rely on control?", "ToE result (IC4)", "Control risk", "Owner"]
r = table_head(ws, r, heads)
rcm = [("C01", "Sales orders above credit limit approved by credit manager", "Revenue", "R01, R05", "Occurrence, Valuation", "Preventive", "IT-dependent manual", "Multiple daily", "Yes", "Yes", "Yes", "Yes", "Yes", "Credit Manager"),
       ("C02", "Monthly bank reconciliations prepared & reviewed", "Cash", "R10", "Existence", "Detective", "Manual", "Monthly", "Yes", "Yes", "Yes", "Yes", "Yes", "Chief Accountant"),
       ("C03", "Three-way match (PO/GRN/invoice) before payment", "Purchases", "R06", "Occurrence, Accuracy", "Preventive", "Automated", "Multiple daily", "Yes", "Yes", "Yes", "Yes", "Yes", "AP Supervisor"),
       ("C04", "Payroll master-file changes authorised by HR manager", "Payroll", "R02", "Occurrence", "Preventive", "Manual", "Monthly", "Yes", "Yes", "Yes", "Yes", "Yes", "HR Manager"),
       ("C05", "Quarterly inventory cycle counts & variance review", "Inventory", "R04", "Existence", "Detective", "Manual", "Quarterly", "Yes", "Yes", "Yes", "Yes", "Yes", "Warehouse Mgr"),
       ("C06", "Quarterly ERP user access reviews (ITGC)", "IT", "R02", "All", "Detective", "Manual", "Quarterly", "Yes", "Yes", "Yes", "Yes", "Yes", "IT Manager"),
       ("C07", "Program changes approved & tested before release (ITGC)", "IT", "R02", "All", "Preventive", "Manual", "Monthly", "Yes", "Yes", "Yes", "Yes", "Yes", "IT Manager"),
       ("C08", "Manual JEs above MNT 50m approved by CFO", "Financial close", "R02", "Occurrence, Accuracy", "Preventive", "Manual", "Weekly", "Yes", "Yes", "Yes", "Yes", "Yes", "CFO"),
       ("C09", "Price master changes restricted & approved", "Revenue", "R01", "Accuracy", "Preventive", "Automated", "Daily", "No", "Yes", "Yes", "Yes", "No", "Sales Director"),
       ("C10", "Monthly inventory ageing review for slow-moving / NRV", "Inventory", "R03", "Valuation", "Detective", "Manual", "Monthly", "Yes", "Yes", "No", "No", "No", "CFO"),
       ("C11", "Monthly management accounts reviewed by CEO (entity-level)", "Financial close", "R02", "All", "Detective", "Manual", "Monthly", "No", "Yes", "Yes", "Yes", "No", "CEO"),
       ("C12", "Segregation: cash receipting vs receivables posting", "Revenue", "R05", "Existence", "Preventive", "Manual", "Daily", "Yes", "Yes", "Yes", "Yes", "No", "CFO")]
m0 = r
for row in rcm:
    vals = list(row)
    for j, v in enumerate(vals[:12]):
        col = chr(65 + j)
        cell(ws, f"{col}{r}", v, is_input=j >= 4, center=j >= 2)
    cell(ws, f"M{r}", f'=IF(AND(K{r}="Yes",L{r}="Yes"),"D&I effective","D&I DEFICIENT - no reliance (ISA 265)")', center=True)
    cell(ws, f"N{r}", vals[12], is_input=True, center=True)
    cell(ws, f"O{r}", f'=IF(N{r}<>"Yes","Not tested (D&I only)",IFERROR(INDEX(ToE_Result,MATCH(A{r},ToE_ID,0)),"Not yet tested"))', link=True, center=True)
    cell(ws, f"P{r}", f'=IF(M{r}<>"D&I effective","High - control deficient",IF(N{r}<>"Yes","High - not relied upon",IF(O{r}="Effective","Lower - rely; reduce substantive","High - reliance failed; extend substantive")))', center=True)
    cell(ws, f"Q{r}", vals[13], center=True)
    ws.row_dimensions[r].height = 30; r += 1
m1 = r - 1
dv(ws, f"I{m0}:L{m1}", "YesNo"); dv(ws, f"N{m0}:N{m1}", "YesNo"); dv(ws, f"F{m0}:F{m1}", "CtlType"); dv(ws, f"H{m0}:H{m1}", "Freq")
dname("RCM_DI", f"IC_Assessment!$M${m0}:$M${m1}"); dname("RCM_CR", f"IC_Assessment!$P${m0}:$P${m1}")
ws.conditional_formatting.add(f"M{m0}:M{m1}", FormulaRule(formula=[f'LEFT(M{m0},12)="D&I DEFICIEN"'], fill=fill("FFCDD2")))
ws.conditional_formatting.add(f"P{m0}:P{m1}", FormulaRule(formula=[f'LEFT(P{m0},5)="Lower"'], fill=fill("C8E6C9")))
ws.conditional_formatting.add(f"O{m0}:O{m1}", CellIsRule(operator="equal", formula=['"Not effective"'], fill=fill("FFCDD2")))
put(ws, f"B{r}", "D&I is evaluated for all controls relevant to the audit (ISA 315R.26(d)) via inquiry + observation/inspection (walkthrough); inquiry alone is insufficient. Only controls with effective D&I may be tested for operating effectiveness.",
    font(8, i=True, c="616161"), al=WRAP); ws.merge_cells(f"B{r}:Q{r}"); ws.row_dimensions[r].height = 24
r += 2

r = section(ws, r, "IC3", "IT general controls (ITGC) assessment", "ISA 315R.26(c) & App. 5-6")
r = table_head(ws, r, ["#", "ITGC domain", "Key controls / procedures", "", "", "Controls tested", "Deviations", "Rating", "Impact on audit"])
ws.merge_cells(f"C{r-1}:E{r-1}"); ws.merge_cells(f"I{r-1}:L{r-1}")
itgc = [("Access to programs & data", "User provisioning/termination, periodic access review (C06), privileged access, password parameters", 4, 1),
        ("Program changes", "Change requests approved, tested & migrated by separate person (C07)", 3, 0),
        ("Program development / acquisition", "No new systems in 2025 - ERP stable since 2021", 1, 0),
        ("Computer operations", "Daily backups, restore tests, batch job monitoring, incident management", 3, 0)]
g0 = r
for i, (dom, ctl, nt, dev) in enumerate(itgc, 1):
    cell(ws, f"A{r}", i, center=True); cell(ws, f"B{r}", dom); cell(ws, f"C{r}", ctl, is_input=True); ws.merge_cells(f"C{r}:E{r}")
    cell(ws, f"F{r}", nt, "0", True, center=True); cell(ws, f"G{r}", dev, "0", True, center=True)
    cell(ws, f"H{r}", f'=IF(F{r}=0,"Not tested",IF(G{r}=0,"Effective","Deficient"))', center=True)
    cell(ws, f"I{r}", f'=IF(H{r}="Deficient","Assess impact on automated controls & system reports (IPE); test compensating controls or extend substantive","No adverse impact on reliance")')
    ws.merge_cells(f"I{r}:L{r}"); ws.row_dimensions[r].height = 30; r += 1
put(ws, f"B{r}", "OVERALL ITGC CONCLUSION", font(9, True), PALE, bd=BOX)
put(ws, f"C{r}", f'=IF(COUNTIF(H{g0}:H{r-1},"Deficient")>0,"ITGCs DEFICIENT in "&COUNTIF(H{g0}:H{r-1},"Deficient")&" domain(s) - automated controls (C03) & IPE need additional evidence","ITGCs effective - automated controls may be relied upon with benchmarking")',
    font(9, True, DK), PALE, bd=BOX); ws.merge_cells(f"C{r}:L{r}")
dname("IC_ITGC", f"IC_Assessment!$C${r}")
ws.conditional_formatting.add(f"H{g0}:H{r-1}", CellIsRule(operator="equal", formula=['"Deficient"'], fill=fill("FFCDD2")))
r += 2

r = section(ws, r, "IC4", "Tests of operating effectiveness - attribute sampling", "ISA 330.8-11, 16-17 | ISA 530")
heads = ["ID", "Control", "Frequency", "Population (occurrences)", "Statistical?", "Confidence", "Tolerable deviation rate",
         "Allowed deviations (k)", "Required sample", "Items tested", "Deviations found", "Sample deviation rate",
         "Upper deviation limit", "Result", "Period covered", "Tested by", "Ref"]
r = table_head(ws, r, heads)
toe = [("C01", "Credit-limit override approval", "Multiple daily", 1186, "=Att_Tested", "=Att_Dev", "Jan-Dec 25", "T. Enkhjin", "IC5"),
       ("C02", "Bank reconciliations", "Monthly", 12, 3, 0, "Mar, Aug, Dec 25", "M. Saruul", "IC4.2"),
       ("C03", "Three-way match", "Multiple daily", 9840, 30, 1, "Jan-Dec 25", "T. Enkhjin", "IC4.3"),
       ("C04", "Payroll changes authorised", "Monthly", 12, 3, 0, "Feb, Jul, Nov 25", "M. Saruul", "IC4.4"),
       ("C05", "Inventory cycle counts", "Quarterly", 4, 2, 0, "Q2, Q4 25", "T. Enkhjin", "IC4.5"),
       ("C06", "ERP user access reviews", "Quarterly", 4, 2, 1, "Q2, Q3 25", "IT auditor", "IC4.6"),
       ("C07", "Program change approval", "Monthly", 12, 3, 0, "Apr, Sep, Nov 25", "IT auditor", "IC4.7"),
       ("C08", "Manual JE approval > MNT 50m", "Weekly", 52, 10, 0, "Jan-Dec 25", "M. Saruul", "IC4.8")]
t0 = r
for (cid, ctl, fq, pop_, tested, dev, per, by, ref) in toe:
    cell(ws, f"A{r}", cid, center=True); cell(ws, f"B{r}", ctl); cell(ws, f"C{r}", fq, is_input=True, center=True)
    cell(ws, f"D{r}", pop_, "#,##0", True, center=True)
    cell(ws, f"E{r}", f'=IF(OR(C{r}="Daily",C{r}="Multiple daily"),"Yes","No")', center=True)
    cell(ws, f"F{r}", "95%", is_input=True, center=True); cell(ws, f"G{r}", 0.10, PCT, True, center=True)
    cell(ws, f"H{r}", 0, "0", True, center=True)
    cell(ws, f"I{r}", f'=IF(E{r}="Yes",MIN(D{r},ROUNDUP(INDEX(RF_Tab,H{r}+1,MATCH(F{r},RF_Conf,0))/G{r},0)),MIN(D{r},INDEX(Freq_Min,MATCH(C{r},Freq_Name,0))))', "0", link=True, center=True)
    cell(ws, f"J{r}", tested, "0", is_input=not str(tested).startswith("="), link=str(tested).startswith("="), center=True)
    cell(ws, f"K{r}", dev, "0", is_input=not str(dev).startswith("="), link=str(dev).startswith("="), center=True)
    cell(ws, f"L{r}", f"=IFERROR(K{r}/J{r},0)", PCT, center=True)
    cell(ws, f"M{r}", f'=IF(E{r}="Yes",IFERROR(INDEX(RF_Tab,MIN(K{r},5)+1,MATCH(F{r},RF_Conf,0))/J{r},0),"n/a")', PCT, center=True)
    cell(ws, f"N{r}", f'=IF(J{r}<I{r},"Sample too small",IF(E{r}="Yes",IF(M{r}<=G{r},"Effective","Not effective"),IF(K{r}=0,"Effective","Not effective")))', center=True)
    cell(ws, f"O{r}", per, is_input=True, center=True); cell(ws, f"P{r}", by, center=True); cell(ws, f"Q{r}", ref, center=True)
    r += 1
t1 = r - 1
dv(ws, f"C{t0}:C{t1}", "Freq"); dv(ws, f"F{t0}:F{t1}", "Confidence")
dname("ToE_ID", f"IC_Assessment!$A${t0}:$A${t1}"); dname("ToE_Result", f"IC_Assessment!$N${t0}:$N${t1}")
dname("Ctl_Result", f"IC_Assessment!$N${t0}:$N${t1}")
dname("ToE_Tested", f"IC_Assessment!$J${t0}:$J${t1}"); dname("ToE_Dev", f"IC_Assessment!$K${t0}:$K${t1}")
ws.conditional_formatting.add(f"N{t0}:N{t1}", CellIsRule(operator="equal", formula=['"Not effective"'], fill=fill("FFCDD2"), font=Font(name=F, bold=True, color="B71C1C")))
ws.conditional_formatting.add(f"N{t0}:N{t1}", CellIsRule(operator="equal", formula=['"Effective"'], fill=fill("C8E6C9")))
put(ws, f"B{r}", "Upper deviation limit = RF(deviations) / items tested. Control is effective when UDL <= tolerable deviation rate (statistical) or no deviations (non-statistical). "
                 "Controls tested at interim need roll-forward evidence to year end (ISA 330.12).", font(8, i=True, c="616161"), al=WRAP)
ws.merge_cells(f"B{r}:Q{r}"); ws.row_dimensions[r].height = 24
r += 2

r = section(ws, r, "IC5", "Attribute testing detail - C01 credit-limit override approval", "ISA 330.10 | ISA 530.8-9")
put(ws, f"B{r}", "Population: 1,186 sales orders exceeding customer credit limit (ERP report RPT-SO-117, completeness tested by re-running query). Selection: random number generator (seed 2026).",
    font(8, i=True, c="616161"), al=WRAP); ws.merge_cells(f"B{r}:Q{r}"); ws.row_dimensions[r].height = 22; r += 1
r = table_head(ws, r, ["#", "Sales order no.", "Date", "Customer", "Order value", "Credit limit", "Excess over limit",
                       "A: Approval evidenced", "B: Approver per DOA", "C: Approved before dispatch", "D: Credit review documented", "Result", "Comment"])
rng2 = random.Random(7)
d_start = D(2025, 1, 2)
x0 = r
for i in range(30):
    dte = d_start + dt.timedelta(days=int(i * 12 + rng2.randint(0, 8)))
    lim = rng2.choice([20000, 30000, 50000, 80000])
    val = int(lim * (1 + rng2.uniform(0.05, 0.6)) // 10 * 10)
    cell(ws, f"A{r}", i + 1, center=True); cell(ws, f"B{r}", f"SO-25-{rng2.randint(10000, 99999)}", center=True)
    cell(ws, f"C{r}", dte, DATE, True, center=True); cell(ws, f"D{r}", f"CUST-{rng2.randint(1, 200):03d}", center=True)
    cell(ws, f"E{r}", val, NUM, True); cell(ws, f"F{r}", lim, NUM, True); cell(ws, f"G{r}", f"=E{r}-F{r}", NUM)
    for col in "HIJK": cell(ws, f"{col}{r}", "Yes", is_input=True, center=True)
    cell(ws, f"L{r}", f'=IF(COUNTA(H{r}:K{r})<4,"Incomplete",IF(COUNTIF(H{r}:K{r},"No")>0,"Deviation","OK"))', center=True)
    cell(ws, f"M{r}", "", is_input=True)
    ws.row_dimensions[r].height = 14; r += 1
x1 = r - 1
dv(ws, f"H{x0}:K{x1}", "YesNo")
ws.conditional_formatting.add(f"L{x0}:L{x1}", CellIsRule(operator="equal", formula=['"Deviation"'], fill=fill("FFCDD2")))
put(ws, f"B{r}", "Items tested", font(9, True), PALE, bd=BOX); put(ws, f"C{r}", f'=COUNTIF(L{x0}:L{x1},"<>Incomplete")', font(9, True), PALE, "0", CEN, BOX)
put(ws, f"D{r}", "Deviations", font(9, True), PALE, bd=BOX); put(ws, f"E{r}", f'=COUNTIF(L{x0}:L{x1},"Deviation")', font(9, True), PALE, "0", CEN, BOX)
dname("Att_Tested", f"IC_Assessment!$C${r}"); dname("Att_Dev", f"IC_Assessment!$E${r}")
r += 2

r = section(ws, r, "IC6", "Evaluation of deviations & effect on audit approach", "ISA 330.16-17 | ISA 265")
r = checklist(ws, r, [
    ("Specific inquiries made to understand each deviation and its potential consequences", "ISA 330.17", "Yes", "C03: AP clerk override during system outage; C06: IT manager vacancy"),
    ("Determined whether deviations are isolated or systematic (not assumed to be anomalies)", "ISA 330.17; ISA 530.13", "Yes", "C03 systematic during outage period"),
    ("Controls with UDL > TDR or deviations not relied upon; control risk reassessed in RCM (IC2)", "ISA 330.17(b)", "Yes", "C03 & C06 reliance withdrawn"),
    ("Nature, timing & extent of substantive procedures revised for failed controls", "ISA 330.17(b); ISA 315R.37", "Yes", "Purchases: extended unrecorded liabilities search; IT: substantive IPE testing"),
    ("Deviations evaluated as control deficiencies & logged in E2 / E3", "ISA 265.7-8", "Yes", "CD02, CD03"),
    ("Interim results rolled forward to year end", "ISA 330.12", "Yes", "Inquiry + re-performance Dec-25 for C01, C02, C08"),
], "IC6")
r = table_head(ws, r, ["", "Summary", "Value"])
icsum = [("Controls in RCM", f"=COUNTA(A{m0}:A{m1})"), ("D&I effective", '=COUNTIF(RCM_DI,"D&I effective")'),
         ("Controls tested for operating effectiveness", "=COUNTA(ToE_ID)"), ("Operating effectively", '=COUNTIF(ToE_Result,"Effective")'),
         ("Controls relied upon (lower control risk)", '=COUNTIF(RCM_CR,"Lower*")'), ("Total items tested", "=SUM(ToE_Tested)"), ("Total deviations", "=SUM(ToE_Dev)")]
for i, (lab, f_) in enumerate(icsum):
    cell(ws, f"B{r+i}", lab); cell(ws, f"C{r+i}", f_, "0", center=True)
dname("IC_DIeff", f"IC_Assessment!$C${r+1}"); dname("IC_RCMn", f"IC_Assessment!$C${r}")
dname("IC_ItemsTested", f"IC_Assessment!$C${r+5}"); dname("IC_Rely", f"IC_Assessment!$C${r+4}")

# ---------------- SUAD ----------------
ws = S["SUAD"]
r = sheet_header(ws, "SUMMARY OF UNADJUSTED DIFFERENCES (ALL MISSTATEMENTS)", "5",
                 "Record every misstatement identified (factual, judgmental, projected; corrected and uncorrected), show the effect on each FS element, and evaluate uncorrected misstatements individually and in aggregate incl. prior-period effects.",
                 "ISA 450.5-14 & A1-A28; ISA 320.12-13; ISA 530.14; ISA 260.16(a); ISA 580 (representation on uncorrected misstatements)")
widths(ws, {"B": 40, "C": 9, "D": 11, "E": 13, "F": 20, "G": 20, "H": 13, "I": 10, "J": 10, "K": 12, "L": 12, "M": 11})
extra_widths(ws, "U", 12)
ws.column_dimensions["S"].width = 26; ws.column_dimensions["T"].width = 26; ws.column_dimensions["U"].width = 26
put(ws, "K2", "Tax rate (CIT)", font(9, True), PALE, bd=BOX); ws.unmerge_cells("K2:L4")
ws["K3"].value = None
c = put(ws, "L2", 0.10, font(9, c="0000FF"), INP, PCT, bd=BOX); dname("TaxRate", "SUAD!$L$2")
put(ws, "K3", "Mongolia CIT 10% (taxable income up to MNT 6bn) - input", font(7, i=True, c="616161"), al=WRAP); ws.merge_cells("K3:L4")

r = section(ws, r, "SUAD1", "Current-year misstatements - schedule of all identified differences (MNT '000)", "ISA 450.5-8")
heads = ["Ref", "Description", "Source WP", "Type", "Nature", "Debit account", "Credit account", "Amount", "Corrected?",
         "Accumulate (> CT)?", "Assets Dr/(Cr)", "Liabilities incr/(decr)", "Equity incr/(decr)", "Income incr/(decr)",
         "Expenses incr/(decr)", "PBT effect", "Tax effect", "Net profit effect", "Root cause", "Qualitative factors (450.A21-A22)", "Mgmt reason if not corrected"]
r = table_head(ws, r, heads)
def has(acct_cell, cls_): return f'(COUNTIFS(TB_Acct,{acct_cell},TB_Class,"{cls_}")>0)'
suad = [("Revenue recorded for goods dispatched after year end (cut-off)", "D3.1", "Factual", "Overstatement", "Revenue", "Trade receivables", 45000, "No", "Period-end target pressure", "Fraud risk area (R01); affects KPI-linked bonus", "Considers immaterial"),
        ("Inventory NRV write-down understated - slow-moving imports", "D3.3 / C4", "Judgmental", "Overstatement", "Cost of sales", "Inventory", 38000, "No", "No formal NRV policy", "Indicator of management bias (ISA 540.32)", "Disagrees with auditor's estimate"),
        ("Unrecorded utility & consulting accruals", "D3.4", "Factual", "Understatement", "Administrative expenses", "Accruals", 22000, "Yes", "Accrual checklist out of date", "-", "Corrected 18-Mar-26"),
        ("Depreciation of warehouse classified in admin instead of cost of sales", "D3.7", "Factual", "Classification", "Cost of sales", "Administrative expenses", 15000, "No", "Cost-centre mapping error", "Affects gross margin KPI", "Classification only"),
        ("Receivables - projected misstatement from MUS sample (incl. known errors)", "S5", "Projected", "Overstatement", "Revenue", "Trade receivables", "=ROUND(Rec_Proj,0)", "No", "Pricing disputes not credited", "Projection - estimate of population", "Statistical estimate"),
        ("Petty cash shortage at branch 2", "D4", "Factual", "Overstatement", "Administrative expenses", "Cash", 1200, "No", "Custodian error", "Below clearly trivial", "Trivial"),
        ("Current portion of borrowings presented as non-current", "D4.1", "Factual", "Classification", "Borrowings", "Borrowings", 1500000, "Yes", "Loan schedule not split", "Affects current ratio & covenant", "Corrected 19-Mar-26"),
        ("Director's loan of 25,000 not disclosed as related party transaction (IAS 24)", "B7", "Factual", "Disclosure", "", "", 25000, "No", "RP register incomplete", "Material by nature - RP & directors (ISA 450.A21)", "Under discussion")]
u0 = r
def suad_row(r, row, blank=False):
    if blank:
        dsc = src = typ = nat = dr = cr = amt = cor = rc = q = why = None
    else:
        dsc, src, typ, nat, dr, cr, amt, cor, rc, q, why = row
    cell(ws, f"A{r}", f"M{r-u0+1:02d}", center=True)
    cell(ws, f"B{r}", dsc, is_input=True); cell(ws, f"C{r}", src, center=True)
    cell(ws, f"D{r}", typ, is_input=True, center=True); cell(ws, f"E{r}", nat, is_input=True, center=True)
    cell(ws, f"F{r}", dr, is_input=True); cell(ws, f"G{r}", cr, is_input=True)
    cell(ws, f"H{r}", amt, NUM, is_input=not str(amt).startswith("="), link=str(amt).startswith("="))
    cell(ws, f"I{r}", cor, is_input=True, center=True)
    cell(ws, f"J{r}", f'=IF(H{r}="","",IF(OR(D{r}="Projected",H{r}>Mat_CT),"Yes","No - trivial"))', center=True)
    cell(ws, f"K{r}", f"=N(H{r})*{has(f'F{r}','Asset')}-N(H{r})*{has(f'G{r}','Asset')}", NUM)
    cell(ws, f"L{r}", f"=N(H{r})*{has(f'G{r}','Liability')}-N(H{r})*{has(f'F{r}','Liability')}", NUM)
    cell(ws, f"M{r}", f"=N(H{r})*{has(f'G{r}','Equity')}-N(H{r})*{has(f'F{r}','Equity')}", NUM)
    cell(ws, f"N{r}", f"=N(H{r})*{has(f'G{r}','Income')}-N(H{r})*{has(f'F{r}','Income')}", NUM)
    cell(ws, f"O{r}", f"=N(H{r})*{has(f'F{r}','Expense')}-N(H{r})*{has(f'G{r}','Expense')}", NUM)
    cell(ws, f"P{r}", f"=N{r}-O{r}", NUM)
    cell(ws, f"Q{r}", f"=-P{r}*TaxRate", NUM)
    cell(ws, f"R{r}", f"=P{r}+Q{r}", NUM)
    cell(ws, f"S{r}", rc, is_input=True); cell(ws, f"T{r}", q, is_input=True); cell(ws, f"U{r}", why, is_input=True)
    ws.row_dimensions[r].height = 32
for row in suad:
    suad_row(r, row); r += 1
for _ in range(4):
    suad_row(r, None, blank=True); r += 1
u1 = r - 1
dv(ws, f"D{u0}:D{u1}", "MSType"); dv(ws, f"E{u0}:E{u1}", "Nature"); dv(ws, f"I{u0}:I{u1}", "YesNo")
acct_dv = DataValidation(type="list", formula1="=TB_Acct", allow_blank=True); ws.add_data_validation(acct_dv); acct_dv.add(f"F{u0}:G{u1}")
dname("MS_Amt", f"SUAD!$H${u0}:$H${u1}"); dname("MS_Dr", f"SUAD!$F${u0}:$F${u1}"); dname("MS_Cr", f"SUAD!$G${u0}:$G${u1}")
dname("MS_Corr", f"SUAD!$I${u0}:$I${u1}"); dname("MS_Acc", f"SUAD!$J${u0}:$J${u1}"); dname("MS_Profit", f"SUAD!$P${u0}:$P${u1}")
dname("MS_Type", f"SUAD!$D${u0}:$D${u1}"); dname("MS_Nature", f"SUAD!$E${u0}:$E${u1}")
ws.conditional_formatting.add(f"I{u0}:I{u1}", CellIsRule(operator="equal", formula=['"No"'], fill=fill("FFE0B2")))
ws.conditional_formatting.add(f"I{u0}:I{u1}", CellIsRule(operator="equal", formula=['"Yes"'], fill=fill("C8E6C9")))
# totals rows
for lab, crit in (("Total - all misstatements", None), ("Total - CORRECTED", '"Yes"'), ("Total - UNCORRECTED (accumulated)", '"No"')):
    put(ws, f"B{r}", lab, font(9, True), PALE, bd=BOX)
    for col in "HKLMNOPQR":
        if crit is None: f_ = f"=SUM({col}{u0}:{col}{u1})"
        elif crit == '"Yes"': f_ = f'=SUMIFS({col}{u0}:{col}{u1},I{u0}:I{u1},"Yes")'
        else: f_ = f'=SUMIFS({col}{u0}:{col}{u1},I{u0}:I{u1},"No",J{u0}:J{u1},"Yes")'
        put(ws, f"{col}{r}", f_, font(9, True), PALE, NUM, bd=BOX)
    if crit == '"No"': UNC_ROW = r
    r += 1
put(ws, f"B{r}", "Check: Assets = Liabilities + Equity + PBT effect for every entry (double entry).", font(8, i=True, c="616161"))
put(ws, f"K{r}", f'=IF(ABS(K{r-3}-L{r-3}-M{r-3}-P{r-3})<1,"Balanced","CHECK ENTRIES")', font(9, True, MD))
ws.merge_cells(f"K{r}:M{r}")
r += 2

r = section(ws, r, "SUAD2", "Prior-period uncorrected misstatements - reversal effects", "ISA 450.11(b) & A20")
r = table_head(ws, r, ["Ref", "Prior-year uncorrected misstatement", "PY PBT effect", "Reversed in CY?", "CY PBT effect of reversal",
                       "Remaining YE balance-sheet effect", "Comment"])
ws.merge_cells(f"G{r-1}:K{r-1}")
py = [("PY inventory NRV under-provision (goods sold in 2025)", -21000, "Yes", "Opening inventory overstated; CY cost of sales overstated on sale"),
      ("PY unrecorded accrual - consulting (invoice booked Jan-25)", -9000, "Yes", "CY expense overstated by the PY cost"),
      ("PY PPE impairment not recognised - idle machine", -6000, "No", "Still unadjusted at YE - carried in balance sheet")]
y0 = r
for i, (dsc, eff, rev, com) in enumerate(py, 1):
    cell(ws, f"A{r}", f"PY{i:02d}", center=True); cell(ws, f"B{r}", dsc, is_input=True)
    cell(ws, f"C{r}", eff, NUM, True); cell(ws, f"D{r}", rev, is_input=True, center=True)
    cell(ws, f"E{r}", f'=IF(D{r}="Yes",-C{r},0)', NUM); cell(ws, f"F{r}", f'=IF(D{r}="Yes",0,C{r})', NUM)
    cell(ws, f"G{r}", com, is_input=True); ws.merge_cells(f"G{r}:K{r}")
    r += 1
y1 = r - 1
dv(ws, f"D{y0}:D{y1}", "YesNo")
put(ws, f"B{r}", "Totals", font(9, True), PALE, bd=BOX)
for col in "CEF": put(ws, f"{col}{r}", f"=SUM({col}{y0}:{col}{y1})", font(9, True), PALE, NUM, bd=BOX)
PYT = r
put(ws, f"B{r+1}", "Sign: negative = overstates profit (correction would reduce profit). A PY overstatement that reverses in CY understates CY profit (positive).", font(8, i=True, c="616161"))
r += 3

r = section(ws, r, "SUAD3", "Evaluation of uncorrected misstatements - individually & in aggregate", "ISA 450.10-12 | ISA 700.11")
ev0 = r
evals = [
    ("Current-year uncorrected misstatements - PBT effect", f"=P{UNC_ROW}", NUM),
    ("Rollover (income-statement) approach = CY + reversal of PY items", f"=C{ev0}+E{PYT}", NUM),
    ("Iron-curtain (balance-sheet) approach = CY + unreversed PY items", f"=C{ev0}+F{PYT}", NUM),
    ("DUAL-APPROACH AGGREGATE (larger absolute value)", f"=IF(ABS(C{ev0+1})>ABS(C{ev0+2}),C{ev0+1},C{ev0+2})", NUM),
    ("Overall materiality (B3)", "=Mat_OM", NUM),
    ("Performance materiality (B3)", "=Mat_PM", NUM),
    ("Aggregate as % of overall materiality", f"=IFERROR(ABS(C{ev0+3})/Mat_OM,0)", PCT),
    ("Largest individual uncorrected misstatement (PBT)", f'=-MIN(SUMIFS(P{u0}:P{u0},I{u0}:I{u0},"No"),0)', NUM),
    ("Uncorrected classification misstatements (gross)", f'=SUMIFS(H{u0}:H{u1},I{u0}:I{u1},"No",E{u0}:E{u1},"Classification")', NUM),
    ("Uncorrected disclosure misstatements (count)", f'=COUNTIFS(I{u0}:I{u1},"No",E{u0}:E{u1},"Disclosure")', "0"),
    ("Net profit effect after tax (uncorrected)", f"=R{UNC_ROW}", NUM),
    ("Reported PBT (after corrected adjustments, D1)", "=TB_PBT_Final", NUM),
    ("PBT if all uncorrected misstatements were corrected", f"=C{ev0+11}+C{ev0}", NUM),
    ("Effect on reported PBT %", f"=IFERROR(C{ev0}/C{ev0+11},0)", PCT),
    ("QUANTITATIVE EVALUATION", f'=IF(ABS(C{ev0+3})>=Mat_OM,"MATERIAL in aggregate - request correction; if refused, modify opinion (ISA 705.7)",'
                                f'IF(ABS(C{ev0+3})>=Mat_PM,"Below OM but above PM - request correction; reassess risk of undetected misstatement (ISA 450.6)",'
                                f'"Not material individually or in aggregate"))', None),
    ("QUALITATIVE EVALUATION", f'=IF(C{ev0+9}>0,C{ev0+9}&" disclosure misstatement(s) - evaluate whether material by nature (ISA 450.A21-A23)","No qualitative matters")', None)]
for i, (lab, v, nf) in enumerate(evals):
    rr = r + i
    if i == 7:  # largest individual: use MIN over range
        v = f'=-MIN(0,SUMPRODUCT(MIN((I{u0}:I{u1}="No")*(J{u0}:J{u1}="Yes")*P{u0}:P{u1})))'
    b = put(ws, f"B{rr}", lab, font(9, lab.isupper() or i == 3), PALE if (lab.isupper() or i == 3) else None, bd=BOX)
    c = put(ws, f"C{rr}", v, font(10 if lab.isupper() or i == 3 else 9, lab.isupper() or i == 3, DK if lab.isupper() or i == 3 else ("008000" if str(v).startswith(("=Mat", "=TB")) else "000000")),
            PALE if (lab.isupper() or i == 3) else None, nf, bd=BOX)
    if lab.isupper() and i >= 14: ws.merge_cells(f"C{rr}:M{rr}")
dname("MS_Uncorr", f"SUAD!$C${ev0}"); dname("MS_Agg", f"SUAD!$C${ev0+3}"); dname("MS_UncorrPct", f"SUAD!$C${ev0+6}")
dname("MS_Eval", f"SUAD!$C${ev0+14}"); dname("MS_Qual", f"SUAD!$C${ev0+15}")
dname("MS_Roll", f"SUAD!$C${ev0+1}"); dname("MS_Iron", f"SUAD!$C${ev0+2}")
r += len(evals) + 1
r = checklist(ws, r, [
    ("All misstatements above CT accumulated & communicated to management on a timely basis", "ISA 450.5, 8", "Yes", "Meeting with CFO 18-Mar-26"),
    ("Management requested to correct all misstatements; reasons for non-correction understood", "ISA 450.8-9", "Yes", "See column U"),
    ("Materiality reassessed in light of actual results before evaluation", "ISA 450.10", "Yes", "Final PBT within 2% of planning benchmark"),
    ("Uncorrected misstatements communicated to TCWG with request for correction", "ISA 450.12-13", None, "F1 item 5"),
    ("Written representation obtained that uncorrected misstatements are immaterial (schedule attached)", "ISA 450.14", None, "F2"),
    ("Documentation: CT amount, all accumulated misstatements, conclusion on materiality", "ISA 450.15", "Yes", "This schedule"),
], "SUAD3")

r = section(ws, r, "SUAD4", "Uncorrected misstatements by FS line item (adjusted vs unadjusted)", "ISA 450.11 & A16-A17")
r = table_head(ws, r, ["#", "Account", "Class", "Reported (D1 final)", "Uncorrected Dr/(Cr)", "If adjusted", "% of line item", "vs PM", "vs OM"])
z0 = r
for i in range(len(tbrows)):
    cell(ws, f"A{r}", i + 1, center=True)
    cell(ws, f"B{r}", f"=INDEX(TB_Acct,{i+1})", link=True); cell(ws, f"C{r}", f"=INDEX(TB_Class,{i+1})", link=True, center=True)
    cell(ws, f"D{r}", f"=INDEX(TB_Final,{i+1})", NUM, link=True)
    cell(ws, f"E{r}", f'=SUMIFS(MS_Amt,MS_Dr,B{r},MS_Corr,"No",MS_Acc,"Yes")-SUMIFS(MS_Amt,MS_Cr,B{r},MS_Corr,"No",MS_Acc,"Yes")', NUM)
    cell(ws, f"F{r}", f"=D{r}+E{r}", NUM)
    cell(ws, f"G{r}", f"=IFERROR(E{r}/ABS(D{r}),0)", PCT, center=True)
    cell(ws, f"H{r}", f'=IF(ABS(E{r})>=Mat_PM,"> PM",IF(E{r}=0,"-","< PM"))', center=True)
    cell(ws, f"I{r}", f'=IF(ABS(E{r})>=Mat_OM,"> OM","-")', center=True)
    r += 1
ws.conditional_formatting.add(f"H{z0}:I{r-1}", FormulaRule(formula=[f'LEFT(H{z0},1)=">"'], fill=fill("FFCDD2"), font=Font(name=F, bold=True, color="B71C1C")))
put(ws, f"B{r}", "Net check (must be zero)", font(9, True), PALE, bd=BOX); put(ws, f"E{r}", f"=SUM(E{z0}:E{r-1})", font(9, True), PALE, NUM, bd=BOX)
ws.freeze_panes = "C8"

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
 ("IC1", "Planning", "Entity-level controls - COSO 17 principles", "ISA 315R.21-25", "Reviewed", 6, 7),
 ("IC2", "Planning", "Risk & control matrix - design & implementation", "ISA 315R.26; ISA 330.8", "Reviewed", 14, 16),
 ("IC3", "Planning", "IT general controls assessment", "ISA 315R.26(c)", "Reviewed", 10, 11),
 ("C1", "Methodology", "Audit programme by FS area & assertion", "ISA 330.6-7", "Reviewed", 8, 8),
 ("C2", "Methodology", "Sampling methodology (MUS)", "ISA 530", "Reviewed", 4, 4),
 ("C3", "Methodology", "Analytical procedures", "ISA 520", "Prepared", 8, 9),
 ("C4", "Methodology", "Accounting estimates", "ISA 540R", "Prepared", 12, 15),
 ("C5", "Methodology", "Internal audit function", "ISA 610", "Closed", 1, 0.5),
 ("C6", "Methodology", "Auditor's expert (IT)", "ISA 620", "Reviewed", 3, 3),
 ("C7", "Methodology", "Group audit considerations", "ISA 600R", "Closed", 1, 0.5),
 ("C8", "Methodology", "Service organisations (payroll)", "ISA 402", "Reviewed", 4, 4),
 ("S1", "Methodology", "Population register & sample sizes", "ISA 530.6-8; ISA 500.9", "Reviewed", 5, 5),
 ("S2", "Methodology", "Statistical factors & attribute calculator", "ISA 530 App. 2-3", "Closed", 1, 1),
 ("S3", "Methodology", "Non-statistical control sample sizes", "ISA 530.A13", "Closed", 1, 1),
 ("S4", "Methodology", "MUS selection - receivables", "ISA 530.8; ISA 505", "Reviewed", 4, 5),
 ("S5", "Methodology", "MUS evaluation - projection & UML", "ISA 530.14-15", "Prepared", 3, 3),
 ("D1", "Execution", "Lead schedule / trial balance", "ISA 500; ISA 230", "Reviewed", 6, 7),
 ("D2", "Execution", "Tests of controls", "ISA 330.8-17", "Reviewed", 40, 46),
 ("D3", "Execution", "Substantive tests of details", "ISA 330.18; ISA 500", "Prepared", 90, 98),
 ("D4", "Execution", "External confirmations", "ISA 505", "In progress", 14, 12),
 ("D5", "Execution", "Inventory count attendance", "ISA 501.4-8", "Reviewed", 16, 18),
 ("D6", "Execution", "Journal entry testing", "ISA 240.33", "Prepared", 12, 14),
 ("D7", "Execution", "Litigation & claims", "ISA 501.9-12", "In progress", 3, 1),
 ("D8", "Execution", "Subsequent events", "ISA 560", "In progress", 4, 2),
 ("D9", "Execution", "Comparative information", "ISA 710", "Reviewed", 2, 2),
 ("IC4", "Execution", "Tests of operating effectiveness (attribute)", "ISA 330.8-11; ISA 530", "Reviewed", 24, 27),
 ("IC5", "Execution", "Attribute testing detail - C01", "ISA 330.10", "Reviewed", 6, 6),
 ("IC6", "Execution", "Evaluation of deviations", "ISA 330.16-17; ISA 265", "Prepared", 3, 3),
 ("E1", "Findings", "Summary of misstatements", "ISA 450", "Prepared", 5, 6),
 ("E2", "Findings", "Control deficiencies evaluation", "ISA 265", "Prepared", 4, 4),
 ("E3", "Findings", "Issue log & remediation tracker", "ISA 230; ISQM 1.42", "In progress", 4, 3),
 ("E4", "Findings", "Management letter", "ISA 265.9-11", "In progress", 5, 3),
 ("SUAD1", "Findings", "SUAD - all current-year misstatements", "ISA 450.5-8", "Prepared", 4, 5),
 ("SUAD2", "Findings", "SUAD - prior-period reversal effects", "ISA 450.11(b)", "Prepared", 1, 1),
 ("SUAD3", "Findings", "SUAD - evaluation (rollover / iron curtain)", "ISA 450.10-12", "In progress", 2, 1),
 ("SUAD4", "Findings", "SUAD - by FS line item", "ISA 450.A16", "Prepared", 1, 1),
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
for i, (lab, f_) in enumerate([("Uncorrected (dual aggregate)", "=ABS(MS_Agg)"), ("Clearly trivial", "=Mat_CT"),
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
        ("EQR", "EQR status", '=INDEX(WP_Status,MATCH("F5",WP_Index!$A$6:$A$200,0))', "@"),
        ("DIpct", "D&I effective %", "=IFERROR(IC_DIeff/IC_RCMn,0)", "0%"),
        ("ITGC", "ITGC status", '=IF(LEFT(IC_ITGC,14)="ITGCs DEFICIEN","Deficient","Effective")', "@")]
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
        ("ISA 265", "Communicating Deficiencies in Internal Control", "Findings", "E2, E4, IC6"),
        ("ISA 300", "Planning an Audit of Financial Statements", "Strategy / Planning", "A6, B10"),
        ("ISA 315 (Revised 2019)", "Identifying and Assessing the Risks of Material Misstatement", "Planning", "B1, B2, B4, IC1-IC3"),
        ("ISA 320", "Materiality in Planning and Performing an Audit", "Planning", "B3"),
        ("ISA 330", "Auditor's Responses to Assessed Risks", "Methodology / Execution", "C1, D2, D3, IC4-IC6"),
        ("ISA 402", "Audit Considerations Relating to Service Organizations", "Methodology", "C8"),
        ("ISA 450", "Evaluation of Misstatements Identified during the Audit", "Findings", "E1, SUAD1-SUAD4"),
        ("ISA 500", "Audit Evidence", "Execution", "D1, D3"),
        ("ISA 501", "Audit Evidence - Specific Considerations", "Execution", "D5, D7"),
        ("ISA 505", "External Confirmations", "Execution", "D4"),
        ("ISA 510", "Initial Audit Engagements - Opening Balances", "Planning", "B9"),
        ("ISA 520", "Analytical Procedures", "Methodology", "C3"),
        ("ISA 530", "Audit Sampling", "Methodology", "C2, S1-S5, IC4-IC5"),
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
                 ["Materiality (B3) & risk register (B4)", "COSO entity-level controls (IC1)", "RCM design & implementation (IC2)", "ITGC assessment (IC3)"], "L2_Planning"),
    "Execution": ("L19", ["Fieldwork execution", "Control testing", "Data validation", "Process walkthrough"],
                  ["Lead schedule (D1)", "Tests of controls - attribute (IC4-IC5)", "Confirmations & MUS selection (D4, S4)", "JE testing (D6)"], "L4_Execution"),
    "Methodology": ("B33", ["Standard procedures", "Sampling techniques", "Evaluation methods", "Quality standards"],
                    ["Audit programme (C1)", "Population register & sizes (S1)", "Statistical factors (S2-S3)", "MUS evaluation / UML (S5)"], "L3_Methodology"),
    "Oversight": ("L33", ["TCWG interaction", "Quality assurance reviews", "Independence monitoring", "Performance evaluation"],
                  ["TCWG log (F1)", "EQ review (F5)", "Opinion matrix (F7)", "ISQM 1 metrics (F10)"], "L6_Oversight"),
    "Findings": ("G47", ["Issue identification", "Root cause analysis", "Recommendation design", "Remediation tracking"],
                 ["SUAD - all misstatements (SUAD1-4)", "Deficiencies (E2, IC6)", "Issue log (E3)", "Management letter (E4)"], "L5_Findings"),
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
for rr in range(6, 55):
    for cc in (2, 3): db.cell(rr, cc).fill = fill(DK)
navs = [("Overview", "LayersMap"), ("Setup", "Setup"), ("WP Index", "WP_Index"), ("1 Strategy", "L1_Strategy"),
        ("2 Planning", "L2_Planning"), ("3 Methodology", "L3_Methodology"), ("   Sampling", "Sampling"), ("4 Execution", "L4_Execution"),
        ("   IC Testing", "IC_Assessment"), ("5 Findings", "L5_Findings"), ("   SUAD", "SUAD"), ("6 Oversight", "L6_Oversight"), ("Standards", "Standards")]
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
      ("Uncorrected Misstatements", "=ABS(MS_Agg)", "#,##0", '=TEXT(MS_UncorrPct,"0%")&" of OM - dual method (SUAD)"'),
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
B2 = [("D&I Effective", "=DD_DIpct", "0%", '=IC_DIeff&" of "&IC_RCMn&" controls (IC2)"'),
      ("ITGC Status", "=DD_ITGC", "@", '="4 domains assessed (IC3)"'),
      ("Control Items Tested", "=IC_ItemsTested", "0", '="attribute samples (IC4)"'),
      ("MUS Sample - AR", "=Rec_Items", "0", '="items | "&TEXT(Rec_Cov,"0%")&" of value (S4)"'),
      ("MUS UML vs TM", "=Rec_UMLpct", "0%", '=IF(Rec_UMLpct<=1,"Population accepted (S5)","Not accepted (S5)")'),
      ("Misstatements Logged (SUAD)", '=COUNT(MS_Amt)&" logged | "&COUNTIF(MS_Corr,"Yes")&" corrected"', "@",
       '="Rollover "&TEXT(ABS(MS_Roll),"#,##0")&" | Iron curtain "&TEXT(ABS(MS_Iron),"#,##0")')]
for k, t in enumerate(B2):
    tile(51, 5 + k * 3, 3 if k < 5 else 6, *t, accent="1565C0")
db["T52"].font = font(12, True, DK)
put(db, "E56", "=\"Misstatement evaluation: \"&MS_Eval", font(9, True, "B71C1C"), BG); db.merge_cells("E56:Y56")
put(db, "E57", "All figures update live from the layer working papers. Illustrative data for teaching - replace in Setup & layer sheets.", font(8, i=True, c="757575"), BG)
db.merge_cells("E57:Y57")
db.freeze_panes = "A5"
db.print_area = "A1:Y57"

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
