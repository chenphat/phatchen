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

order = ["Dashboard", "LayersMap", "Setup", "WP_Index", "L1_Strategy", "Acceptance_QM", "Team_Resources", "PBC_List", "L2_Planning",
         "Risk_315R", "Assertions", "Fraud_240R", "GoingConcern_570R", "L3_Methodology", "Sampling", "L4_Execution", "IC_Assessment", "L5_Findings", "Deficiencies_265", "SUAD", "L6_Oversight", "AI_Agents", "DashData", "Standards", "Lists"]
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
    "DefType": ["Design", "Implementation", "Operating"],
    "NewCont": ["Continuing", "New"],
    "DataClass": ["Public", "Anonymised", "Confidential"],
    "PBCStatus": ["Not received", "Partially received", "Received", "Accepted", "Re-request"],
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
    ("Early adoption: ISA 240 (Revised) & ISA 570 (Revised 2024)", "Yes", "Effective for periods beginning on/after 15-Dec-2026; IAASB encourages early adoption", "EarlyAdopt"),
    ("Engagement status (Continuing / New)", "Continuing", "Drives new-client procedures (AC4)", "EngStatus"),
    ("Years as audit client", 4, "Long association considerations (IESBA 540)", "YearsClient"),
    ("Standard working hours per month", 168, "Capacity planning (TR4)", "StdHrs"),
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
        from openpyxl.utils import get_column_letter as _gl, column_index_from_string as _ci; col = _gl(_ci(start) + j)
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
r = section(ws, r, "A1", "Engagement acceptance & continuance - summary (full: Acceptance_QM AC1-AC7)", "ISQM 1.30 | ISA 220R.22-24")
r = table_head(ws, r, ["", "Summary", "Value"])
for lab, f_ in [("Engagement risk rating", "=AC_Rating"), ("Fatal factors rated high", "=AC_Fatal"), ("Decision", "=AC_Decision"),
                ("Team resources conclusion", "=TR_Conclusion")]:
    cell(ws, f"B{r}", lab); cell(ws, f"C{r}", f_, link=True); ws.merge_cells(f"C{r}:J{r}"); r += 1
c = cell(ws, f"B{r}", '=HYPERLINK("#Acceptance_QM!A1","Open acceptance & continuance working paper >")'); c.font = Font(name=F, size=9, color="1565C0", underline="single")
r += 2

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
dname("Ind_Assess", f"L1_Strategy!$K${t0}:$K${r-1}"); dname("Ind_Name", f"L1_Strategy!$B${t0}:$B${r-1}")
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
                 "ISA 300; ISA 315 (Revised 2019); ISA 320; ISA 240 (Revised); ISA 250; ISA 550; ISA 570 (Revised 2024); ISA 510")
r = section(ws, r, "B1", "Understanding the entity, its environment & framework", "ISA 315R.19")
r = checklist(ws, r, [
    ("Organisational structure, ownership, governance and business model understood", "ISA 315R.19(a)(i)", "Yes", "Family-owned; Board of 5; retail of consumer goods"),
    ("Industry, regulatory and other external factors (incl. tax, inflation, FX) understood", "ISA 315R.19(a)(ii)", "Yes", "MNT depreciation affects imported inventory cost"),
    ("Measures used to assess financial performance (KPIs, bonuses) understood", "ISA 315R.19(a)(iii)", "Yes", "Mgmt bonus linked to EBITDA - fraud risk factor"),
    ("Applicable framework, accounting policies & reasons for changes understood", "ISA 315R.19(b)", "Yes", "IFRS 16 & IFRS 15 policies unchanged"),
    ("Inherent risk factors considered (complexity, subjectivity, change, uncertainty, bias)", "ISA 315R.19(c)", "Yes", "Applied in Risk_315R RA3 (IRF ratings)"),
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

r = section(ws, r, "B4", "Risk assessment - summary (full register: Risk_315R RA1-RA6)", "ISA 315 (Revised 2019).28-38")
r = table_head(ws, r, ["", "Summary", "Value"])
for lab, f_ in [('Risks of material misstatement assessed', '=COUNTA(Risk_ID)'), ('FS-level risks', '=COUNTIF(Risk_Level,"FS level")'), ('Significant risks', '=COUNTIF(Risk_Spectrum,"Significant")'), ('High / Elevated', '=COUNTIF(Risk_Spectrum,"High")&" / "&COUNTIF(Risk_Spectrum,"Elevated")'), ('Moderate / Low', '=COUNTIF(Risk_Spectrum,"Moderate")&" / "&COUNTIF(Risk_Spectrum,"Low")'), ('Risks with an inherent risk factor rated 3 (high)', '=COUNTIF(Risk_MaxIRF,3)'), ('Fraud risks', '=COUNTIF(Risk_Fraud,"Yes")'), ('Assessments needing justification (consistency check)', '=COUNTIF(Risk_Check,"<>OK")'), ('Significant SCOTABDs / material-not-significant (stand-back)', '=COUNTIF(SCOT_Class,"Significant*")&" / "&COUNTIF(SCOT_Class,"Material*")')]:
    cell(ws, f"B{r}", lab); cell(ws, f"C{r}", f_, link=True); ws.merge_cells(f"C{r}:J{r}"); r += 1
c = cell(ws, f"B{r}", '=HYPERLINK("#Risk_315R!A1","Open risk register - inherent risk factors & spectrum >")'); c.font = Font(name=F, size=9, color="1565C0", underline="single")
r += 2

r = section(ws, r, "B5", "Fraud - summary (full working paper: Fraud_240R FR1-FR7)", "ISA 240 (Revised)")
r = table_head(ws, r, ["", "Summary", "Value"])
for lab, f_ in [('Fraud risk factors present', '=COUNTIF(FRF_Present,"Yes")'), ('Risks of material misstatement due to fraud', '=COUNTIF(Risk_Fraud,"Yes")'), ('Journal entries tested / exceptions', '=JE_Tested&" / "&JE_Exc'), ('Fraud or suspected fraud matters (incl. allegations)', '=COUNTA(Fraud_Status)'), ('Open fraud matters', '=COUNTIF(Fraud_Status,"<>Closed")')]:
    cell(ws, f"B{r}", lab); cell(ws, f"C{r}", f_, link=True); ws.merge_cells(f"C{r}:J{r}"); r += 1
c = cell(ws, f"B{r}", '=HYPERLINK("#Fraud_240R!A1","Open ISA 240 (Revised) working paper >")'); c.font = Font(name=F, size=9, color="1565C0", underline="single")
r += 2

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
r = section(ws, r, "B8", "Going concern - risk assessment summary (full: GoingConcern_570R)", "ISA 570 (Revised 2024)")
r = table_head(ws, r, ["", "Summary", "Value"])
for lab, f_ in [('Events or conditions identified', '=GC_EventsN'), ("Months covered by management's assessment from FS approval", '=GC_Months'), ('Reporting outcome', '=GC_Outcome')]:
    cell(ws, f"B{r}", lab); cell(ws, f"C{r}", f_, link=True); ws.merge_cells(f"C{r}:J{r}"); r += 1
c = cell(ws, f"B{r}", '=HYPERLINK("#GoingConcern_570R!A1","Open ISA 570 (Revised 2024) working paper >")'); c.font = Font(name=F, size=9, color="1565C0", underline="single")
r += 2

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
                 "ISA 230; ISA 330; ISA 500; ISA 501; ISA 505; ISA 240 (Revised); ISA 560; ISA 710")
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
r = section(ws, r, "D6", "Journal entry testing (management override)", "ISA 240 (Revised) - mgmt override (see FR4) | ISA 330.20")
r = checklist(ws, r, [
    ("Completeness of JE population reconciled to TB", "ISA 240R - JE testing", "Yes", "48,320 JEs; agreed to TB movement"),
    ("Inquiries of individuals about inappropriate/unusual activity in JE process", "ISA 240R - JE testing", "Yes", ""),
    ("JEs & adjustments at period end selected & tested", "ISA 240R - JE testing", "Yes", "All post-closing entries tested"),
    ("Risk criteria applied: weekend/holiday, round amounts, unusual users, seldom-used accounts, post-close", "ISA 240R - JE testing (FR4)", "Yes", "112 entries selected via data analytics"),
    ("Significant transactions outside normal course of business evaluated for business rationale", "ISA 240R - unusual transactions", "Yes", "None noted"),
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

r = section(ws, r, "E2", "Control deficiencies - summary (full ISA 265 working paper: Deficiencies_265)", "ISA 265.7-11")
r = table_head(ws, r, ["", "Summary", "Value"])
for lab, f_ in [("Deficiencies identified", "=CD_N"), ("Significant deficiencies", '=COUNTIF(CD_Class,"Significant deficiency")'),
                ("Significant deficiencies communicated in writing to TCWG", "=CD_SigComm"),
                ("Other deficiencies of sufficient importance (management)", '=COUNTIF(CD_Class,"Other deficiency - sufficient importance")'),
                ("Remediated", "=CD_Closed")]:
    cell(ws, f"B{r}", lab); cell(ws, f"C{r}", f_, "0", link=True, center=True); r += 1
c = cell(ws, f"B{r}", '=HYPERLINK("#Deficiencies_265!A1","Open ISA 265 deficiencies working paper >")'); c.font = Font(name=F, size=9, color="1565C0", underline="single")
r += 2

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
                 "ISA 260; ISA 580; ISA 570 (Revised 2024); ISA 701; ISA 700/705/706; ISA 720; ISA 220 (Revised); ISQM 1; ISQM 2")
r = section(ws, r, "F1", "Communications with those charged with governance", "ISA 260 (Revised) | ISA 265")
r = table_head(ws, r, ["#", "Matter required to be communicated", "Standard ref", "Required?", "Form", "Date", "Status", "Ref"])
comms = [("Auditor's responsibilities in relation to the FS audit", "ISA 260.14", "Yes", "Written", D(2025, 10, 20), "Done"),
         ("Planned scope & timing incl. significant risks", "ISA 260.15", "Yes", "Written", D(2025, 10, 20), "Done"),
         ("Significant qualitative aspects of accounting practices", "ISA 260.16(a)", "Yes", "Meeting", None, "Pending"),
         ("Significant difficulties encountered during the audit", "ISA 260.16(b)", "Yes", "Meeting", None, "Pending"),
         ("Uncorrected misstatements & request for correction", "ISA 450.12-13", "Yes", "Written", None, "Pending"),
         ("Significant deficiencies in internal control", "ISA 265.9", "Yes", "Written", D(2026, 3, 20), "Done"),
         ("Statement on independence & threats/safeguards", "ISA 260.17; IESBA 400", "Yes", "Written", None, "Pending"),
         ("Fraud-related matters incl. suspected fraud & allegations (FR6)", "ISA 240R - communications", "Yes", "Written", None, "Pending"),
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
    ("Fraud, laws & regulations, related parties, subsequent events, going concern", "ISA 240R; 250.17; 550.26; 560.9; 570R", None, ""),
    ("Dated as near as practicable to, but not after, the auditor's report date", "ISA 580.14", None, ""),
], "F2")
r = section(ws, r, "F3", "Going concern - final conclusion (detail: GoingConcern_570R GC4-GC6)", "ISA 570 (Revised 2024)")
r = table_head(ws, r, ["", "Summary", "Value"])
for lab, f_ in [('Going concern basis appropriate?', '=GC_Basis'), ('Material uncertainty exists?', '=GC_MU'), ('Close call?', '=GC_Close'), ('Reporting outcome', '=GC_Outcome')]:
    cell(ws, f"B{r}", lab); cell(ws, f"C{r}", f_, link=True); ws.merge_cells(f"C{r}:J{r}"); r += 1
c = cell(ws, f"B{r}", '=HYPERLINK("#GoingConcern_570R!A1","Open going concern working paper >")'); c.font = Font(name=F, size=9, color="1565C0", underline="single")
r += 2

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
        ("Material uncertainty related to going concern? (GC4)", "=GC_MU", "ISA 570 (Revised 2024)", False),
        ("Emphasis of matter required?", "No", "ISA 706.8", True),
        ("PROPOSED OPINION", None, "ISA 705 decision matrix", False),
        ("Additional report section", None, "", False)]
for i, (lab, v, basis, isin) in enumerate(opin):
    rr = r + i
    if lab == "PROPOSED OPINION":
        v = (f'=IF(GC_Basis="No","Adverse opinion (ISA 570R)",IF(AND(C{o0+1}="Yes",C{o0+2}="Yes"),"Disclaimer of opinion",IF(AND(C{o0}="Yes",C{o0+2}="Yes"),"Adverse opinion",'
             f'IF(OR(C{o0}="Yes",C{o0+1}="Yes"),"Qualified opinion","Unmodified opinion"))))')
    if lab.startswith("Additional"):
        v = (f'=IF(C{o0+3}="Yes","Material Uncertainty Related to Going Concern section",IF(C{o0+4}="Yes","Emphasis of Matter paragraph","Going Concern section - no material uncertainty (ISA 570R)"))')
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
      ("Partner hours as % of total actual hours", 0.1, "=TR_PartnerPct", PCT, "TR3"),
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
    ("Deviations evaluated as control deficiencies & logged in Deficiencies_265 / E3", "ISA 265.7-8", "Yes", "CD02, CD03"),
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

# ======================================================================
# ISA 315 (Revised 2019) / ISA 240 (Revised) / ISA 570 (Revised 2024)
# ======================================================================
from openpyxl.utils import column_index_from_string as ci, get_column_letter as gl
SPECTRUM = ["Low", "Moderate", "Elevated", "High", "Significant"]
SPEC_COL = {"Low": "C8E6C9", "Moderate": "FFF59D", "Elevated": "FFCC80", "High": "EF6C00", "Significant": "C62828"}

def note(ws, r, text, last="L", h=24):
    put(ws, f"B{r}", text, font(8, i=True, c="616161"), al=WRAP); ws.merge_cells(f"B{r}:{last}{r}"); ws.row_dimensions[r].height = h

# ---------------- RISK 315R ----------------
ws = S["Risk_315R"]
r = sheet_header(ws, "RISK IDENTIFICATION & ASSESSMENT - INHERENT RISK FACTORS & SPECTRUM", "2",
                 "Identify risks of material misstatement at FS and assertion level, assess inherent risk using the inherent risk factors and the spectrum of inherent risk, assess control risk separately, and apply the stand-back requirement.",
                 "ISA 315 (Revised 2019) paras 12 (definitions), 28-38 & Appendix 2 (inherent risk factors); ISA 330.5-7, 18; ISA 240 (Revised) fraud lens")
for k in range(ci("M"), ci("Y") + 1): ws.column_dimensions[gl(k)].width = 11
widths(ws, {"B": 38, "C": 18, "D": 16, "E": 10, "F": 8, "X": 34, "T": 26, "W": 18})

r = section(ws, r, "RA1", "Inherent risk factors (IRFs) & spectrum of inherent risk - methodology", "ISA 315R.12(f), 31 & App. 2")
r = table_head(ws, r, ["", "Inherent risk factor", "What it means (ISA 315R App. 2)", "", "", "", "", "Rating scale"])
ws.merge_cells(f"C{r-1}:G{r-1}"); ws.merge_cells(f"H{r-1}:L{r-1}")
irfs = [("Complexity", "Difficulty in applying the framework or processing data - many sources, complex calculations or models"),
        ("Subjectivity", "Inherent limitations in preparing information objectively - range of possible outcomes requiring management judgement"),
        ("Change", "Events or conditions changing the business, industry, framework or IT that affect amounts or disclosures"),
        ("Uncertainty", "Information cannot be measured precisely from direct observation - estimation uncertainty"),
        ("Susceptibility to management bias / other fraud risk factors", "Conditions creating opportunity for intentional or unintentional bias; incentives, pressures, opportunities, attitudes (link Fraud_240R FR2)")]
for i, (n_, d_) in enumerate(irfs):
    cell(ws, f"B{r}", n_); cell(ws, f"C{r}", d_); ws.merge_cells(f"C{r}:G{r}")
    cell(ws, f"H{r}", "0 = not present | 1 = low | 2 = moderate | 3 = high" if i == 0 else ""); ws.merge_cells(f"H{r}:L{r}")
    ws.row_dimensions[r].height = 28; r += 1
r += 1
r = table_head(ws, r, ["", "Spectrum position", "Likelihood x Magnitude (1-25)", "Meaning / typical response"])
ws.merge_cells(f"D{r-1}:L{r-1}")
for band, rng, mean in [("Low", "1 - 4", "Not a RMM, or remote; may still need ISA 330.18 procedures if material"),
                        ("Moderate", "5 - 9", "Standard substantive procedures; controls reliance possible"),
                        ("Elevated", "10 - 14", "More persuasive evidence; extend sample sizes / reduce reliance on analytics"),
                        ("High", "15 - 19", "Most persuasive evidence; tests of details close to period end; experienced staff"),
                        ("Significant", "Auditor judgement: close to upper end; ALWAYS for fraud risks", "Specific response required: understand controls (315R.26(a)(i)), substantive procedures specifically responsive (330.21), no reliance on PY control evidence (330.15)")]:
    c = cell(ws, f"B{r}", band, center=True); c.fill = fill(SPEC_COL[band])
    if band in ("High", "Significant"): c.font = font(9, True, "FFFFFF")
    cell(ws, f"C{r}", rng, center=True); cell(ws, f"D{r}", mean); ws.merge_cells(f"D{r}:L{r}"); ws.row_dimensions[r].height = 26; r += 1
note(ws, r, "Inherent risk is assessed by assessing the likelihood AND magnitude of misstatement, taking into account how and the degree to which IRFs affect susceptibility (ISA 315R.31). "
            "Significant risk = inherent risk close to the upper end of the spectrum (315R.12(l)). Control risk is assessed separately (315R.34); where controls are not tested, control risk is maximum and RMM = inherent risk.")
r += 2

r = section(ws, r, "RA2", "Significant classes of transactions, account balances & disclosures (SCOTABDs)", "ISA 315R.29 | ISA 330.18")
r = table_head(ws, r, ["#", "Account / disclosure", "Class", "Final balance (abs)", "> PM?", "Qualitative significance", "Risks mapped (RA3)",
                       "Relevant assertions", "", "Classification", "", ""])
ws.merge_cells(f"H{r-1}:I{r-1}"); ws.merge_cells(f"J{r-1}:L{r-1}")
qual = {"Revenue": "KPI-linked bonus; fraud presumption", "Inventory": "Imported goods; NRV judgement", "Property, plant & equipment": "IFRS 16 right-of-use assets",
        "Retained earnings (opening)": "Agreed to PY", "Income tax payable": "Tax inspection history"}
asr = {"Revenue": "Occurrence, Cut-off, Accuracy", "Trade receivables": "Existence, Valuation", "Inventory": "Existence, Valuation",
       "Property, plant & equipment": "Existence, Valuation, Classification", "Accruals": "Completeness", "Borrowings": "Presentation, Completeness",
       "Cash": "Existence", "Cost of sales": "Accuracy, Classification", "Trade payables": "Completeness"}
s0 = r
for i, row in enumerate(tbrows):
    acct = row[1]
    cell(ws, f"A{r}", i + 1, center=True); cell(ws, f"B{r}", f"=INDEX(TB_Acct,{i+1})", link=True)
    cell(ws, f"C{r}", f"=INDEX(TB_Class,{i+1})", link=True, center=True)
    cell(ws, f"D{r}", f"=ABS(INDEX(TB_Final,{i+1}))", NUM, link=True)
    cell(ws, f"E{r}", f'=IF(D{r}>Mat_PM,"Yes","No")', center=True)
    cell(ws, f"F{r}", qual.get(acct, ""), is_input=True)
    cell(ws, f"G{r}", f"=COUNTIF(Risk_SCOT,B{r})", "0", center=True)
    cell(ws, f"H{r}", asr.get(acct, ""), is_input=True); ws.merge_cells(f"H{r}:I{r}")
    cell(ws, f"J{r}", f'=IF(G{r}>0,"Significant - RMM identified",IF(E{r}="Yes","Material, not significant - stand-back & ISA 330.18","Not material"))', center=True)
    ws.merge_cells(f"J{r}:L{r}"); r += 1
cell(ws, f"A{r}", len(tbrows) + 1, center=True); cell(ws, f"B{r}", "Disclosures - related parties"); cell(ws, f"C{r}", "Disclosure", center=True)
cell(ws, f"D{r}", 25000, NUM, True); cell(ws, f"E{r}", "Qualitative", center=True); cell(ws, f"F{r}", "Material by nature (IAS 24)", is_input=True)
cell(ws, f"G{r}", f"=COUNTIF(Risk_SCOT,B{r})", "0", center=True); cell(ws, f"H{r}", "Completeness, Presentation", is_input=True); ws.merge_cells(f"H{r}:I{r}")
cell(ws, f"J{r}", f'=IF(G{r}>0,"Significant - RMM identified","Material, not significant - stand-back & ISA 330.18")', center=True); ws.merge_cells(f"J{r}:L{r}")
r += 1
dname("SCOT_Class", f"Risk_315R!$J${s0}:$J${r-1}")
ws.conditional_formatting.add(f"J{s0}:J{r-1}", FormulaRule(formula=[f'LEFT($J{s0},11)="Significant"'], fill=fill("FFE0B2")))
ws.conditional_formatting.add(f"J{s0}:J{r-1}", FormulaRule(formula=[f'LEFT($J{s0},8)="Material"'], fill=fill("FFF9C4")))
r += 1

r = section(ws, r, "RA3", "Risk register - inherent risk factors, spectrum & combined RMM", "ISA 315R.28, 31-34 | ISA 330.7")
heads = ["ID", "Risk of material misstatement (what could go wrong)", "SCOTABD", "Relevant assertion(s)", "Level", "Fraud risk?",
         "IRF: Complexity", "IRF: Subjectivity", "IRF: Change", "IRF: Uncertainty", "IRF: Bias / fraud", "IRF score (/15)", "Max IRF",
         "Dominant IRF", "Likelihood (1-5)", "Magnitude (1-5)", "L x M", "Judged significant? (upper end)", "Spectrum of inherent risk",
         "Consistency check", "Substantive alone insufficient? (315R.33)", "Control risk", "Combined RMM", "Planned response (ISA 330)", "WP ref"]
r = table_head(ws, r, heads)
HDR = r - 1
risks2 = [
    ("Revenue recorded before goods dispatched to meet targets (cut-off) - presumed fraud risk", "Revenue", "Occurrence, Cut-off", "Assertion", "Yes", (1, 1, 1, 0, 3), 4, 5, "Yes", "No", "Not tested", "Extended cut-off 10 days either side; credit notes post-YE; analytics by day", "D3, FR3"),
    ("Management override of controls - journals, estimates, unusual transactions (presumed)", "All (FS level)", "All", "FS level", "Yes", (1, 2, 0, 1, 3), 4, 5, "Yes", "No", "Not tested", "JE testing with CAATs; retrospective estimates review; business rationale of unusual transactions", "D6, FR4"),
    ("Inventory NRV understated for slow-moving imported goods (MNT depreciation)", "Inventory", "Valuation", "Assertion", "No", (2, 3, 2, 3, 2), 4, 4, "No", "No", "Not tested", "Ageing, post-YE selling prices, independent NRV range (ISA 540R)", "D3, C4"),
    ("Inventory does not exist at 3 branches & warehouse", "Inventory", "Existence", "Assertion", "No", (1, 0, 1, 0, 1), 3, 4, "No", "No", "Rely", "Count attendance + test counts; cycle-count controls (C05)", "D5"),
    ("Receivables overstated - ECL under-provided (IFRS 9)", "Trade receivables", "Valuation", "Assertion", "No", (2, 2, 1, 2, 2), 3, 3, "No", "No", "Not tested", "MUS confirmations; after-date cash; ECL matrix recompute", "D4, S4"),
    ("Accruals & payables incomplete at year end", "Accruals", "Completeness", "Assertion", "No", (1, 1, 0, 1, 2), 3, 3, "No", "No", "Rely", "Search for unrecorded liabilities (S1.5)", "D3"),
    ("Repairs capitalised as PPE / additions not existing", "Property, plant & equipment", "Existence, Classification", "Assertion", "No", (1, 1, 1, 0, 1), 2, 3, "No", "No", "Not tested", "MUS on additions; inspection", "D3, S1.3"),
    ("Borrowings misclassified; covenant breach not disclosed", "Borrowings", "Presentation", "Assertion", "No", (2, 1, 1, 1, 1), 2, 4, "No", "No", "Not tested", "Bank confirmation; covenant recalculation (GC3)", "D4, GC3"),
    ("Related party transactions / balances not identified or disclosed", "Disclosures - related parties", "Completeness, Presentation", "Assertion", "No", (1, 1, 0, 0, 3), 3, 2, "No", "No", "Not tested", "RP inquiries, minutes, ledger scans", "B7"),
    ("Cash balances do not exist", "Cash", "Existence", "Assertion", "No", (0, 0, 0, 0, 1), 1, 2, "No", "No", "Rely", "Bank confirmations & reconciliations", "D4"),
    ("Right-of-use assets & lease liabilities misstated (IFRS 16 - branch leases)", "Property, plant & equipment", "Valuation, Accuracy", "Assertion", "No", (3, 2, 1, 1, 1), 2, 3, "No", "No", "Not tested", "Recompute lease models; discount rate review", "D3"),
]
q0 = r
for i, (dsc, scot, asr_, lvl, frd, irf, L_, M_, sig, subo, cr, resp, wp) in enumerate(risks2, 1):
    cell(ws, f"A{r}", f"R{i:02d}", center=True); cell(ws, f"B{r}", dsc, is_input=True)
    cell(ws, f"C{r}", scot, is_input=True); cell(ws, f"D{r}", asr_, is_input=True)
    cell(ws, f"E{r}", lvl, is_input=True, center=True); cell(ws, f"F{r}", frd, is_input=True, center=True)
    for j, v in enumerate(irf): cell(ws, f"{chr(71+j)}{r}", v, "0", True, center=True)
    cell(ws, f"L{r}", f"=SUM(G{r}:K{r})", "0", center=True); cell(ws, f"M{r}", f"=MAX(G{r}:K{r})", "0", center=True)
    cell(ws, f"N{r}", f'=IF(M{r}=0,"None",SUBSTITUTE(INDEX($G${HDR}:$K${HDR},MATCH(M{r},G{r}:K{r},0)),"IRF: ",""))', center=True)
    cell(ws, f"O{r}", L_, "0", True, center=True); cell(ws, f"P{r}", M_, "0", True, center=True)
    cell(ws, f"Q{r}", f"=O{r}*P{r}", "0", center=True)
    cell(ws, f"R{r}", sig, is_input=True, center=True)
    cell(ws, f"S{r}", f'=IF(R{r}="Yes","Significant",IF(Q{r}>=15,"High",IF(Q{r}>=10,"Elevated",IF(Q{r}>=5,"Moderate","Low"))))', center=True)
    cell(ws, f"T{r}", f'=IF(AND(F{r}="Yes",R{r}<>"Yes"),"Fraud risk must be a significant risk (ISA 240R)",IF(AND(R{r}<>"Yes",Q{r}>=20),"Score at upper end - consider significant",'
                      f'IF(AND(M{r}=3,Q{r}<5),"High IRF but low IR - justify",IF(AND(R{r}="Yes",Q{r}<10),"Significant but low L x M - justify","OK"))))')
    cell(ws, f"U{r}", subo, is_input=True, center=True); cell(ws, f"V{r}", cr, is_input=True, center=True)
    cell(ws, f"W{r}", f'=IF(S{r}="Significant","High - significant risk",IF(OR(S{r}="High",S{r}="Elevated"),IF(V{r}="Rely","Medium","High"),IF(S{r}="Moderate",IF(V{r}="Rely","Low","Medium"),"Low")))', center=True)
    cell(ws, f"X{r}", resp, is_input=True); cell(ws, f"Y{r}", wp, center=True)
    ws.row_dimensions[r].height = 40; r += 1
q1 = r - 1
for col in ("F", "R", "U"): dv(ws, f"{col}{q0}:{col}{q1}", "YesNo")
dname("Risk_ID", f"Risk_315R!$A${q0}:$A${q1}"); dname("Risk_SCOT", f"Risk_315R!$C${q0}:$C${q1}"); dname("Risk_Assert", f"Risk_315R!$D${q0}:$D${q1}")
dname("Risk_Fraud", f"Risk_315R!$F${q0}:$F${q1}"); dname("Risk_MaxIRF", f"Risk_315R!$M${q0}:$M${q1}")
dname("Risk_L", f"Risk_315R!$O${q0}:$O${q1}"); dname("Risk_M", f"Risk_315R!$P${q0}:$P${q1}")
dname("Risk_Spectrum", f"Risk_315R!$S${q0}:$S${q1}"); dname("Risk_Check", f"Risk_315R!$T${q0}:$T${q1}")
dname("Risk_RMM", f"Risk_315R!$W${q0}:$W${q1}"); dname("Risk_Level", f"Risk_315R!$E${q0}:$E${q1}")
for band, colr in SPEC_COL.items():
    ws.conditional_formatting.add(f"S{q0}:S{q1}", CellIsRule(operator="equal", formula=[f'"{band}"'], fill=fill(colr),
                                  font=Font(name=F, bold=True, color="FFFFFF" if band in ("High", "Significant") else "000000")))
ws.conditional_formatting.add(f"G{q0}:K{q1}", CellIsRule(operator="equal", formula=["3"], fill=fill("FFCDD2"), font=Font(name=F, bold=True, color="B71C1C")))
ws.conditional_formatting.add(f"T{q0}:T{q1}", CellIsRule(operator="notEqual", formula=['"OK"'], fill=fill("FFE0B2")))
note(ws, r, "Enter IRF ratings 0-3, likelihood & magnitude 1-5, and your judgement on significance. 'Consistency check' flags assessments that need justification (e.g. fraud risks not treated as significant; high IRFs with low inherent risk). "
            "Control risk 'Rely' only where operating effectiveness is tested (IC4).", "Y")
r += 2

r = section(ws, r, "RA4", "Inherent risk heat map - likelihood x magnitude (number of risks)", "ISA 315R.31")
put(ws, f"B{r}", "Likelihood  ↓   /   Magnitude  →", font(9, True)); r += 1
put(ws, f"B{r}", "", font(9));
for m in range(1, 6): put(ws, f"{chr(66+m)}{r}", m, font(9, True, "FFFFFF"), DK, al=CEN, bd=BOX)
r += 1
for L_ in range(5, 0, -1):
    put(ws, f"B{r}", f"Likelihood {L_}", font(9, True, "FFFFFF"), DK, al=CEN, bd=BOX)
    for m in range(1, 6):
        p = L_ * m
        band = "High" if p >= 15 else ("Elevated" if p >= 10 else ("Moderate" if p >= 5 else "Low"))
        if p >= 20: band = "Significant"
        put(ws, f"{chr(66+m)}{r}", f'=IF(COUNTIFS(Risk_L,{L_},Risk_M,{m})=0,"",COUNTIFS(Risk_L,{L_},Risk_M,{m}))',
            font(12, True, "FFFFFF" if band in ("High", "Significant") else "000000"), SPEC_COL[band], "0", CEN, BOX)
    ws.row_dimensions[r].height = 26; r += 1
note(ws, r, "Cell colours show the spectrum band for each likelihood x magnitude combination; numbers are the count of risks in the register (RA3) at that position.", "L")
r += 2

r = section(ws, r, "RA5", "Risks at the financial statement level (pervasive) & overall responses", "ISA 315R.28(a), 30 | ISA 330.5")
r = table_head(ws, r, ["#", "FS-level risk", "Source / evidence", "", "Affects assessment of assertion-level risks?", "Overall response (ISA 330.5, A1-A3)", "", "", "", "Linked"])
ws.merge_cells(f"C{r-1}:D{r-1}"); ws.merge_cells(f"F{r-1}:I{r-1}")
for i, (a, b, c_, d_, e_) in enumerate([
        ("Management override of controls", "Owner-managed; no audit committee", "Yes", "Increased skepticism; unpredictable procedures; senior staff on JE testing", "R02, FR3"),
        ("Weak entity risk assessment process & monitoring", "IC1 - COSO principles 6-9, 16", "Yes", "Less reliance on entity-level controls; more substantive evidence", "IC1, CD01"),
        ("ITGC deficiency - user access", "IC3 / C06 deviation", "Yes", "Test IPE substantively; no reliance on automated control C03", "IC3, IC4"),
        ("Pressure from EBITDA-linked bonus", "FR2 fraud risk factor", "Yes", "Focus on revenue cut-off & estimates bias; partner involvement", "FR2")], 1):
    cell(ws, f"A{r}", i, center=True); cell(ws, f"B{r}", a, is_input=True); cell(ws, f"C{r}", b, is_input=True); ws.merge_cells(f"C{r}:D{r}")
    cell(ws, f"E{r}", c_, is_input=True, center=True); cell(ws, f"F{r}", d_, is_input=True); ws.merge_cells(f"F{r}:I{r}")
    cell(ws, f"J{r}", e_, center=True); ws.row_dimensions[r].height = 30; r += 1
r += 1

r = section(ws, r, "RA6", "Stand-back, evaluation of evidence & documentation", "ISA 315R.35-38")
r = checklist(ws, r, [
    ("Audit evidence from risk assessment procedures provides an appropriate basis for identifying & assessing RMM", "ISA 315R.35", "Yes", "Corroborative and contradictory evidence considered"),
    ("Stand-back: material classes/balances/disclosures NOT identified as significant re-evaluated (see RA2 'Material, not significant')", "ISA 315R.36", "Yes", "Cost of sales, trade payables, share capital - conclusion appropriate"),
    ("Substantive procedures designed for each material class/balance/disclosure regardless of assessed RMM", "ISA 330.18", "Yes", "C1 programme"),
    ("Risk assessment revised where new information is inconsistent with the original assessment", "ISA 315R.37", "Yes", "C03 & C06 reliance withdrawn after IC4"),
    ("Documentation: discussion, understanding, controls evaluated, RMM identified & assessed with rationale for significant judgements", "ISA 315R.38", "Yes", "This sheet + IC_Assessment"),
], "RA6")

# ---------------- FRAUD 240R ----------------
ws = S["Fraud_240R"]
r = sheet_header(ws, "FRAUD - ISA 240 (REVISED)", "2 & 4",
                 "Apply a fraud lens to risk assessment, respond to assessed risks of material misstatement due to fraud, address fraud or suspected fraud (incl. allegations) and communicate fraud-related matters throughout the audit.",
                 "ISA 240 (Revised) - issued Jul-2025, effective for periods beginning on/after 15-Dec-2026 (early adoption encouraged); ISA 315R; ISA 330; ISA 260; ISA 701; ISA 580")
note(ws, 7, "Paragraph numbers of ISA 240 (Revised) differ from extant ISA 240 - references below cite the section of the revised standard. Verify against the IAASB Handbook edition adopted locally.", "L", 14)
r = 9
r = section(ws, r, "FR1", "Professional skepticism & engagement team discussion", "ISA 240R - Professional skepticism; Engagement team discussion")
r = checklist(ws, r, [
    ("Professional skepticism maintained throughout - remain alert to information inconsistent with records or indicating fraud", "ISA 240R - Professional skepticism", "Yes", "Briefed at kick-off; reinforced at each review"),
    ("Records & documents not accepted as genuine without consideration; authenticity investigated where conditions indicate", "ISA 240R; ISA 500", "Yes", "Bank confirmations sent directly; signatures checked"),
    ("Responses to inquiries of management/TCWG investigated where inconsistent", "ISA 240R", "Yes", ""),
    ("Engagement team discussion held: how & where FS susceptible to fraud, incl. management override & use of technology to commit/conceal fraud", "ISA 240R - Engagement team discussion", "Yes", "15-Oct-25; minutes on file; IT auditor attended"),
    ("Engagement partner determined matters to communicate to members not at the discussion", "ISA 240R; ISA 220R", "Yes", "Late joiner briefed 3-Nov-25"),
], "FR1")

r = section(ws, r, "FR2", "Risk assessment with a fraud lens - fraud risk factors & inquiries", "ISA 240R - Risk assessment procedures & related activities")
r = table_head(ws, r, ["#", "Fraud risk factor", "Category", "Source", "Related IRF", "Linked RMM", "Present?", "Effect on assessment", "", "", ""])
ws.merge_cells(f"H{r-1}:K{r-1}")
frf = [("Management bonus linked to EBITDA target", "Incentive / pressure", "Inquiry; remuneration policy", "Bias / fraud", "R01, R03", "Yes", "Revenue & NRV estimates - bias focus"),
       ("Bank covenant (net debt / EBITDA) creates pressure on reported results", "Incentive / pressure", "Loan agreement", "Bias / fraud", "R08", "Yes", "Covenant headroom tested (GC3)"),
       ("Owner-manager dominance; no audit committee", "Opportunity", "Governance review", "Bias / fraud", "R02", "Yes", "FS-level risk (RA5)"),
       ("ERP user access reviews lapsed (C06)", "Opportunity", "IC3 / IC4", "Complexity", "R02", "Yes", "Unauthorised JE risk - JE criteria expanded"),
       ("Cash handling at branches with limited segregation", "Opportunity", "Walkthrough", "Bias / fraud", "R10", "Yes", "Surprise cash counts"),
       ("Undisclosed director loan - attitude toward disclosure", "Attitude / rationalisation", "Ledger scan", "Bias / fraud", "R09", "Yes", "RP procedures extended"),
       ("Aggressive revenue targets communicated to sales staff", "Incentive / pressure", "Inquiry of sales staff", "Bias / fraud", "R01", "Yes", "Cut-off extended to 10 days")]
f0 = r
for i, row in enumerate(frf, 1):
    cell(ws, f"A{r}", i, center=True)
    for j, v in enumerate(row[:6]): cell(ws, f"{chr(66+j)}{r}", v, is_input=True, center=j in (1, 3, 4, 5))
    cell(ws, f"H{r}", row[6], is_input=True); ws.merge_cells(f"H{r}:K{r}"); ws.row_dimensions[r].height = 28; r += 1
dv(ws, f"G{f0}:G{r-1}", "YesNo")
dname("FRF_Present", f"Fraud_240R!$G${f0}:$G${r-1}")
r += 1
put(ws, f"B{r}", "Inquiries regarding fraud (management, TCWG, others within the entity)", font(9, True, DK)); r += 1
r = table_head(ws, r, ["#", "Person inquired", "Role", "Date", "Matters covered", "", "", "Response / knowledge of fraud or allegations", "", "", ""])
ws.merge_cells(f"E{r-1}:G{r-1}"); ws.merge_cells(f"H{r-1}:K{r-1}")
for i, (p, role, d_, m, resp) in enumerate([
        ("CEO", "Management", D(2025, 10, 16), "Fraud risk assessment process; known/suspected fraud; whistleblower reports", "No known fraud; informal process"),
        ("Board chair", "TCWG", D(2025, 10, 20), "Oversight of fraud risk; allegations received", "Aware of branch 2 petty cash allegation (FR5)"),
        ("Chief accountant", "Financial reporting", D(2025, 10, 22), "Unusual JEs; pressure to adjust results", "None reported"),
        ("Sales staff (2)", "Others - operations", D(2025, 11, 20), "Pressure to meet targets; side agreements", "Year-end push acknowledged"),
        ("Warehouse manager", "Others - operations", D(2025, 12, 31), "Inventory losses; write-offs", "Damaged stock not yet written off")], 1):
    cell(ws, f"A{r}", i, center=True); cell(ws, f"B{r}", p); cell(ws, f"C{r}", role, center=True); cell(ws, f"D{r}", d_, DATE, True, center=True)
    cell(ws, f"E{r}", m, is_input=True); ws.merge_cells(f"E{r}:G{r}"); cell(ws, f"H{r}", resp, is_input=True); ws.merge_cells(f"H{r}:K{r}")
    ws.row_dimensions[r].height = 28; r += 1
r += 1
r = checklist(ws, r, [
    ("Understanding of entity's culture, management's commitment to integrity, and TCWG oversight of fraud risk obtained", "ISA 240R; ISA 315R.21", "Yes", "Code of conduct; no whistleblower policy - recommended"),
    ("Entity's process for identifying, assessing & responding to fraud risks (incl. whistleblower programme) understood", "ISA 240R; ISA 315R.22", "Yes", "Informal - deficiency CD01"),
    ("Unusual or unexpected relationships from analytical procedures evaluated for fraud indicators", "ISA 240R; ISA 520", "Yes", "Dec-25 revenue spike +38% vs Nov"),
    ("Fraud risk factors evaluated as part of inherent risk factor 'susceptibility to bias/fraud' (Risk_315R RA3)", "ISA 240R; ISA 315R.31", "Yes", "Linked above"),
], "FR2")

r = section(ws, r, "FR3", "Identified & assessed risks of material misstatement due to fraud", "ISA 240R - Identifying & assessing RMM due to fraud")
r = table_head(ws, r, ["", "Item", "Value / decision", "Rationale"])
ws.merge_cells(f"D{r-1}:L{r-1}")
g0 = r
fr3 = [("Risks flagged as fraud risks in register (RA3)", '=COUNTIF(Risk_Fraud,"Yes")', "0", "", False),
       ("...of which assessed as significant risks", '=COUNTIFS(Risk_Fraud,"Yes",Risk_Spectrum,"Significant")', "0", "All fraud RMMs must be treated as significant risks", False),
       ("Presumed fraud risk in revenue recognition rebutted?", "No", None, "Not rebutted - incentives (bonus, targets) and cut-off history (PY & CY misstatement M01)", True),
       ("Management override - always a significant risk (cannot be rebutted)", "Significant", None, "R02", False),
       ("Fraud risk factors identified (FR2)", '=COUNTIF(FRF_Present,"Yes")', "0", "", False),
       ("CONCLUSION", f'=IF(C{g0+1}<C{g0},"ERROR - fraud RMM not assessed as significant: revise RA3",IF(C{g0+2}="Yes","Rebuttal documented - re-evaluate at completion","Fraud RMMs appropriately identified & treated as significant"))', None, "", False)]
for i, (lab, v, nf, rat, isin) in enumerate(fr3):
    b = cell(ws, f"B{r}", lab); c = cell(ws, f"C{r}", v, nf, isin, center=True)
    if lab == "CONCLUSION": b.font = font(9, True); c.font = font(9, True, DK); ws.merge_cells(f"C{r}:L{r}")
    else: cell(ws, f"D{r}", rat, is_input=isin); ws.merge_cells(f"D{r}:L{r}")
    r += 1
dv(ws, f"C{g0+2}", "YesNo")
r += 1

r = section(ws, r, "FR4", "Responses - overall, management override (JE testing, estimates, unusual transactions)", "ISA 240R - Responses to assessed RMM due to fraud | ISA 330")
r = checklist(ws, r, [
    ("Overall responses: assignment & supervision, element of unpredictability, evaluation of accounting policies (esp. subjective measurements)", "ISA 240R; ISA 330.5", "Yes", "Unannounced branch 3 cash count; new sample stratification"),
    ("Retrospective review of management judgements & assumptions in PY estimates for bias", "ISA 240R; ISA 540R", "Yes", "PY NRV under-provided 21,000 (SUAD2)"),
    ("Business rationale of significant transactions outside normal course of business / otherwise unusual evaluated", "ISA 240R", "Yes", "None identified"),
    ("Use of automated tools & techniques (CAATs) considered for JE testing & data analytics", "ISA 240R; ISA 500", "Yes", "Full-population JE analytics in data tool"),
], "FR4")
put(ws, f"B{r}", "Journal entry testing - risk criteria (population: 48,320 entries reconciled to TB)", font(9, True, DK)); r += 1
r = table_head(ws, r, ["#", "Risk criterion", "Entries meeting criterion", "Selected for testing", "Tested", "Exceptions", "Result"])
jec = [("Post-closing entries (after 31-Dec-25)", 18, 18, 18, 0), ("Entries on weekends / public holidays", 64, 22, 22, 0),
       ("Round-amount entries >= MNT 10m", 41, 15, 15, 0), ("Entries by unusual or senior users (CEO, IT admin)", 9, 9, 9, 1),
       ("Entries to seldom-used accounts", 27, 11, 11, 0), ("Revenue / receivables entries with unusual offsets (cash, suspense)", 33, 20, 20, 0),
       ("Descriptions with keywords: 'adjust', 'correction', 'per CEO', blank", 52, 17, 17, 0)]
j0 = r
for i, (crit, n_, sel, tst, exc) in enumerate(jec, 1):
    cell(ws, f"A{r}", i, center=True); cell(ws, f"B{r}", crit)
    for col, v in zip("CDEF", (n_, sel, tst, exc)): cell(ws, f"{col}{r}", v, "#,##0", True, center=True)
    cell(ws, f"G{r}", f'=IF(E{r}<D{r},"Incomplete",IF(F{r}=0,"No exceptions","Exception - investigate (FR5)"))', center=True); r += 1
put(ws, f"B{r}", "TOTAL", font(9, True), PALE, bd=BOX)
for col in "CDEF": put(ws, f"{col}{r}", f"=SUM({col}{j0}:{col}{r-1})", font(9, True), PALE, "#,##0", CEN, BOX)
dname("JE_Tested", f"Fraud_240R!$E${r}"); dname("JE_Exc", f"Fraud_240R!$F${r}")
ws.conditional_formatting.add(f"G{j0}:G{r-1}", FormulaRule(formula=[f'LEFT($G{j0},9)="Exception"'], fill=fill("FFCDD2")))
r += 1
note(ws, r, "Exception (criterion 4): IT admin posted a reclassification of 3,500 - supported by approved request; no fraud indication. Documented in D6.", "L", 14)
r += 2

r = section(ws, r, "FR5", "Fraud or suspected fraud identified (incl. allegations)", "ISA 240R - Fraud or suspected fraud")
r = table_head(ws, r, ["ID", "Date identified", "Source", "Description (allegations treated as suspected fraud)", "Involves mgmt / key IC roles?",
                       "Procedures performed to understand", "Effect on RMM & audit", "Misstatement (SUAD)", "Communicated to mgmt", "Communicated to TCWG", "Report to authority?", "Status"])
frows = [("FR5.1", D(2025, 10, 20), "Allegation (whistleblower email to Board)", "Branch 2 petty cash custodian misappropriating cash", "No",
          "Surprise count; reconciliation; inquiry of custodian & branch manager", "Isolated; cash RMM unchanged; ISA 450 accumulation", "M06 (1,200 - trivial)", D(2025, 11, 5), D(2025, 11, 7), "No", "Closed"),
         ("FR5.2", D(2026, 2, 24), "Audit procedure (JE criterion 4)", "IT admin posting journal entries", "Yes",
          "Inspected approval, change ticket; inquiry of CFO", "No fraud; access rights restricted (C06)", "-", D(2026, 2, 26), D(2026, 3, 20), "No", "Closed"),
         ("FR5.3", D(2026, 3, 10), "Audit procedure (RP ledger scan)", "Undisclosed director loan - possible concealment", "Yes",
          "Inquiry of director & Board chair; loan agreement inspected", "Evaluate integrity of mgmt reps; qualitative materiality", "M08 (disclosure)", D(2026, 3, 12), None, "No", "Open")]
h0 = r
for row in frows:
    for j, v in enumerate(row):
        col = chr(65 + j)
        cell(ws, f"{col}{r}", v, DATE if j in (1, 8, 9) else None, is_input=j > 0, center=j in (0, 1, 4, 7, 8, 9, 10, 11))
    ws.row_dimensions[r].height = 40; r += 1
dv(ws, f"E{h0}:E{r-1}", "YesNo"); dv(ws, f"K{h0}:K{r-1}", "YesNo"); dv(ws, f"L{h0}:L{r-1}", "IssueStatus")
dname("Fraud_Status", f"Fraud_240R!$L${h0}:$L${r-1}"); dname("Fraud_TCWG", f"Fraud_240R!$J${h0}:$J${r-1}")
dname("Fraud_Mgmt", f"Fraud_240R!$E${h0}:$E${r-1}")
ws.conditional_formatting.add(f"J{h0}:J{r-1}", FormulaRule(formula=[f'AND($E{h0}="Yes",$J{h0}="")'], fill=fill("FFCDD2")))
note(ws, r, "Where fraud or suspected fraud involves management, employees with significant roles in internal control, or results in a material misstatement, communicate with TCWG on a timely basis; consider legal/regulatory reporting duties and whether to withdraw. "
            "Red cell = management involvement not yet communicated to TCWG.", "L")
r += 2

r = section(ws, r, "FR6", "Ongoing communication of fraud-related matters with management & TCWG", "ISA 240R - Communications | ISA 260")
r = table_head(ws, r, ["#", "Date", "With", "Matter communicated", "", "", "", "Form", "Ref"])
ws.merge_cells(f"D{r-1}:G{r-1}")
for i, (d_, w, m, frm, ref) in enumerate([
        (D(2025, 10, 20), "TCWG", "Planned approach to fraud risks incl. presumed revenue & override risks", "Written", "A7"),
        (D(2025, 11, 7), "TCWG", "Results of branch 2 allegation investigation", "Meeting", "FR5.1"),
        (D(2026, 2, 26), "Management", "JE exception - IT admin access to post journals", "Oral", "FR5.2"),
        (D(2026, 3, 12), "Management", "Undisclosed director loan", "Meeting", "FR5.3"),
        (None, "TCWG", "Fraud-related matters at completion incl. FR5.3 and control deficiencies", "Written", "F1")], 1):
    cell(ws, f"A{r}", i, center=True); cell(ws, f"B{r}", d_, DATE, True, center=True); cell(ws, f"C{r}", w, center=True)
    cell(ws, f"D{r}", m, is_input=True); ws.merge_cells(f"D{r}:G{r}"); cell(ws, f"H{r}", frm, center=True); cell(ws, f"I{r}", ref, center=True); r += 1
r += 1

r = section(ws, r, "FR7", "Auditor's report, written representations & documentation", "ISA 240R; ISA 701; ISA 580; ISA 230")
r = checklist(ws, r, [
    ("Auditor's report describes the auditor's responsibilities relating to fraud", "ISA 240R; ISA 700R", None, "Use revised illustrative wording"),
    ("Listed / publicly traded entity: determine which fraud-related matters are KAMs and describe them", "ISA 240R; ISA 701", "N/A", "Non-listed - voluntary KAM on revenue cut-off (F4)"),
    ("Written representations: responsibility for fraud prevention/detection; disclosure of fraud, suspected fraud & allegations", "ISA 240R; ISA 580", None, "F2"),
    ("Evaluated whether misstatements indicate fraud and implications for reliability of representations", "ISA 240R; ISA 450", "Yes", "M01 cut-off - pressure-driven; reps scrutinised"),
    ("Considered whether to withdraw from the engagement where legally possible", "ISA 240R", "N/A", ""),
    ("Documentation: discussion, fraud risk factors, RMM due to fraud & responses, JE testing, fraud identified, communications", "ISA 240R; ISA 230", "Yes", "This sheet"),
], "FR7")

# ---------------- GOING CONCERN 570R ----------------
ws = S["GoingConcern_570R"]
r = sheet_header(ws, "GOING CONCERN - ISA 570 (REVISED 2024)", "2 & 6",
                 "Identify events or conditions through risk assessment, evaluate management's assessment (at least 12 months from approval of the FS), conclude on the going concern basis & material uncertainty, communicate and report.",
                 "ISA 570 (Revised 2024) - issued Apr-2025, effective for periods beginning on/after 15-Dec-2026 (early adoption encouraged); ISA 315R; ISA 540R; ISA 260; ISA 700R; ISA 580")
note(ws, 7, "References cite sections of ISA 570 (Revised 2024); paragraph numbering differs from extant ISA 570 (Revised). Verify against the locally adopted handbook.", "L", 14)
r = 9
r = section(ws, r, "GC1", "Risk assessment - events or conditions & entity's going concern process", "ISA 570R - Risk assessment procedures | ISA 315R")
r = table_head(ws, r, ["#", "Event or condition", "Category", "Present?", "Evidence / comment", "", "", "", "", "Source"])
ws.merge_cells(f"E{r-1}:I{r-1}")
evc = [("Net liability or net current liability position", "Financial", "No", "Current ratio 2.3x"),
       ("Borrowings approaching maturity without realistic renewal prospects", "Financial", "No", "Loan matures 2028"),
       ("Negative operating cash flows (historical or forecast)", "Financial", "No", "Operating cash inflow 1.9bn MNT"),
       ("Adverse key financial ratios / covenant breach or near-breach", "Financial", "Yes", "Net debt/EBITDA headroom narrow under stress (GC3)"),
       ("Substantial operating losses or significant deterioration in asset values", "Financial", "No", ""),
       ("Inability to pay creditors on due dates / arrears", "Financial", "No", "Payable days stable"),
       ("Loss of key management, major market, key customer or supplier", "Operating", "Yes", "Key supplier (35% of imports) under sanctions review"),
       ("Labour difficulties or shortages of important supplies", "Operating", "No", ""),
       ("Emergence of a highly successful competitor", "Operating", "No", ""),
       ("Non-compliance with capital or other statutory/regulatory requirements", "Other", "No", ""),
       ("Pending legal proceedings that could result in claims the entity cannot satisfy", "Other", "No", "One customer claim - immaterial (D7)"),
       ("Changes in law, regulation or government policy (e.g. import duties, FX controls)", "Other", "No", "")]
e0 = r
for i, (ev_, cat, pres, com) in enumerate(evc, 1):
    cell(ws, f"A{r}", i, center=True); cell(ws, f"B{r}", ev_); cell(ws, f"C{r}", cat, center=True)
    cell(ws, f"D{r}", pres, is_input=True, center=True); cell(ws, f"E{r}", com, is_input=True); ws.merge_cells(f"E{r}:I{r}")
    cell(ws, f"J{r}", "Inquiry / analytics", center=True); ws.row_dimensions[r].height = 26; r += 1
dv(ws, f"D{e0}:D{r-1}", "YesNo")
ws.conditional_formatting.add(f"D{e0}:D{r-1}", CellIsRule(operator="equal", formula=['"Yes"'], fill=fill("FFE0B2")))
put(ws, f"B{r}", "Events or conditions identified", font(9, True), PALE, bd=BOX)
put(ws, f"D{r}", f'=COUNTIF(D{e0}:D{r-1},"Yes")', font(10, True, DK), PALE, "0", CEN, BOX)
dname("GC_EventsN", f"GoingConcern_570R!$D${r}")
r += 2
r = checklist(ws, r, [
    ("Understanding of entity's process for its going concern assessment obtained (method, assumptions, data, controls, use of experts)", "ISA 570R; ISA 315R.19-26", "Yes", "Board-approved 15-month cash forecast"),
    ("Considered whether events or conditions exist that may cast significant doubt - throughout the audit, not only at planning", "ISA 570R", "Yes", "2 items identified above"),
    ("Discussed going concern basis and events/conditions with management; obtained management's preliminary assessment", "ISA 570R", "Yes", "CFO meeting 14-Nov-25"),
    ("Engagement team discussion included susceptibility to going concern issues & management bias", "ISA 570R; ISA 315R.17", "Yes", "15-Oct-25"),
], "GC1")

r = section(ws, r, "GC2", "Evaluating management's assessment of going concern", "ISA 570R - Evaluating management's assessment | ISA 540R")
g0 = r
dates = [("Date of approval of the financial statements", "=Set_RptDate", DATE, False),
         ("End of period covered by management's assessment", D(2027, 6, 30), DATE, True),
         ("Months covered from date of approval", f"=(YEAR(C{g0+1})-YEAR(C{g0}))*12+MONTH(C{g0+1})-MONTH(C{g0})", "0", False),
         ("Assessment period test (>= 12 months from approval)", f'=IF(C{g0+2}>=12,"Meets requirement","SHORTER THAN 12 MONTHS - request management to extend")', None, False)]
for i, (lab, v, nf, isin) in enumerate(dates):
    cell(ws, f"B{r}", lab); c = cell(ws, f"C{r}", v, nf, isin, link=(i == 0), center=True)
    if i == 3: ws.merge_cells(f"C{r}:H{r}"); c.font = font(9, True, DK)
    r += 1
dname("GC_Months", f"GoingConcern_570R!$C${g0+2}")
r += 1
r = table_head(ws, r, ["#", "Significant assumption", "Management's assumption", "Auditor's evidence", "", "", "Sensitivity tested", "", "Reasonable?"])
ws.merge_cells(f"D{r-1}:F{r-1}"); ws.merge_cells(f"G{r-1}:H{r-1}")
for i, (a, m, ev_, sens, ok) in enumerate([
        ("Revenue growth FY26", "+6%", "Q1-26 actual +5%; order book; market data", "-10% receipts (GC3)", "Yes"),
        ("Gross margin", "35%", "3-year average 34.6%", "-3 pp", "Yes"),
        ("Bank facility renewal / availability", "Undrawn 1,000,000 available", "Bank confirmation; facility letter to 2028", "Facility withdrawn", "Yes"),
        ("Alternative supplier if key supplier sanctioned", "Switch within 3 months, +4% cost", "Draft contract with second supplier", "6-month disruption", "Yes"),
        ("MNT/USD exchange rate", "3,450", "Central bank forecast range", "+10% depreciation", "Yes")], 1):
    cell(ws, f"A{r}", i, center=True); cell(ws, f"B{r}", a); cell(ws, f"C{r}", m, is_input=True, center=True)
    cell(ws, f"D{r}", ev_, is_input=True); ws.merge_cells(f"D{r}:F{r}"); cell(ws, f"G{r}", sens, is_input=True); ws.merge_cells(f"G{r}:H{r}")
    cell(ws, f"I{r}", ok, is_input=True, center=True); ws.row_dimensions[r].height = 26; r += 1
r += 1
r = checklist(ws, r, [
    ("Method used by management is appropriate and consistently applied", "ISA 570R; ISA 540R", "Yes", ""),
    ("Significant assumptions reasonable & supportable; sensitivity / stress tests evaluated", "ISA 570R", "Yes", "See GC3"),
    ("Data used is relevant & reliable (incl. IPE)", "ISA 570R; ISA 500", "Yes", "Forecast agreed to budgets & actuals"),
    ("Management's plans for future actions feasible; intent & ability evaluated", "ISA 570R", "Yes", "Supplier switch plan evidenced"),
    ("Financial support by third parties / owners: intent AND ability of the supporter evaluated", "ISA 570R", "N/A", "No support letter relied upon"),
    ("Indicators of possible management bias evaluated; contradictory evidence considered", "ISA 570R; ISA 540R.32", "Yes", "Optimistic growth vs. sector - within range"),
    ("Additional procedures performed where events/conditions identified (ISA 570R - when events or conditions identified)", "ISA 570R", "Yes", "Stress test, bank confirmation, post-YE trading"),
], "GC2")

r = section(ws, r, "GC3", "Cash-flow forecast, headroom & stress test (MNT '000)", "ISA 570R | management's forecast")
r = table_head(ws, r, ["", "Line", "Q2-26", "Q3-26", "Q4-26", "Q1-27", "Q2-27", "Total"])
c0 = r
lines = [("Opening cash", None), ("Operating receipts", [4900000, 4700000, 5300000, 4500000, 4950000]),
         ("Operating payments", [-4450000, -4300000, -4750000, -4200000, -4500000]), ("Capital expenditure", [-150000, -150000, -150000, -150000, -150000]),
         ("Debt service (interest + principal)", [-95000, -95000, -595000, -95000, -95000]), ("Closing cash - BASE", None),
         ("Undrawn facility", [1000000] * 5), ("Liquidity headroom - BASE", None),
         ("Stress: reduction in operating receipts", None), ("Closing cash - STRESS (cumulative)", None), ("Liquidity headroom - STRESS", None)]
rows_ = {}
for i, (lab, vals) in enumerate(lines):
    rr = r + i; rows_[lab] = rr
    b = cell(ws, f"B{rr}", lab)
    for k in range(5):
        col = chr(67 + k)
        if lab == "Opening cash": v = '=INDEX(TB_Final,MATCH("Cash",TB_Acct,0))' if k == 0 else f"={chr(66+k)}{c0+5}"
        elif lab == "Closing cash - BASE": v = f"=SUM({col}{c0}:{col}{c0+4})"
        elif lab == "Liquidity headroom - BASE": v = f"={col}{c0+5}+{col}{c0+6}"
        elif lab == "Stress: reduction in operating receipts": v = f"=-{col}{c0+1}*$C${c0+12}"
        elif lab == "Closing cash - STRESS (cumulative)": v = f"={col}{c0+5}+SUM($C{c0+8}:{col}{c0+8})"
        elif lab == "Liquidity headroom - STRESS": v = f"={col}{c0+9}+{col}{c0+6}"
        else: v = vals[k]
        cc = cell(ws, f"{col}{rr}", v, NUM, is_input=vals is not None and lab != "Undrawn facility" or lab == "Undrawn facility", link=(lab == "Opening cash" and k == 0))
        if "headroom" in lab.lower() or "Closing" in lab: cc.font = font(9, True)
    if lab in ("Operating receipts", "Operating payments", "Capital expenditure", "Debt service (interest + principal)", "Stress: reduction in operating receipts"):
        cell(ws, f"H{rr}", f"=SUM(C{rr}:G{rr})", NUM)
    if "headroom" in lab.lower() or "Closing" in lab: b.font = font(9, True); b.fill = fill(PALE)
r += len(lines) + 1
st0 = r
stress = [("Stress: operating receipts reduction %", 0.10, PCT, True),
          ("Minimum liquidity headroom - base", f"=MIN(C{c0+7}:G{c0+7})", NUM, False),
          ("Minimum liquidity headroom - stress", f"=MIN(C{c0+10}:G{c0+10})", NUM, False),
          ("EBITDA (PBT + finance costs + depreciation)", f'=TB_PBT_Final+INDEX(TB_Final,MATCH("Finance costs",TB_Acct,0))+C{r+4}', NUM, False),
          ("Depreciation & amortisation (input)", 850000, NUM, True),
          ("Net debt (borrowings - cash)", '=-INDEX(TB_Final,MATCH("Borrowings",TB_Acct,0))-INDEX(TB_Final,MATCH("Cash",TB_Acct,0))', NUM, False),
          ("Net debt / EBITDA (x)", f"=IFERROR(C{r+5}/C{r+3},0)", "0.00", False),
          ("Covenant maximum (x)", 1.5, "0.00", True),
          ("Covenant headroom", f'=IF(C{r+6}<=C{r+7},"Met - headroom "&TEXT(1-C{r+6}/C{r+7},"0%"),"BREACH - event/condition (ISA 570R)")', None, False)]
for i, (lab, v, nf, isin) in enumerate(stress):
    cell(ws, f"B{r}", lab); c = cell(ws, f"C{r}", v, nf, isin, center=True)
    if i == 8: ws.merge_cells(f"C{r}:H{r}"); c.font = font(9, True, DK)
    r += 1
assert st0 == c0 + 12
dname("GC_MinStress", f"GoingConcern_570R!$C${st0+2}")
ws.conditional_formatting.add(f"C{c0+10}:G{c0+10}", CellIsRule(operator="lessThan", formula=["0"], fill=fill("FFCDD2")))
r += 1

r = section(ws, r, "GC4", "Conclusions - going concern basis & material uncertainty", "ISA 570R - Auditor conclusions")
k0 = r
concl = [("Events or conditions identified (GC1)?", '=IF(GC_EventsN>0,"Yes","No")', False),
         ("Use of going concern basis of accounting appropriate?", "Yes", True),
         ("Material uncertainty (MU) related to going concern exists?", "No", True),
         ("Disclosures adequate (incl. MU or 'close call' disclosures)?", "Yes", True),
         ("Close call (events/conditions identified but no MU)?", f'=IF(AND(C{k0}="Yes",C{k0+2}="No"),"Yes","No")', False),
         ("Management unwilling to make / extend assessment?", "No", True),
         ("REPORTING OUTCOME", f'=IF(C{k0+1}="No","Adverse opinion - going concern basis inappropriate",IF(C{k0+5}="Yes","Consider qualified / disclaimer - scope limitation",'
                               f'IF(C{k0+2}="Yes",IF(C{k0+3}="Yes","Unmodified + Material Uncertainty Related to Going Concern section","Qualified / adverse - inadequate MU disclosure"),'
                               f'IF(C{k0+3}="Yes","Unmodified - Going Concern section (no MU identified)","Qualified - inadequate close-call disclosure"))))', False)]
for i, (lab, v, isin) in enumerate(concl):
    b = cell(ws, f"B{r}", lab); c = cell(ws, f"C{r}", v, None, isin, center=(i < 6))
    if i == 6: b.font = font(9, True); c.font = font(10, True, DK); b.fill = fill(PALE); c.fill = fill(PALE); ws.merge_cells(f"C{r}:J{r}")
    r += 1
for k in (1, 2, 3, 5): dv(ws, f"C{k0+k}", "YesNo")
dname("GC_MU", f"GoingConcern_570R!$C${k0+2}"); dname("GC_Outcome", f"GoingConcern_570R!$C${k0+6}")
dname("GC_Close", f"GoingConcern_570R!$C${k0+4}"); dname("GC_Basis", f"GoingConcern_570R!$C${k0+1}")
r += 1

r = section(ws, r, "GC5", "Draft 'Going Concern' section of the auditor's report (auto-generated)", "ISA 570R - Implications for the auditor's report | ISA 700R")
txt = (f'=IF(GC_MU="Yes","Material Uncertainty Related to Going Concern - We draw attention to Note [X] in the financial statements, which indicates that [events/conditions]. '
       f'As stated in Note [X], these events or conditions indicate that a material uncertainty exists that may cast significant doubt on the Entity\'s ability to continue as a going concern. '
       f'Our opinion is not modified in respect of this matter.",'
       f'"Going Concern - We have concluded that management\'s use of the going concern basis of accounting in the preparation of the financial statements is appropriate. '
       f'Based on the audit evidence obtained, we have not identified a material uncertainty related to events or conditions that, individually or collectively, may cast significant doubt on the Entity\'s ability to continue as a going concern '
       f'for a period of at least twelve months from the date on which the financial statements are authorised for issue."'
       f')&IF(AND(GC_Close="Yes",ISNUMBER(SEARCH("listed",Set_EntityType))),"  [Listed/PTE: describe how we evaluated management\'s assessment, incl. key events or conditions identified.]","")')
c = put(ws, f"B{r}", txt, font(9, i=True), "F1F8E9", al=WRAP, bd=BOX); ws.merge_cells(f"B{r}:L{r+3}")
for rr in range(r, r + 4): ws.row_dimensions[rr].height = 22
r += 5
note(ws, r, "Under ISA 570 (Revised 2024) every auditor's report includes a going concern section. Illustrative wording only - align with the revised ISA 700/570 illustrative reports and local law.", "L", 14)
r += 2

r = section(ws, r, "GC6", "Communication with TCWG, written representations & documentation", "ISA 570R; ISA 260; ISA 580; ISA 230")
r = checklist(ws, r, [
    ("Communicated to TCWG: whether events/conditions identified that may cast significant doubt (incl. close calls)", "ISA 570R; ISA 260", None, "Completion meeting"),
    ("Communicated: whether MU exists, appropriateness of GC basis, adequacy of disclosures, and implications for the report", "ISA 570R", None, ""),
    ("Communicated any significant delay in approval of the FS and reasons", "ISA 570R", "N/A", ""),
    ("Written representations on plans for future action & feasibility, and on going concern assessment", "ISA 570R; ISA 580", None, "F2"),
    ("Documentation: events/conditions, evaluation of management's assessment, judgements on MU & disclosures, reporting decision", "ISA 570R; ISA 230", "Yes", "This sheet"),
], "GC6")

# ---------------- DEFICIENCIES 265 ----------------
ws = S["Deficiencies_265"]
r = sheet_header(ws, "DEFICIENCIES IN INTERNAL CONTROL - ISA 265", "5",
                 "Determine whether deficiencies in internal control exist, evaluate whether they are significant individually or in combination, and communicate them appropriately and on a timely basis to TCWG and management.",
                 "ISA 265 paras 5-11 & A1-A30; ISA 315 (Revised 2019); ISA 330.17; ISA 260; ISA 230")
for k in range(ci("M"), ci("AA") + 1): ws.column_dimensions[gl(k)].width = 12
widths(ws, {"B": 40, "C": 12, "D": 16, "E": 13, "F": 14, "G": 11, "S": 34, "T": 34})

r = section(ws, r, "DF1", "Definitions & evaluation framework", "ISA 265.6, 8 & A5-A11")
defs_ = [("Deficiency in internal control (265.6(a))", "A control is designed, implemented or operated so that it cannot prevent, or detect and correct, misstatements on a timely basis; OR a control necessary to do so is missing."),
         ("Significant deficiency (265.6(b))", "A deficiency or combination of deficiencies that, in the auditor's professional judgement, is of sufficient importance to merit the attention of TCWG."),
         ("Evaluation (265.8, A6)", "Consider: likelihood of leading to material misstatement; susceptibility of the asset/liability to loss or fraud; subjectivity & complexity; amounts exposed; volume of activity; importance of the control; cause & frequency of exceptions; interaction with other deficiencies."),
         ("Indicators of significant deficiency (A7)", "Ineffective control environment (e.g. fraud by management, significant transactions with related parties not appropriately scrutinised); absence of risk assessment process where expected; ineffective response to significant risks; "
          "misstatements detected by the auditor not prevented/detected by IC; restatement for error/fraud; management's failure to remediate previously communicated significant deficiencies.")]
for lab, txt_ in defs_:
    cell(ws, f"B{r}", lab).font = font(9, True); cell(ws, f"C{r}", txt_); ws.merge_cells(f"C{r}:L{r}"); ws.row_dimensions[r].height = 36; r += 1
note(ws, r, "Scoring used in DF2: Likelihood that the deficiency leads to a material misstatement (1 remote, 2 reasonably possible, 3 probable) x Magnitude of potential misstatement (1 clearly below PM, 2 approaching PM, 3 could exceed OM). "
            "Classification: SIGNIFICANT if an A7 indicator is present, OR score >= 6 with no effective compensating control. Other deficiencies scoring >= 3 are 'of sufficient importance' to communicate to management (265.10(b)).")
r += 2

r = section(ws, r, "DF2", "Deficiency register - identification & evaluation", "ISA 265.7-8")
heads = ["Ref", "Deficiency", "Source WP", "COSO component", "Type", "Process / area", "Linked risk / misst.", "Likelihood (1-3)",
         "Magnitude (1-3)", "Score", "Compensating control effective?", "A7 indicator present?", "Repeat from PY?", "Combination group",
         "Classification (265.8)", "Communicate to", "Form", "Potential effect on FS", "Recommendation", "Date communicated", "Timeliness",
         "Mgmt response", "Owner", "Remediation due", "Status", "Letter rank", "Letter key"]
r = table_head(ws, r, heads)
dfs = [("No formal entity risk assessment process", "IC1", "Risk assessment", "Design", "Entity level", "R02", 2, 2, "No", "No", "No", "Governance",
        "Business & fraud risks may not be identified and addressed in a timely manner", "Board to hold an annual documented risk workshop and maintain a risk register", None, "Agreed - 2026 workshop", "CEO", D(2026, 6, 30), "In progress"),
       ("ERP user access reviews not performed in Q2 & Q3 2025 (ITGC)", "IC3 / IC4 C06", "Control activities", "Operating", "IT", "R02", 3, 3, "No", "No", "No", "IT & access",
        "Unauthorised access could permit unapproved journals or master-data changes, undermining automated controls and system reports", "Perform and evidence quarterly access reviews; remove leavers within 24 hours; restrict IT admin posting rights", D(2026, 3, 20), "Agreed", "IT Manager", D(2026, 3, 31), "In progress"),
       ("Three-way match overridden without approval during system outage (1 of 30)", "IC4 C03", "Control activities", "Operating", "Purchases", "R06", 2, 2, "Yes", "No", "No", "IT & access",
        "Payments for goods not received or at incorrect prices", "Require CFO approval for any manual override; log outages", D(2026, 3, 20), "Agreed", "AP Supervisor", D(2026, 4, 30), "Open"),
       ("No documented NRV / slow-moving inventory policy; ageing review not performed (C10)", "IC2 C10 / SUAD M02", "Control activities", "Design", "Inventory", "R03 / M02", 3, 3, "No", "Yes", "Yes", "Inventory",
        "Inventory overstated; PY and CY NRV under-provisions (21,000 / 38,000)", "Adopt NRV policy; monthly ageing report reviewed by CFO with documented provisioning", D(2026, 3, 20), "Agreed", "CFO", D(2026, 4, 30), "Open"),
       ("Accrual checklist not updated for new contracts", "SUAD M03", "Control activities", "Operating", "Financial close", "R06 / M03", 2, 2, "Yes", "No", "No", "Financial close",
        "Expenses & liabilities understated at period end", "Update checklist monthly from contract register", D(2026, 3, 20), "Agreed", "Chief Accountant", D(2026, 3, 31), "Closed"),
       ("Revenue cut-off: dispatch notes not matched to invoices at year end", "SUAD M01", "Control activities", "Design", "Revenue", "R01 / M01", 3, 3, "No", "Yes", "No", "Revenue",
        "Revenue and receivables overstated (M01 45,000 detected by auditor, not by entity)", "System lock: invoice date = GDN date; YE cut-off review by Financial Controller", D(2026, 3, 20), "Agreed", "CFO", D(2026, 3, 31), "In progress"),
       ("Inadequate segregation: cash receipting & receivables posting (C12)", "IC2 C12", "Control activities", "Design", "Revenue", "R05", 2, 2, "Yes", "No", "No", "Revenue",
        "Receipts could be misappropriated and concealed (lapping)", "Separate duties or monthly independent review of customer statements", None, "", "CFO", D(2026, 6, 30), "Open"),
       ("Branch petty cash custody & counts", "FR5.1 / SUAD M06", "Control activities", "Operating", "Cash", "R10 / M06", 1, 1, "No", "No", "No", "Cash",
        "Small cash losses", "Surprise monthly counts by branch manager", None, "", "Branch Mgr", D(2026, 4, 30), "Open"),
       ("Related party register incomplete; director loan not identified for disclosure", "FR5.3 / SUAD M08", "Control environment", "Design", "Disclosures", "R09 / M08", 2, 3, "No", "Yes", "No", "Governance",
        "Related party disclosures (IAS 24) incomplete - material by nature", "Annual RP declarations by directors; Board approval of all RP transactions", D(2026, 3, 20), "Under discussion", "Board", D(2026, 3, 31), "Open"),
       ("No audit committee; limited Board oversight of financial reporting", "IC1 principle 2", "Control environment", "Design", "Entity level", "R02", 2, 2, "No", "No", "No", "Governance",
        "Reduced oversight of management override & financial reporting", "Consider audit committee or finance sub-committee of the Board", None, "", "Board", D(2026, 12, 31), "Open")]
d0 = r
for i, row in enumerate(dfs, 1):
    (dsc, src, comp, typ, proc, lnk, L_, M_, cmp_, a7, rep_, grp, eff, rec, dcom, resp, own, due, stt) = row
    cell(ws, f"A{r}", f"CD{i:02d}", center=True); cell(ws, f"B{r}", dsc, is_input=True); cell(ws, f"C{r}", src, center=True)
    cell(ws, f"D{r}", comp, is_input=True, center=True); cell(ws, f"E{r}", typ, is_input=True, center=True)
    cell(ws, f"F{r}", proc, is_input=True, center=True); cell(ws, f"G{r}", lnk, center=True)
    cell(ws, f"H{r}", L_, "0", True, center=True); cell(ws, f"I{r}", M_, "0", True, center=True)
    cell(ws, f"J{r}", f"=H{r}*I{r}", "0", center=True)
    cell(ws, f"K{r}", cmp_, is_input=True, center=True); cell(ws, f"L{r}", a7, is_input=True, center=True); cell(ws, f"M{r}", rep_, is_input=True, center=True)
    cell(ws, f"N{r}", grp, is_input=True, center=True)
    cell(ws, f"O{r}", f'=IF(OR(L{r}="Yes",AND(J{r}>=6,K{r}<>"Yes")),"Significant deficiency",IF(J{r}>=3,"Other deficiency - sufficient importance","Other deficiency - minor"))', center=True)
    cell(ws, f"P{r}", f'=IF(O{r}="Significant deficiency","TCWG & management",IF(O{r}="Other deficiency - sufficient importance","Management","Management (oral)"))', center=True)
    cell(ws, f"Q{r}", f'=IF(O{r}="Significant deficiency","Written - 265.9 & 10(a)",IF(O{r}="Other deficiency - sufficient importance","Written or oral - 265.10(b)","Oral"))', center=True)
    cell(ws, f"R{r}", eff, is_input=True); cell(ws, f"S{r}", rec, is_input=True)
    cell(ws, f"T{r}", dcom, DATE, True, center=True)
    cell(ws, f"U{r}", f'=IF(T{r}="",IF(O{r}="Significant deficiency",IF(Set_AsOf>=Set_RptDate-7,"DUE - before report date","Pending"),"Pending"),IF(T{r}<=Set_RptDate,"Timely","LATE - after report date"))', center=True)
    cell(ws, f"V{r}", resp, is_input=True); cell(ws, f"W{r}", own, is_input=True, center=True)
    cell(ws, f"X{r}", due, DATE, True, center=True); cell(ws, f"Y{r}", stt, is_input=True, center=True)
    cell(ws, f"Z{r}", f'=IF(O{r}="Other deficiency - minor","",COUNTIF($O${d0}:O{r},O{r}))', "0", center=True)
    cell(ws, f"AA{r}", f'=IF(Z{r}="","",O{r}&"|"&Z{r})', center=True)
    ws.row_dimensions[r].height = 42; r += 1
d1 = r - 1
for col in "KLM": dv(ws, f"{col}{d0}:{col}{d1}", "YesNo")
dv(ws, f"Y{d0}:Y{d1}", "IssueStatus"); dv(ws, f"E{d0}:E{d1}", "DefType")
dname("CD_ID", f"Deficiencies_265!$A${d0}:$A${d1}"); dname("CD_Class", f"Deficiencies_265!$O${d0}:$O${d1}")
dname("CD_Group", f"Deficiencies_265!$N${d0}:$N${d1}"); dname("CD_Proc", f"Deficiencies_265!$F${d0}:$F${d1}"); dname("CD_Score", f"Deficiencies_265!$J${d0}:$J${d1}")
dname("CD_Date", f"Deficiencies_265!$T${d0}:$T${d1}"); dname("CD_Status", f"Deficiencies_265!$Y${d0}:$Y${d1}")
dname("CD_Rank", f"Deficiencies_265!$Z${d0}:$Z${d1}"); dname("CD_Key", f"Deficiencies_265!$AA${d0}:$AA${d1}"); dname("CD_Timely", f"Deficiencies_265!$U${d0}:$U${d1}")
dname("CD_Desc", f"Deficiencies_265!$B${d0}:$B${d1}"); dname("CD_Eff", f"Deficiencies_265!$R${d0}:$R${d1}"); dname("CD_Rec", f"Deficiencies_265!$S${d0}:$S${d1}")
ws.conditional_formatting.add(f"O{d0}:O{d1}", CellIsRule(operator="equal", formula=['"Significant deficiency"'], fill=fill("FFCDD2"), font=Font(name=F, bold=True, color="B71C1C")))
ws.conditional_formatting.add(f"O{d0}:O{d1}", CellIsRule(operator="equal", formula=['"Other deficiency - sufficient importance"'], fill=fill("FFE0B2")))
ws.conditional_formatting.add(f"U{d0}:U{d1}", FormulaRule(formula=[f'OR(LEFT($U{d0},3)="DUE",LEFT($U{d0},4)="LATE")'], fill=fill("FFCDD2")))
ws.conditional_formatting.add(f"Y{d0}:Y{d1}", CellIsRule(operator="equal", formula=['"Closed"'], fill=fill("C8E6C9")))
ws.conditional_formatting.add(f"L{d0}:M{d1}", CellIsRule(operator="equal", formula=['"Yes"'], fill=fill("FFE0B2")))
note(ws, r, "Sources consolidated: COSO assessment (IC1), design & implementation (IC2), ITGCs (IC3), tests of operating effectiveness (IC4), misstatements not prevented/detected by the entity (SUAD), fraud matters (FR5) and prior-year points (DF6). "
            "Classification is a starting point - override with documented professional judgement where appropriate.", "Z")
r += 2

r = section(ws, r, "DF3", "Evaluation in combination - deficiencies affecting the same area or assertion", "ISA 265.8 & A6(h)")
r = table_head(ws, r, ["#", "Combination group", "Deficiencies", "Significant included", "Combined score", "Conclusion", "", "", "", "Auditor's judgement / rationale"])
ws.merge_cells(f"F{r-1}:I{r-1}"); ws.merge_cells(f"J{r-1}:L{r-1}")
for i, (g, why) in enumerate([("Governance", "Control environment weaknesses interact - Board oversight, risk assessment and RP identification"),
                              ("IT & access", "C03 automated control depends on effective access controls (C06)"),
                              ("Revenue", "Cut-off design gap and segregation weakness both affect revenue/receivables"),
                              ("Inventory", "Single deficiency - assessed individually"),
                              ("Financial close", "Single deficiency - remediated"), ("Cash", "Single minor deficiency")], 1):
    cell(ws, f"A{r}", i, center=True); cell(ws, f"B{r}", g)
    cell(ws, f"C{r}", f"=COUNTIF(CD_Group,B{r})", "0", center=True)
    cell(ws, f"D{r}", f'=COUNTIFS(CD_Group,B{r},CD_Class,"Significant deficiency")', "0", center=True)
    cell(ws, f"E{r}", f"=SUMIF(CD_Group,B{r},CD_Score)", "0", center=True)
    cell(ws, f"F{r}", f'=IF(D{r}>0,"Contains significant deficiency - communicate combined effect",IF(AND(C{r}>1,E{r}>=8),"Combination may be a SIGNIFICANT deficiency - document judgement","Not significant in combination"))')
    ws.merge_cells(f"F{r}:I{r}"); cell(ws, f"J{r}", why, is_input=True); ws.merge_cells(f"J{r}:L{r}")
    ws.row_dimensions[r].height = 28; r += 1
r += 1

r = section(ws, r, "DF4", "Communication requirements & tracking", "ISA 265.9-11 & A12-A30")
r = checklist(ws, r, [
    ("Significant deficiencies communicated IN WRITING to TCWG on a timely basis (before the report date)", "ISA 265.9; A12-A14", "Yes", "Letter dated 20-Mar-26 (DF5)"),
    ("Significant deficiencies communicated in writing to management at an appropriate level (unless inappropriate, e.g. management integrity in question)", "ISA 265.10(a); A19-A20", "Yes", "CEO & CFO copied"),
    ("Other deficiencies of sufficient importance communicated to management (unless already communicated by others)", "ISA 265.10(b); A21-A26", "Yes", "Management letter section 2 / oral for minor items"),
    ("Written communication includes description of deficiencies and explanation of their potential effects", "ISA 265.11(a)", "Yes", "Columns R-S of DF2"),
    ("Communication explains: purpose of audit was to express an opinion on the FS", "ISA 265.11(b)(i)", "Yes", "Letter para 2"),
    ("...the audit included consideration of IC relevant to FS preparation to design procedures, not to express an opinion on IC effectiveness", "ISA 265.11(b)(ii)", "Yes", "Letter para 2"),
    ("...matters reported are limited to deficiencies identified during the audit and concluded to be of sufficient importance", "ISA 265.11(b)(iii)", "Yes", "Letter para 2"),
    ("Previously communicated significant deficiencies not yet remediated are communicated again (or referred to)", "ISA 265.A28-A29", "Yes", "CD04 repeat - DF6"),
    ("Communication of significant deficiencies recorded in TCWG log (F1) and documented in the audit file", "ISA 260; ISA 230", "Yes", "F1 item 6"),
], "DF4")
t0 = r
tsum = [("Deficiencies identified", "=COUNTA(CD_ID)"), ("Significant deficiencies", '=COUNTIF(CD_Class,"Significant deficiency")'),
        ("Other - sufficient importance", '=COUNTIF(CD_Class,"Other deficiency - sufficient importance")'), ("Other - minor", '=COUNTIF(CD_Class,"Other deficiency - minor")'),
        ("Significant deficiencies communicated in writing", '=COUNTIFS(CD_Class,"Significant deficiency",CD_Date,">0")'),
        ("Communications late / due", '=COUNTIF(CD_Timely,"LATE*")+COUNTIF(CD_Timely,"DUE*")'), ("Remediated (closed)", '=COUNTIF(CD_Status,"Closed")')]
r = table_head(ws, r, ["", "Summary", "Count"])
for lab, f_ in tsum:
    cell(ws, f"B{r}", lab); cell(ws, f"C{r}", f_, "0", center=True); r += 1
dname("CD_N", f"Deficiencies_265!$C${t0+1}"); dname("CD_SigComm", f"Deficiencies_265!$C${t0+5}")
dname("CD_Closed", f"Deficiencies_265!$C${t0+7}")
r += 1

r = section(ws, r, "DF5", "Draft written communication to TCWG & management (auto-generated)", "ISA 265.9-11 | illustrative")
L0 = r
letter = ['="To: The Board of Directors, "&Set_Client',
          '="Communication of deficiencies in internal control identified during the audit of the financial statements for the year ended "&TEXT(Set_FYE,"dd mmmm yyyy")',
          '="The purpose of our audit was to express an opinion on the financial statements. Our audit included consideration of internal control relevant to the preparation of the financial statements in order to design audit procedures that are appropriate in the circumstances, but not for the purpose of expressing an opinion on the effectiveness of internal control. '
          'The matters reported are limited to those deficiencies that we identified during the audit and that we have concluded are of sufficient importance to merit being reported to you."',
          '="1. SIGNIFICANT DEFICIENCIES IN INTERNAL CONTROL (ISA 265.9)"']
for t in letter:
    c = put(ws, f"B{r}", t, font(9, t.startswith('="1.')), "F1F8E9", al=WRAP); ws.merge_cells(f"B{r}:L{r}")
    ws.row_dimensions[r].height = 48 if len(t) > 200 else 18; r += 1
def letter_items(r, cls, n):
    for k in range(1, n + 1):
        m = f'MATCH("{cls}|{k}",CD_Key,0)'
        f_ = f'=IFERROR("1.{k} "&INDEX(CD_Desc,{m})&". Potential effect: "&INDEX(CD_Eff,{m})&". Recommendation: "&INDEX(CD_Rec,{m})&".","")'
        if cls != "Significant deficiency": f_ = f_.replace('"1.', '"2.')
        put(ws, f"B{r}", f_, font(9), "F1F8E9", al=WRAP); ws.merge_cells(f"B{r}:L{r}"); ws.row_dimensions[r].height = 30; r += 1
    return r
r = letter_items(r, "Significant deficiency", 6)
put(ws, f"B{r}", '="2. OTHER DEFICIENCIES OF SUFFICIENT IMPORTANCE (ISA 265.10(b))"', font(9, True), "F1F8E9"); ws.merge_cells(f"B{r}:L{r}"); r += 1
r = letter_items(r, "Other deficiency - sufficient importance", 6)
for t in ['="We would be pleased to discuss these matters with you. Management\'s responses are set out in the attached appendix."', '="[Audit firm]   |   Engagement partner: "&Set_EP&"   |   Date: ____________"']:
    put(ws, f"B{r}", t, font(9), "F1F8E9", al=WRAP); ws.merge_cells(f"B{r}:L{r}"); ws.row_dimensions[r].height = 18; r += 1
note(ws, r, "Letter items fill automatically from DF2 in register order. Blank lines mean fewer items than slots. Edit the final wording before issue; minor deficiencies may be communicated orally.", "L", 14)
r += 2

r = section(ws, r, "DF6", "Follow-up of prior-year deficiencies (management letter FY2024)", "ISA 265.A28-A29 | ISA 315R.A6")
r = table_head(ws, r, ["#", "Prior-year deficiency", "PY classification", "Management's PY response", "", "CY status", "Evidence", "", "Action", "", "CY ref"])
ws.merge_cells(f"D{r-1}:E{r-1}"); ws.merge_cells(f"G{r-1}:H{r-1}"); ws.merge_cells(f"I{r-1}:J{r-1}")
py_ = [("NRV / slow-moving inventory not reviewed", "Significant deficiency", "Policy to be adopted by Jun-25", "Not remediated", "C10 D&I deficient; M02", "CD04"),
       ("Bank reconciliations not reviewed timely", "Other deficiency", "Review sign-off introduced", "Remediated", "C02 tested effective", "-"),
       ("Payroll changes not authorised", "Other deficiency", "HR approval implemented", "Remediated", "C04 tested effective", "-"),
       ("Weak password parameters in ERP", "Other deficiency", "Parameters updated Mar-25", "Partially remediated", "IC3 access domain still deficient", "CD02")]
for i, (d_, cls, resp, stt, ev_, ref) in enumerate(py_, 1):
    cell(ws, f"A{r}", i, center=True); cell(ws, f"B{r}", d_); cell(ws, f"C{r}", cls, center=True)
    cell(ws, f"D{r}", resp, is_input=True); ws.merge_cells(f"D{r}:E{r}"); cell(ws, f"F{r}", stt, is_input=True, center=True)
    cell(ws, f"G{r}", ev_, is_input=True); ws.merge_cells(f"G{r}:H{r}")
    cell(ws, f"I{r}", f'=IF(F{r}="Remediated","Closed - no further communication",IF(C{r}="Significant deficiency","Re-communicate in writing; A7 indicator (failure to remediate)","Re-communicate to management"))')
    ws.merge_cells(f"I{r}:J{r}"); cell(ws, f"K{r}", ref, center=True); ws.row_dimensions[r].height = 28; r += 1
ws.freeze_panes = "C8"

# ======================================================================
# ACCEPTANCE / TEAM RESOURCES / PBC
# ======================================================================
from openpyxl.chart import BarChart as _Bar, Reference as _Ref
from openpyxl.chart.label import DataLabelList as _DLL

def mini_chart(ws, title, cats, data, anchor, color=LT, horizontal=True, w=14, h=7, pct=False):
    ch = _Bar(); ch.type = "bar" if horizontal else "col"; ch.title = title; ch.roundedCorners = False
    ch.add_data(data, titles_from_data=True); ch.set_categories(cats); ch.legend = None
    ch.width, ch.height = w, h; ch.x_axis.delete = False; ch.y_axis.delete = False; ch.y_axis.majorGridlines = None
    if horizontal: ch.x_axis.scaling.orientation = "maxMin"
    if pct: ch.y_axis.number_format = "0%"
    ch.series[0].graphicalProperties.solidFill = color; ch.series[0].graphicalProperties.line.solidFill = color
    ch.dataLabels = _DLL(showVal=True, showSerName=False, showCatName=False, showLegendKey=False, showPercent=False)
    ch.title.overlay = False
    ws.add_chart(ch, anchor)

# ---------------- ACCEPTANCE ----------------
ws = S["Acceptance_QM"]
r = sheet_header(ws, "CLIENT ACCEPTANCE & CONTINUANCE", "1",
                 "Decide whether to accept or continue the client relationship and the specific engagement, based on integrity, competence & resources, ethics, and engagement risk, with documented approval before agreeing terms.",
                 "ISQM 1.30-31 & A64-A75; ISA 220 (Revised).22-24; ISA 210.6-12; ISA 300.6; IESBA Code 320 (client & engagement acceptance), 360 (NOCLAR), 400 series")
r = section(ws, r, "AC1", "Engagement profile", "ISQM 1.30(a) | ISA 220R.22")
prof = [("Client", "=Set_Client"), ("Engagement status", "=Set_EngStatus"), ("Years as audit client", "=Set_YearsClient"),
        ("Entity type", "=Set_EntityType"), ("Financial reporting framework", "=Set_Framework"), ("Services provided", "Statutory audit; corporate tax compliance (separate team)"),
        ("Audit fee / total client fees (MNT '000)", '=TEXT(Set_AuditFee,"#,##0")&" / "&TEXT(Set_AuditFee+Set_NASFee,"#,##0")'),
        ("Engagement partner / EQ reviewer", '=Set_EP&" / "&Set_EQR')]
for lab, v in prof:
    cell(ws, f"B{r}", lab); c = cell(ws, f"C{r}", v, link=str(v).startswith("="), is_input=not str(v).startswith("=")); ws.merge_cells(f"C{r}:H{r}"); r += 1
r += 1

r = section(ws, r, "AC2", "Engagement risk assessment - acceptance/continuance factors", "ISQM 1.30(c), A67-A70")
r = table_head(ws, r, ["#", "Risk factor", "Weight", "Rating (1 low - 3 high)", "Weighted", "Fatal if high?", "Evidence / basis", "", "", "", "Ref"])
ws.merge_cells(f"G{r-1}:J{r-1}")
facs = [("Integrity of owners, key management & TCWG", 3, 1, "Yes", "Background & media checks clean; no adverse PY matters", "AC3"),
        ("Financial condition & going concern indicators", 2, 2, "No", "Profitable; covenant headroom narrows under stress", "GC3"),
        ("Industry & business risk (retail, imported goods, FX)", 2, 2, "No", "MNT depreciation; key supplier sanctions review", "B1"),
        ("Complexity of accounting & estimates (NRV, ECL, IFRS 16)", 2, 2, "No", "NRV judgemental; leases at branches", "RA3"),
        ("Quality of governance & internal control environment", 2, 2, "No", "No audit committee; informal risk assessment", "IC1"),
        ("Regulatory, legal & litigation exposure", 1, 1, "No", "One immaterial customer claim", "D7"),
        ("Public interest / PIE status & users of FS", 2, 1, "No", "Non-PIE; bank is main external user", "Setup"),
        ("Related parties & unusual transactions", 2, 2, "No", "Director loan identified in CY", "B7"),
        ("AML / KYC & sanctions concerns", 3, 1, "Yes", "Beneficial owners verified; no PEP/sanctions hits", "AC3"),
        ("Fee level, recoverability & overdue fees", 1, 1, "No", "PY fees settled 30-Sep-25", "A4"),
        ("Pressure on deadline / scope limitations imposed", 2, 1, "No", "Reasonable timetable; full access", "A2"),
        ("Independence & ethical threats", 3, 1, "Yes", "Threats reduced to acceptable level (A4)", "A3-A5"),
        ("Firm competence, capabilities & capacity (Team_Resources)", 3, 1, "Yes", "See TR6 conclusion", "TR6")]
f0 = r
for i, (fac, w, rt, fatal, ev_, ref) in enumerate(facs, 1):
    cell(ws, f"A{r}", i, center=True); cell(ws, f"B{r}", fac); cell(ws, f"C{r}", w, "0", True, center=True)
    cell(ws, f"D{r}", rt, "0", True, center=True); cell(ws, f"E{r}", f"=C{r}*D{r}", "0", center=True)
    cell(ws, f"F{r}", fatal, is_input=True, center=True); cell(ws, f"G{r}", ev_, is_input=True); ws.merge_cells(f"G{r}:J{r}")
    cell(ws, f"K{r}", ref, center=True); ws.row_dimensions[r].height = 26; r += 1
f1 = r - 1
dv(ws, f"F{f0}:F{f1}", "YesNo")
ws.conditional_formatting.add(f"D{f0}:D{f1}", CellIsRule(operator="equal", formula=["3"], fill=fill("FFCDD2")))
ws.conditional_formatting.add(f"D{f0}:D{f1}", CellIsRule(operator="equal", formula=["2"], fill=fill("FFE0B2")))
s0 = r
summ = [("Weighted score", f"=SUM(E{f0}:E{f1})", "0"), ("Maximum possible", f"=SUM(C{f0}:C{f1})*3", "0"),
        ("Risk score %", f"=IFERROR((C{s0}-SUM(C{f0}:C{f1}))/(C{s0+1}-SUM(C{f0}:C{f1})),0)", "0%"),
        ("ENGAGEMENT RISK RATING", f'=IF(C{s0+2}<0.4,"Low",IF(C{s0+2}<0.65,"Medium","High"))', None),
        ("Fatal factors rated high", f'=COUNTIFS(F{f0}:F{f1},"Yes",D{f0}:D{f1},3)', "0")]
for i, (lab, v, nf) in enumerate(summ):
    b = put(ws, f"B{r}", lab, font(9, True), PALE, bd=BOX); c = put(ws, f"C{r}", v, font(10, True, DK), PALE, nf, CEN, BOX); r += 1
dname("AC_Rating", f"Acceptance_QM!$C${s0+3}"); dname("AC_Fatal", f"Acceptance_QM!$C${s0+4}"); dname("AC_ScorePct", f"Acceptance_QM!$C${s0+2}")
ws.conditional_formatting.add(f"C{s0+3}", CellIsRule(operator="equal", formula=['"High"'], fill=fill("FFCDD2")))
ws.conditional_formatting.add(f"C{s0+3}", CellIsRule(operator="equal", formula=['"Medium"'], fill=fill("FFE0B2")))
note(ws, r, "Risk score % rescales the weighted score between minimum (all 1s) and maximum (all 3s). Low < 40% | Medium 40-65% | High > 65%. A 'fatal' factor rated 3 requires the engagement to be declined or escalated to the risk/managing partner.")
r += 2

r = section(ws, r, "AC3", "Know-your-client (KYC), AML & sanctions screening", "IESBA 320.3 | local AML law | ISQM 1.30(a)(ii)")
r = checklist(ws, r, [
    ("Legal identity verified (state registration certificate, charter)", "AML / KYC", "Yes", "Certificate no. on file"),
    ("Ultimate beneficial owners (>= 25%) identified & verified", "AML / KYC", "Yes", "2 family shareholders - IDs verified"),
    ("Directors & key management screened for PEP status", "AML / KYC", "Yes", "No PEPs"),
    ("Entity, owners & key counterparties screened against sanctions lists", "AML / KYC", "Yes", "Key supplier under review - monitored (GC1)"),
    ("Adverse media / litigation search performed", "ISQM 1.A65", "Yes", "No adverse findings"),
    ("Nature of business & source of funds understood and consistent", "AML / KYC", "Yes", ""),
], "AC3")

r = section(ws, r, "AC4", "New-client procedures (predecessor auditor & opening balances)", "IESBA 320.8 | ISA 510 | ISA 300.13")
put(ws, f"B{r}", '=IF(Set_EngStatus="New","NEW ENGAGEMENT - complete all items below","CONTINUING ENGAGEMENT - section not applicable (mark N/A)")', font(9, True, "B71C1C")); ws.merge_cells(f"B{r}:J{r}"); r += 1
r = checklist(ws, r, [
    ("Client permission obtained to communicate with predecessor auditor", "IESBA 320.8", "N/A", "Continuing - year 4"),
    ("Predecessor asked about facts/circumstances relevant to acceptance (integrity, disagreements, NOCLAR)", "IESBA 320.8 & A1-A2; 360", "N/A", ""),
    ("Reasons for change of auditor understood", "IESBA 320", "N/A", ""),
    ("Access to predecessor's working papers arranged for opening balances", "ISA 510.6; ISA 300.13", "N/A", ""),
    ("Prior-period audit report reviewed (modifications, EOM)", "ISA 510.6(c)", "N/A", ""),
], "AC4")

r = section(ws, r, "AC5", "Continuance-specific considerations", "ISQM 1.30 | ISA 220R.23")
r = checklist(ws, r, [
    ("Significant matters arising in the current or previous engagement & implications for continuance considered", "ISA 220R.23", "Yes", "PY NRV adjustment; CY cut-off - not affecting continuance"),
    ("Changes in ownership, management, governance or nature of business assessed", "ISQM 1.A66", "Yes", "No changes"),
    ("Disagreements with management or scope limitations in PY", "ISQM 1.30(c)", "Yes", "None"),
    ("Long association with engagement partner / senior personnel evaluated", "IESBA 540", "Yes", "EP year 4 - non-PIE"),
    ("Fee dispute or significant overdue fees", "IESBA 410.7", "Yes", "Resolved"),
    ("Previously communicated significant deficiencies remediated? (affects risk)", "ISA 265.A28", "Yes", "NRV point not remediated - CD04"),
], "AC5")

r = section(ws, r, "AC6", "Competence, capabilities & resources check", "ISQM 1.30(a)(i), 32 | ISA 220R.25-28")
for lab, v in [("Team resources conclusion (TR6)", "=TR_Conclusion"), ("Competency gaps (TR2)", "=TR_Gaps"), ("Team members over-allocated (TR4)", "=TR_Over"),
               ("Specialists required", "IT auditor (ISA 620 - firm employee); technical dept consultation on NRV")]:
    cell(ws, f"B{r}", lab); c = cell(ws, f"C{r}", v, link=str(v).startswith("="), is_input=not str(v).startswith("=")); ws.merge_cells(f"C{r}:J{r}"); r += 1
r += 1

r = section(ws, r, "AC7", "Decision & approvals", "ISQM 1.30-31 | ISA 220R.22, 24")
a0 = r
dec = [("Preconditions for an audit present (A2)?", "Yes", True), ("Independence & ethical requirements can be met (A3-A5)?", "Yes", True),
       ("Engagement risk rating (AC2)", "=AC_Rating", False), ("Required approval level", f'=IF(C{a0+2}="High","EP + risk partner + managing partner",IF(C{a0+2}="Medium","EP + risk partner","Engagement partner"))', False),
       ("EQ review required?", f'=IF(OR(ISNUMBER(SEARCH("PIE",Set_EntityType))*NOT(ISNUMBER(SEARCH("Non-PIE",Set_EntityType))),C{a0+2}="High"),"Required (ISQM 1.34(f))","Firm policy: "&IF(Set_EQR<>"","appointed - "&Set_EQR,"not required"))', False),
       ("DECISION", f'=IF(OR(C{a0}="No",C{a0+1}="No",AC_Fatal>0),"DECLINE / escalate - fatal factor or precondition not met",IF(C{a0+2}="High","ACCEPT subject to enhanced approval & safeguards","ACCEPT / CONTINUE"))', False),
       ("Approved by (EP)", "B. Bat-Erdene", True), ("Approved by (risk partner, if required)", "D. Tsetseg", True),
       ("Approval date", D(2025, 9, 10), True), ("Engagement letter date (A2)", D(2025, 9, 22), True),
       ("Sequence check", f'=IF(C{a0+8}="","Approval pending",IF(C{a0+8}<=C{a0+9},"Approved before terms agreed - OK","APPROVAL AFTER ENGAGEMENT LETTER - breach of ISQM 1.30"))', False)]
for i, (lab, v, isin) in enumerate(dec):
    b = cell(ws, f"B{r}", lab); c = cell(ws, f"C{r}", v, DATE if isinstance(v, D) else None, isin, link=(i == 2), center=True)
    if lab == "DECISION": b.font = font(9, True); c.font = font(11, True, DK); b.fill = fill(PALE); c.fill = fill(PALE)
    ws.merge_cells(f"C{r}:H{r}"); r += 1
for k in (0, 1): dv(ws, f"C{a0+k}", "YesNo")
dname("AC_Decision", f"Acceptance_QM!$C${a0+5}")

# ---------------- TEAM RESOURCES ----------------
ws = S["Team_Resources"]
r = sheet_header(ws, "ENGAGEMENT TEAM - COMPETENCIES, CAPACITY & AVAILABILITY", "1",
                 "Confirm the engagement team collectively has the competence, capabilities and sufficient time to perform the engagement, and plan direction, supervision and review.",
                 "ISQM 1.32 (resources) & A92-A105; ISA 220 (Revised).25-28 (resources), 29-34 (direction, supervision & review); IESBA 113 (competence); IES 8")
for k in range(ci("M"), ci("U") + 1): ws.column_dimensions[gl(k)].width = 11
widths(ws, {"B": 22, "C": 16})
team_rows = [("B. Bat-Erdene", "Engagement partner", "Partner", "CPA (MN), ACCA", 18, 9, 5, 5, 3, 3, 4, 4, 46, 40),
             ("D. Ganbold", "EQ reviewer", "Partner", "CPA (MN)", 20, 6, 5, 5, 3, 3, 4, 4, 42, 40),
             ("S. Oyunaa", "Manager", "Manager", "ACCA", 9, 5, 4, 4, 3, 3, 4, 3, 38, 40),
             ("T. Enkhjin", "Senior", "Senior", "ACCA (part-qualified)", 4, 3, 3, 3, 3, 3, 3, 3, 31, 30),
             ("M. Saruul", "Staff", "Associate", "Graduate", 2, 1, 2, 2, 2, 3, 1, 2, 18, 20),
             ("N. Temuulen", "Staff", "Associate", "Graduate", 1, 0, 2, 2, 1, 2, 1, 1, 12, 20),
             ("IT auditor (firm)", "Specialist", "Manager", "CISA", 8, 2, 2, 3, 5, 4, 1, 3, 44, 40)]
r = section(ws, r, "TR1", "Competency matrix (proficiency 1 = basic ... 5 = expert)", "ISA 220R.25-26 | ISQM 1.32(a)")
heads = ["#", "Team member", "Role", "Grade", "Qualification", "Years audit exp.", "Years retail exp.", "IFRS", "ISA / methodology",
         "IT audit", "Data analytics", "Estimates / valuation", "Fraud & forensic", "Avg proficiency", "Required (role)", "Meets role?",
         "CPD hrs done", "CPD required", "CPD status", "Independence (A3)"]
r = table_head(ws, r, heads)
HDR1 = r - 1
m0 = r
for i, row in enumerate(team_rows, 1):
    cell(ws, f"A{r}", i, center=True)
    for j, v in enumerate(row[:4]): cell(ws, f"{chr(66+j)}{r}", v, is_input=True, center=j > 0)
    for j, v in enumerate(row[4:12]): cell(ws, f"{chr(70+j)}{r}", v, "0", True, center=True)
    cell(ws, f"N{r}", f"=AVERAGE(H{r}:M{r})", "0.0", center=True)
    cell(ws, f"O{r}", f'=IFERROR(INDEX(RoleReq_Lvl,MATCH(C{r},RoleReq_Role,0)),3)', "0.0", center=True)
    cell(ws, f"P{r}", f'=IF(C{r}="Specialist",IF(MAX(H{r}:M{r})>=4,"Yes - specialist","Gap"),IF(N{r}>=O{r},"Yes","Development need"))', center=True)
    cell(ws, f"Q{r}", row[12], "0", True, center=True); cell(ws, f"R{r}", row[13], "0", True, center=True)
    cell(ws, f"S{r}", f'=IF(Q{r}>=R{r},"Met","Short "&R{r}-Q{r}&" hrs")', center=True)
    cell(ws, f"T{r}", f'=IFERROR(INDEX(Ind_Assess,MATCH(B{r},Ind_Name,0)),"Not on A3")', link=True, center=True)
    r += 1
m1 = r - 1
for col in "HIJKLM":
    for lv, colr in ((5, "1B5E20"), (4, "66BB6A"), (3, "C8E6C9"), (2, "FFF9C4"), (1, "FFE0B2")):
        ws.conditional_formatting.add(f"{col}{m0}:{col}{m1}", CellIsRule(operator="equal", formula=[str(lv)], fill=fill(colr), font=Font(name=F, bold=lv == 5, color="FFFFFF" if lv == 5 else "000000")))
ws.conditional_formatting.add(f"P{m0}:P{m1}", FormulaRule(formula=[f'LEFT($P{m0},3)<>"Yes"'], fill=fill("FFE0B2")))
ws.conditional_formatting.add(f"S{m0}:S{m1}", FormulaRule(formula=[f'LEFT($S{m0},5)="Short"'], fill=fill("FFCDD2")))
ws.conditional_formatting.add(f"T{m0}:T{m1}", FormulaRule(formula=[f'$T{m0}<>"Independent"'], fill=fill("FFE0B2")))
dname("TM_Name", f"Team_Resources!$B${m0}:$B${m1}"); dname("TM_Meets", f"Team_Resources!$P${m0}:$P${m1}"); dname("TM_CPD", f"Team_Resources!$S${m0}:$S${m1}")
r += 1
put(ws, f"B{r}", "Required average proficiency by role (firm policy)", font(9, True)); r += 1
rq0 = r
for role, lv in [("Engagement partner", 4), ("EQ reviewer", 4), ("Manager", 3.5), ("Senior", 3), ("Staff", 2), ("Specialist", 4)]:
    cell(ws, f"B{r}", role); cell(ws, f"C{r}", lv, "0.0", True, center=True); r += 1
dname("RoleReq_Role", f"Team_Resources!$B${rq0}:$B${r-1}"); dname("RoleReq_Lvl", f"Team_Resources!$C${rq0}:$C${r-1}")
r += 1

r = section(ws, r, "TR2", "Engagement competency needs vs team coverage", "ISA 220R.26 | ISQM 1.32 | ISA 620")
r = table_head(ws, r, ["#", "Competency needed", "Linked risk / area", "Skill column", "Required level", "Best in team", "Covered by", "Status", "Mitigation (training, specialist, consultation)", "", "", ""])
ws.merge_cells(f"I{r-1}:L{r-1}")
needs = [("IFRS 15 revenue & cut-off", "R01", "IFRS", 4, ""), ("IFRS 16 leases (branches)", "R11", "IFRS", 4, ""),
         ("IFRS 9 expected credit losses", "R05", "Estimates / valuation", 3, ""),
         ("Inventory NRV estimates (ISA 540R)", "R03", "Estimates / valuation", 5, "Consult firm technical department; independent NRV range (C4)"),
         ("ITGCs & IPE testing", "IC3", "IT audit", 4, ""), ("JE analytics / CAATs", "FR4", "Data analytics", 3, ""),
         ("Fraud risk & suspected fraud (ISA 240R)", "FR1-FR7", "Fraud & forensic", 4, ""), ("ISA methodology & ISQM", "All", "ISA / methodology", 4, "")]
n0 = r
for i, (nd, lnk, col_name, req, mit) in enumerate(needs, 1):
    cell(ws, f"A{r}", i, center=True); cell(ws, f"B{r}", nd); cell(ws, f"C{r}", lnk, center=True)
    cell(ws, f"D{r}", col_name, is_input=True, center=True); cell(ws, f"E{r}", req, "0", True, center=True)
    cell(ws, f"F{r}", f"=MAX(INDEX($H${m0}:$M${m1},0,MATCH(D{r},$H${HDR1}:$M${HDR1},0)))", "0", center=True)
    cell(ws, f"G{r}", f"=INDEX($B${m0}:$B${m1},MATCH(F{r},INDEX($H${m0}:$M${m1},0,MATCH(D{r},$H${HDR1}:$M${HDR1},0)),0))", center=True)
    cell(ws, f"H{r}", f'=IF(F{r}>=E{r},"Covered",IF(I{r}<>"","Gap - mitigated","GAP - action required"))', center=True)
    cell(ws, f"I{r}", mit, is_input=True); ws.merge_cells(f"I{r}:L{r}"); r += 1
ws.conditional_formatting.add(f"H{n0}:H{r-1}", CellIsRule(operator="equal", formula=['"GAP - action required"'], fill=fill("FFCDD2"), font=Font(name=F, bold=True, color="B71C1C")))
ws.conditional_formatting.add(f"H{n0}:H{r-1}", CellIsRule(operator="equal", formula=['"Gap - mitigated"'], fill=fill("FFE0B2")))
ws.conditional_formatting.add(f"H{n0}:H{r-1}", CellIsRule(operator="equal", formula=['"Covered"'], fill=fill("C8E6C9")))
dname("TR_NeedStatus", f"Team_Resources!$H${n0}:$H${r-1}")
r += 1

r = section(ws, r, "TR3", "Budget allocation, actual hours & utilisation by member", "ISA 220R.25(b) | ISQM 1.32")
r = table_head(ws, r, ["#", "Team member", "Role", "Budget share %", "Budget hours", "Actual share %", "Actual hours", "Remaining", "Burn % (actual/budget)", "Status", "Charge-out rate (MNT '000/hr)", "Cost at actual"])
alloc = [(0.09, 0.09, 450), (0.03, 0.01, 450), (0.18, 0.19, 280), (0.32, 0.35, 160), (0.18, 0.20, 90), (0.12, 0.08, 90), (0.08, 0.08, 300)]
b0 = r
for i, (row, (bs, as_, rate)) in enumerate(zip(team_rows, alloc), 1):
    cell(ws, f"A{r}", i, center=True); cell(ws, f"B{r}", f"=INDEX(TM_Name,{i})", link=True); cell(ws, f"C{r}", row[1], center=True)
    cell(ws, f"D{r}", bs, "0%", True, center=True); cell(ws, f"E{r}", f"=D{r}*DD_HrsBud", "0", center=True)
    cell(ws, f"F{r}", as_, "0%", True, center=True); cell(ws, f"G{r}", f"=F{r}*DD_HrsAct", "0", center=True)
    cell(ws, f"H{r}", f"=E{r}-G{r}", '0;(0);"-"', center=True); cell(ws, f"I{r}", f"=IFERROR(G{r}/E{r},0)", "0%", center=True)
    cell(ws, f"J{r}", f'=IF(I{r}>1.1,"Overrun",IF(I{r}>0.95,"At budget",IF(I{r}<0.5,"Behind / not started","On track")))', center=True)
    cell(ws, f"K{r}", rate, NUM, True, center=True); cell(ws, f"L{r}", f"=G{r}*K{r}", NUM); r += 1
put(ws, f"B{r}", "TOTAL / reconciliation to WP_Index", font(9, True), PALE, bd=BOX)
for col, f_, nf in (("D", f"=SUM(D{b0}:D{r-1})", "0%"), ("E", f"=SUM(E{b0}:E{r-1})", "0"), ("F", f"=SUM(F{b0}:F{r-1})", "0%"),
                    ("G", f"=SUM(G{b0}:G{r-1})", "0"), ("L", f"=SUM(L{b0}:L{r-1})", NUM)):
    put(ws, f"{col}{r}", f_, font(9, True), PALE, nf, CEN, BOX)
put(ws, f"J{r}", f'=IF(AND(ABS(D{r}-1)<0.001,ABS(F{r}-1)<0.001),"Shares = 100% - reconciled","SHARES MUST TOTAL 100%")', font(9, True, MD), PALE, bd=BOX)
put(ws, f"I{r+1}", "Fee recovery (audit fee / cost at actual)", font(9), bd=BOX); put(ws, f"L{r+1}", f"=IFERROR(Set_AuditFee/L{r},0)", font(9, True), nf="0%", bd=BOX)
dname("TR_PartnerPct", f"Team_Resources!$F${b0}")
ws.conditional_formatting.add(f"J{b0}:J{r-1}", CellIsRule(operator="equal", formula=['"Overrun"'], fill=fill("FFCDD2")))
r += 3

r = section(ws, r, "TR4", "Capacity plan - monthly hours scheduled & capacity usage (incl. other engagements)", "ISA 220R.25 | ISQM 1.32(a)")
months = ["Sep-25", "Oct-25", "Nov-25", "Dec-25", "Jan-26", "Feb-26", "Mar-26", "Apr-26", "May-26"]
prof_std = [0.08, 0.12, 0.12, 0.14, 0.06, 0.22, 0.20, 0.03, 0.03]
prof = {"EQ reviewer": [0, 0, 0, 0, 0, 0, 0.9, 0.1, 0], "Specialist": [0, 0.3, 0.4, 0.3, 0, 0, 0, 0, 0],
        "Engagement partner": [0.15, 0.1, 0.05, 0.05, 0.05, 0.15, 0.35, 0.05, 0.05]}
other = [0.75, 0.80, 0.70, 0.65, 0.60, 0.65, 0.75]
avail = [0.90, 0.90, 0.90, 0.85, 0.90, 0.90, 0.90]
heads = ["#", "Team member", "Availability % (leave/training)", "Other engagements % of std"] + months + ["Total scheduled", "Peak usage", "Status"]
r = table_head(ws, r, heads)
cap0 = r
TOT = 545
for i, (row, (bs, _, _)) in enumerate(zip(team_rows, alloc)):
    cell(ws, f"A{r}", i + 1, center=True); cell(ws, f"B{r}", f"=INDEX(TM_Name,{i+1})", link=True)
    cell(ws, f"C{r}", avail[i], "0%", True, center=True); cell(ws, f"D{r}", other[i], "0%", True, center=True)
    p_ = prof.get(row[1], prof_std)
    for k in range(9): cell(ws, f"{gl(5+k)}{r}", round(bs * TOT * p_[k]), "0", True, center=True)
    cell(ws, f"N{r}", f"=SUM(E{r}:M{r})", "0", center=True); r += 1
cap1 = r - 1
# usage grid
r += 1
put(ws, f"B{r}", "Capacity usage % = (scheduled on this engagement + other engagements) / available hours   [standard hours per month: Setup]", font(9, True)); ws.merge_cells(f"B{r}:P{r}"); r += 1
r = table_head(ws, r, ["#", "Team member", "", ""] + months + ["", "Peak usage", "Status"])
u0 = r
for i in range(cap1 - cap0 + 1):
    src = cap0 + i
    cell(ws, f"A{r}", i + 1, center=True); cell(ws, f"B{r}", f"=B{src}", link=True)
    for k in range(9):
        col = gl(5 + k)
        cell(ws, f"{col}{r}", f"=IFERROR(({col}{src}+$D{src}*Set_StdHrs)/($C{src}*Set_StdHrs),0)", "0%", center=True)
    cell(ws, f"O{r}", f"=MAX(E{r}:M{r})", "0%", center=True)
    cell(ws, f"P{r}", f'=IF(O{r}>1,"Over-allocated",IF(O{r}>0.85,"Near capacity","Available capacity"))', center=True)
    cell(ws, f"O{src}", f"=O{r}", "0%", center=True); cell(ws, f"P{src}", f"=P{r}", center=True)
    r += 1
u1 = r - 1
ws.conditional_formatting.add(f"E{u0}:M{u1}", CellIsRule(operator="greaterThan", formula=["1"], fill=fill("EF9A9A"), font=Font(name=F, bold=True, color="B71C1C")))
ws.conditional_formatting.add(f"E{u0}:M{u1}", CellIsRule(operator="between", formula=["0.85", "1"], fill=fill("FFE0B2")))
ws.conditional_formatting.add(f"E{u0}:M{u1}", CellIsRule(operator="lessThan", formula=["0.85"], fill=fill("C8E6C9")))
for rng in (f"P{u0}:P{u1}", f"P{cap0}:P{cap1}"):
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Over-allocated"'], fill=fill("FFCDD2"), font=Font(name=F, bold=True, color="B71C1C")))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Available capacity"'], fill=fill("C8E6C9")))
dname("TR_Peak", f"Team_Resources!$O${u0}:$O${u1}"); dname("TR_PeakStatus", f"Team_Resources!$P${u0}:$P${u1}")
put(ws, f"B{r}", "Total scheduled vs budget hours", font(9, True), PALE, bd=BOX)
put(ws, f"N{r}", f"=SUM(N{cap0}:N{cap1})", font(9, True), PALE, "0", CEN, BOX)
put(ws, f"O{r}", "=DD_HrsBud", font(9, True, "008000"), PALE, "0", CEN, BOX)
put(ws, f"P{r}", f'=IF(ABS(N{r}-O{r})/O{r}<=0.05,"Within 5% of budget","Re-plan schedule vs budget")', font(9, True, MD), PALE, bd=BOX)
r += 2
mini_chart(ws, "Peak capacity usage by team member", _Ref(ws, min_col=2, min_row=u0, max_row=u1), _Ref(ws, min_col=15, min_row=u0 - 1, max_row=u1),
           f"R{cap0}", color="2E7D32", pct=True, w=13, h=7)

r = section(ws, r, "TR5", "Direction, supervision & review plan", "ISA 220R.29-34 | ISA 230.9")
r = table_head(ws, r, ["#", "Area / layer", "Preparer", "First-level reviewer", "Partner review", "Timing", "", "Review evidence"])
ws.merge_cells(f"F{r-1}:G{r-1}")
for i, (a, p_, rv, pr, tm, ev_) in enumerate([("Strategy & acceptance", "S. Oyunaa", "B. Bat-Erdene", "Yes", "Before engagement letter", "Sign-off A1-A7, AC7"),
        ("Planning, risk & fraud", "S. Oyunaa", "B. Bat-Erdene", "Yes", "Before fieldwork", "RA3, FR3, B3"),
        ("Controls testing & sampling", "T. Enkhjin", "S. Oyunaa", "Key judgements", "End of interim", "IC4, S1"),
        ("Substantive procedures", "T. Enkhjin / staff", "S. Oyunaa", "Significant risks", "During final", "D1-D9"),
        ("Findings & misstatements", "S. Oyunaa", "B. Bat-Erdene", "Yes", "Before TCWG meeting", "SUAD, DF2"),
        ("Completion & report", "S. Oyunaa", "B. Bat-Erdene + EQR", "Yes", "Before report date", "F5-F7")], 1):
    cell(ws, f"A{r}", i, center=True); cell(ws, f"B{r}", a); cell(ws, f"C{r}", p_, is_input=True, center=True); cell(ws, f"D{r}", rv, is_input=True, center=True)
    cell(ws, f"E{r}", pr, is_input=True, center=True); cell(ws, f"F{r}", tm, is_input=True); ws.merge_cells(f"F{r}:G{r}"); cell(ws, f"H{r}", ev_, center=True); r += 1
put(ws, f"B{r}", "Partner hours as % of actual (TR3)", font(9, True), PALE, bd=BOX); put(ws, f"C{r}", "=TR_PartnerPct", font(9, True, DK), PALE, "0%", CEN, BOX)
r += 2

r = section(ws, r, "TR6", "Conclusion on sufficiency & appropriateness of resources", "ISA 220R.25-28 | ISQM 1.32")
c0 = r
cons = [("Competency gaps requiring action (TR2)", '=COUNTIF(TR_NeedStatus,"GAP - action required")', "0"),
        ("Gaps mitigated (specialist / consultation)", '=COUNTIF(TR_NeedStatus,"Gap - mitigated")', "0"),
        ("Members below role proficiency", '=COUNTIF(TM_Meets,"Development need")+COUNTIF(TM_Meets,"Gap")', "0"),
        ("Members short of CPD", '=COUNTIF(TM_CPD,"Short*")', "0"),
        ("Members over-allocated in any month (TR4)", '=COUNTIF(TR_PeakStatus,"Over-allocated")', "0"),
        ("CONCLUSION", f'=IF(OR(C{c0}>0,C{c0+4}>1),"RESOURCES INSUFFICIENT - reassign staff / reschedule before proceeding",IF(OR(C{c0+1}>0,C{c0+2}>0,C{c0+3}>0,C{c0+4}>0),"Sufficient with actions: mitigate gaps, coach juniors, rebalance peak month","Sufficient & appropriate"))', None)]
for lab, v, nf in cons:
    b = cell(ws, f"B{r}", lab); c = cell(ws, f"C{r}", v, nf, center=(nf is not None))
    if lab == "CONCLUSION": b.font = font(9, True); c.font = font(10, True, DK); b.fill = fill(PALE); c.fill = fill(PALE); ws.merge_cells(f"C{r}:L{r}")
    r += 1
dname("TR_Gaps", f"Team_Resources!$C${c0}"); dname("TR_Over", f"Team_Resources!$C${c0+4}"); dname("TR_Conclusion", f"Team_Resources!$C${c0+5}")

# ---------------- PBC ----------------
ws = S["PBC_List"]
r = sheet_header(ws, "PREPARED-BY-CLIENT (PBC) REQUEST LIST & PROGRESS", "4",
                 "Agree with management the information and documents to be provided, track receipt against due dates, assess reliability of information produced by the entity, and escalate delays.",
                 "ISA 210.6(b)(iii) & A23 (access to information); ISA 500.7-9 (information produced by the entity); ISA 300.A13; ISA 260.16(b) (significant difficulties); ISA 705.7 (scope limitation)")
for k in range(ci("M"), ci("P") + 1): ws.column_dimensions[gl(k)].width = 12
widths(ws, {"B": 34, "C": 14, "D": 9, "N": 26})
r = section(ws, r, "PBC1", "PBC request list & tracker", "ISA 210.6(b)(iii) | ISA 500.9")
r = table_head(ws, r, ["#", "Document / information requested", "Area", "WP ref", "Client owner", "Auditor owner", "Requested", "Due", "Received",
                       "Status", "IPE tested? (ISA 500.9)", "Ageing", "Follow-ups", "Comment"])
R_, I_, Y, N_ = D(2025, 10, 20), D(2026, 1, 12), "Yes", "No"
pbc = [("Final trial balance & general ledger (CY)", "General", "D1", "Chief Accountant", "T. Enkhjin", I_, D(2026, 1, 30), D(2026, 1, 29), "Accepted", Y, 0, ""),
       ("Draft financial statements & notes (IFRS)", "General", "F6", "CFO", "S. Oyunaa", I_, D(2026, 2, 20), D(2026, 2, 27), "Accepted", "N/A", 1, "1 week late"),
       ("Bank statements & reconciliations at 31-Dec", "Cash", "D4", "Chief Accountant", "M. Saruul", I_, D(2026, 1, 20), D(2026, 1, 19), "Accepted", Y, 0, ""),
       ("Aged receivables listing (by customer)", "Receivables", "S4", "Credit Manager", "T. Enkhjin", I_, D(2026, 1, 20), D(2026, 1, 21), "Accepted", Y, 0, "Footed & agreed to GL"),
       ("ECL provision matrix & loss-rate support", "Receivables", "C4", "CFO", "S. Oyunaa", I_, D(2026, 2, 6), D(2026, 2, 16), "Accepted", Y, 2, ""),
       ("Inventory count sheets & final priced listing", "Inventory", "D5", "Warehouse Mgr", "T. Enkhjin", R_, D(2026, 1, 15), D(2026, 1, 14), "Accepted", Y, 0, ""),
       ("Inventory ageing & NRV calculation", "Inventory", "C4", "CFO", "S. Oyunaa", I_, D(2026, 2, 6), D(2026, 2, 24), "Accepted", Y, 3, "Late - escalated to CFO"),
       ("Fixed asset register & additions invoices", "PPE", "D3", "Chief Accountant", "M. Saruul", I_, D(2026, 1, 30), D(2026, 2, 2), "Accepted", Y, 1, ""),
       ("Lease contracts & IFRS 16 calculations", "PPE", "D3", "CFO", "S. Oyunaa", I_, D(2026, 2, 6), D(2026, 2, 6), "Accepted", Y, 0, ""),
       ("Aged payables & supplier statements", "Payables", "D3", "AP Supervisor", "N. Temuulen", I_, D(2026, 1, 30), D(2026, 1, 30), "Accepted", Y, 0, ""),
       ("Post year-end payments listing (Jan-Feb 26)", "Payables", "S1.5", "AP Supervisor", "N. Temuulen", I_, D(2026, 3, 2), D(2026, 3, 4), "Accepted", Y, 1, ""),
       ("Loan agreements & covenant calculations", "Borrowings", "D4", "CFO", "S. Oyunaa", I_, D(2026, 1, 30), D(2026, 1, 28), "Accepted", "N/A", 0, ""),
       ("Payroll registers & bureau reports (Dec-25)", "Payroll", "C8", "HR Manager", "M. Saruul", I_, D(2026, 1, 30), D(2026, 2, 3), "Accepted", Y, 1, ""),
       ("Sales invoices & dispatch notes - cut-off sample", "Revenue", "D3", "Sales Admin", "T. Enkhjin", I_, D(2026, 2, 13), D(2026, 2, 13), "Accepted", "N/A", 0, ""),
       ("Credit notes issued after year end", "Revenue", "D3", "Sales Admin", "T. Enkhjin", I_, D(2026, 2, 13), D(2026, 2, 18), "Accepted", Y, 1, ""),
       ("Manual journal entry listing (full year) & user list", "Journals", "FR4", "IT Manager", "IT auditor (firm)", I_, D(2026, 1, 30), D(2026, 2, 9), "Accepted", Y, 2, "Completeness re-run"),
       ("ERP user access listing & access review evidence", "IT", "IC3", "IT Manager", "IT auditor (firm)", R_, D(2025, 11, 14), D(2025, 11, 21), "Accepted", Y, 1, "Q2/Q3 reviews missing - CD02"),
       ("Board & shareholder minutes (2025 to date)", "Governance", "B6", "CEO", "S. Oyunaa", I_, D(2026, 1, 30), D(2026, 1, 30), "Partially received", "N/A", 1, "Mar-26 minutes pending"),
       ("Related party list & directors' declarations", "Related parties", "B7", "CEO", "S. Oyunaa", I_, D(2026, 2, 13), None, "Re-request", "N/A", 3, "Director loan not on list - FR5.3"),
       ("Corporate income tax return & computations", "Tax", "D3", "Chief Accountant", "M. Saruul", I_, D(2026, 2, 20), D(2026, 2, 19), "Received", "N/A", 0, ""),
       ("Going concern assessment & 15-month cash forecast", "Going concern", "GC2", "CFO", "S. Oyunaa", I_, D(2026, 2, 27), D(2026, 3, 3), "Accepted", Y, 1, ""),
       ("Legal confirmations - list of lawyers & claims", "Litigation", "D7", "CEO", "S. Oyunaa", I_, D(2026, 2, 27), D(2026, 2, 27), "Received", "N/A", 0, "Letter to counsel outstanding"),
       ("Subsequent events schedule & Jan-Feb management accounts", "Subsequent events", "D8", "CFO", "T. Enkhjin", I_, D(2026, 3, 13), D(2026, 3, 16), "Received", Y, 1, ""),
       ("Management responses to deficiencies letter", "Findings", "DF2", "CFO", "S. Oyunaa", D(2026, 3, 20), D(2026, 3, 27), None, "Not received", "N/A", 0, ""),
       ("Signed management representation letter", "Completion", "F2", "CEO / CFO", "S. Oyunaa", D(2026, 3, 18), D(2026, 3, 30), None, "Not received", "N/A", 0, "Dated at report date"),
       ("Other information - Directors' report", "Completion", "F8", "CEO", "S. Oyunaa", I_, D(2026, 3, 13), None, "Not received", "N/A", 2, "")]
p0 = r
for i, row in enumerate(pbc, 1):
    (doc, area, wp, cown, aown, rq, due, rec, stt, ipe, fu, com) = row
    cell(ws, f"A{r}", i, center=True); cell(ws, f"B{r}", doc, is_input=True); cell(ws, f"C{r}", area, is_input=True, center=True)
    cell(ws, f"D{r}", wp, center=True); cell(ws, f"E{r}", cown, is_input=True, center=True); cell(ws, f"F{r}", aown, is_input=True, center=True)
    cell(ws, f"G{r}", rq, DATE, True, center=True); cell(ws, f"H{r}", due, DATE, True, center=True); cell(ws, f"I{r}", rec, DATE, True, center=True)
    cell(ws, f"J{r}", stt, is_input=True, center=True); cell(ws, f"K{r}", ipe, is_input=True, center=True)
    cell(ws, f"L{r}", f'=IF(I{r}<>"",IF(I{r}<=H{r},"On time","Late "&(I{r}-H{r})&"d"),IF(Set_AsOf>H{r},"OVERDUE "&(Set_AsOf-H{r})&"d","Not yet due"))', center=True)
    cell(ws, f"M{r}", fu, "0", True, center=True); cell(ws, f"N{r}", com, is_input=True)
    ws.row_dimensions[r].height = 26; r += 1
p1 = r - 1
dv(ws, f"J{p0}:J{p1}", "PBCStatus"); dv(ws, f"K{p0}:K{p1}", "YesNo")
dname("PBC_Area", f"PBC_List!$C${p0}:$C${p1}"); dname("PBC_Status", f"PBC_List!$J${p0}:$J${p1}"); dname("PBC_Age", f"PBC_List!$L${p0}:$L${p1}")
dname("PBC_IPE", f"PBC_List!$K${p0}:$K${p1}"); dname("PBC_Doc", f"PBC_List!$B${p0}:$B${p1}")
stc = {"Accepted": "81C784", "Received": "C8E6C9", "Partially received": "FFE0B2", "Re-request": "FFCDD2", "Not received": "EEEEEE"}
for k, v in stc.items():
    ws.conditional_formatting.add(f"J{p0}:J{p1}", CellIsRule(operator="equal", formula=[f'"{k}"'], fill=fill(v)))
ws.conditional_formatting.add(f"L{p0}:L{p1}", FormulaRule(formula=[f'LEFT($L{p0},7)="OVERDUE"'], fill=fill("FFCDD2"), font=Font(name=F, bold=True, color="B71C1C")))
ws.conditional_formatting.add(f"L{p0}:L{p1}", FormulaRule(formula=[f'LEFT($L{p0},4)="Late"'], fill=fill("FFE0B2")))
note(ws, r, "Status: Not received > Partially received > Received (logged) > Accepted (reviewed as complete & appropriate) | Re-request = rejected / incomplete. "
            "Information produced by the entity (IPE) used as audit evidence must be evaluated for accuracy & completeness (ISA 500.9).", "N")
r += 2

r = section(ws, r, "PBC2", "PBC procedures - planning, quality, follow-up & escalation", "ISA 210 | ISA 300 | ISA 500.9 | ISA 260")
r = checklist(ws, r, [
    ("PBC list agreed with management at planning, with owners, formats & due dates (linked to timeline B10)", "ISA 300.A13; ISA 210.A23", "Yes", "Issued 20-Oct-25 and 12-Jan-26"),
    ("Engagement letter confirms management's responsibility to provide information & unrestricted access", "ISA 210.6(b)(iii)", "Yes", "A2"),
    ("Each item reviewed on receipt for completeness, period, format & consistency with TB before acceptance", "ISA 500.7", "Yes", "Status 'Accepted' only after review"),
    ("IPE accuracy & completeness tested (footing, agreement to GL, re-running reports, ITGC reliance)", "ISA 500.9", "Yes", "See column K; ITGC access deficiency -> substantive IPE tests"),
    ("Overdue items followed up; repeated delays escalated to CFO, then to TCWG", "ISA 260.16(b)", "Yes", "NRV & RP items escalated"),
    ("Delays/refusals evaluated as potential scope limitation or indicator of fraud risk", "ISA 705.7, 11; ISA 240R", None, "RP list re-request outstanding"),
    ("Client documents retained securely; confidential data handled per firm policy", "ISQM 1; IESBA 114", "Yes", "Secure client portal"),
], "PBC2")

r = section(ws, r, "PBC3", "PBC progress dashboard", "Engagement management")
r = table_head(ws, r, ["", "Status", "Count", "% of total"])
s0 = r
for st_ in ["Accepted", "Received", "Partially received", "Re-request", "Not received"]:
    cell(ws, f"B{r}", st_); cell(ws, f"C{r}", f"=COUNTIF(PBC_Status,B{r})", "0", center=True)
    cell(ws, f"D{r}", f"=IFERROR(C{r}/COUNTA(PBC_Doc),0)", "0%", center=True); r += 1
s1 = r - 1
k0 = r + 1
kp = [("Total items requested", "=COUNTA(PBC_Doc)", "0"), ("Received or accepted", '=COUNTIF(PBC_Status,"Accepted")+COUNTIF(PBC_Status,"Received")', "0"),
      ("% received", f"=IFERROR(C{k0+1}/C{k0},0)", "0%"), ("Received on time", '=COUNTIF(PBC_Age,"On time")', "0"),
      ("On-time rate (of received)", f'=IFERROR(C{k0+3}/(C{k0+3}+COUNTIF(PBC_Age,"Late*")),0)', "0%"),
      ("Overdue now", '=COUNTIF(PBC_Age,"OVERDUE*")', "0"), ("IPE items tested", '=COUNTIF(PBC_IPE,"Yes")', "0")]
r += 1
for lab, v, nf in kp:
    cell(ws, f"B{r}", lab); cell(ws, f"C{r}", v, nf, center=True); r += 1
dname("PBC_N", f"PBC_List!$C${k0}"); dname("PBC_PctRec", f"PBC_List!$C${k0+2}"); dname("PBC_Overdue", f"PBC_List!$C${k0+5}")
dname("PBC_OnTime", f"PBC_List!$C${k0+4}"); dname("PBC_IPEn", f"PBC_List!$C${k0+6}")
r += 1
r = table_head(ws, r, ["", "Area", "Items", "Outstanding"])
areas = []
for row in pbc:
    if row[1] not in areas: areas.append(row[1])
for a in areas:
    cell(ws, f"B{r}", a); cell(ws, f"C{r}", f"=COUNTIF(PBC_Area,B{r})", "0", center=True)
    cell(ws, f"D{r}", f'=COUNTIFS(PBC_Area,B{r},PBC_Status,"<>Accepted",PBC_Status,"<>Received")', "0", center=True); r += 1
mini_chart(ws, "PBC items by status", _Ref(ws, min_col=2, min_row=s0, max_row=s1), _Ref(ws, min_col=3, min_row=s0 - 1, max_row=s1), f"F{s0}", color="1565C0", w=12, h=6.5)
ws.freeze_panes = "C8"

# ======================================================================
# ASSERTIONS & AI AGENTS
# ======================================================================
# ---------------- ASSERTIONS ----------------
ws = S["Assertions"]
r = sheet_header(ws, "ASSERTIONS AT RISK - FROM CONTROL GAP TO FINANCIAL IMPACT", "2 & 5",
                 "Identify which financial statement assertion could be misstated, link each control gap to the assertion at risk and its financial impact, and drive specific recommendations.",
                 "ISA 315 (Revised 2019).12(a) & A190 (assertions); ISA 330.6-7; ISA 265.11(a); ISA 450")
for k in range(ci("M"), ci("P") + 1): ws.column_dimensions[gl(k)].width = 13
widths(ws, {"B": 26, "C": 22, "D": 34})
r = section(ws, r, "AS1", "Assertion framework - what could go wrong?", "ISA 315R.A190")
r = table_head(ws, r, ["#", "Assertion", "Applies to", "Key question", "", "", "Typical audit procedure", "", "", "", "Direction of test"])
ws.merge_cells(f"D{r-1}:F{r-1}"); ws.merge_cells(f"G{r-1}:J{r-1}")
fw = [("Occurrence", "P&L / transactions", "Did the transaction actually happen and pertain to the entity?", "Vouch from ledger to source documents (orders, GDNs, cash)", "Ledger > source"),
      ("Completeness", "P&L & balance sheet", "Have all transactions / balances / disclosures been recorded?", "Trace from source to ledger; search for unrecorded liabilities; analytics", "Source > ledger"),
      ("Accuracy", "P&L & balance sheet", "Are amounts and data recorded correctly?", "Recalculate; agree to contracts, price lists, rates", "Both"),
      ("Cut-off", "P&L / transactions", "Recorded in the correct accounting period?", "Test transactions either side of year end", "Both"),
      ("Classification", "P&L & balance sheet", "Recorded in the proper accounts / current vs non-current?", "Review account mapping & reclassifications", "Both"),
      ("Existence", "Balance sheet", "Do the assets, liabilities and equity interests exist?", "Confirmations; physical inspection; count attendance", "Ledger > asset"),
      ("Rights & obligations", "Balance sheet", "Does the entity hold/control the rights to assets; are liabilities obligations of the entity?", "Inspect title deeds, contracts, loan agreements", "Ledger > documents"),
      ("Valuation & allocation", "Balance sheet", "Included at appropriate amounts, with valuation adjustments recorded?", "NRV, ECL, impairment testing; estimates (ISA 540R)", "Both"),
      ("Presentation & disclosure", "P&L & balance sheet", "Appropriately aggregated, described and disclosed per the framework?", "Disclosure checklist; RP & GC disclosures", "FS > framework")]
for i, (a, ap, q, pr, dirn) in enumerate(fw, 1):
    cell(ws, f"A{r}", i, center=True); c = cell(ws, f"B{r}", a); c.font = font(9, True)
    cell(ws, f"C{r}", ap, center=True); cell(ws, f"D{r}", q); ws.merge_cells(f"D{r}:F{r}")
    cell(ws, f"G{r}", pr); ws.merge_cells(f"G{r}:J{r}"); cell(ws, f"K{r}", dirn, center=True); ws.row_dimensions[r].height = 28; r += 1
note(ws, r, "ISA 315R groups balance-sheet 'accuracy, valuation and allocation' together; they are shown separately here for teaching. Use the same assertion names in Risk_315R RA3 so the matrix below picks them up.")
r += 2

r = section(ws, r, "AS2", "Assertion risk matrix - risks of material misstatement by SCOTABD & assertion (from RA3)", "ISA 315R.28-29 | ISA 330.6")
cols = [("Occurrence / Existence", ["*Occurrence*", "*Existence*"]), ("Completeness", ["*Completeness*"]), ("Accuracy / Valuation", ["*Accuracy*", "*Valuation*"]),
        ("Cut-off", ["*Cut-off*"]), ("Classification", ["*Classification*"]), ("Rights & obligations", ["*Rights*"]), ("Presentation & disclosure", ["*Presentation*"])]
r = table_head(ws, r, ["#", "SCOTABD", "Class"] + [c_[0] for c_ in cols] + ["Assertions at risk", "Highest spectrum"])
HDR2 = r - 1
x0 = r
acct_rows = [f"=INDEX(TB_Acct,{i+1})" for i in range(len(tbrows))] + ["Disclosures - related parties", "All (FS level)"]
for i, a in enumerate(acct_rows, 1):
    cell(ws, f"A{r}", i, center=True); cell(ws, f"B{r}", a, link=a.startswith("="))
    cell(ws, f"C{r}", f"=INDEX(TB_Class,{i})" if i <= len(tbrows) else ("Disclosure" if "Disclosures" in a else "Pervasive"), link=i <= len(tbrows), center=True)
    for j, (nm, pats) in enumerate(cols):
        col = gl(4 + j)
        if a == "All (FS level)":
            f_ = f'=COUNTIFS(Risk_SCOT,$B{r})'
        else:
            f_ = "=" + "+".join(f'COUNTIFS(Risk_SCOT,$B{r},Risk_Assert,"{p}")' for p in pats)
        put(ws, f"{col}{r}", f_ + "" , font(10, True), None, '0;0;""', CEN, BOX)
    cell(ws, f"K{r}", f'=COUNTIF(D{r}:J{r},">0")', '0;0;"-"', center=True)
    cell(ws, f"L{r}", f'=IF(COUNTIFS(Risk_SCOT,B{r},Risk_Spectrum,"Significant")>0,"Significant",IF(COUNTIFS(Risk_SCOT,B{r},Risk_Spectrum,"High")>0,"High",IF(COUNTIFS(Risk_SCOT,B{r},Risk_Spectrum,"Elevated")>0,"Elevated",IF(COUNTIFS(Risk_SCOT,B{r},Risk_Spectrum,"Moderate")>0,"Moderate",IF(COUNTIF(Risk_SCOT,B{r})>0,"Low","-")))))', center=True)
    r += 1
x1 = r - 1
ws.conditional_formatting.add(f"D{x0}:J{x1}", CellIsRule(operator="greaterThan", formula=["0"], fill=fill("FFCC80"), font=Font(name=F, bold=True, color="B71C1C")))
for band, colr in SPEC_COL.items():
    ws.conditional_formatting.add(f"L{x0}:L{x1}", CellIsRule(operator="equal", formula=[f'"{band}"'], fill=fill(colr), font=Font(name=F, bold=True, color="FFFFFF" if band in ("High", "Significant") else "000000")))
put(ws, f"B{r}", "Risks per assertion", font(9, True), PALE, bd=BOX)
for j in range(7):
    col = gl(4 + j); put(ws, f"{col}{r}", f"=SUM({col}{x0}:{col}{x1})", font(9, True), PALE, "0", CEN, BOX)
put(ws, f"K{r}", f"=SUM(K{x0}:K{x1})", font(9, True), PALE, "0", CEN, BOX)
dname("AS_Total", f"Assertions!$K${r}")
TOTR = r
r += 2
# transposed totals for chart
put(ws, f"N{x0-1}", "Assertion", font(9, True, "FFFFFF"), DK, al=CEN); put(ws, f"O{x0-1}", "Risks", font(9, True, "FFFFFF"), DK, al=CEN)
for j, (nm, _) in enumerate(cols):
    put(ws, f"N{x0+j}", nm, font(9), bd=BOX); put(ws, f"O{x0+j}", f"={gl(4+j)}{TOTR}", font(9), nf="0", al=CEN, bd=BOX)
mini_chart(ws, "Risks of material misstatement by assertion", _Ref(ws, min_col=14, min_row=x0, max_row=x0 + 6), _Ref(ws, min_col=15, min_row=x0 - 1, max_row=x0 + 6),
           f"N{x0+9}", color="C62828", w=12, h=7)

r = section(ws, r, "AS3", "Process -> Control gap -> Assertion at risk -> Financial impact -> Recommendation", "ISA 265.11(a) | ISA 315R | ISA 450")
r = table_head(ws, r, ["Ref", "Process", "Control gap (DF2)", "Assertion at risk", "FS line affected", "Financial impact (MNT '000)", "Impact basis",
                       "Impact vs PM", "Impact rating", "Classification (265)", "Recommendation (DF2)", ""])
ws.merge_cells(f"K{r-1}:L{r-1}")
chain = [("All - FS level", "All", 0, "Pervasive - not quantifiable"),
         ("Occurrence, Accuracy (journals)", "All / journals", 0, "Pervasive - IPE & automated controls"),
         ("Occurrence, Accuracy", "Trade payables / Cost of sales", 2100, "Exception value in IC4 sample"),
         ("Valuation & allocation", "Inventory", "=INDEX(MS_Amt,2)", "SUAD M02 (uncorrected)"),
         ("Completeness", "Accruals", "=INDEX(MS_Amt,3)", "SUAD M03 (corrected)"),
         ("Cut-off, Occurrence", "Revenue / Trade receivables", "=INDEX(MS_Amt,1)", "SUAD M01 (uncorrected)"),
         ("Existence", "Trade receivables / Cash", 0, "Potential - no misstatement found"),
         ("Existence", "Cash", "=INDEX(MS_Amt,6)", "SUAD M06 (trivial)"),
         ("Presentation & disclosure", "Disclosures - related parties", "=INDEX(MS_Amt,8)", "SUAD M08 (qualitative)"),
         ("All - FS level", "All", 0, "Pervasive - governance")]
c0 = r
for i, (asr_, line, amt, basis) in enumerate(chain, 1):
    cell(ws, f"A{r}", f"=INDEX(CD_ID,{i})", link=True, center=True)
    cell(ws, f"B{r}", f"=INDEX(CD_Proc,{i})", link=True, center=True)
    cell(ws, f"C{r}", f"=INDEX(CD_Desc,{i})", link=True)
    cell(ws, f"D{r}", asr_, is_input=True, center=True); cell(ws, f"E{r}", line, is_input=True, center=True)
    cell(ws, f"F{r}", amt, NUM, is_input=not str(amt).startswith("="), link=str(amt).startswith("="))
    cell(ws, f"G{r}", basis, is_input=True)
    cell(ws, f"H{r}", f"=IFERROR(F{r}/Mat_PM,0)", PCT, center=True)
    cell(ws, f"I{r}", f'=IF(F{r}>=Mat_PM,"Material",IF(F{r}>=Mat_CT,"Above CT",IF(F{r}>0,"Trivial",IF(LEFT(G{r},9)="Pervasive","Pervasive / qualitative","Potential"))))', center=True)
    cell(ws, f"J{r}", f"=INDEX(CD_Class,{i})", link=True, center=True)
    cell(ws, f"K{r}", f"=INDEX(CD_Rec,{i})", link=True); ws.merge_cells(f"K{r}:L{r}")
    ws.row_dimensions[r].height = 40; r += 1
c1 = r - 1
for k, v in (("Material", "FFCDD2"), ("Above CT", "FFE0B2"), ("Pervasive / qualitative", "E1BEE7")):
    ws.conditional_formatting.add(f"I{c0}:I{c1}", CellIsRule(operator="equal", formula=[f'"{k}"'], fill=fill(v)))
put(ws, f"B{r}", "Total quantified impact of control gaps", font(9, True), PALE, bd=BOX); ws.merge_cells(f"B{r}:E{r}")
put(ws, f"F{r}", f"=SUM(F{c0}:F{c1})", font(9, True), PALE, NUM, bd=BOX)
dname("AS_Impact", f"Assertions!$F${r}"); dname("AS_Quant", f"Assertions!$I${c0}:$I${c1}")
r += 1
note(ws, r, "Because effective auditing is about understanding WHAT could be misstated and WHY it matters: every control gap is expressed in terms of the assertion at risk and its potential financial impact before a recommendation is made.", "L", 20)

# ---------------- AI AGENTS ----------------
ws = S["AI_Agents"]
r = sheet_header(ws, "AI-ASSISTED AUDIT - PROCEDURES, REVIEW & INSPECTION-READINESS AGENTS", "6",
                 "Use rule-based 'agents' built into this file (and, where the firm allows, AI tools such as Microsoft Copilot, ChatGPT or Claude) to test, review and inspection-check the working papers - with traceable references, data protection, verification and professional skepticism.",
                 "ISQM 1 (resources - technological resources; information & communication; monitoring); ISA 220R.29-34 (review); ISA 230 (documentation); ISA 500.7-9 (relevance & reliability); IESBA 114 (confidentiality), 120 (skepticism)")
for k in range(ci("M"), ci("N") + 1): ws.column_dimensions[gl(k)].width = 14
widths(ws, {"B": 44, "C": 14, "D": 12, "E": 14})

r = section(ws, r, "AI1", "Responsible AI protocol (complete before using any AI tool on this engagement)", "ISQM 1 | IESBA 114 & 120 | ISA 230 | ISA 500")
r = checklist(ws, r, [
    ("Only firm-approved AI tools / enterprise tenants used (Copilot, ChatGPT Enterprise, Claude for Work - per firm policy)", "ISQM 1 - technological resources", "Yes", "Firm AI policy v2.1"),
    ("No confidential client data entered into non-approved tools; data anonymised or synthetic where required", "IESBA 114 (confidentiality)", "Yes", "Client name & IDs replaced with codes"),
    ("AI output treated as a draft - never as audit evidence by itself; relevance & reliability evaluated", "ISA 500.7-9", "Yes", ""),
    ("Every AI-assisted output verified by a qualified team member and traced to source documents / WP references", "ISA 220R.30-31; ISA 230.9", "Yes", "Reviewer column in AI5"),
    ("Professional skepticism applied - challenge AI conclusions, look for contradictory evidence and hallucinated references", "ISA 200.15; IESBA 120", "Yes", ""),
    ("AI use documented: tool, purpose, prompt, data class, verification, WP ref (AI5 log)", "ISA 230.8", "Yes", ""),
    ("Agents (prompts) shared & reused across teams only via firm's approved library (AI6)", "ISQM 1 - intellectual resources", "Yes", ""),
], "AI1")

r = section(ws, r, "AI2", "Audit Procedures Agent - automated exception tests across the file", "ISA 330 | ISA 500 | ISA 530")
r = table_head(ws, r, ["#", "Automated test", "Exceptions", "Threshold", "Result", "Source WP", "Next step for the auditor", "", "", "", "Link"])
ws.merge_cells(f"G{r-1}:J{r-1}")
apa = [("Lead schedule movements > PM and > 10% needing explanation", '=COUNTIF(L4_Execution!K:K,"Explain")', 0, "D1", "Document explanations & corroborate", "L4_Execution"),
       ("Trial balance out of balance", "=IF(ABS(SUM(TB_CY))>0.5,1,0)", 0, "D1", "Reconcile TB to GL", "L4_Execution"),
       ("MUS sample misstatements (receivables)", '=COUNTIF(Rec_Mis,"<>0")', 0, "S4-S5", "Investigate cause; project; record in SUAD", "Sampling"),
       ("Controls not operating effectively", '=COUNTIF(ToE_Result,"Not effective")', 0, "IC4", "Withdraw reliance; extend substantive (ISA 330.17)", "IC_Assessment"),
       ("Journal entry exceptions", "=JE_Exc", 0, "FR4", "Obtain support; evaluate fraud indicators", "Fraud_240R"),
       ("Confirmation differences / non-responses", '=COUNTIF(Conf_Status,"Received - reconciled")+COUNTIF(Conf_Status,"Not received*")+COUNTIF(Conf_Status,"Outstanding")', 0, "D4", "Reconcile differences; alternative procedures", "L4_Execution"),
       ("Analytical procedures outside threshold", '=COUNTIF(L3_Methodology!I:I,"Investigate")', 0, "C3", "Obtain & corroborate explanations (ISA 520.7)", "L3_Methodology"),
       ("PBC items overdue (information not yet obtained)", "=PBC_Overdue", 0, "PBC1", "Escalate; consider scope limitation", "PBC_List")]
a0 = r
for i, (t, f_, th, src, nxt, sh) in enumerate(apa, 1):
    cell(ws, f"A{r}", i, center=True); cell(ws, f"B{r}", t); cell(ws, f"C{r}", f_, "0", center=True); cell(ws, f"D{r}", th, "0", True, center=True)
    cell(ws, f"E{r}", f'=IF(C{r}<=D{r},"Clear","Exceptions ("&C{r}&")")', center=True); cell(ws, f"F{r}", src, center=True)
    cell(ws, f"G{r}", nxt); ws.merge_cells(f"G{r}:J{r}")
    c = cell(ws, f"K{r}", f'=HYPERLINK("#{sh}!A1","Open >")', center=True); c.font = Font(name=F, size=9, color="1565C0", underline="single"); r += 1
ws.conditional_formatting.add(f"E{a0}:E{r-1}", FormulaRule(formula=[f'LEFT($E{a0},10)="Exceptions"'], fill=fill("FFE0B2")))
ws.conditional_formatting.add(f"E{a0}:E{r-1}", CellIsRule(operator="equal", formula=['"Clear"'], fill=fill("C8E6C9")))
dname("AI_ProcExc", f"AI_Agents!$E${a0}:$E${r-1}")
r += 1

r = section(ws, r, "AI3", "Manager / Partner Review Agent - automated review points", "ISA 220R.30-34 | ISA 230.9")
r = table_head(ws, r, ["#", "Review check", "Issues found", "Severity", "Result", "Standard", "What the reviewer should do", "", "", "", "Link"])
ws.merge_cells(f"G{r-1}:J{r-1}")
rva = [("Prepared/reviewed WPs without preparer sign-off", '=SUMPRODUCT((WP_Status<>"Not started")*(WP_Status<>"In progress")*(WP_Prep=""))', "High", "ISA 230.9(a)", "Obtain preparer sign-off & date", "WP_Index"),
       ("Reviewed/closed WPs without reviewer sign-off", '=SUMPRODUCT(((WP_Status="Reviewed")+(WP_Status="Closed"))*(WP_Rev=""))', "High", "ISA 230.9(c)", "Evidence review with name & date", "WP_Index"),
       ("Risk assessments needing justification (RA3 consistency check)", '=COUNTIF(Risk_Check,"<>OK")', "Medium", "ISA 315R.31-32", "Document rationale or revise", "Risk_315R"),
       ("Fraud risks not assessed as significant", '=COUNTIFS(Risk_Fraud,"Yes",Risk_Spectrum,"<>Significant")', "High", "ISA 240R", "Treat as significant risk", "Risk_315R"),
       ("Uncorrected misstatements above PM (aggregate)", "=IF(ABS(MS_Agg)>=Mat_PM,1,0)", "High", "ISA 450.8-11", "Request correction; reassess", "SUAD"),
       ("Significant deficiencies not communicated in writing", '=COUNTIF(CD_Class,"Significant deficiency")-CD_SigComm', "High", "ISA 265.9", "Issue written communication to TCWG", "Deficiencies_265"),
       ("Fraud matters involving management not yet reported to TCWG", '=COUNTIFS(Fraud_Mgmt,"Yes",Fraud_TCWG,"")', "High", "ISA 240R", "Communicate on a timely basis", "Fraud_240R"),
       ("Open high-severity issues", "=DD_OpenHigh", "Medium", "ISQM 1.42", "Resolve or escalate before sign-off", "L5_Findings"),
       ("Independence declarations outstanding", "=DD_IndTot-DD_IndRecv", "High", "IESBA 400; ISA 220R.17", "Obtain before further work", "L1_Strategy"),
       ("TCWG communications pending", '=COUNTIF(Comm_Status,"Pending")', "Medium", "ISA 260", "Schedule completion meeting", "L6_Oversight"),
       ("EQ review not completed", '=IF(OR(DD_EQR="Reviewed",DD_EQR="Closed"),0,1)', "High", "ISQM 2; ISA 220R.36", "Complete before report is dated", "L6_Oversight"),
       ("Going concern assessment < 12 months from FS approval", "=IF(GC_Months<12,1,0)", "High", "ISA 570R", "Request extension", "GoingConcern_570R"),
       ("MUS upper misstatement limit exceeds tolerable", "=IF(Rec_UMLpct>1,1,0)", "Medium", "ISA 530.14", "Extend testing / adjust", "Sampling"),
       ("Team members over-allocated", "=TR_Over", "Low", "ISA 220R.25", "Rebalance schedule", "Team_Resources"),
       ("PBC items overdue", "=PBC_Overdue", "Medium", "ISA 210.6(b)(iii)", "Escalate to CFO / TCWG", "PBC_List")]
v0 = r
for i, (t, f_, sev, std_, act, sh) in enumerate(rva, 1):
    cell(ws, f"A{r}", i, center=True); cell(ws, f"B{r}", t); cell(ws, f"C{r}", f_, "0", center=True)
    cell(ws, f"D{r}", sev, is_input=True, center=True); cell(ws, f"E{r}", f'=IF(C{r}=0,"Pass","Flag ("&C{r}&")")', center=True)
    cell(ws, f"F{r}", std_, center=True); cell(ws, f"G{r}", act); ws.merge_cells(f"G{r}:J{r}")
    c = cell(ws, f"K{r}", f'=HYPERLINK("#{sh}!A1","Open >")', center=True); c.font = Font(name=F, size=9, color="1565C0", underline="single"); r += 1
v1 = r - 1
dv(ws, f"D{v0}:D{v1}", "Severity")
ws.conditional_formatting.add(f"E{v0}:E{v1}", FormulaRule(formula=[f'AND(LEFT($E{v0},4)="Flag",$D{v0}="High")'], fill=fill("FFCDD2"), font=Font(name=F, bold=True, color="B71C1C")))
ws.conditional_formatting.add(f"E{v0}:E{v1}", FormulaRule(formula=[f'LEFT($E{v0},4)="Flag"'], fill=fill("FFE0B2")))
ws.conditional_formatting.add(f"E{v0}:E{v1}", CellIsRule(operator="equal", formula=['"Pass"'], fill=fill("C8E6C9")))
put(ws, f"B{r}", "Review flags open (high)", font(9, True), PALE, bd=BOX)
put(ws, f"C{r}", f'=COUNTIF(E{v0}:E{v1},"Flag*")', font(10, True, DK), PALE, "0", CEN, BOX)
put(ws, f"D{r}", f'=COUNTIFS(E{v0}:E{v1},"Flag*",D{v0}:D{v1},"High")', font(10, True, "B71C1C"), PALE, "0", CEN, BOX)
dname("AI_Flags", f"AI_Agents!$C${r}"); dname("AI_FlagsHigh", f"AI_Agents!$D${r}"); dname("AI_Checks", f"AI_Agents!$A${v1}")
r += 2

r = section(ws, r, "AI4", "Audit Inspection Readiness Agent - simulated inspection review", "ISQM 1.35-47 (monitoring) | regulator inspection themes")
r = table_head(ws, r, ["#", "Inspection theme", "Weight", "Score (0-100%)", "Weighted", "Standard", "Likely inspection finding if < 100%", "", "", "", "Link"])
ws.merge_cells(f"G{r-1}:J{r-1}")
ins = [("Documentation & sign-offs complete", 3, '=1-IFERROR(C' + str(v0) + '/COUNTIFS(WP_Status,"<>Not started",WP_Status,"<>In progress"),0)', "ISA 230", "Unsigned / undated working papers", "WP_Index"),
       ("Risk assessment documented & consistent (IRFs, spectrum)", 3, "=1-IFERROR(COUNTIF(Risk_Check,\"<>OK\")/COUNTA(Risk_ID),0)", "ISA 315R", "Insufficient rationale for risk assessments", "Risk_315R"),
       ("Fraud risks & responses (incl. JE testing)", 3, f'=IF(AND(C{v0+3}=0,JE_Exc<=JE_Tested),IF(C{v0+6}=0,1,0.5),0)', "ISA 240R", "Fraud procedures / communications not evidenced", "Fraud_240R"),
       ("Controls reliance supported by testing", 2, "=DD_CtlPct", "ISA 330.8-17", "Reliance on controls without sufficient testing", "IC_Assessment"),
       ("Misstatements evaluated & communicated", 2, f"=IF(C{v0+4}=0,1,0.6)", "ISA 450", "Aggregate uncorrected misstatements near materiality", "SUAD"),
       ("Deficiencies communicated (ISA 265)", 2, '=IFERROR(CD_SigComm/COUNTIF(CD_Class,"Significant deficiency"),1)', "ISA 265", "Significant deficiencies not communicated in writing", "Deficiencies_265"),
       ("Going concern evaluation (12-month period)", 2, "=IF(GC_Months>=12,1,0)", "ISA 570R", "Assessment period too short", "GoingConcern_570R"),
       ("Communication with TCWG", 2, "=DD_CommPct", "ISA 260", "Required communications not made", "L6_Oversight"),
       ("Independence confirmations", 2, "=DD_IndPct", "IESBA 400", "Missing independence confirmations", "L1_Strategy"),
       ("Engagement quality review", 3, f"=IF(C{v0+10}=0,1,0)", "ISQM 2", "EQR not completed before report date", "L6_Oversight"),
       ("Audit evidence obtained (PBC / IPE)", 1, "=PBC_PctRec", "ISA 500.9", "IPE not tested; evidence outstanding", "PBC_List")]
n0 = r
for i, (t, w, f_, std_, fnd, sh) in enumerate(ins, 1):
    cell(ws, f"A{r}", i, center=True); cell(ws, f"B{r}", t); cell(ws, f"C{r}", w, "0", True, center=True)
    cell(ws, f"D{r}", f_, "0%", center=True); cell(ws, f"E{r}", f"=C{r}*D{r}", "0.0", center=True); cell(ws, f"F{r}", std_, center=True)
    cell(ws, f"G{r}", f'=IF(D{r}>=0.999,"-","{fnd}")'); ws.merge_cells(f"G{r}:J{r}")
    c = cell(ws, f"K{r}", f'=HYPERLINK("#{sh}!A1","Open >")', center=True); c.font = Font(name=F, size=9, color="1565C0", underline="single"); r += 1
n1 = r - 1
ws.conditional_formatting.add(f"D{n0}:D{n1}", CellIsRule(operator="lessThan", formula=["0.75"], fill=fill("FFCDD2")))
ws.conditional_formatting.add(f"D{n0}:D{n1}", CellIsRule(operator="between", formula=["0.75", "0.999"], fill=fill("FFE0B2")))
ws.conditional_formatting.add(f"D{n0}:D{n1}", CellIsRule(operator="greaterThanOrEqual", formula=["0.999"], fill=fill("C8E6C9")))
put(ws, f"B{r}", "INSPECTION READINESS SCORE", font(9, True), PALE, bd=BOX)
put(ws, f"D{r}", f"=IFERROR(SUM(E{n0}:E{n1})/SUM(C{n0}:C{n1}),0)", font(12, True, DK), PALE, "0%", CEN, BOX)
put(ws, f"E{r}", f'=IF(D{r}>=0.9,"Inspection ready",IF(D{r}>=0.75,"Minor findings likely - remediate before sign-off","Significant findings likely - do not sign"))', font(9, True, DK), PALE, bd=BOX)
ws.merge_cells(f"E{r}:J{r}")
dname("AI_Ready", f"AI_Agents!$D${r}"); dname("AI_Rating", f"AI_Agents!$E${r}")
r += 2

r = section(ws, r, "AI5", "AI use log - traceability & verification", "ISA 230.8 | ISQM 1 | IESBA 114")
r = table_head(ws, r, ["#", "Purpose / agent", "Date", "Team member", "AI tool", "Data classification", "Approved tool?", "Data anonymised?", "Verified by", "Verification method", "WP ref", "Status"])
logs = [("Audit Procedures Agent - summarise exceptions in JE listing", D(2026, 2, 12), "T. Enkhjin", "Microsoft Copilot (tenant)", "Confidential", "Yes", "Yes", "S. Oyunaa", "Re-performed 20 items to source", "FR4"),
        ("Draft ISA 265 letter wording from DF2 register", D(2026, 3, 18), "S. Oyunaa", "Claude for Work", "Anonymised", "Yes", "Yes", "B. Bat-Erdene", "Compared to DF2 & ISA 265.11", "DF5"),
        ("Review Agent - check WP index for missing sign-offs", D(2026, 3, 19), "S. Oyunaa", "Workbook rules (AI3)", "Confidential", "Yes", "N/A", "B. Bat-Erdene", "Spot-check 10 WPs", "AI3"),
        ("Summarise IFRS 16 lease contracts", D(2026, 2, 5), "M. Saruul", "ChatGPT Enterprise", "Confidential", "Yes", "No", "", "", "D3"),
        ("Inspection readiness pre-check", D(2026, 3, 20), "S. Oyunaa", "Workbook rules (AI4)", "Confidential", "Yes", "N/A", "B. Bat-Erdene", "Reviewed results", "AI4")]
g0 = r
for i, (pur, d_, who, tool, cls, app, anon, ver, meth, wp) in enumerate(logs, 1):
    cell(ws, f"A{r}", i, center=True); cell(ws, f"B{r}", pur, is_input=True); cell(ws, f"C{r}", d_, DATE, True, center=True)
    for col, v in zip("DEFGHIJ", (who, tool, cls, app, anon, ver, meth)): cell(ws, f"{col}{r}", v, is_input=True, center=True)
    cell(ws, f"K{r}", wp, center=True)
    cell(ws, f"L{r}", f'=IF(G{r}<>"Yes","NON-COMPLIANT - unapproved tool",IF(AND(F{r}="Confidential",H{r}="No"),"NON-COMPLIANT - confidential data not anonymised",IF(I{r}="","Pending verification","Compliant")))', center=True)
    ws.row_dimensions[r].height = 28; r += 1
dv(ws, f"F{g0}:F{r-1}", "DataClass"); dv(ws, f"G{g0}:H{r-1}", "YesNo")
ws.conditional_formatting.add(f"L{g0}:L{r-1}", FormulaRule(formula=[f'LEFT($L{g0},3)="NON"'], fill=fill("FFCDD2"), font=Font(name=F, bold=True, color="B71C1C")))
ws.conditional_formatting.add(f"L{g0}:L{r-1}", CellIsRule(operator="equal", formula=['"Compliant"'], fill=fill("C8E6C9")))
dname("AI_LogStatus", f"AI_Agents!$L${g0}:$L${r-1}")
note(ws, r, "Entry 4 shows a typical breach: confidential lease data entered without anonymisation and no verification - report under firm policy and re-perform the work without AI.", "L", 14)
r += 2

r = section(ws, r, "AI6", "Agent prompt library (reusable across teams; no client identifiers)", "ISQM 1 - intellectual & technological resources")
r = table_head(ws, r, ["#", "Agent", "Prompt template (paste into firm-approved Copilot / ChatGPT / Claude)", "", "", "", "", "", "", "Human verification required"])
ws.merge_cells(f"C{r-1}:I{r-1}"); ws.merge_cells(f"J{r-1}:L{r-1}")
prompts = [("Audit Procedures Agent", "You are an audit senior applying ISA 330/500. Using the attached anonymised listing [population], test each item against [attributes]. List exceptions with item ID, attribute failed and amount. Cite the source row for every exception. Do not infer missing data - flag it.", "Re-perform a sample of exceptions & clears to source documents; record in IC5/S4."),
           ("Manager / Partner Review Agent", "Act as an engagement manager reviewing working paper [WP ref] under ISA 220 (Revised) and ISA 230. Identify: missing sign-offs, unsupported conclusions, inconsistencies with the risk assessment [RA3 ref], and missing cross-references. Output a review-notes table: WP ref, issue, ISA ref, required action.", "Reviewer confirms each point against the WP before raising; clears notes in E3."),
           ("Audit Inspection Readiness Agent", "Simulate an audit regulator inspection of this engagement file index [WP_Index export]. For each theme (ISA 230, 315R, 240R, 330, 450, 265, 570R, 260, ISQM 2, independence) state whether documentation appears sufficient, likely findings, and evidence a reviewer should look for.", "Partner evaluates results; AI4 score remains the documented assessment."),
           ("Assertion & impact coach", "For the control gap [description] in process [process], identify the financial statement assertion(s) at risk, the FS line(s) affected, how a misstatement could arise, an estimate method for financial impact, and a specific SMART recommendation.", "Auditor quantifies impact from SUAD / sample data in AS3; never accept AI figures without source.")]
for i, (ag, pr, ver) in enumerate(prompts, 1):
    cell(ws, f"A{r}", i, center=True); cell(ws, f"B{r}", ag).font = font(9, True)
    cell(ws, f"C{r}", pr, is_input=True); ws.merge_cells(f"C{r}:I{r}"); cell(ws, f"J{r}", ver); ws.merge_cells(f"J{r}:L{r}")
    ws.row_dimensions[r].height = 58; r += 1

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
 ("A1", "Strategy", "Acceptance & continuance summary", "ISQM 1.30; ISA 220R.22-24", "Reviewed", 1, 1),
 ("AC1", "Strategy", "Engagement profile", "ISQM 1.30(a)", "Closed", 1, 1),
 ("AC2", "Strategy", "Engagement risk assessment (weighted)", "ISQM 1.30(c)", "Reviewed", 2, 2),
 ("AC3", "Strategy", "KYC / AML & sanctions screening", "IESBA 320; AML law", "Reviewed", 2, 2),
 ("AC4", "Strategy", "New-client & predecessor procedures", "IESBA 320.8; ISA 510", "Closed", 0.5, 0.5),
 ("AC5", "Strategy", "Continuance considerations", "ISA 220R.23", "Reviewed", 1, 1),
 ("AC6", "Strategy", "Competence & resources check", "ISQM 1.30(a)(i)", "Reviewed", 0.5, 0.5),
 ("AC7", "Strategy", "Decision & approvals", "ISQM 1.30-31", "Closed", 1, 1),
 ("TR1", "Strategy", "Team competency matrix", "ISA 220R.25-26; IES 8", "Reviewed", 1, 1),
 ("TR2", "Strategy", "Competency needs vs coverage", "ISA 220R.26", "Reviewed", 1, 1),
 ("TR3", "Strategy", "Budget, actual hours & utilisation", "ISA 220R.25(b)", "In progress", 1, 1),
 ("TR4", "Strategy", "Monthly capacity plan", "ISQM 1.32", "Reviewed", 2, 2),
 ("TR5", "Strategy", "Direction, supervision & review plan", "ISA 220R.29-34", "Reviewed", 1, 1),
 ("TR6", "Strategy", "Resources conclusion", "ISA 220R.25-28", "Reviewed", 0.5, 0.5),
 ("PBC1", "Planning", "PBC request list & tracker", "ISA 210.6(b)(iii); ISA 500.9", "In progress", 4, 4),
 ("PBC2", "Planning", "PBC procedures & escalation", "ISA 300; ISA 260.16(b)", "Prepared", 1, 1),
 ("PBC3", "Execution", "PBC progress dashboard", "Engagement management", "In progress", 1, 1),
 ("A2", "Strategy", "Terms of engagement & preconditions", "ISA 210", "Closed", 3, 2.5),
 ("A3", "Strategy", "Independence declarations - team", "IESBA Part 4A; ISA 220R.16-21", "Prepared", 4, 5),
 ("A4", "Strategy", "Threats & safeguards register", "IESBA 120; 300", "Reviewed", 3, 3),
 ("A5", "Strategy", "Fee dependency & non-assurance services", "IESBA 410; 600", "Reviewed", 2, 1.5),
 ("A6", "Strategy", "Overall audit strategy", "ISA 300.7-8", "Reviewed", 8, 9),
 ("A7", "Strategy", "Planned scope communicated to TCWG", "ISA 260.15", "Closed", 2, 2),
 ("B1", "Planning", "Understanding the entity & environment", "ISA 315R.19", "Reviewed", 12, 14),
 ("B2", "Planning", "Understanding system of internal control", "ISA 315R.21-26", "Reviewed", 18, 21),
 ("B3", "Planning", "Materiality", "ISA 320; ISA 450", "Closed", 3, 3),
 ("B4", "Planning", "Risk assessment summary", "ISA 315R.28-38", "Reviewed", 2, 2),
 ("RA1", "Planning", "IRF & spectrum methodology", "ISA 315R.12, 31; App. 2", "Closed", 1, 1),
 ("RA2", "Planning", "Significant classes, balances & disclosures", "ISA 315R.29; ISA 330.18", "Reviewed", 3, 3),
 ("RA3", "Planning", "Risk register - IRFs, spectrum, combined RMM", "ISA 315R.28, 31-34", "Reviewed", 10, 12),
 ("RA4", "Planning", "Inherent risk heat map", "ISA 315R.31", "Reviewed", 1, 1),
 ("RA5", "Planning", "Financial-statement-level risks", "ISA 315R.30; ISA 330.5", "Reviewed", 2, 2),
 ("RA6", "Planning", "Stand-back & documentation", "ISA 315R.35-38", "Prepared", 2, 2),
 ("AS1", "Planning", "Assertion framework", "ISA 315R.A190", "Closed", 0.5, 0.5),
 ("AS2", "Planning", "Assertion risk matrix by SCOTABD", "ISA 315R.28-29", "Reviewed", 1, 1),
 ("AS3", "Findings", "Control gap > assertion > impact > recommendation", "ISA 265.11; ISA 450", "Prepared", 2, 2),
 ("B5", "Planning", "Fraud summary", "ISA 240 (Revised)", "Reviewed", 1, 1),
 ("FR1", "Planning", "Skepticism & engagement team discussion", "ISA 240R", "Reviewed", 2, 2),
 ("FR2", "Planning", "Fraud lens - risk factors & inquiries", "ISA 240R; ISA 315R", "Reviewed", 5, 6),
 ("FR3", "Planning", "Risks of material misstatement due to fraud", "ISA 240R", "Reviewed", 2, 2),
 ("FR4", "Execution", "Responses - JE testing, estimates, unusual txns", "ISA 240R; ISA 330", "Prepared", 14, 15),
 ("FR5", "Execution", "Fraud or suspected fraud (incl. allegations)", "ISA 240R", "In progress", 4, 5),
 ("FR6", "Oversight", "Ongoing fraud communications", "ISA 240R; ISA 260", "In progress", 1, 1),
 ("FR7", "Oversight", "Report, representations & documentation", "ISA 240R; ISA 701; ISA 580", "Not started", 2, 0),
 ("B6", "Planning", "Laws & regulations", "ISA 250", "Prepared", 4, 3),
 ("B7", "Planning", "Related parties", "ISA 550", "In progress", 5, 3),
 ("B8", "Planning", "Going concern summary", "ISA 570 (Revised 2024)", "Reviewed", 1, 1),
 ("GC1", "Planning", "Events or conditions & entity's process", "ISA 570R; ISA 315R", "Reviewed", 3, 3),
 ("GC2", "Planning", "Evaluating management's assessment (12 months)", "ISA 570R; ISA 540R", "Prepared", 6, 7),
 ("GC3", "Planning", "Cash-flow forecast & stress test", "ISA 570R", "Prepared", 4, 5),
 ("GC4", "Oversight", "Conclusions - basis & material uncertainty", "ISA 570R", "In progress", 2, 1),
 ("GC5", "Oversight", "Draft going concern report section", "ISA 570R; ISA 700R", "In progress", 1, 1),
 ("GC6", "Oversight", "TCWG, representations & documentation", "ISA 570R; ISA 260; ISA 580", "Not started", 1, 0),
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
 ("D6", "Execution", "Journal entry testing", "ISA 240 (Revised)", "Prepared", 12, 14),
 ("D7", "Execution", "Litigation & claims", "ISA 501.9-12", "In progress", 3, 1),
 ("D8", "Execution", "Subsequent events", "ISA 560", "In progress", 4, 2),
 ("D9", "Execution", "Comparative information", "ISA 710", "Reviewed", 2, 2),
 ("IC4", "Execution", "Tests of operating effectiveness (attribute)", "ISA 330.8-11; ISA 530", "Reviewed", 24, 27),
 ("IC5", "Execution", "Attribute testing detail - C01", "ISA 330.10", "Reviewed", 6, 6),
 ("IC6", "Execution", "Evaluation of deviations", "ISA 330.16-17; ISA 265", "Prepared", 3, 3),
 ("E1", "Findings", "Summary of misstatements", "ISA 450", "Prepared", 5, 6),
 ("E2", "Findings", "Control deficiencies summary", "ISA 265", "Prepared", 1, 1),
 ("DF1", "Findings", "ISA 265 definitions & evaluation framework", "ISA 265.6, 8; A5-A11", "Closed", 1, 1),
 ("DF2", "Findings", "Deficiency register & evaluation", "ISA 265.7-8", "Reviewed", 5, 6),
 ("DF3", "Findings", "Evaluation in combination", "ISA 265.8; A6", "Reviewed", 1, 1),
 ("DF4", "Findings", "Communication requirements & tracking", "ISA 265.9-11", "Prepared", 1, 1),
 ("DF5", "Findings", "Written communication to TCWG & management", "ISA 265.9-11", "Prepared", 2, 2),
 ("DF6", "Findings", "Follow-up of prior-year deficiencies", "ISA 265.A28-A29", "Reviewed", 1, 1),
 ("E3", "Findings", "Issue log & remediation tracker", "ISA 230; ISQM 1.42", "In progress", 4, 3),
 ("E4", "Findings", "Management letter", "ISA 265.9-11", "In progress", 5, 3),
 ("SUAD1", "Findings", "SUAD - all current-year misstatements", "ISA 450.5-8", "Prepared", 4, 5),
 ("SUAD2", "Findings", "SUAD - prior-period reversal effects", "ISA 450.11(b)", "Prepared", 1, 1),
 ("SUAD3", "Findings", "SUAD - evaluation (rollover / iron curtain)", "ISA 450.10-12", "In progress", 2, 1),
 ("SUAD4", "Findings", "SUAD - by FS line item", "ISA 450.A16", "Prepared", 1, 1),
 ("F1", "Oversight", "Communications with TCWG", "ISA 260; ISA 265", "In progress", 6, 3),
 ("F2", "Oversight", "Written representations", "ISA 580", "Not started", 2, 0),
 ("F3", "Oversight", "Going concern - final conclusion", "ISA 570 (Revised 2024)", "Prepared", 1, 1),
 ("F4", "Oversight", "Key audit matters", "ISA 701", "In progress", 4, 2),
 ("F5", "Oversight", "Engagement quality review", "ISQM 2; ISA 220R.36", "Not started", 10, 0),
 ("F6", "Oversight", "Completion checklist & partner sign-off", "ISA 220R.37-40", "Not started", 6, 0),
 ("F7", "Oversight", "Forming the opinion", "ISA 700R; 705R; 706R", "Not started", 4, 0),
 ("F8", "Oversight", "Other information", "ISA 720R", "Not started", 2, 0),
 ("F9", "Oversight", "Assembly of final file", "ISA 230.14; ISQM 1.A83", "Not started", 4, 0),
 ("F10", "Oversight", "ISQM 1 monitoring & quality indicators", "ISQM 1.35-47", "In progress", 3, 1),
 ("AI1", "Oversight", "Responsible AI protocol", "ISQM 1; IESBA 114, 120", "Reviewed", 0.5, 0.5),
 ("AI2", "Oversight", "Audit Procedures Agent - exception tests", "ISA 330; ISA 500", "Reviewed", 1, 1),
 ("AI3", "Oversight", "Manager/Partner Review Agent", "ISA 220R.30-34; ISA 230", "In progress", 1, 1),
 ("AI4", "Oversight", "Inspection Readiness Agent", "ISQM 1.35-47", "In progress", 1, 1),
 ("AI5", "Oversight", "AI use log & verification", "ISA 230.8", "In progress", 0.5, 0.5),
 ("AI6", "Oversight", "Agent prompt library", "ISQM 1 resources", "Closed", 0.5, 0.5)]
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
dname("WP_Bud", f"WP_Index!$J${w0}:$J${w1}"); dname("WP_Prep", f"WP_Index!$E${w0}:$E${w1}"); dname("WP_Rev", f"WP_Index!$G${w0}:$G${w1}"); dname("WP_Act", f"WP_Index!$K${w0}:$K${w1}")

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

put(dd, "A19", "Spectrum of inherent risk (ISA 315R)", font(9, True, "FFFFFF"), DK); put(dd, "B19", "Count", font(9, True, "FFFFFF"), DK)
for i, s in enumerate(["Low", "Moderate", "Elevated", "High", "Significant"]):
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
        ("HighIRF", "Risks with an IRF rated 3", "=COUNTIF(Risk_MaxIRF,3)", "0"),
        ("RiskFlags", "Risk assessments needing justification", '=COUNTIF(Risk_Check,"<>OK")', "0"),
        ("SigSCOT", "Significant SCOTABDs", '=COUNTIF(SCOT_Class,"Significant*")', "0"),
        ("MatNotSig", "Material - not significant", '=COUNTIF(SCOT_Class,"Material*")', "0"),
        ("FraudN", "Fraud matters logged", "=COUNTA(Fraud_Status)", "0"),
        ("FraudOpen", "Fraud matters open", '=COUNTIF(Fraud_Status,"<>Closed")', "0"),
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
stds = [("ISQM 1", "Quality Management for Firms that Perform Audits or Reviews", "Strategy / Oversight", "A1, AC1-AC7, TR1-TR6, E3, F10"),
        ("ISQM 2", "Engagement Quality Reviews", "Oversight", "F5"),
        ("IESBA Code", "International Code of Ethics for Professional Accountants (incl. Independence Standards)", "Strategy", "A3, A4, A5"),
        ("ISA 200", "Overall Objectives of the Independent Auditor", "All", "-"),
        ("ISA 210", "Agreeing the Terms of Audit Engagements", "Strategy", "A2, AC7, PBC1-PBC2"),
        ("ISA 220 (Revised)", "Quality Management for an Audit of Financial Statements", "Strategy / Oversight", "A1, A3, AC1-AC7, TR1-TR6, B10, F5, F6"),
        ("ISA 230", "Audit Documentation", "All", "WP_Index, D1, E3, F9"),
        ("ISA 240 (Revised)", "Auditor's Responsibilities Relating to Fraud in an Audit of FS (issued Jul-2025; effective periods beginning on/after 15-Dec-2026)", "Planning / Execution / Oversight", "B5, FR1-FR7, D6"),
        ("ISA 250 (Revised)", "Consideration of Laws and Regulations", "Planning", "B6"),
        ("ISA 260 (Revised)", "Communication with Those Charged with Governance", "Strategy / Oversight", "A7, F1"),
        ("ISA 265", "Communicating Deficiencies in Internal Control to TCWG and Management", "Findings", "E2, E4, IC6, DF1-DF6"),
        ("ISA 300", "Planning an Audit of Financial Statements", "Strategy / Planning", "A6, B10"),
        ("ISA 315 (Revised 2019)", "Identifying and Assessing the Risks of Material Misstatement", "Planning", "B1, B2, B4, RA1-RA6, AS1-AS3, IC1-IC3"),
        ("ISA 320", "Materiality in Planning and Performing an Audit", "Planning", "B3"),
        ("ISA 330", "Auditor's Responses to Assessed Risks", "Methodology / Execution", "C1, D2, D3, IC4-IC6"),
        ("ISA 402", "Audit Considerations Relating to Service Organizations", "Methodology", "C8"),
        ("ISA 450", "Evaluation of Misstatements Identified during the Audit", "Findings", "E1, SUAD1-SUAD4"),
        ("ISA 500", "Audit Evidence (incl. 500.9 information produced by the entity)", "Execution", "D1, D3, PBC1-PBC3"),
        ("ISA 501", "Audit Evidence - Specific Considerations", "Execution", "D5, D7"),
        ("ISA 505", "External Confirmations", "Execution", "D4"),
        ("ISA 510", "Initial Audit Engagements - Opening Balances", "Planning", "B9"),
        ("ISA 520", "Analytical Procedures", "Methodology", "C3"),
        ("ISA 530", "Audit Sampling", "Methodology", "C2, S1-S5, IC4-IC5"),
        ("ISA 540 (Revised)", "Auditing Accounting Estimates and Related Disclosures", "Methodology", "C4"),
        ("ISA 550", "Related Parties", "Planning", "B7"),
        ("ISA 560", "Subsequent Events", "Execution", "D8"),
        ("ISA 570 (Revised 2024)", "Going Concern (issued Apr-2025; effective periods beginning on/after 15-Dec-2026)", "Planning / Oversight", "B8, F3, GC1-GC6"),
        ("ISA 580", "Written Representations", "Oversight", "F2"),
        ("ISA 600 (Revised)", "Audits of Group Financial Statements", "Methodology", "C7"),
        ("ISA 610 (Revised 2013)", "Using the Work of Internal Auditors", "Methodology", "C5"),
        ("ISA 620", "Using the Work of an Auditor's Expert", "Methodology", "C6"),
        ("ISA 700 (Revised)", "Forming an Opinion and Reporting on Financial Statements", "Oversight", "F7"),
        ("ISA 701", "Communicating Key Audit Matters", "Oversight", "F4"),
        ("ISA 705 (Revised)", "Modifications to the Opinion", "Oversight", "F7"),
        ("ISA 706 (Revised)", "Emphasis of Matter & Other Matter Paragraphs", "Oversight", "F7"),
        ("ISA 710", "Comparative Information", "Execution", "D9"),
        ("ISA 720 (Revised)", "Auditor's Responsibilities Relating to Other Information", "Oversight", "F8"),
        ("ISQM 1 - technological resources / firm AI policy", "Use of AI tools: approved tools, confidentiality, verification, documentation", "Oversight", "AI1-AI6")]
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
                 ["Acceptance & continuance (AC1-AC7)", "Engagement letter (A2) & independence (A3-A5)", "Team competency & capacity (TR1-TR6)", "Overall strategy (A6) & PBC list"], "L1_Strategy"),
    "Planning": ("B19", ["Understanding entity & IC", "Materiality setting", "Inherent risk factors & spectrum", "Fraud & going concern lens"],
                 ["IRF risk register & spectrum (RA1-RA6)", "Fraud lens - ISA 240R (FR1-FR3)", "Going concern - ISA 570R (GC1-GC3)", "Internal controls (IC1-IC3)"], "L2_Planning"),
    "Execution": ("L19", ["Fieldwork execution", "Control testing", "Data validation", "Process walkthrough"],
                  ["Lead schedule (D1)", "Tests of controls - attribute (IC4-IC5)", "Confirmations & MUS selection (D4, S4)", "JE testing (D6)"], "L4_Execution"),
    "Methodology": ("B33", ["Standard procedures", "Sampling techniques", "Evaluation methods", "Quality standards"],
                    ["Audit programme (C1)", "Population register & sizes (S1)", "Statistical factors (S2-S3)", "MUS evaluation / UML (S5)"], "L3_Methodology"),
    "Oversight": ("L33", ["TCWG interaction", "Quality assurance reviews", "Independence monitoring", "Performance evaluation"],
                  ["TCWG log (F1)", "EQ review (F5)", "Opinion matrix (F7)", "AI review & inspection readiness (AI1-AI6)"], "L6_Oversight"),
    "Findings": ("G47", ["Issue identification", "Root cause analysis", "Recommendation design", "Remediation tracking"],
                 ["SUAD - all misstatements (SUAD1-4)", "ISA 265 deficiencies & letter (DF1-DF6)", "Issue log (E3)", "Management letter (E4)"], "L5_Findings"),
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
for rr in range(6, 76):
    for cc in (2, 3): db.cell(rr, cc).fill = fill(DK)
navs = [("Overview", "LayersMap"), ("Setup", "Setup"), ("WP Index", "WP_Index"), ("1 Strategy", "L1_Strategy"), ("   Acceptance", "Acceptance_QM"), ("   Team", "Team_Resources"), ("   PBC", "PBC_List"),
        ("2 Planning", "L2_Planning"), ("   Risk 315R", "Risk_315R"), ("   Assertions", "Assertions"), ("   Fraud 240R", "Fraud_240R"), ("   GC 570R", "GoingConcern_570R"), ("3 Methodology", "L3_Methodology"), ("   Sampling", "Sampling"), ("4 Execution", "L4_Execution"),
        ("   IC Testing", "IC_Assessment"), ("5 Findings", "L5_Findings"), ("   Deficiencies", "Deficiencies_265"), ("   SUAD", "SUAD"), ("6 Oversight", "L6_Oversight"), ("AI Agents", "AI_Agents"), ("Standards", "Standards")]
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
c3 = bar_chart("Spectrum of Inherent Risk", Reference(dd, min_col=1, min_row=20, max_row=24),
               Reference(dd, min_col=2, min_row=19, max_row=24), colors=[LT], w=8, h=7, legend=False)
pts = ["81C784", "FFF176", "FFB300", "EF6C00", "C62828"]
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
     ("Signif. Deficiencies", "=DD_SigDef", "0", '="of "&CD_N&" | "&CD_SigComm&" in writing (ISA 265)"'),
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
B3 = [("Risks with IRF = 3", "=DD_HighIRF", "0", '="of "&DD_TotalRisks&" | "&DD_RiskFlags&" to justify"'),
      ("Significant SCOTABDs", "=DD_SigSCOT", "0", '=DD_MatNotSig&" material, not significant"'),
      ("Fraud Matters (240R)", "=DD_FraudN", "0", '=DD_FraudOpen&" open | "&COUNT(Fraud_TCWG)&" to TCWG"'),
      ("JE Tested (240R)", "=JE_Tested", "#,##0", '=JE_Exc&" exception(s) (FR4)"'),
      ("GC Period (months)", "=GC_Months", "0", '="from FS approval (570R)"'),
      ("Going Concern Conclusion (570R)", "=GC_Outcome", "@", '="Close call: "&GC_Close&" | events/conditions: "&GC_EventsN')]
for k, t in enumerate(B3):
    tile(56, 5 + k * 3, 3 if k < 5 else 6, *t, accent="6A1B9A")
db["T57"].font = font(9, True, DK)
B4 = [("Engagement Risk", "=AC_Rating", "@", '="Acceptance score "&TEXT(AC_ScorePct,"0%")&" (AC2)"'),
      ("Acceptance Decision", '=IF(LEFT(AC_Decision,6)="ACCEPT","Accepted","Decline")', "@", '="Approved before EL (AC7)"'),
      ("Peak Capacity Use", "=MAX(TR_Peak)", "0%", '=TR_Over&" member(s) over-allocated (TR4)"'),
      ("Competency Gaps", "=TR_Gaps", "0", '=COUNTIF(TR_NeedStatus,"Gap - mitigated")&" gap(s) mitigated (TR2)"'),
      ("PBC Received", "=PBC_PctRec", "0%", '=PBC_Overdue&" overdue | on-time "&TEXT(PBC_OnTime,"0%")'),
      ("Resources Conclusion (TR6)", "=TR_Conclusion", "@", '="Partner involvement "&TEXT(TR_PartnerPct,"0%")&" of actual hours"')]
for k, t in enumerate(B4):
    tile(61, 5 + k * 3, 3 if k < 5 else 6, *t, accent="F9A825")
db["T62"].font = font(8, True, DK)
B5 = [("Review Agent Flags", "=AI_Flags", "0", '=AI_FlagsHigh&" high | "&AI_Checks&" checks (AI3)"'),
      ("Inspection Readiness", "=AI_Ready", "0%", '="simulated inspection (AI4)"'),
      ("Assertion Risks", "=AS_Total", "0", '="SCOTABD x assertion cells (AS2)"'),
      ("Control-Gap Impact", "=AS_Impact", "#,##0", '=COUNTIF(AS_Quant,"Material")&" material | "&COUNTIF(AS_Quant,"Pervasive*")&" pervasive"'),
      ("AI Uses Logged", "=COUNTA(AI_LogStatus)", "0", '=COUNTIF(AI_LogStatus,"NON*")&" non-compliant (AI5)"'),
      ("Inspection Readiness Rating (AI4)", "=AI_Rating", "@", '="Responsible AI: approved tools, anonymised data, human verification"')]
for k, t in enumerate(B5):
    tile(66, 5 + k * 3, 3 if k < 5 else 6, *t, accent="00838F")
db["T67"].font = font(8, True, DK)
put(db, "E71", "=\"Misstatement evaluation: \"&MS_Eval", font(9, True, "B71C1C"), BG); db.merge_cells("E71:Y71")
put(db, "E72", "All figures update live from the layer working papers. Illustrative data for teaching - replace in Setup & layer sheets.", font(8, i=True, c="757575"), BG)
db.merge_cells("E72:Y72")
db.freeze_panes = "A5"
db.print_area = "A1:Y72"

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
