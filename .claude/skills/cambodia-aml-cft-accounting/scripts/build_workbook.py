"""Build the Cambodia accounting-firm AML/CFT Business Risk Assessment workbook.

Usage:
    python build_workbook.py [output.xlsx] [--firm "Firm name"]

Method adapted from the ADGM business risk assessment guide (six steps:
method -> inherent risk by factor -> control rating -> residual risk ->
separate TF/PF/sanctions component -> feed into CRA, policies, MI) and
re-based on the Cambodian AML/CFT framework for accountants and auditors.
"""
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.comments import Comment

OUT = "Cambodia_Accounting_Firm_AML_CFT_BRA.xlsx"
FIRM = "[Firm name] Co., Ltd."
args = sys.argv[1:]
if args and not args[0].startswith("--"):
    OUT = args.pop(0)
if "--firm" in args:
    FIRM = args[args.index("--firm") + 1]

# ---------------------------------------------------------------- styles
F = "Arial"
NAVY = "1F3864"
f_base = Font(name=F, size=10)
f_bold = Font(name=F, size=10, bold=True)
f_title = Font(name=F, size=16, bold=True, color=NAVY)
f_sub = Font(name=F, size=11, bold=True, color=NAVY)
f_hdr = Font(name=F, size=10, bold=True, color="FFFFFF")
f_input = Font(name=F, size=10, color="0000FF")
f_link = Font(name=F, size=10, color="008000")
f_note = Font(name=F, size=9, italic=True, color="595959")
fill_hdr = PatternFill("solid", fgColor=NAVY)
fill_input = PatternFill("solid", fgColor="FFF2CC")
fill_sec = PatternFill("solid", fgColor="D9E1F2")
fill_grey = PatternFill("solid", fgColor="F2F2F2")
thin = Side(style="thin", color="BFBFBF")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
wrap = Alignment(wrap_text=True, vertical="top")
center = Alignment(horizontal="center", vertical="center", wrap_text=True)

LEVEL_FILL = {
    "Low": "C6EFCE", "Medium": "FFEB9C", "Medium-High": "F8CBAD", "High": "FF7C80",
}

wb = Workbook()
wb.remove(wb.active)


def sheet(name, widths, tab=None):
    ws = wb.create_sheet(name)
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.sheet_view.showGridLines = False
    if tab:
        ws.sheet_properties.tabColor = tab
    return ws


def title(ws, text, sub=None):
    ws["A1"] = text
    ws["A1"].font = f_title
    if sub:
        ws["A2"] = sub
        ws["A2"].font = f_note


def header(ws, row, labels, col=1):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=lab)
        c.font, c.fill, c.alignment, c.border = f_hdr, fill_hdr, center, box
    ws.row_dimensions[row].height = 32


def put(ws, ref, value, font=None, fill=None, align=wrap, border=True, fmt=None):
    c = ws[ref]
    c.value = value
    c.font = font or f_base
    if fill:
        c.fill = fill
    c.alignment = align
    if border:
        c.border = box
    if fmt:
        c.number_format = fmt
    return c


def inp(ws, ref, value=None, fmt=None):
    return put(ws, ref, value, font=f_input, fill=fill_input, fmt=fmt)


def section(ws, ref, text, span=None):
    c = ws[ref]
    c.value = text
    c.font = f_sub
    c.fill = fill_sec
    if span:
        ws.merge_cells(f"{ref}:{span}")


def level_cf(ws, rng):
    for lvl, col in LEVEL_FILL.items():
        ws.conditional_formatting.add(
            rng, CellIsRule(operator="equal", formula=[f'"{lvl}"'],
                            fill=PatternFill("solid", fgColor=col)))


def dv_list(ws, formula, rng, prompt=None):
    dv = DataValidation(type="list", formula1=formula, allow_blank=True)
    if prompt:
        dv.promptTitle, dv.prompt, dv.showInputMessage = "Select", prompt, True
    ws.add_data_validation(dv)
    dv.add(rng)


# =================================================================== Scales
sc = sheet("Scales", {"A": 16, "B": 9, "C": 3, "D": 16, "E": 6, "F": 44, "G": 3,
                      "H": 12, "I": 6, "J": 44, "K": 3, "L": 18, "M": 11, "N": 46},
           tab="7F7F7F")
title(sc, "Rating scales, risk matrix and control rule",
      "Method: rate INHERENT risk first (ignore controls), then apply the control rule. "
      "Scales are firm practice, not a regulatory scale - adjust only with MLRO/management approval.")
section(sc, "A4", "Risk levels")
header(sc, 4, ["Level", "Score"], col=1)
for i, lvl in enumerate(["Low", "Medium", "Medium-High", "High"]):
    put(sc, f"A{5+i}", lvl)
    put(sc, f"B{5+i}", i + 1, align=center)
level_cf(sc, "A5:A8")

header(sc, 4, ["Likelihood", "Score", "Meaning (next 12 months)"], col=4)
lik = [("Unlikely", "Not expected; no known cases in the firm or sector peers."),
       ("Possible", "Could occur; isolated indicators in the client base or sector typologies."),
       ("Likely", "Expected to occur; recurring indicators, NRA/sector typologies match our clients."),
       ("Almost certain", "Already occurring or near certain given our client base and services.")]
for i, (n, d) in enumerate(lik):
    put(sc, f"D{5+i}", n)
    put(sc, f"E{5+i}", i + 1, align=center)
    put(sc, f"F{5+i}", d)

header(sc, 4, ["Impact", "Score", "Meaning"], col=8)
imp = [("Low", "Small value; no regulatory, legal or reputational consequence."),
       ("Moderate", "Some value; CAFIU/ACAR query, minor remediation, local reputational harm."),
       ("Major", "Material value; supervisory finding, fine, loss of key clients or audit licence review."),
       ("Severe", "Facilitating serious crime/TF/PF; criminal liability, licence loss, sanctions exposure.")]
for i, (n, d) in enumerate(imp):
    put(sc, f"H{5+i}", n)
    put(sc, f"I{5+i}", i + 1, align=center)
    put(sc, f"J{5+i}", d)

header(sc, 4, ["Control effectiveness", "Levels reduced", "Meaning"], col=12)
ctl = [("Effective", 1, "Designed to address the risk AND tested in operation with no/insignificant exceptions."),
       ("Partly effective", 0, "Designed well but gaps in operation, or not yet tested. No reduction."),
       ("Not effective", 0, "Missing, failed testing, or design does not address the risk.")]
for i, (n, r, d) in enumerate(ctl):
    put(sc, f"L{5+i}", n)
    put(sc, f"M{5+i}", r, align=center)
    put(sc, f"N{5+i}", d)
sc["L9"] = "Rule: controls lower risk by one level at most; residual never exceeds inherent."
sc["L9"].font = f_note

section(sc, "A11", "Inherent risk matrix (score) - rows = Likelihood, columns = Impact")
put(sc, "A12", "Likelihood \\ Impact", font=f_bold, fill=fill_grey)
for j, (n, _) in enumerate(imp):
    put(sc, f"{'BCDE'[j]}12", n, font=f_bold, fill=fill_grey, align=center)
# matrix[likelihood][impact] -> 1..4  (consistent with ADGM guide Table 7)
M = [[1, 1, 2, 2],   # Unlikely
     [1, 2, 3, 3],   # Possible
     [2, 3, 4, 4],   # Likely
     [3, 4, 4, 4]]   # Almost certain
for i, (n, _) in enumerate(lik):
    put(sc, f"A{13+i}", n, font=f_bold, fill=fill_grey)
    for j in range(4):
        put(sc, f"{'BCDE'[j]}{13+i}", M[i][j], font=f_input, align=center)

section(sc, "A18", "Same matrix shown as levels")
put(sc, "A19", "Likelihood \\ Impact", font=f_bold, fill=fill_grey)
for j, (n, _) in enumerate(imp):
    put(sc, f"{'BCDE'[j]}19", n, font=f_bold, fill=fill_grey, align=center)
for i, (n, _) in enumerate(lik):
    put(sc, f"A{20+i}", n, font=f_bold, fill=fill_grey)
    for j in range(4):
        L = "BCDE"[j]
        put(sc, f"{L}{20+i}", f"=INDEX($A$5:$A$8,{L}{13+i})", align=center)
level_cf(sc, "B20:E23")
sc.column_dimensions["C"].width = 12
sc.column_dimensions["E"].width = 12
sc.column_dimensions["B"].width = 12

section(sc, "A26", "Residual = control rule applied to inherent")
header(sc, 27, ["Inherent", "If Effective", "If Partly", "If Not effective"], col=1)
for i in range(4):
    put(sc, f"A{28+i}", f"=A{5+i}")
    put(sc, f"B{28+i}", f"=INDEX($A$5:$A$8,MAX(1,{i+1}-$M$5))")
    put(sc, f"C{28+i}", f"=INDEX($A$5:$A$8,MAX(1,{i+1}-$M$6))")
    put(sc, f"D{28+i}", f"=INDEX($A$5:$A$8,MAX(1,{i+1}-$M$7))")
level_cf(sc, "A28:D31")
sc["A33"] = ("Blue cells on this sheet are the firm's chosen scales. Changing them re-rates every "
             "register automatically - record any change in Triggers_Versions.")
sc["A33"].font = f_note

LEVELS = "Scales!$A$5:$A$8"
LIK = "Scales!$D$5:$D$8"
IMP = "Scales!$H$5:$H$8"
CTL = "Scales!$L$5:$L$7"


def inh_num(lc, ic, r):
    return (f'=IF(OR({lc}{r}="",{ic}{r}=""),"",INDEX(Scales!$B$13:$E$16,'
            f'MATCH({lc}{r},{LIK},0),MATCH({ic}{r},{IMP},0)))')


def res_num(hc, kc, r):
    return (f'=IF(OR({hc}{r}="",{kc}{r}=""),"",MAX(1,{hc}{r}-'
            f'INDEX(Scales!$M$5:$M$7,MATCH({kc}{r},{CTL},0))))')


def lvl_name(nc, r):
    return f'=IF({nc}{r}="","",INDEX({LEVELS},{nc}{r}))'


# ============================================================ Risk appetite
ra = sheet("Risk_Appetite", {"A": 26, "B": 18, "C": 70, "D": 16, "E": 12, "F": 12, "G": 12, "H": 14},
           tab="C00000")
# ============================================================ BRA register
FACTORS = ["Customers", "Geography", "Services", "Delivery channels", "Transactions",
           "New technology & products", "Fraud, cyber & data", "Internal / people"]

RISKS = [
    # id, factor, risk, Cambodia driver, lik, imp, controls, ctl
    ("C01", "Customers", "Clients in higher-risk sectors (casino/gaming, real estate & property development, construction, "
     "precious metals & stones, money changers, import/export) use the firm's services to make illicit funds look legitimate.",
     "NRA and APG typologies point to real estate, casinos and cash-based trade; Sihanoukville / border-zone growth.",
     "Likely", "Major", "Sector screening at acceptance; EDD for listed sectors; partner approval.", "Partly effective"),
    ("C02", "Customers", "Beneficial owner hidden behind nominee shareholders, layered or foreign holding structures.",
     "Foreign persons cannot own land: Khmer nominee shareholding (51/49) structures are common; MOC beneficial-ownership "
     "data still maturing.", "Likely", "Major",
     "BO verification to natural person; MOC registry check; nominee declaration; refuse if BO unverifiable.", "Partly effective"),
    ("C03", "Customers", "Domestic or foreign PEPs, their family members and close associates (incl. businesses they control).",
     "Corruption identified as a major proceeds-generating crime; politically connected business groups.",
     "Possible", "Major", "PEP screening tool; source-of-wealth check; senior management approval; annual review.", "Partly effective"),
    ("C04", "Customers", "Cash-intensive SMEs with weak or incomplete records mix illicit cash with business income.",
     "Large informal economy; high USD cash usage (dollarisation); many SMEs keep minimal records.",
     "Likely", "Moderate", "Revenue reasonableness review; bank-to-books reconciliation; escalation of unexplained cash.", "Partly effective"),
    ("C05", "Customers", "NPO / association clients misused to raise or move funds for terrorism or other crime.",
     "LANGO-registered NPOs, foreign-funded NGOs; NPO sector risk assessment flagged as pending in APG follow-up reports.",
     "Possible", "Major", "NPO questionnaire; donor and beneficiary review; TFS screening of officers.", "Partly effective"),
    ("C06", "Customers", "Clients linked to online scam compounds, cyber-fraud, human trafficking or illegal online gambling.",
     "International reporting and 2025 US/UK sanctions on Cambodia-linked conglomerates; FinCEN action on a Cambodian "
     "payments group.", "Possible", "Severe",
     "Adverse-media search; OFAC/UK/UN screening; refuse SEZ-based clients with unclear business; MLRO sign-off.", "Partly effective"),
    ("G01", "Geography", "Clients, UBOs or counterparties in FATF 'call for action' / 'increased monitoring' or sanctioned jurisdictions.",
     "Inbound investment and nominee structures from multiple jurisdictions; regional transit routes.",
     "Possible", "Major", "Country risk list updated after each FATF plenary; EDD for listed countries.", "Effective"),
    ("G02", "Geography", "Cross-border flows through border towns and SEZs (e.g. Sihanoukville, Bavet, Poipet) and neighbouring markets.",
     "Casinos and cash trade concentrated in border/SEZ areas; smuggling and trade-based ML typologies.",
     "Likely", "Major", "Location flag in CRA; trade-documentation review; enhanced monitoring.", "Partly effective"),
    ("G03", "Geography", "Proceeds of domestic predicate crimes identified in the NRA (e.g. corruption, drug trafficking, fraud, "
     "illegal logging/wildlife, tax evasion) enter the client base.",
     "Cambodia ML/TF National Risk Assessment (sanitised version published by ACAR).",
     "Possible", "Major", "NRA briefing in training; red-flag list; STR procedure.", "Partly effective"),
    ("S01", "Services", "Company formation, registered office, nominee director/secretary or corporate-maintenance services create "
     "vehicles for ML.", "Firm acts as company service provider at MOC; shelf companies; foreign investor demand.",
     "Likely", "Major", "Purpose-of-entity check; no nominee-director service without partner approval; annual BO refresh.", "Partly effective"),
    ("S02", "Services", "Handling client money, paying suppliers/taxes or operating client bank accounts on the client's behalf.",
     "Outsourced finance-function / bookkeeping engagements with payment authority.",
     "Possible", "Major", "Segregated client account; dual authorisation; no third-party payments; monthly reconciliation.", "Effective"),
    ("S03", "Services", "Tax advisory and GDT filings used to facilitate tax evasion (a predicate offence) or to legitimise income.",
     "Self-assessment regime, transfer-pricing rules, high informal activity.",
     "Possible", "Major", "Tax-engagement acceptance review; documentation for positions; refusal policy for aggressive schemes.", "Partly effective"),
    ("S04", "Services", "Statutory audit / review: audit opinion lends credibility to a business that launders funds, or ML "
     "indicators found in the audit are not reported.", "CIFRS/CISA audits for qualifying entities; limited audit-team AML awareness.",
     "Possible", "Major", "AML step in audit programme (CISA 240/250 link); escalation to MLRO; STR procedure.", "Partly effective"),
    ("S05", "Services", "Valuation, M&A, buying/selling businesses or real estate advisory used to move or integrate illicit value.",
     "Property-market activity; private share transfers between related parties.",
     "Possible", "Major", "Source-of-funds check on buyer; pricing reasonableness; partner review.", "Partly effective"),
    ("S06", "Services", "Payroll services used to pay ghost employees or move funds to third parties.",
     "Outsourced payroll for garment, construction and service clients; NSSF registration gaps.",
     "Unlikely", "Moderate", "Payroll vs NSSF/HR records; new-employee verification; variance review.", "Effective"),
    ("D01", "Delivery channels", "Non-face-to-face onboarding (Telegram, WhatsApp, email) with copies of ID only.",
     "Common business messaging channels; remote foreign investors.",
     "Likely", "Moderate", "Video-call ID verification; certified copies; e-KYC tool; extra checks for remote clients.", "Partly effective"),
    ("D02", "Delivery channels", "Reliance on introducers, agents or foreign referral firms without verifying their CDD.",
     "Referrals from regional advisory firms and investment agents.",
     "Possible", "Moderate", "Introducer due diligence; firm obtains CDD itself; introducer register.", "Partly effective"),
    ("T01", "Transactions", "Fees paid in cash, by unrelated third parties, or through e-wallets/QR (e.g. Bakong, bank apps) that "
     "obscure the payer.", "Wide use of KHQR/e-wallets and USD cash for professional fees.",
     "Possible", "Moderate", "Fee-payment policy (bank transfer from client only); cash limit; third-party payment refusal.", "Effective"),
    ("T02", "Transactions", "Unusual transactions in client records: round-tripping, unexplained shareholder loans/capital injections, "
     "related-party transfers.", "Shareholder loans from abroad common; weak related-party disclosure.",
     "Likely", "Major", "Analytical review triggers; source-of-funds request; MLRO escalation log.", "Partly effective"),
    ("N01", "New technology & products", "Clients using virtual assets/crypto to move or hide funds; firm asked to account for or "
     "advise on crypto holdings.", "Restricted and evolving VA rules in Cambodia; regional VA-based fraud typologies.",
     "Possible", "Major", "New-service risk assessment before acceptance; VA wallet tracing request; MLRO approval.", "Not effective"),
    ("N02", "New technology & products", "New tools (cloud accounting, e-KYC, AI) introduced without assessing ML, fraud and data risk.",
     "Fast adoption of cloud software and messaging apps for document exchange.",
     "Possible", "Moderate", "New-technology risk assessment (pre-launch); vendor due diligence.", "Not effective"),
    ("F01", "Fraud, cyber & data", "Client or third-party fraud, including financial-statement fraud and invoice fraud.",
     "Fraud listed as a major predicate offence; weak internal controls at SMEs.",
     "Possible", "Major", "Fraud-risk procedures in engagements; professional scepticism training.", "Partly effective"),
    ("F02", "Fraud, cyber & data", "Theft of CDD files/personal data, ransomware or business email compromise affecting the firm.",
     "Scanned IDs and passports stored on shared drives and chat apps; rising regional cyber-crime.",
     "Possible", "Major", "Access control; MFA; encrypted storage; no CDD on chat apps; backup and incident plan.", "Partly effective"),
    ("P01", "Internal / people", "Staff complicity, tipping-off, or low AML/CFT awareness leading to missed red flags and no STRs.",
     "Rapid staff turnover; AML training not yet embedded across the profession.",
     "Possible", "Moderate", "Annual role-based AML training; fit-and-proper hiring; whistle-blowing channel; MLRO independence.", "Partly effective"),
]
N_ROWS = len(RISKS) + 10  # spare rows for firm-specific risks

br = sheet("BRA_Register", {"A": 6, "B": 16, "C": 46, "D": 38, "E": 16, "F": 13, "G": 11, "H": 6, "I": 13,
                            "J": 38, "K": 15, "L": 6, "M": 13, "N": 44, "O": 13, "P": 10, "Q": 22,
                            "R": 18, "S": 14, "T": 12, "U": 6},
           tab="2F5597")
title(br, f"Business Risk Assessment register - {FIRM}",
      "Rate inherent risk BEFORE thinking about controls. Every rating needs a written reason in column N, in the firm's "
      "own words. Pre-filled rows are a STARTING POINT ONLY - replace drivers, ratings and reasons with your own data.")
H = ["ID", "Risk factor", "Risk - what could happen", "Cambodia driver / evidence", "Data source (evidence ref)",
     "Likelihood", "Impact", "Inh. score", "INHERENT", "Key controls (map to Control_Tests)", "Control effectiveness",
     "Res. score", "RESIDUAL", "Reason for ratings (firm's own words)", "Appetite (max)", "Within appetite?",
     "Decision", "Action needed?", "Owner", "Target date", "Rank"]
header(br, 4, H)
br.freeze_panes = "C5"
for i in range(N_ROWS):
    r = 5 + i
    if i < len(RISKS):
        rid, fac, risk, drv, li, im, cont, ce = RISKS[i]
        put(br, f"A{r}", rid)
        put(br, f"B{r}", fac)
        put(br, f"C{r}", risk)
        put(br, f"D{r}", drv)
        inp(br, f"E{r}", "[e.g. client list Q3, NRA p.x]")
        inp(br, f"F{r}", li)
        inp(br, f"G{r}", im)
        put(br, f"J{r}", cont)
        inp(br, f"K{r}", ce)
        inp(br, f"N{r}", "[Write why THIS firm is rated so - client numbers, % of fees, cases seen.]")
        inp(br, f"S{r}", "MLRO")
        inp(br, f"T{r}", None, fmt="dd-mmm-yyyy")
    else:
        for col in "ABCDEFGJKNST":
            inp(br, f"{col}{r}")
        br[f"T{r}"].number_format = "dd-mmm-yyyy"
    put(br, f"H{r}", inh_num("F", "G", r), align=center)
    put(br, f"I{r}", lvl_name("H", r), align=center)
    put(br, f"L{r}", res_num("H", "K", r), align=center)
    put(br, f"M{r}", lvl_name("L", r), align=center)
    put(br, f"O{r}", f'=IF(B{r}="","",IFERROR(INDEX(Risk_Appetite!$B$6:$B$13,MATCH(B{r},Risk_Appetite!$A$6:$A$13,0)),""))',
        font=f_link, align=center)
    put(br, f"P{r}", f'=IF(OR(L{r}="",O{r}=""),"",IF(L{r}<=MATCH(O{r},{LEVELS},0),"Yes","No"))', align=center)
    put(br, f"Q{r}", f'=IF(P{r}="","",IF(P{r}="Yes",IF(K{r}="Effective","Maintain","Maintain + close control gap"),'
                     f'IF(K{r}="Effective","Limit or stop activity","Strengthen controls")))')
    put(br, f"R{r}", f'=IF(L{r}="","",IF(OR(L{r}>=3,P{r}="No"),"Yes - Action_Plan","No"))')
    put(br, f"U{r}", f'=IF(L{r}="","",L{r}+H{r}/10+(1000-ROW())/1000000)', fmt="0.00")
    br.row_dimensions[r].height = 64
LAST = 4 + N_ROWS
dv_list(br, f"={LIK}", f"F5:F{LAST}", "Likelihood - inherent, ignoring controls")
dv_list(br, f"={IMP}", f"G5:G{LAST}", "Impact if the risk happens")
dv_list(br, f"={CTL}", f"K5:K{LAST}", "Based on TESTED evidence (Control_Tests), not on policy existence")
dv_list(br, '"' + ",".join(FACTORS) + '"', f"B{5+len(RISKS)}:B{LAST}")
level_cf(br, f"I5:I{LAST}")
level_cf(br, f"M5:M{LAST}")
br.conditional_formatting.add(f"P5:P{LAST}", CellIsRule(operator="equal", formula=['"No"'],
                              fill=PatternFill("solid", fgColor="FF7C80")))
br.column_dimensions["U"].hidden = True
br["E3"].value = None
br["A3"] = "Legend: blue text / yellow fill = input; black = formula; green = linked from another sheet."
br["A3"].font = f_note

# ---------------------------------------------------------- Risk appetite
title(ra, "Risk appetite statement (ML / TF / PF / sanctions)",
      "Appetite turns BRA ratings into decisions. It cannot lower a legal duty: CDD, TFS screening, STR/CTR reporting "
      "and record keeping remain mandatory whatever the appetite.")
section(ra, "A4", "1. Maximum acceptable RESIDUAL risk by factor (feeds BRA_Register column O)", "H4")
header(ra, 5, ["Risk factor", "Max residual", "Statement (testable - a junior can apply it without asking)"])
app = [("Customers", "Medium", "We accept higher-risk sector clients only with EDD and partner approval. We do not accept clients "
        "whose beneficial owner we cannot verify to a natural person."),
       ("Geography", "Medium", "We do not act for clients or UBOs resident in FATF 'call for action' countries. Increased-monitoring "
        "countries require EDD and MLRO approval."),
       ("Services", "Medium", "We do not provide nominee director/shareholder services to foreign clients without partner approval "
        "and annual BO refresh. We never pay third parties from client money."),
       ("Delivery channels", "Medium", "Remote clients must pass video-call or e-KYC verification before work starts."),
       ("Transactions", "Low", "We accept fees only by bank transfer from the client's own account; cash fees above the limit are refused."),
       ("New technology & products", "Medium", "No new service or tool goes live without a documented ML/TF, fraud and data-risk "
        "assessment approved by the MLRO."),
       ("Fraud, cyber & data", "Medium", "CDD documents are never stored on personal devices or chat apps; MFA on all systems."),
       ("Internal / people", "Low", "Every staff member completes role-based AML/CFT training each year and before client work.")]
for i, (a, b, c) in enumerate(app):
    r = 6 + i
    put(ra, f"A{r}", a)
    inp(ra, f"B{r}", b)
    inp(ra, f"C{r}", c)
    ra.row_dimensions[r].height = 42
dv_list(ra, f"={LEVELS}", "B6:B13")
level_cf(ra, "B6:B13")

section(ra, "A15", "2. Separate statements for ML, TF, PF and sanctions (expected by supervisors)", "H15")
header(ra, 16, ["Risk type", "Appetite", "Statement"])
st = [("Money laundering", "Low", "We will not knowingly act for any client where we suspect proceeds of crime; suspicion is "
       "reported to CAFIU, not managed commercially."),
      ("Terrorist financing", "Zero tolerance", "No business with any person or entity designated under UNSC resolutions or "
       "Cambodian designations; any match is frozen and reported without delay."),
      ("Proliferation financing", "Zero tolerance", "No services connected to DPRK/Iran-related designated persons, dual-use goods "
       "trade or front companies; screen all parties."),
      ("Targeted financial sanctions", "Zero tolerance", "Screen clients, UBOs, directors and payees at onboarding, on list change, "
       "and before each payment. We also screen OFAC/UK/EU lists because of USD clearing exposure.")]
for i, (a, b, c) in enumerate(st):
    r = 17 + i
    put(ra, f"A{r}", a)
    inp(ra, f"B{r}", b)
    inp(ra, f"C{r}", c)
    ra.row_dimensions[r].height = 42

section(ra, "A22", "3. Will NOT accept (exit / refuse)", "H22")
no = ["Clients whose beneficial owner cannot be verified to a natural person.",
      "Any person/entity on UN, Cambodian, OFAC or UK sanctions lists, or owned/controlled by one.",
      "Shell companies with no clear economic purpose, especially in SEZs with unclear activity.",
      "Requests to receive, hold or pay out funds unrelated to an accounting/audit service.",
      "Clients refusing to provide source-of-funds/wealth information when EDD applies."]
for i, t in enumerate(no):
    inp(ra, f"A{23+i}", f"{i+1}. {t}")
    ra.merge_cells(f"A{23+i}:C{23+i}")

section(ra, "A29", "4. Appetite measures / key risk indicators (report quarterly to management)", "H29")
header(ra, 30, ["Indicator", "Risk tracked", "How measured", "Limit", "Amber at", "Actual", "RAG", "Owner"])
kri = [("High-risk clients as % of active clients", "Customers", "CRA_Clients High / all rated (auto)", 0.15, 0.10,
        "=IFERROR(CRA_Clients!$Q$4,0)", "MLRO"),
       ("Overdue CDD / BO refreshes (count)", "Customers", "Clients past next-review date (auto)", 3, 1,
        "=CRA_Clients!$Q$5", "Client service"),
       ("Sanctions screening alerts open > 2 business days", "TFS / sanctions", "Screening log", 0, 0, None, "MLRO"),
       ("Controls rated Partly / Not effective", "Control gaps", "Control_Tests (auto)", 3, 2,
        "=Control_Tests!$J$4", "Compliance"),
       ("Mitigation actions past due date", "Follow-through", "Action_Plan (auto)", 0, 0, "=Action_Plan!$J$3", "MLRO"),
       ("Staff trained on current BRA version", "Awareness", "Training record", 1.0, 0.95, None, "HR / MLRO"),
       ("Internal suspicion reports reviewed within 5 days", "Reporting", "MLRO log", 1.0, 0.9, None, "MLRO")]
for i, (a, b, c, lim, amb, act, own) in enumerate(kri):
    r = 31 + i
    put(ra, f"A{r}", a)
    put(ra, f"B{r}", b)
    put(ra, f"C{r}", c)
    pct = isinstance(lim, float)
    fmt = "0%" if pct else "0"
    inp(ra, f"D{r}", lim, fmt=fmt)
    inp(ra, f"E{r}", amb, fmt=fmt)
    if act:
        put(ra, f"F{r}", act, font=f_link, fmt=fmt)
    else:
        inp(ra, f"F{r}", None, fmt=fmt)
    higher_bad = i not in (5, 6)
    if higher_bad:
        rag = f'=IF(F{r}="","",IF(F{r}>D{r},"Red",IF(AND(F{r}>=E{r},F{r}>0),"Amber","Green")))'
    else:
        rag = f'=IF(F{r}="","",IF(F{r}<E{r},"Red",IF(F{r}<D{r},"Amber","Green")))'
    put(ra, f"G{r}", rag, align=center)
    inp(ra, f"H{r}", own)
    ra.row_dimensions[r].height = 30
for txt, col in (("Red", "FF7C80"), ("Amber", "FFEB9C"), ("Green", "C6EFCE")):
    ra.conditional_formatting.add("G31:G37", CellIsRule(operator="equal", formula=[f'"{txt}"'],
                                  fill=PatternFill("solid", fgColor=col)))
ra["A39"] = "Decision rule: within appetite + effective = Maintain; above appetite with a fix = Strengthen; above appetite, fix not enough = Limit or stop."
ra["A39"].font = f_note

# ============================================================== TFS / TF / PF
tf = sheet("TFS_TF_PF", {"A": 7, "B": 20, "C": 50, "D": 13, "E": 11, "F": 6, "G": 13, "H": 40, "I": 15,
                         "J": 6, "K": 13, "L": 44}, tab="7030A0")
title(tf, "Separate component: Terrorist financing, Proliferation financing & Targeted financial sanctions",
      "Supervisors expect this as a clearly identifiable part of the BRA with its OWN rating - not a paragraph in the "
      "general register. Ask, for each factor: how could a designated person or a proliferation network reach us?")
H2 = ["ID", "Factor", "How a designated person / network could reach the firm", "Likelihood", "Impact", "Inh.",
      "INHERENT", "Key controls", "Control effectiveness", "Res.", "RESIDUAL", "Reason (firm's own words)"]
tf_rows = {
    "Terrorist financing (TF)": [
        ("Customers", "NPO or charity clients diverting donations; individuals linked to designated persons."),
        ("Business activity", "Company formation used to set up front entities that raise or move small TF amounts."),
        ("Geography", "Clients/UBOs/counterparties in or near conflict zones or with links to regional extremist networks."),
        ("Products / services", "Handling client money or payroll used to channel funds to third parties."),
        ("Delivery channels", "Remote onboarding of persons using forged or borrowed IDs."),
        ("Cyber", "Use of crypto or e-wallets by clients to move funds outside banking scrutiny."),
    ],
    "Proliferation financing (PF)": [
        ("Customers", "Front or shell companies acting for DPRK/Iran-linked designated persons (incl. seafarers, IT workers)."),
        ("Business activity", "Import/export, shipping, ship registration or dual-use goods trading clients."),
        ("Geography", "Trade routes and transhipment through regional ports; links to DPRK-related jurisdictions."),
        ("Products / services", "Corporate services and nominee arrangements hiding ultimate controllers."),
        ("Delivery channels", "Introducers or agents acting for unseen foreign principals."),
        ("Cyber", "DPRK-linked cyber theft and crypto laundering affecting client funds or the firm itself."),
    ],
    "Targeted financial sanctions (TFS) evasion": [
        ("Customers", "A listed person, or an entity owned/controlled by one, becomes or already is a client/UBO."),
        ("Business activity", "Late screening: lists change but existing clients are not re-screened."),
        ("Geography", "USD-denominated business exposes the firm/clients to OFAC secondary-sanctions risk."),
        ("Products / services", "Payments made for clients without screening the payee."),
        ("Delivery channels", "Name variations / Khmer-Latin transliteration cause missed matches."),
        ("Cyber", "Screening tool not updated or failing silently."),
    ],
}
r = 4
blocks = []
seq = 0
for blk, items in tf_rows.items():
    section(tf, f"A{r}", blk, f"L{r}")
    r += 1
    header(tf, r, H2)
    r += 1
    start = r
    for fac, desc in items:
        seq += 1
        prefix = {"Te": "TF", "Pr": "PF", "Ta": "TS"}[blk[:2]]
        put(tf, f"A{r}", f"{prefix}{seq:02d}")
        put(tf, f"B{r}", fac)
        put(tf, f"C{r}", desc)
        inp(tf, f"D{r}", "Possible")
        inp(tf, f"E{r}", "Severe" if blk.startswith(("Prolif", "Target")) else "Major")
        put(tf, f"F{r}", inh_num("D", "E", r), align=center)
        put(tf, f"G{r}", lvl_name("F", r), align=center)
        inp(tf, f"H{r}", "[Screening / EDD / training control]")
        inp(tf, f"I{r}", "Partly effective")
        put(tf, f"J{r}", res_num("F", "I", r), align=center)
        put(tf, f"K{r}", lvl_name("J", r), align=center)
        inp(tf, f"L{r}", "[Reason in firm's words]")
        tf.row_dimensions[r].height = 40
        r += 1
    blocks.append((blk, start, r - 1))
    seq = 0
    r += 1
dv_list(tf, f"={LIK}", f"D5:D{r}")
dv_list(tf, f"={IMP}", f"E5:E{r}")
dv_list(tf, f"={CTL}", f"I5:I{r}")
level_cf(tf, f"G5:G{r}")
level_cf(tf, f"K5:K{r}")

section(tf, f"A{r}", "Component ratings (highest residual in each block = component rating)", f"L{r}")
r += 1
header(tf, r, ["", "Component", "Overall RESIDUAL rating"], col=1)
r += 1
comp_first = r
for blk, s, e in blocks:
    put(tf, f"B{r}", blk)
    put(tf, f"C{r}", f'=IF(COUNT(J{s}:J{e})=0,"",INDEX({LEVELS},MAX(J{s}:J{e})))', align=center)
    r += 1
put(tf, f"B{r}", "TF / PF / TFS COMPONENT", font=f_bold)
put(tf, f"C{r}", f'=IF(COUNT(J5:J{r-2})=0,"",INDEX({LEVELS},MAX(J5:J{r-2})))', font=f_bold, align=center)
TF_OVERALL = f"TFS_TF_PF!$C${r}"
level_cf(tf, f"C{comp_first}:C{r}")
r += 2

section(tf, f"A{r}", "TFS control checklist (answer Yes / No / Partly; keep evidence reference)", f"L{r}")
r += 1
header(tf, r, ["#", "Control", "Requirement / good practice", "Answer", "Score", "", "Evidence ref"], col=1)
r += 1
chk_first = r
checks = [
    ("Lists screened", "UN Security Council Consolidated List and Cambodian domestic designations screened (verify current "
     "list source with CAFIU / NCC)."),
    ("Voluntary lists", "OFAC SDN, UK and EU lists also screened given USD exposure (firm policy)."),
    ("Who is screened", "Clients, beneficial owners, directors, authorised signatories and payees."),
    ("When", "At onboarding, without delay on every list update, and before payments made for clients."),
    ("Fuzzy matching", "Tool handles Khmer/Latin transliteration and name order; threshold documented."),
    ("Freeze without delay", "Written procedure to freeze funds/assets of a true match and stop the service immediately."),
    ("Report", "True/potential matches reported to CAFIU (and other competent authority as required) - verify current channel."),
    ("No tipping-off", "Staff trained not to alert the client."),
    ("Alert log", "Every alert logged with decision, reviewer and date; alerts cleared within 2 business days."),
    ("Testing", "Screening tool tested at least annually with known test names."),
    ("Training", "Staff trained on TF/PF typologies and dual-use/front-company red flags."),
]
for i, (a, b) in enumerate(checks):
    put(tf, f"A{r}", i + 1, align=center)
    put(tf, f"B{r}", a)
    put(tf, f"C{r}", b)
    inp(tf, f"D{r}", "Partly")
    put(tf, f"E{r}", f'=IF(D{r}="Yes",1,IF(D{r}="Partly",0.5,IF(D{r}="No",0,"")))', fmt="0.0", align=center)
    inp(tf, f"G{r}")
    tf.merge_cells(f"G{r}:H{r}")
    tf.row_dimensions[r].height = 32
    r += 1
dv_list(tf, '"Yes,Partly,No,N/A"', f"D{chk_first}:D{r-1}")
put(tf, f"C{r}", "TFS checklist compliance", font=f_bold)
put(tf, f"D{r}", f'=IFERROR(AVERAGE(E{chk_first}:E{r-1}),"")', font=f_bold, fmt="0%")
TF_CHECK = f"TFS_TF_PF!$D${r}"

# ================================================================ Heat map
hm = sheet("Heat_Map", {"A": 18, "B": 14, "C": 14, "D": 14, "E": 14, "F": 3, "G": 8, "H": 50, "I": 14, "J": 14,
                        "K": 26}, tab="FFC000")
title(hm, "One-page summary for senior management",
      "Heat map, top five residual risks with actions, and the TF/PF/TFS component. This is the page management reads.")
section(hm, "A4", "Inherent risk heat map - number of BRA risks in each cell", "E4")
put(hm, "A5", "Likelihood \\ Impact", font=f_bold, fill=fill_grey)
for j, (n, _) in enumerate(imp):
    put(hm, f"{'BCDE'[j]}5", n, font=f_bold, fill=fill_grey, align=center)
for i in range(4):
    li_name = lik[3 - i][0]
    rr = 6 + i
    put(hm, f"A{rr}", li_name, font=f_bold, fill=fill_grey)
    for j in range(4):
        L = "BCDE"[j]
        put(hm, f"{L}{rr}",
            f'=COUNTIFS(BRA_Register!$F$5:$F${LAST},$A{rr},BRA_Register!$G$5:$G${LAST},{L}$5)', align=center)
        lvl = ["Low", "Medium", "Medium-High", "High"][M[3 - i][j] - 1]
        hm[f"{L}{rr}"].fill = PatternFill("solid", fgColor=LEVEL_FILL[lvl])

section(hm, "A11", "Risk count by level", "E11")
header(hm, 12, ["Level", "Inherent", "Residual", "TF/PF/TFS residual"], col=1)
for i, lvl in enumerate(["High", "Medium-High", "Medium", "Low"]):
    rr = 13 + i
    put(hm, f"A{rr}", lvl)
    put(hm, f"B{rr}", f'=COUNTIF(BRA_Register!$I$5:$I${LAST},A{rr})', align=center)
    put(hm, f"C{rr}", f'=COUNTIF(BRA_Register!$M$5:$M${LAST},A{rr})', align=center)
    put(hm, f"D{rr}", f'=COUNTIF(TFS_TF_PF!$K$5:$K$60,A{rr})', align=center)
level_cf(hm, "A13:A16")
put(hm, "A18", "Risks outside appetite", font=f_bold)
put(hm, "B18", f'=COUNTIF(BRA_Register!$P$5:$P${LAST},"No")', font=f_bold, align=center)
put(hm, "A19", "TF/PF/TFS component", font=f_bold)
put(hm, "B19", f"={TF_OVERALL}", font=f_link, align=center)
put(hm, "A20", "TFS checklist score", font=f_bold)
put(hm, "B20", f"={TF_CHECK}", font=f_link, fmt="0%", align=center)
level_cf(hm, "B19")

section(hm, "G4", "Top 5 residual risks (ranked by residual, then inherent)", "K4")
header(hm, 5, ["ID", "Risk", "Residual", "Appetite", "Decision"], col=7)
for k in range(5):
    rr = 6 + k
    m = f"MATCH(LARGE(BRA_Register!$U$5:$U${LAST},{k+1}),BRA_Register!$U$5:$U${LAST},0)"
    put(hm, f"G{rr}", f"=IFERROR(INDEX(BRA_Register!$A$5:$A${LAST},{m}),\"\")", font=f_link)
    put(hm, f"H{rr}", f"=IFERROR(INDEX(BRA_Register!$C$5:$C${LAST},{m}),\"\")", font=f_link)
    put(hm, f"I{rr}", f"=IFERROR(INDEX(BRA_Register!$M$5:$M${LAST},{m}),\"\")", font=f_link, align=center)
    put(hm, f"J{rr}", f"=IFERROR(INDEX(BRA_Register!$O$5:$O${LAST},{m}),\"\")", font=f_link, align=center)
    put(hm, f"K{rr}", f"=IFERROR(INDEX(BRA_Register!$Q$5:$Q${LAST},{m}),\"\")", font=f_link)
    hm.row_dimensions[rr].height = 52
level_cf(hm, "I6:J10")
section(hm, "G12", "What changed since the last version / decision requested", "K12")
inp(hm, "G13", "[e.g. New crypto-advisory requests rated High; management asked to approve 'Limit' decision for N01.]")
hm.merge_cells("G13:K16")
section(hm, "G18", "Approval", "K18")
for i, lab in enumerate(["Approved by (senior management)", "Date", "Minute reference", "Version"]):
    put(hm, f"G{19+i}", lab, font=f_bold)
    hm.merge_cells(f"G{19+i}:H{19+i}")
    put(hm, f"I{19+i}", f"=Cover!$C${12+i}", font=f_link, fmt="dd-mmm-yyyy" if lab == "Date" else "General")
    hm.merge_cells(f"I{19+i}:K{19+i}")

# ============================================================ CRA (client)
cr = sheet("CRA_Clients", {"A": 8, "B": 24, "C": 14, "D": 12, "E": 12, "F": 12, "G": 12, "H": 12, "I": 12,
                           "J": 8, "K": 10, "L": 8, "M": 15, "N": 22, "O": 12, "P": 12, "Q": 12, "R": 30},
           tab="00B050")
title(cr, "Customer risk assessment (CRA) - one row per client",
      "The BRA rates the whole firm; the CRA rates each client. Factor weights below come from the BRA: raise a weight "
      "when that factor is High/Medium-High residual in BRA_Register.")
put(cr, "A3", "Weights", font=f_bold)
wts = [1, 1.5, 1, 1, 1, 1]
for j, w in enumerate(wts):
    inp(cr, f"{'DEFGHI'[j]}3", w, fmt="0.0")
put(cr, "K3", "Thresholds", font=f_bold, border=False)
put(cr, "M3", "Medium from", border=False)
inp(cr, "N3", 1.7, fmt="0.00")
put(cr, "M4", "High from", border=False)
inp(cr, "N4", 2.3, fmt="0.00")
put(cr, "P4", "% High", font=f_bold)
put(cr, "P5", "Overdue reviews", font=f_bold)
put(cr, "P6", "Review months: High / Med / Low", font=f_bold)
inp(cr, "Q6", 12)
inp(cr, "R6", 24)
put(cr, "R7", None, border=False)
inp(cr, "S6", 36)
cr.column_dimensions["S"].width = 8
HC = ["Client ID", "Client name", "Client type", "Customer / sector", "Ownership complexity", "Geography",
      "Service", "Delivery channel", "Payment / transactions", "PEP?", "Sanctions / adverse media hit?",
      "Score", "RATING", "CDD level", "Last review", "Next review", "Overdue?", "MLRO notes / override reason"]
header(cr, 8, HC)
cr.freeze_panes = "C9"
examples = [
    ("CL-001", "Example: Local garment exporter", "Company", "Medium", "Low", "Low", "Medium", "Low", "Low", "No", "No"),
    ("CL-002", "Example: Foreign-owned property developer (51/49 nominee)", "Company", "High", "High", "Medium", "High",
     "Medium", "Medium", "No", "No"),
    ("CL-003", "Example: NGO (LANGO registered)", "NPO", "Medium", "Low", "Medium", "Low", "Low", "Medium", "Yes", "No"),
]
CRA_N = 60
for i in range(CRA_N):
    r = 9 + i
    if i < len(examples):
        for j, v in enumerate(examples[i]):
            inp(cr, f"{'ABCDEFGHIJK'[j]}{r}", v)
        inp(cr, f"O{r}", f"=DATE(2026,{3+i*2},15)", fmt="dd-mmm-yyyy")
    else:
        for col in "ABCDEFGHIJK":
            inp(cr, f"{col}{r}")
        inp(cr, f"O{r}", None, fmt="dd-mmm-yyyy")
    sc_parts = "+".join(
        f'{"DEFGHI"[j]}$3*MATCH({"DEFGHI"[j]}{r},{{"Low","Medium","High"}},0)' for j in range(6))
    put(cr, f"L{r}", f'=IF(COUNTA(D{r}:I{r})<6,"",({sc_parts})/SUM($D$3:$I$3))', fmt="0.00", align=center)
    put(cr, f"M{r}", f'=IF(L{r}="","",IF(K{r}="Yes","Prohibited / escalate",IF(OR(J{r}="Yes",L{r}>=$N$4),"High",'
                     f'IF(L{r}>=$N$3,"Medium","Low"))))', align=center)
    put(cr, f"N{r}", f'=IF(M{r}="","",IF(M{r}="High","Enhanced CDD + senior approval",IF(M{r}="Medium","Standard CDD",'
                     f'IF(M{r}="Low","Simplified CDD (if permitted)","STOP - MLRO decision; consider STR / freeze"))))')
    put(cr, f"P{r}", f'=IF(OR(O{r}="",M{r}=""),"",IF(M{r}="Prohibited / escalate",O{r},'
                     f'EDATE(O{r},IF(M{r}="High",$Q$6,IF(M{r}="Medium",$R$6,$S$6)))))', fmt="dd-mmm-yyyy")
    put(cr, f"Q{r}", f'=IF(P{r}="","",IF(P{r}<Cover!$C$9,"Yes","No"))', align=center)
    inp(cr, f"R{r}")
CL = 8 + CRA_N
put(cr, "Q4", f'=IFERROR(COUNTIF(M9:M{CL},"High")/COUNT(L9:L{CL}),0)', fmt="0%")
put(cr, "Q5", f'=COUNTIF(Q9:Q{CL},"Yes")')
dv_list(cr, '"Low,Medium,High"', f"D9:I{CL}")
dv_list(cr, '"Yes,No"', f"J9:K{CL}")
dv_list(cr, '"Individual,Company,Partnership,NPO,Trust/foreign arrangement,Government/SOE"', f"C9:C{CL}")
level_cf(cr, f"M9:M{CL}")
cr.conditional_formatting.add(f"M9:M{CL}", CellIsRule(operator="equal", formula=['"Prohibited / escalate"'],
                              fill=PatternFill("solid", fgColor="C00000"), font=Font(name=F, color="FFFFFF", bold=True)))
cr.conditional_formatting.add(f"Q9:Q{CL}", CellIsRule(operator="equal", formula=['"Yes"'],
                              fill=PatternFill("solid", fgColor="FF7C80")))
cr["A7"] = ("Rules: sanctions/adverse-media hit overrides to Prohibited/escalate; PEP = at least High. "
            "Factor guide: Customer/sector High = casino, real estate, precious metals, money changer, SEZ trading; "
            "Ownership High = nominee/foreign layered; Geography High = FATF-listed or border/SEZ; "
            "Service High = company formation / nominee / client money; Delivery High = non-face-to-face or introducer; "
            "Payment High = cash/third-party/e-wallet from unknown payer.")
cr["A7"].font = f_note
cr["A7"].alignment = wrap
cr.merge_cells("A7:R7")
cr.row_dimensions[7].height = 40

# =========================================================== Control tests
ct = sheet("Control_Tests", {"A": 7, "B": 40, "C": 16, "D": 14, "E": 10, "F": 9, "G": 10, "H": 10, "I": 13,
                             "J": 16, "K": 40}, tab="00B0F0")
title(ct, "Control testing - design and operation",
      "A control that exists but has not been tested has not been shown to work. Reviewers ask for the test, not the policy.")
put(ct, "I4", "Partly/Not effective:", font=f_bold)
header(ct, 5, ["Ctrl ID", "Control", "Linked risk IDs", "Owner", "Design adequate?", "Sample size", "Exceptions",
               "Exception rate", "Test date", "Proposed rating", "Test notes / evidence ref"])
ctrls = [("K01", "Client acceptance & sector screening", "C01, C06"),
         ("K02", "Beneficial-owner verification to natural person (MOC check)", "C02, S01"),
         ("K03", "PEP screening and senior-management approval", "C03"),
         ("K04", "TFS screening (UN / domestic / OFAC / UK) incl. re-screening on list change", "TF/PF/TS"),
         ("K05", "Enhanced due diligence: source of funds / wealth", "C01-C06, G01"),
         ("K06", "Remote onboarding verification (video call / e-KYC)", "D01"),
         ("K07", "Fee-payment policy (no cash above limit, no third-party payers)", "T01"),
         ("K08", "Client-money segregation and dual authorisation", "S02"),
         ("K09", "Unusual-transaction escalation to MLRO and STR filing to CAFIU", "T02, S04"),
         ("K10", "Record keeping (CDD and transaction records retained per law)", "All"),
         ("K11", "Annual role-based AML/CFT training", "P01, G03"),
         ("K12", "New service / technology risk assessment before launch", "N01, N02"),
         ("K13", "Information-security controls over CDD data (MFA, access, backup)", "F02")]
for i in range(len(ctrls) + 7):
    r = 6 + i
    if i < len(ctrls):
        a, b, c = ctrls[i]
        put(ct, f"A{r}", a)
        put(ct, f"B{r}", b)
        put(ct, f"C{r}", c)
    else:
        for col in "ABC":
            inp(ct, f"{col}{r}")
    inp(ct, f"D{r}")
    inp(ct, f"E{r}")
    inp(ct, f"F{r}", fmt="0")
    inp(ct, f"G{r}", fmt="0")
    put(ct, f"H{r}", f'=IF(OR(F{r}="",F{r}=0),"",G{r}/F{r})', fmt="0.0%", align=center)
    inp(ct, f"I{r}", fmt="dd-mmm-yyyy")
    put(ct, f"J{r}", f'=IF(E{r}="","",IF(E{r}="No","Not effective",IF(H{r}="","Partly effective",'
                     f'IF(H{r}=0,"Effective",IF(H{r}<=0.1,"Partly effective","Not effective")))))')
    inp(ct, f"K{r}")
    ct.row_dimensions[r].height = 30
CTL_LAST = 5 + len(ctrls) + 7
put(ct, "J4", f'=COUNTIF(J6:J{CTL_LAST},"Partly effective")+COUNTIF(J6:J{CTL_LAST},"Not effective")', align=center)
dv_list(ct, '"Yes,No"', f"E6:E{CTL_LAST}")
for txt, col in (("Effective", "C6EFCE"), ("Partly effective", "FFEB9C"), ("Not effective", "FF7C80")):
    ct.conditional_formatting.add(f"J6:J{CTL_LAST}", CellIsRule(operator="equal", formula=[f'"{txt}"'],
                                  fill=PatternFill("solid", fgColor=col)))
ct[f"A{CTL_LAST+2}"] = ("Rule used: design 'No' = Not effective; untested = Partly effective; 0% exceptions = Effective; "
                        "<=10% = Partly effective; >10% = Not effective. Copy the result into BRA_Register column K.")
ct[f"A{CTL_LAST+2}"].font = f_note

# ============================================================== Action plan
ap = sheet("Action_Plan", {"A": 7, "B": 12, "C": 40, "D": 44, "E": 14, "F": 13, "G": 13, "H": 12, "I": 4, "J": 8},
           tab="ED7D31")
title(ap, "Action plan - every Medium-High / High residual risk and every risk outside appetite")
put(ap, "H3", "Overdue:", font=f_bold)
put(ap, "J3", None)
header(ap, 4, ["Act #", "Risk / Ctrl ID", "Issue", "Action", "Owner", "Due date", "Status", "Days overdue"])
acts = [("A01", "N01", "No process to assess crypto-related engagements.", "Draft VA engagement policy; add VA questions to CRA.",
         "MLRO", "=DATE(2027,1,31)", "Open"),
        ("A02", "C02", "BO verification partly effective - 3 of 20 files lacked MOC extract.", "Re-verify all nominee-structure "
         "clients; add MOC extract to checklist.", "Client service", "=DATE(2027,2,28)", "Open")]
for i in range(30):
    r = 5 + i
    if i < len(acts):
        for j, v in enumerate(acts[i]):
            inp(ap, f"{'ABCDEFG'[j]}{r}", v, fmt="dd-mmm-yyyy" if j == 5 else None)
    else:
        for col in "ABCDEFG":
            inp(ap, f"{col}{r}", fmt="dd-mmm-yyyy" if col == "F" else None)
    put(ap, f"H{r}", f'=IF(OR(F{r}="",G{r}="Closed"),"",MAX(0,Cover!$C$9-F{r}))', fmt="0", align=center)
put(ap, "J3", '=COUNTIF(H5:H34,">0")', align=center)
dv_list(ap, '"Open,In progress,Closed"', "G5:G34")

# ======================================================== Triggers / versions
tv = sheet("Triggers_Versions", {"A": 10, "B": 14, "C": 44, "D": 22, "E": 16, "F": 30}, tab="808080")
title(tv, "Version log and trigger register",
      "Review the BRA at least annually AND whenever a trigger occurs. Log every trigger with a decision, even 'no change'.")
section(tv, "A4", "Version log", "F4")
header(tv, 5, ["Version", "Date", "Summary of changes", "Approved by", "Minute ref", "Next scheduled review"])
inp(tv, "A6", "v1.0")
inp(tv, "B6", "=DATE(2026,11,26)", fmt="dd-mmm-yyyy")
inp(tv, "C6", "First BRA under Cambodian AML/CFT framework.")
inp(tv, "D6", "[Managing Partner]")
inp(tv, "E6", "[MM-2026-xx]")
put(tv, "F6", '=IF(B6="","",EDATE(B6,12))', fmt="dd-mmm-yyyy")
for r in range(7, 16):
    for col in "ABCDE":
        inp(tv, f"{col}{r}", fmt="dd-mmm-yyyy" if col == "B" else None)
    put(tv, f"F{r}", f'=IF(B{r}="","",EDATE(B{r},12))', fmt="dd-mmm-yyyy")
section(tv, "A18", "Trigger register", "F18")
header(tv, 19, ["Trig #", "Date", "Trigger event", "Type", "Decision", "BRA rows affected"])
trig = [("New/amended Cambodian AML/CFT law, prakas or CAFIU/ACAR guideline", "Regulatory"),
        ("Updated National Risk Assessment or sector risk assessment", "NRA"),
        ("FATF plenary list changes (Feb / Jun / Oct)", "Geography"),
        ("New UN / domestic / OFAC designations relevant to clients", "Sanctions"),
        ("New service, technology or delivery channel", "New product"),
        ("Supervisory inspection, audit or thematic review finding", "Findings"),
        ("Material change in client base (e.g. new sector, large foreign client)", "Business change")]
for i, (t, ty) in enumerate(trig):
    r = 20 + i
    put(tv, f"A{r}", f"T{i+1:02d}")
    inp(tv, f"B{r}", fmt="dd-mmm-yyyy")
    put(tv, f"C{r}", t)
    put(tv, f"D{r}", ty)
    inp(tv, f"E{r}")
    inp(tv, f"F{r}")
dv_list(tv, '"No change,Re-rate,New risk added,Escalate to management"', "E20:E40")

# ================================================================ Red flags
rf = sheet("Red_Flags", {"A": 22, "B": 70, "C": 40}, tab="FF0000")
title(rf, "Red flags for accounting & audit engagements in Cambodia",
      "If you see one: do not tip off the client, record it, and escalate to the MLRO the same day. The MLRO decides on an STR to CAFIU.")
header(rf, 4, ["Category", "Red flag", "Typical response"])
flags = [
    ("Client identity / BO", "Client reluctant to identify the beneficial owner, or Khmer shareholder appears to have no real "
     "control or funds (nominee).", "EDD; nominee declaration; refuse if BO not verified."),
    ("Client identity / BO", "Frequent changes of shareholders/directors at MOC without commercial reason.", "Ask for rationale; re-verify BO."),
    ("Client identity / BO", "ID documents inconsistent, poor-quality copies, or provided only through chat apps.", "Video verification; certified copies."),
    ("Business activity", "Business located in an SEZ/compound with unclear activity, high security, or large foreign workforce "
     "with no visible operations.", "Adverse-media and sanctions check; MLRO review."),
    ("Business activity", "Revenue far above what the premises, staff or sector would support.", "Analytical review; source-of-funds request."),
    ("Business activity", "Client operates in casino/gaming, real estate, gold or money-changing but cannot document fund sources.", "EDD; consider declining."),
    ("Transactions", "Large cash deposits in USD followed by quick transfers abroad.", "Escalate; consider STR."),
    ("Transactions", "Shareholder loans or capital injections from unrelated offshore entities.", "Source-of-funds evidence; MLRO review."),
    ("Transactions", "Round-tripping: funds leave and return as 'foreign investment'.", "Escalate; consider STR."),
    ("Transactions", "Invoices that do not match goods/services, over/under-invoicing in import-export.", "Trade-based ML check; escalate."),
    ("Transactions", "Use of multiple e-wallets or personal accounts of staff for business receipts.", "Ask for explanation; reconcile; escalate."),
    ("Transactions", "Crypto/virtual-asset holdings or payments without clear source.", "Wallet information; MLRO decision."),
    ("Fees & engagement", "Client offers to pay fees in cash, in advance, through third parties, or above normal rates.", "Refuse; follow fee policy."),
    ("Fees & engagement", "Request to use the firm's bank account to receive or pass on funds.", "Refuse - not an accounting service."),
    ("Fees & engagement", "Pressure to issue an audit opinion or tax filing quickly without supporting records.", "Partner review; consider withdrawal."),
    ("Fees & engagement", "Request for company formation with nominee directors for foreigners who will not meet the firm.", "Partner approval; EDD; usually decline."),
    ("PEP / corruption", "Client is or is linked to a PEP and funds cannot be explained by known income.", "EDD; senior approval; consider STR."),
    ("PEP / corruption", "Government contract income with unusual commissions or 'consultancy' payments.", "Escalate; corruption-risk review."),
    ("NPO / TF", "NPO funds transferred to individuals or regions with no link to its stated programmes.", "TF review; screening; escalate."),
    ("NPO / TF", "Donors or beneficiaries match or resemble designated names.", "Freeze/stop; report without delay."),
    ("PF / sanctions", "Client trades dual-use goods, ships or registers vessels, or has DPRK/Iran-linked counterparties.", "PF screening; MLRO; report."),
    ("PF / sanctions", "Complex structures used to hide the end-user of goods or the true counterparty.", "EDD; consider declining."),
    ("Tax", "Two sets of books, or unrecorded sales discovered during bookkeeping or audit.", "Escalate - tax evasion is a predicate offence."),
    ("Tax", "Client asks for fictitious invoices or VAT credits.", "Refuse; escalate; consider STR."),
    ("Cyber / fraud", "Email requests to change bank details for client payments (possible BEC).", "Call-back verification; incident log."),
]
for i, (a, b, c) in enumerate(flags):
    r = 5 + i
    put(rf, f"A{r}", a)
    put(rf, f"B{r}", b)
    put(rf, f"C{r}", c)
    rf.row_dimensions[r].height = 32

# ============================================================ Traceability
trc = sheet("Traceability", {"A": 8, "B": 30, "C": 28, "D": 28, "E": 28, "F": 24, "G": 24}, tab="00B050")
title(trc, "Traceability map - follow one High risk from assessment to action",
      "Pick a High / Medium-High residual risk and follow it through. Wherever the thread breaks is where a reviewer will stop.")
header(trc, 4, ["Risk ID", "Residual risk", "CRA effect", "Policy / procedure section", "Monitoring / alert",
                "Training module", "MI / KRI"])
trace = [("C02", "Nominee / layered ownership", "Ownership complexity weight 1.5; nominee = High",
          "AML Manual s.4.3 BO verification", "Annual BO refresh; MOC change alert", "BO & nominee module", "Overdue BO checks"),
         ("C06", "Scam / cyber-fraud linked clients", "Adverse media hit = Prohibited", "Client acceptance s.3.2",
          "Ongoing adverse-media screening", "Scam-compound typologies", "High-risk client %"),
         ("N01", "Virtual assets", "Payment factor High for VA", "New-service policy s.9", "VA questions at onboarding",
          "VA red flags", "Actions past due")]
for i in range(12):
    r = 5 + i
    if i < len(trace):
        for j, v in enumerate(trace[i]):
            inp(trc, f"{'ABCDEFG'[j]}{r}", v)
    else:
        for col in "ABCDEFG":
            inp(trc, f"{col}{r}")
    trc.row_dimensions[r].height = 32

# ===================================================== Evidence / 30-day plan
ev = sheet("Evidence_Checklist", {"A": 6, "B": 70, "C": 12, "D": 30}, tab="808080")
title(ev, "Evidence file checklist, 30-day build plan and common mistakes")
section(ev, "A3", "A. Evidence file - what a CAFIU / sector supervisor reviewer should be able to open and follow", "D3")
header(ev, 4, ["#", "Item", "Done?", "File / location"])
evid = ["Method note (scales, matrix, control rule) signed and dated",
        "Firm data used: client list by type/sector/country, services and fee split, payment methods",
        "Every rating has a written, firm-specific reason",
        "TF / PF / TFS component with its own rating (TFS_TF_PF sheet)",
        "Control test results with sample sizes and dates",
        "Heat map and top-five actions on one page",
        "Risk appetite statement approved and minuted",
        "Senior-management approval: approver, date, version on the cover",
        "Version log and trigger register maintained",
        "CRA methodology links back to BRA factors",
        "Training records: date, attendee, role, BRA version covered",
        "STR / internal-report log and MLRO decisions (kept confidential)",
        "Record-retention evidence (CDD and transaction records kept for the period required by law)",
        "Independent review / audit of the AML/CFT programme"]
for i, t in enumerate(evid):
    r = 5 + i
    put(ev, f"A{r}", i + 1, align=center)
    put(ev, f"B{r}", t)
    inp(ev, f"C{r}", "No")
    inp(ev, f"D{r}")
dv_list(ev, '"Yes,No,Partly"', f"C5:C{4+len(evid)}")
r = 6 + len(evid)
put(ev, f"B{r}", "Completion", font=f_bold)
put(ev, f"C{r}", f'=COUNTIF(C5:C{4+len(evid)},"Yes")/{len(evid)}', font=f_bold, fmt="0%")
r += 2
section(ev, f"A{r}", "B. 30-day plan for a small or medium firm starting from scratch", f"D{r}")
r += 1
header(ev, r, ["Week", "Task", "", "Owner"])
for w, t in [(1, "Agree method and scales (Scales sheet); name data owners; pull client list"),
             (2, "Collect data; list risks by factor (BRA_Register); draft reasons"),
             (3, "Test controls (Control_Tests); rate residual risk; build TF/PF/TFS component"),
             (4, "Set risk appetite; draft one-page pack (Heat_Map); challenge, approve and minute")]:
    r += 1
    put(ev, f"A{r}", w, align=center)
    put(ev, f"B{r}", t)
    inp(ev, f"D{r}", "MLRO")
r += 2
section(ev, f"A{r}", "C. Common mistakes supervisors find", f"D{r}")
for t in ["Copying a template without putting the firm's own data in it",
          "Ratings with no written reason",
          "Folding sanctions, TF and PF into a general paragraph",
          "No approval date, approver or version history",
          "Assessing once and never refreshing, even after launching something new",
          "Treating a score as the rule instead of the reasoning behind it",
          "Keeping the BRA apart from customer risk ratings and controls",
          "Rating a control 'Effective' because a policy exists, without testing it",
          "Leaving out fraud, theft of personal data and cyber-security risk"]:
    r += 1
    put(ev, f"B{r}", "- " + t, border=False)

# ==================================================================== Cover
cv = wb.create_sheet("Cover", 0)
cv.sheet_view.showGridLines = False
cv.sheet_properties.tabColor = NAVY
for col, w in {"A": 3, "B": 34, "C": 34, "D": 24, "E": 48}.items():
    cv.column_dimensions[col].width = w
cv["B2"] = "AML / CFT / CPF Business Risk Assessment Toolkit"
cv["B2"].font = f_title
cv["B3"] = "Accounting, audit & tax firms in Cambodia (reporting entities under the Law on AML/CFT, 2020)"
cv["B3"].font = f_sub
cv["B4"] = ("Method adapted from the ADGM business risk assessment guide (six steps) and FATF Guidance for a Risk-Based "
            "Approach for the Accounting Profession (2019). Re-based on the Cambodian framework.")
cv["B4"].font = f_note
cv["B4"].alignment = wrap
cv.merge_cells("B4:E4")
cv.row_dimensions[4].height = 28

section(cv, "B6", "Firm details & approval (fill in)", "E6")
fields = [("Firm name", FIRM, None), ("Registration / ACAR licence no.", "[ ]", None),
          ("Assessment date (as-of)", "=DATE(2026,11,26)", "dd-mmm-yyyy"),
          ("MLRO / compliance officer", "[Name]", None), ("Prepared by", "[Name]", None),
          ("Approved by (senior management)", "[Managing Partner]", None),
          ("Approval date", None, "dd-mmm-yyyy"), ("Minute reference", "[MM-2026-xx]", None),
          ("Version", "v1.0", None), ("Next scheduled review", '=IF(C13="","",EDATE(C13,12))', "dd-mmm-yyyy")]
for i, (lab, val, fmt) in enumerate(fields):
    r = 7 + i
    put(cv, f"B{r}", lab, font=f_bold)
    if lab == "Next scheduled review":
        put(cv, f"C{r}", val, fmt=fmt)
    else:
        inp(cv, f"C{r}", val, fmt=fmt)
# fix references: Assessment date is C9; Approval date C13 -> map used elsewhere
# Heat_Map approval block expects Cover!C12..C15 = Approved by, Approval date, Minute ref, Version
section(cv, "B18", "Headline results (live)", "E18")
heads = [("BRA risks rated", f'=COUNT(BRA_Register!$H$5:$H${LAST})', "0"),
         ("High / Medium-High residual risks", f'=Heat_Map!$C$13+Heat_Map!$C$14', "0"),
         ("Risks outside appetite", "=Heat_Map!$B$18", "0"),
         ("TF / PF / TFS component rating", f"={TF_OVERALL}", None),
         ("TFS checklist score", f"={TF_CHECK}", "0%"),
         ("High-risk clients (CRA)", "=CRA_Clients!$Q$4", "0%"),
         ("Overdue mitigation actions", "=Action_Plan!$J$3", "0")]
for i, (lab, f, fmt) in enumerate(heads):
    r = 19 + i
    put(cv, f"B{r}", lab, font=f_bold)
    put(cv, f"C{r}", f, font=f_link, fmt=fmt, align=center)
level_cf(cv, "C22")

section(cv, "B27", "How to use (six steps)", "E27")
steps = [("1. Method", "Scales", "Confirm likelihood/impact scales, matrix and control rule; record approval."),
         ("2. Inherent risk", "BRA_Register", "Rate each factor ignoring controls; write the firm-specific reason."),
         ("3. Controls", "Control_Tests", "Test design and operation on a sample; copy rating into the register."),
         ("4. Residual risk", "BRA_Register / Heat_Map", "Residual calculates automatically; compare with appetite; log actions."),
         ("5. TF / PF / TFS", "TFS_TF_PF", "Separate component with its own rating and screening checklist."),
         ("6. Feed everything else", "CRA_Clients / Traceability", "Weights in the CRA, policies, monitoring, training and MI follow the BRA."),
         ("Then", "Risk_Appetite / Action_Plan / Triggers_Versions", "Approve appetite, track actions, re-assess on triggers and at least annually.")]
header(cv, 28, ["Step", "Sheet", "What to do"], col=2)
cv.merge_cells("D28:E28")
for i, (a, b, c) in enumerate(steps):
    r = 29 + i
    put(cv, f"B{r}", a, font=f_bold)
    put(cv, f"C{r}", b)
    put(cv, f"D{r}", c)
    cv.merge_cells(f"D{r}:E{r}")
    cv.row_dimensions[r].height = 28

section(cv, "B37", "Legend", "E37")
put(cv, "B38", "Input (edit)", font=f_input, fill=fill_input)
put(cv, "C38", "Blue text, yellow fill - the only cells you should type in.")
put(cv, "B39", "Formula", font=f_base)
put(cv, "C39", "Black text - calculated, do not overwrite.")
put(cv, "B40", "Link", font=f_link)
put(cv, "C40", "Green text - pulled from another sheet.")
for i, lvl in enumerate(["Low", "Medium", "Medium-High", "High"]):
    put(cv, f"{'BCDE'[i]}41", lvl, fill=PatternFill("solid", fgColor=LEVEL_FILL[lvl]), align=center)

section(cv, "B43", "Cambodian legal & regulatory basis (verify against official Khmer texts before relying on it)", "E43")
header(cv, 44, ["Source", "Relevance to accounting firms", "Status", "Note"], col=2)
basis = [
    ("Law on Anti-Money Laundering and Combating the Financing of Terrorism (promulgated June 2020; replaced the 2007 law)",
     "Accountants and auditors are reporting entities: CDD/EDD, BO, record keeping, internal controls, STR/CTR to CAFIU.",
     "Binding", "Confirm article numbers, STR timing and cash-reporting threshold in the official text."),
    ("Law on Counter Financing of the Proliferation of Weapons of Mass Destruction (2020)",
     "Targeted financial sanctions for PF; freeze and report obligations.", "Binding", "Verify title/date and implementing rules."),
    ("Cambodia Financial Intelligence Unit (CAFIU), under the National Bank of Cambodia",
     "Receives STRs/CTRs; issues guidelines (e.g. STR guideline); AML supervision on a risk basis.", "Authority",
     "Check cafiu.gov.kh for current guidelines and reporting channel."),
    ("CAFIU Prakas on AML/CFT for reporting entities not regulated by the NBC (2010)",
     "DNFBP-specific CDD, reporting and compliance-officer requirements.", "Check if still in force",
     "May be superseded/updated after the 2020 law."),
    ("Accounting and Auditing Regulator (ACAR), Ministry of Economy and Finance; KICPAA",
     "Licensing/oversight of the profession; published the sanitised ML/TF National Risk Assessment.", "Sector body",
     "Confirm current AML supervisory role for accountants/auditors."),
    ("National Risk Assessment (2016, sanitised) and later updates; National Coordination Committee on AML/CFT",
     "Firm BRA must take NRA results into account.", "Input", "Use the latest NRA / sector assessment available."),
    ("FATF Recommendations 1, 10, 22, 23 and FATF RBA Guidance for the Accounting Profession (June 2019)",
     "International standard for DNFBP risk assessment and CDD.", "Standard", "Cambodia removed from FATF grey list Feb 2023."),
    ("APG Mutual Evaluation of Cambodia (2017) and follow-up reports (2018-2023)",
     "Sector weaknesses and supervisory expectations.", "Reference", "Next-round evaluation findings will supersede."),
]
for i, (a, b, c, d) in enumerate(basis):
    r = 45 + i
    put(cv, f"B{r}", a)
    put(cv, f"C{r}", b)
    put(cv, f"D{r}", c)
    put(cv, f"E{r}", d)
    cv.row_dimensions[r].height = 54
r = 45 + len(basis) + 1
cv[f"B{r}"] = ("Disclaimer: teaching and practice tool. Pre-filled risks, ratings and statements are illustrative starting points "
               "and must be replaced with the firm's own data and reasons. It is not legal advice.")
cv[f"B{r}"].font = f_note
cv[f"B{r}"].alignment = wrap
cv.merge_cells(f"B{r}:E{r}")
cv.row_dimensions[r].height = 30

# Cover field rows: 7 Firm, 8 Reg, 9 Assessment date, 10 MLRO, 11 Prepared, 12 Approved by,
# 13 Approval date, 14 Minute, 15 Version, 16 Next review  -> Heat_Map uses C12..C15. Good.

# sheet order
order = ["Cover", "Scales", "BRA_Register", "TFS_TF_PF", "Risk_Appetite", "Heat_Map", "Control_Tests",
         "CRA_Clients", "Action_Plan", "Traceability", "Red_Flags", "Triggers_Versions", "Evidence_Checklist"]
wb._sheets = [wb[n] for n in order]
for ws in wb.worksheets:
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
wb.active = 0
wb.save(OUT)
print("saved", OUT)
