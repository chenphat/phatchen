#!/usr/bin/env python3
"""Build the companion Excel workbook (registers, calculators, checklists).

Usage: python build_xlsx.py <content.json> <output.xlsx>
Blue font = input cell; black = formula. All calculations are live Excel formulas.
"""
import json
import sys

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

NAVY, TEAL = "1F3A5F", "1A7F86"
HDR = PatternFill("solid", fgColor=NAVY)
INPUT_FILL = PatternFill("solid", fgColor="FFF9E5")
ALT = PatternFill("solid", fgColor="EEF3F8")
RED = PatternFill("solid", fgColor="F4CCCC")
AMBER = PatternFill("solid", fgColor="FFF2CC")
GREEN = PatternFill("solid", fgColor="D9EAD3")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")
BLUE = Font(color="0000FF")
BOLD = Font(bold=True)


def title(ws, text, sub=None):
    ws["A1"] = text
    ws["A1"].font = Font(bold=True, size=16, color=NAVY)
    if sub:
        ws["A2"] = sub
        ws["A2"].font = Font(italic=True, color="595959")


def header(ws, row, headers, widths=None):
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = HDR
        c.alignment = Alignment(wrap_text=True, vertical="center")
        c.border = BOX
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w


def rows(ws, start, data, input_cols=()):
    for r, rec in enumerate(data):
        for c, v in enumerate(rec, 1):
            cell = ws.cell(row=start + r, column=c, value=v)
            cell.alignment = WRAP
            cell.border = BOX
            if c in input_cols:
                cell.fill = INPUT_FILL
                cell.font = BLUE
            elif r % 2:
                cell.fill = ALT
    return start + len(data)


def inp(ws, ref, value, fmt=None):
    ws[ref] = value
    ws[ref].font = BLUE
    ws[ref].fill = INPUT_FILL
    ws[ref].border = BOX
    if fmt:
        ws[ref].number_format = fmt


def calc(ws, ref, formula, fmt=None, bold=False):
    ws[ref] = formula
    ws[ref].border = BOX
    if bold:
        ws[ref].font = BOLD
    if fmt:
        ws[ref].number_format = fmt


def dv_list(ws, rng, items):
    dv = DataValidation(type="list", formula1='"' + ",".join(items) + '"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(rng)


def build(content, out):
    m, a, pols = content["meta"], content["analysis"], content["policies"]
    wb = Workbook()

    # 1 Cover
    ws = wb.active
    ws.title = "Cover"
    title(ws, m["title"] + " - Companion Workbook", m["subtitle"])
    info = [["Organisation", m["org"]], ["Document ref", m["doc_ref"]], ["Version", m["version"]],
            ["Effective date", m["effective_date"]], ["Owner", m["owner"]], ["Approved by", m["approved_by"]],
            ["Source", m["source"]], ["Engagement / client", ""], ["Year end", ""], ["Prepared by / date", ""], ["Reviewed by / date", ""]]
    header(ws, 4, ["Field", "Detail"], [28, 90])
    end = rows(ws, 5, info, input_cols=())
    for r in range(12, end):
        ws.cell(row=r, column=2).fill = INPUT_FILL
        ws.cell(row=r, column=2).font = BLUE
    sheets = [
        ["Policy_Register", "All 16 policies with standards, purpose and statement counts"],
        ["Procedures", "Every procedure step: responsibility, WP ref, timing, plus status tracking"],
        ["RACI", "Role involvement by policy (count of procedure steps)"],
        ["WP_Index", "Working paper reference structure (Ac / C-O / Ab / Aa / A-B)"],
        ["Acceptance", "Three-test acceptance and six-trigger continuance checklist (POL-02)"],
        ["Materiality", "Overall, performance and clearly trivial materiality calculator (POL-07)"],
        ["Sample_Size", "Transaction table lookup, substantive formula and controls sizes (POL-08)"],
        ["Risk_Register", "What/where/why risk register with separate IR and CR (POL-05)"],
        ["Misstatements", "ISA 450 summary of misstatements vs materiality (POL-12)"],
        ["Gap_Analysis", "Review findings on the source APM with action tracking"],
        ["Implementation", "Roadmap and Dec 2026 standards tracker (POL-16)"]]
    header(ws, end + 1, ["Sheet", "Purpose"])
    rows(ws, end + 2, sheets)
    ws.cell(row=end + 2 + len(sheets) + 1, column=1, value="Key: blue text on cream = input; black = formula.").font = Font(italic=True)

    # 2 Policy register
    ws = wb.create_sheet("Policy_Register")
    title(ws, "Policy Register")
    header(ws, 3, ["Policy", "Title", "Phase", "Standards", "Purpose", "# Statements", "# Procedures", "Policy owner", "Last reviewed"], [9, 34, 16, 34, 60, 12, 12, 18, 14])
    end = rows(ws, 4, [[p["id"], p["title"], p["phase"], p["isa"], p["purpose"], len(p["statements"]), len(p["procedures"]), "", ""] for p in pols], input_cols=(8, 9))
    ws.freeze_panes = "C4"
    ws.auto_filter.ref = f"A3:I{end - 1}"

    # 3 Procedures
    ws = wb.create_sheet("Procedures")
    title(ws, "Procedures Register and Engagement Tracker", "Use per engagement: set status, completed by and date for each step.")
    header(ws, 3, ["Policy", "Step", "Procedure", "Resp.", "WP ref", "Timing", "Status", "Completed by", "Date", "Comments"], [9, 6, 70, 10, 13, 18, 14, 14, 12, 30])
    data = [[p["id"], i, *s, "Not started", "", "", ""] for p in pols for i, s in enumerate(p["procedures"], 1)]
    end = rows(ws, 4, data, input_cols=(7, 8, 9, 10))
    dv_list(ws, f"G4:G{end - 1}", ["Not started", "In progress", "Complete", "N/A - justified"])
    ws.conditional_formatting.add(f"G4:G{end - 1}", CellIsRule(operator="equal", formula=['"Complete"'], fill=GREEN))
    ws.conditional_formatting.add(f"G4:G{end - 1}", CellIsRule(operator="equal", formula=['"In progress"'], fill=AMBER))
    ws.conditional_formatting.add(f"G4:G{end - 1}", CellIsRule(operator="equal", formula=['"Not started"'], fill=RED))
    ws.freeze_panes = "D4"
    ws.auto_filter.ref = f"A3:J{end - 1}"
    last = end - 1
    ws.cell(row=end + 1, column=3, value="Completion %").font = BOLD
    calc(ws, f"D{end + 1}", f'=IFERROR((COUNTIF(G4:G{last},"Complete")+COUNTIF(G4:G{last},"N/A - justified"))/COUNTA(C4:C{last}),0)', "0%", True)
    proc_last = last

    # 4 RACI
    ws = wb.create_sheet("RACI")
    title(ws, "Role Involvement Matrix", "Number of procedure steps each role is responsible for, by policy (formula-driven from Procedures).")
    codes = [r[0] for r in content["roles"]]
    header(ws, 3, ["Policy", "Title"] + codes, [9, 40] + [8] * len(codes))
    for r, p in enumerate(pols, 4):
        ws.cell(row=r, column=1, value=p["id"]).border = BOX
        ws.cell(row=r, column=2, value=p["title"]).border = BOX
        for c, code in enumerate(codes, 3):
            col = get_column_letter(c)
            calc(ws, f"{col}{r}", f'=COUNTIFS(Procedures!$A$4:$A${proc_last},$A{r},Procedures!$D$4:$D${proc_last},"*"&{col}$3&"*")')
    rlast = 3 + len(pols)
    ws.conditional_formatting.add(f"C4:{get_column_letter(2 + len(codes))}{rlast}", CellIsRule(operator="greaterThan", formula=["0"], fill=GREEN))
    ws.cell(row=rlast + 2, column=1, value="Roles").font = BOLD
    rows(ws, rlast + 3, content["roles"])

    # 5 WP index
    ws = wb.create_sheet("WP_Index")
    title(ws, "Working Paper Index")
    header(ws, 3, ["Ref", "Working paper", "Phase", "Standard", "Prepared by", "Reviewed by", "Status"], [10, 60, 20, 22, 14, 14, 14])
    end = rows(ws, 4, [w + ["", "", ""] for w in content["working_papers"]], input_cols=(5, 6, 7))
    dv_list(ws, f"G4:G{end - 1}", ["Not started", "Prepared", "Reviewed", "N/A"])
    ws.freeze_panes = "B4"

    # 6 Acceptance
    ws = wb.create_sheet("Acceptance")
    title(ws, "Client Acceptance & Continuance Checklist (POL-02)", "All three tests must pass. A pass on two and a fail on one is a fail.")
    header(ws, 4, ["Test", "Question", "Answer (Yes/No)", "Evidence / WP ref"], [16, 70, 16, 40])
    qs = [
        ["Integrity", "Background research on owners and directors is clear (media, legal, regulatory)"],
        ["Integrity", "Previous auditor replied with no professional reason not to accept"],
        ["Integrity", "Ownership is clear and not hidden"],
        ["Integrity", "Accounting approach is not aggressive; no known fraud or regulatory issues"],
        ["Capability", "Team includes named person(s) with industry knowledge"],
        ["Capability", "Staff available and hours sufficient before the deadline"],
        ["Capability", "Fee supports the hours required (see time budget Ac17)"],
        ["Capability", "Engagement is within the manual's scope (not group / publicly traded / bank / insurer / significant IT)"],
        ["Independence", "No financial interest or family relationship threat, or safeguard removes it"],
        ["Independence", "Non-audit services assessed (Ac3) and permitted"],
        ["Independence", "No fee dependence threat, or safeguard removes it"]]
    end = rows(ws, 5, [q + ["", ""] for q in qs], input_cols=(3, 4))
    dv_list(ws, f"C5:C{end - 1}", ["Yes", "No"])
    ws.conditional_formatting.add(f"C5:C{end - 1}", CellIsRule(operator="equal", formula=['"No"'], fill=RED))
    ws.conditional_formatting.add(f"C5:C{end - 1}", CellIsRule(operator="equal", formula=['"Yes"'], fill=GREEN))
    r0 = end + 1
    ws.cell(row=r0, column=1, value="Result").font = BOLD
    for i, t in enumerate(["Integrity", "Capability", "Independence"]):
        ws.cell(row=r0 + 1 + i, column=1, value=t)
        calc(ws, f"B{r0 + 1 + i}", f'=IF(COUNTIFS($A$5:$A${end - 1},"{t}",$C$5:$C${end - 1},"")>0,"Incomplete",IF(COUNTIFS($A$5:$A${end - 1},"{t}",$C$5:$C${end - 1},"No")>0,"FAIL","PASS"))', bold=True)
    ws.cell(row=r0 + 4, column=1, value="DECISION").font = BOLD
    calc(ws, f"B{r0 + 4}", f'=IF(COUNTIF(B{r0 + 1}:B{r0 + 3},"Incomplete")>0,"Complete all questions",IF(COUNTIF(B{r0 + 1}:B{r0 + 3},"FAIL")>0,"DECLINE or condition - escalate to EP","ACCEPT - EP to sign Ac2 before fieldwork"))', bold=True)
    for rr in (f"B{r0 + 1}:B{r0 + 4}",):
        ws.conditional_formatting.add(rr, FormulaRule(formula=[f'ISNUMBER(SEARCH("FAIL",B{r0 + 1}))'], fill=RED))
        ws.conditional_formatting.add(rr, FormulaRule(formula=[f'ISNUMBER(SEARCH("DECLINE",B{r0 + 1}))'], fill=RED))
        ws.conditional_formatting.add(rr, FormulaRule(formula=[f'OR(B{r0 + 1}="PASS",LEFT(B{r0 + 1},6)="ACCEPT")'], fill=GREEN))
    t0 = r0 + 6
    ws.cell(row=t0, column=1, value="Continuance: six triggers (any 'Yes' = fresh decision, not a rollforward)").font = Font(bold=True, color=TEAL)
    header(ws, t0 + 1, ["Trigger", "Why it matters", "Changed? (Yes/No)", "Explanation / evidence"])
    tend = rows(ws, t0 + 2, [t + ["", ""] for t in content["continuance_triggers"]], input_cols=(3, 4))
    dv_list(ws, f"C{t0 + 2}:C{tend - 1}", ["Yes", "No"])
    ws.cell(row=tend + 1, column=1, value="Continuance basis").font = BOLD
    calc(ws, f"B{tend + 1}", f'=IF(COUNTIF(C{t0 + 2}:C{tend - 1},"Yes")>0,"FRESH acceptance decision required",IF(COUNTBLANK(C{t0 + 2}:C{tend - 1})>0,"Complete all triggers","Rollforward permitted - evidence that nothing changed"))', bold=True)

    # 7 Materiality
    ws = wb.create_sheet("Materiality")
    title(ws, "Materiality Calculator (POL-07, Ac13)", "Choose the benchmark, then justify the benchmark AND the percentage.")
    ws.column_dimensions["A"].width = 34
    for col in "BCDE":
        ws.column_dimensions[col].width = 18
    ws.column_dimensions["F"].width = 50
    header(ws, 4, ["Benchmark", "Amount (input)", "Adjustments (unusual items)", "Adjusted amount", "Low % / High %", "Best used for"])
    benches = content["materiality_benchmarks"]
    for i, b in enumerate(benches):
        r = 5 + i
        ws.cell(row=r, column=1, value=b[0]).border = BOX
        inp(ws, f"B{r}", None, "#,##0")
        inp(ws, f"C{r}", None, "#,##0")
        calc(ws, f"D{r}", f"=B{r}+C{r}", "#,##0")
        ws.cell(row=r, column=5, value=f"{b[1]*100:g}% - {b[2]*100:g}%").border = BOX
        ws.cell(row=r, column=6, value=b[3]).border = BOX
        ws.cell(row=r, column=6).alignment = WRAP
        ws.cell(row=r, column=8, value=b[1])
        ws.cell(row=r, column=9, value=b[2])
    ws["H4"], ws["I4"] = "Low", "High"
    lastb = 4 + len(benches)
    labels = [("Selected benchmark", "A12"), ("Benchmark amount", "A13"), ("Percentage applied", "A14"),
              ("Guide range", "A15"), ("OVERALL MATERIALITY", "A16"), ("Performance materiality %", "A17"),
              ("PERFORMANCE MATERIALITY", "A18"), ("Clearly trivial % of overall", "A19"), ("CLEARLY TRIVIAL THRESHOLD", "A20"),
              ("Justification of benchmark", "A22"), ("Justification of percentage", "A23"),
              ("Completion: reassessed materiality (actual results)", "A25"), ("Change vs planning", "A26"), ("Action", "A27")]
    for text, ref in labels:
        ws[ref] = text
        ws[ref].font = BOLD
    inp(ws, "B12", benches[0][0])
    dv_list(ws, "B12", [b[0] for b in benches])
    calc(ws, "B13", f"=INDEX($D$5:$D${lastb},MATCH(B12,$A$5:$A${lastb},0))", "#,##0")
    inp(ws, "B14", 0.05, "0.00%")
    calc(ws, "B15", f'=TEXT(INDEX($H$5:$H${lastb},MATCH(B12,$A$5:$A${lastb},0)),"0.0%")&" - "&TEXT(INDEX($I$5:$I${lastb},MATCH(B12,$A$5:$A${lastb},0)),"0.0%")')
    calc(ws, "C14", f'=IF(OR(B14<INDEX($H$5:$H${lastb},MATCH(B12,$A$5:$A${lastb},0)),B14>INDEX($I$5:$I${lastb},MATCH(B12,$A$5:$A${lastb},0))),"Outside guide range - justify","Within guide range")')
    calc(ws, "B16", "=ROUND(B13*B14,0)", "#,##0", True)
    inp(ws, "B17", 0.75, "0%")
    calc(ws, "B18", "=ROUND(B16*B17,0)", "#,##0", True)
    inp(ws, "B19", 0.05, "0%")
    calc(ws, "B20", "=ROUND(B16*B19,0)", "#,##0", True)
    ws.merge_cells("B22:F22")
    ws.merge_cells("B23:F23")
    inp(ws, "B22", "")
    inp(ws, "B23", "")
    inp(ws, "B25", 0, "#,##0")
    calc(ws, "B26", '=IF(B25=0,"",(B25-B16)/B16)', "0.0%")
    calc(ws, "B27", '=IF(B25=0,"Enter reassessed amount at completion",IF(B25<B16,"Materiality FELL - reduce PM and reconsider sufficiency of work done","No reduction - document reassessment"))')
    for ref in ("B16", "B18", "B20"):
        ws[ref].fill = GREEN
    ws.conditional_formatting.add("C14", FormulaRule(formula=['LEFT(C14,7)="Outside"'], fill=AMBER))
    pbt_warn = 'AND(B12="Profit before tax",ISNUMBER($B$5),B13<=0)'
    ws.conditional_formatting.add("B13", FormulaRule(formula=[pbt_warn], fill=RED))
    ws["C13"] = f'=IF({pbt_warn},"PBT negative/nil - choose another benchmark","")'
    ws["C13"].font = Font(color="C00000", bold=True)

    # 8 Sample size
    ws = wb.create_sheet("Sample_Size")
    title(ws, "Sample Size Calculator (POL-08, Ac15)", "Choose the approach first: I / P / % need no sample. Only T and S produce a sample.")
    ws.column_dimensions["A"].width = 44
    for col in "BCDEF":
        ws.column_dimensions[col].width = 16
    ws["A4"] = "A. TRANSACTION TESTING - minimum sizes"
    ws["A4"].font = Font(bold=True, color=TEAL)
    header(ws, 5, ["Population band", "From", "To", "High", "Medium", "Low"])
    ts = content["transaction_samples"]
    tsorted = sorted(ts, key=lambda t: t[1])
    rows(ws, 6, [[t[0], t[1], t[2], t[3], t[4], t[5]] for t in tsorted])
    tl = 5 + len(ts)
    ws["A13"], ws["A14"], ws["A15"] = "Population size (number of transactions)", "Assessed risk", "MINIMUM SAMPLE SIZE"
    for r in ("A13", "A14", "A15"):
        ws[r].font = BOLD
    inp(ws, "B13", 1200, "#,##0")
    inp(ws, "B14", "Medium")
    dv_list(ws, "B14", ["High", "Medium", "Low"])
    calc(ws, "B15", f'=IF(B13<1,0,INDEX($D$6:$F${tl},MATCH(B13,$B$6:$B${tl},1),MATCH(B14,$D$5:$F$5,0)))', "0", True)
    ws["B15"].fill = GREEN
    ws["C15"] = "High/Medium sizes are floors (+), not targets."
    ex = content["substantive_example"]
    ws["A18"] = "B. SUBSTANTIVE TESTING - residual population method"
    ws["A18"].font = Font(bold=True, color=TEAL)
    items = [("Population per listing", ex["population"]), ("Overall materiality", ex["overall_materiality"]),
             ("Performance materiality", ex["performance_materiality"]), ("Less: items above PM (tested 100%)", ex["large_items"]),
             ("Less: key items e.g. related party, credits (tested 100%)", ex["key_items"]), ("Risk factor (firm table)", ex["risk_factor"])]
    for i, (lab, val) in enumerate(items):
        ws[f"A{19 + i}"] = lab
        inp(ws, f"B{19 + i}", val, "0.0" if i == 5 else "#,##0")
    ws["A25"], ws["A26"], ws["A27"] = "Residual population", "Unrounded size", "RESIDUAL SAMPLE SIZE"
    for r in ("A25", "A26", "A27"):
        ws[r].font = BOLD
    calc(ws, "B25", "=B19-B22-B23", "#,##0")
    calc(ws, "B26", "=IFERROR(B25/B20*B24,0)", "0.00")
    calc(ws, "B27", "=ROUNDUP(B26,0)", "0", True)
    ws["B27"].fill = GREEN
    ws["C24"] = "APM shows 1.2 for low risk; adopt firm's full table for medium/high (Gap 8)."
    ws["C27"] = f"Worked example from the APM gives {ex['expected_sample']}."
    ws["A30"] = "C. TESTS OF CONTROLS - only where reliance planned (ISA 330.8)"
    ws["A30"].font = Font(bold=True, color=TEAL)
    header(ws, 31, ["Control frequency", "Sample", "Note"])
    rows(ws, 32, content["controls_samples"])
    ws.column_dimensions["C"].width = 40

    # 9 Risk register
    ws = wb.create_sheet("Risk_Register")
    title(ws, "Risk Register (POL-05, Ac10/Ac11)", "A heading is not a risk: name the WHAT, the WHERE and the WHY. Assess IR and CR separately (1 = Low, 2 = Moderate, 3 = High).")
    hdrs = ["Risk ID", "WHERE: class / balance / disclosure", "WHAT: assertion", "WHY: inherent risk factor", "Risk description (client-specific)",
            "IR (1-3)", "Controls tested?", "CR (1-3)", "RMM score", "RMM rating", "Significant risk?", "Response / procedure", "WP ref", "Conclusion"]
    header(ws, 4, hdrs, [8, 22, 16, 18, 50, 8, 10, 8, 9, 11, 11, 40, 9, 30])
    examples = [
        ["R01", "Revenue - export sales", "Occurrence; Cut-off", "Bias or fraud", "Export revenue may be recognised before control transfers, because Incoterms vary by customer and the sales team is bonused on shipped volume.", 3, "No", "", "", "", "Yes", "Test cut-off against Incoterms per contract for shipments either side of year end", "O", ""],
        ["R02", "Journal entries - all", "All assertions", "Bias or fraud", "Management override of controls through manual period-end journals (presumed significant risk, ISA 240).", 3, "No", "", "", "", "Yes", "Criteria-based journal testing; estimates bias review; unpredictable procedure", "O / Ab", ""]]
    blanks = [[f"R{n:02d}", "", "", "", "", "", "No", "", "", "", "", "", "", ""] for n in range(3, 26)]
    end = rows(ws, 5, examples + blanks, input_cols=(2, 3, 4, 5, 6, 7, 11, 12, 13, 14))
    for r in range(5, end):
        inp(ws, f"H{r}", 3)  # CR defaults to HIGH; lower only if controls tested and effective
        calc(ws, f"I{r}", f'=IF(OR(F{r}="",NOT(ISNUMBER(H{r}))),"",F{r}*H{r})')
        calc(ws, f"J{r}", f'=IF(I{r}="","",IF(I{r}>=6,"High",IF(I{r}>=3,"Moderate","Low")))')
    dv_list(ws, f"D5:D{end - 1}", [f[0] for f in content["inherent_risk_factors"]])
    dv_list(ws, f"F5:F{end - 1}", ["1", "2", "3"])
    dv_list(ws, f"H5:H{end - 1}", ["1", "2", "3"])
    # Flag CR below high without effective controls testing
    ws.conditional_formatting.add(f"H5:H{end - 1}", FormulaRule(formula=['AND(H5<3,G5<>"Yes - effective")'], fill=RED))
    dv_list(ws, f"G5:G{end - 1}", ["No", "Yes - effective", "Yes - not effective"])
    dv_list(ws, f"K5:K{end - 1}", ["Yes", "No"])
    ws.conditional_formatting.add(f"J5:J{end - 1}", CellIsRule(operator="equal", formula=['"High"'], fill=RED))
    ws.conditional_formatting.add(f"J5:J{end - 1}", CellIsRule(operator="equal", formula=['"Moderate"'], fill=AMBER))
    ws.conditional_formatting.add(f"J5:J{end - 1}", CellIsRule(operator="equal", formula=['"Low"'], fill=GREEN))
    ws.conditional_formatting.add(f"E5:E{end - 1}", FormulaRule(formula=['AND(B5<>"",E5="")'], fill=RED))
    ws.freeze_panes = "C5"
    ws.cell(row=end + 1, column=1, value="Note: if controls are not tested, CR is HIGH (3) by definition and RMM = IR. Overwrite CR only where controls are tested and effective.").font = Font(italic=True)

    # 10 Misstatements
    ws = wb.create_sheet("Misstatements")
    title(ws, "Summary of Misstatements (POL-12, ISA 450)", "Record items above the clearly trivial threshold. Positive = increases profit / assets.")
    header(ws, 4, ["#", "Description", "WP ref", "Type", "Profit effect", "Net assets effect", "Corrected?", "Material by nature? (trigger)", "Comment"], [5, 44, 9, 14, 15, 15, 12, 30, 30])
    for r in range(5, 25):
        ws.cell(row=r, column=1, value=r - 4).border = BOX
        for c in range(2, 10):
            cell = ws.cell(row=r, column=c)
            cell.fill, cell.font, cell.border = INPUT_FILL, BLUE, BOX
        ws.cell(row=r, column=5).number_format = ws.cell(row=r, column=6).number_format = "#,##0;(#,##0)"
    dv_list(ws, "D5:D24", ["Factual", "Judgemental", "Projected"])
    dv_list(ws, "G5:G24", ["Yes", "No"])
    dv_list(ws, "H5:H24", ["None"] + [t[0] for t in content["material_by_nature"]])
    lab = [("Total uncorrected - profit effect", '=SUMIFS(E5:E24,G5:G24,"No")'),
           ("Total uncorrected - net assets effect", '=SUMIFS(F5:F24,G5:G24,"No")'),
           ("Overall materiality (from Materiality sheet)", "=Materiality!B16"),
           ("Uncorrected as % of materiality", '=IFERROR(MAX(ABS(B26),ABS(B27))/B28,0)'),
           ("Items material by nature (uncorrected)", '=COUNTIFS(G5:G24,"No",H5:H24,"<>None",H5:H24,"<>")'),
           ("CONCLUSION", '=IF(OR(B29>=1,B30>0),"MATERIAL - adjustment required or opinion modified",IF(B29>=0.5,"Approaching materiality - consider further work","Not material individually or in aggregate - document qualitative assessment"))')]
    ws.column_dimensions["A"].width = 44
    ws.column_dimensions["B"].width = 44
    fmts = ["#,##0;(#,##0)", "#,##0;(#,##0)", "#,##0", "0%", "0", None]
    for i, (t, f) in enumerate(lab):
        ws[f"A{26 + i}"] = t
        ws[f"A{26 + i}"].font = BOLD
        calc(ws, f"B{26 + i}", f, fmts[i], True)
    ws.conditional_formatting.add("B31", FormulaRule(formula=['LEFT(B31,8)="MATERIAL"'], fill=RED))
    ws.conditional_formatting.add("B31", FormulaRule(formula=['LEFT(B31,5)="Appro"'], fill=AMBER))
    ws.conditional_formatting.add("B31", FormulaRule(formula=['LEFT(B31,3)="Not"'], fill=GREEN))
    ws["A33"] = "Six material-by-nature triggers"
    ws["A33"].font = Font(bold=True, color=TEAL)
    rows(ws, 34, content["material_by_nature"])

    # 11 Gap analysis
    ws = wb.create_sheet("Gap_Analysis")
    title(ws, "Review of Source APM - Gap Analysis", a["summary"][:250] + "...")
    header(ws, 4, ["#", "Area", "Finding", "Rating", "Recommendation", "Action owner", "Due date", "Status"], [5, 22, 55, 10, 55, 16, 12, 14])
    end = rows(ws, 5, [[i + 1, g["area"], g["finding"], g["risk"], g["recommendation"], "", "", "Open"] for i, g in enumerate(a["gaps"])], input_cols=(6, 7, 8))
    dv_list(ws, f"H5:H{end - 1}", ["Open", "In progress", "Closed"])
    for k, f in (("High", RED), ("Medium", AMBER), ("Low", GREEN)):
        ws.conditional_formatting.add(f"D5:D{end - 1}", CellIsRule(operator="equal", formula=[f'"{k}"'], fill=f))
    ws.freeze_panes = "C5"

    # 12 Implementation
    ws = wb.create_sheet("Implementation")
    title(ws, "Implementation Roadmap & Dec 2026 Standards Tracker (POL-16)")
    header(ws, 3, ["When", "Action", "Owner", "Status", "Evidence"], [18, 70, 16, 14, 30])
    pol16 = next(p for p in pols if p["id"] == "POL-16")
    data = [[r[0], r[1], "", "Not started", ""] for r in content["roadmap"]] + [[s[3], s[0], s[1], "Not started", ""] for s in pol16["procedures"]]
    end = rows(ws, 4, data, input_cols=(3, 4, 5))
    dv_list(ws, f"D4:D{end - 1}", ["Not started", "In progress", "Complete"])
    ws.conditional_formatting.add(f"D4:D{end - 1}", CellIsRule(operator="equal", formula=['"Complete"'], fill=GREEN))

    for s in wb.worksheets:
        s.sheet_view.showGridLines = False
        s.page_setup.orientation = "landscape"
        s.page_setup.fitToWidth = 1
    wb.save(out)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    with open(sys.argv[1], encoding="utf-8") as f:
        build(json.load(f), sys.argv[2])
    print("wrote", sys.argv[2])
