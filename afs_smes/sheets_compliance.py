"""IFRS for SMEs (third edition) compliance checklist: Compliance_Summary (by section) and Compliance (every item).

The items come from afs_smes/data/ifrs_smes_3e_checklist.json, extracted from the text of the Standard by
scripts/extract_checklist.py in the skill. Usable inside the AFS workbook (applicability pre-answered from the
ledger) or as a standalone workbook (scripts/compliance_checklist.py in the skill).
"""
import json
from pathlib import Path

from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

from .xl import B_TOP, CENTER, F_BOLD, F_GREY, F_INPUT, F_LINK, F_NORMAL, F_SUB, FILL_INPUT, WRAP, header_row, put, \
    sheet_title, widths

DATA = Path(__file__).with_name("data") / "ifrs_smes_3e_checklist.json"
GREEN = PatternFill("solid", bgColor="C6EFCE")
RED = PatternFill("solid", bgColor="FFC7CE")
AMBER = PatternFill("solid", bgColor="FFEB9C")
GREYF = PatternFill("solid", bgColor="EDEDED")
MAXLEN = 1500

# section -> (applicability question, default answer key). Keys are resolved by default_applicability().
QUESTIONS = {
    2: "Has the entity used an undue cost or effort exemption offered by a specific paragraph (2.30-2.31)?",
    3: "Always applicable - financial statement presentation",
    4: "Always applicable - statement of financial position",
    5: "Always applicable - statement of comprehensive income",
    6: "Always applicable - statement of changes in equity / income and retained earnings",
    7: "Always applicable - statement of cash flows",
    8: "Always applicable - notes",
    9: "Does the entity have subsidiaries, or present consolidated, combined or separate financial statements?",
    10: "Always applicable - accounting policies, estimates and errors",
    11: "Does the entity hold financial assets or financial liabilities (cash, receivables, payables, loans, investments)?",
    12: "Is any item measured at fair value (revaluation, investment property, biological assets, Part II instruments)?",
    13: "Does the entity hold inventories?",
    14: "Does the entity have investments in associates?",
    15: "Does the entity have joint arrangements (jointly controlled operations, assets or entities)?",
    16: "Does the entity hold investment property?",
    17: "Does the entity hold property, plant and equipment?",
    18: "Does the entity hold intangible assets other than goodwill?",
    19: "Did the entity enter into a business combination, or does it carry goodwill?",
    20: "Is the entity a lessee or lessor under any lease?",
    21: "Does the entity have provisions, contingent liabilities or contingent assets?",
    22: "Has the entity issued equity or compound instruments, or made distributions to owners?",
    23: "Does the entity have revenue from contracts with customers?",
    24: "Has the entity received government grants or other government assistance?",
    25: "Does the entity incur borrowing costs?",
    26: "Does the entity have share-based payment arrangements?",
    27: "Always applicable - impairment indicators are assessed for non-financial assets each reporting date",
    28: "Does the entity have employees (short-term, post-employment, other long-term or termination benefits)?",
    29: "Is the entity subject to income tax?",
    30: "Does the entity have foreign-currency transactions, foreign operations or a presentation currency different from its functional currency?",
    31: "Is the functional currency the currency of a hyperinflationary economy?",
    32: "Always applicable - events after the end of the reporting period",
    33: "Always applicable - related party disclosures (incl. key management personnel compensation)",
    34: "Does the entity have agriculture, extractive activities or service concession arrangements?",
    35: "Is this the entity's first set of financial statements under the IFRS for SMEs Accounting Standard?",
    "A": "Is the third edition applied for the first time this year (mandatory from 1 January 2027, or early)?",
}
ALWAYS = {3, 4, 5, 6, 7, 8, 10, 27, 32, 33}


def load_items():
    return json.load(open(DATA, encoding="utf-8"))


def default_applicability(active_faces=frozenset(), active=frozenset(), natures=frozenset(), cfg=None):
    """Answer each section's question from the ledger where possible ("Yes"/"No"); "?" means unknown."""
    cfg = cfg or {}
    f, a, n = set(active_faces), set(active), set(natures)
    ans = {s: "Yes" for s in ALWAYS}
    ans.update({
        2: "Yes" if cfg.get("undue_cost_or_effort") else "No", 9: "No", 11: "Yes" if f & {"TOR", "CASH", "TOP", "BORR_NC", "BORR_C", "OFA_NC", "OFA_C"} else "No",
        12: "Yes" if a & {"RESV.REVAL", "INVPROP", "OI.FV"} else "No", 13: "Yes" if "INV" in f else "No",
        14: "Yes" if "INVASSOC" in f else "No", 15: "No", 16: "Yes" if "INVPROP" in f else "No",
        17: "Yes" if "PPE" in f else "No", 18: "Yes" if "INTANG" in f else "No", 19: "No",
        20: "Yes" if "RENT" in n or a & {"BORR_NC.LEASE", "BORR_C.LEASE"} else "?",
        21: "Yes" if f & {"PROV_C", "PROV_NC"} else "?", 22: "Yes", 23: "Yes" if "REV" in f else "No", 24: "No",
        25: "Yes" if "FINCOST" in f else "No", 26: "No", 28: "Yes" if "EMP" in n or "EBO_NC" in f else "?",
        29: "Yes" if "TAX" in f else "No",
        30: "Yes" if cfg.get("pres_currency", cfg.get("currency")) != cfg.get("currency") else "?",
        31: "No", 34: "No", 35: "Yes" if cfg.get("first_time_adopter") else "No",
        "A": "Yes" if cfg.get("early_application") or str(cfg.get("reporting_date", "")) >= "2027" else "?",
    })
    return ans


def where_map(notes):
    """Section -> where the requirement is usually addressed in this workbook's set."""
    n = lambda *keys: ", ".join(f"Note {notes[k]}" for k in keys if k in notes)
    m = {3: "Notes 1-2; all statements", 4: "SoFP; " + n("equity"), 5: "SoCI; " + n("expenses"), 6: "SoCE",
         7: "SoCF; " + n("cash", "borr"), 8: n("basis", "policies", "judgements"), 10: n("basis", "policies"),
         11: n("tor", "fininst", "borr", "finance"), 12: n("ppe", "fininst"), 13: n("inv"), 14: n("invassoc"),
         16: n("invprop"), 17: n("ppe"), 18: n("intang"), 20: n("commitments"), 21: n("prov", "contingencies"),
         22: n("equity", "dividends"), 23: n("revenue"), 25: n("finance"), 27: n("ppe", "tor", "inv"),
         28: n("pbt", "ebo"), 29: n("tax", "dtax"), 30: n("basis") + "; SoCI OCI; SoCE", 32: n("events", "approval"),
         33: n("related"), 35: n("basis"), "A": n("basis")}
    return {k: v.strip("; ") for k, v in m.items()}


def build_compliance(wb, items=None, applies=None, where=None, entity_ref=None, period_ref=None):
    """Add Compliance_Summary and Compliance sheets to workbook `wb`. Returns (summary_ws, items_ws)."""
    items = items or load_items()
    applies = applies or {}
    where = where or {}
    sections = []
    for it in items:
        if it["section"] not in [s for s, _ in sections]:
            sections.append((it["section"], it["section_title"]))
    sm = wb.create_sheet("Compliance_Summary")
    cl = wb.create_sheet("Compliance")

    # ---------------- items sheet
    sheet_title(cl, "IFRS for SMEs Accounting Standard (third edition) - compliance checklist",
                "One row per requirement (lettered sub-paragraphs separately). Applicability follows the section answer on "
                "Compliance_Summary; override per item if needed. Complete Status, Where addressed and the sign-off columns (blue).")
    if entity_ref:
        put(cl, "A3", f'={entity_ref}&" - "&{period_ref}' if period_ref else f"={entity_ref}", F_BOLD)
    widths(cl, [6, 8, 11, 30, 15, 90, 11, 12, 24, 30, 14])
    header_row(cl, 5, ["ID", "Section", "Paragraph", "Topic", "Area", "Requirement (text of the Standard)",
                       "Applicable?", "Status", "Where addressed", "Preparer comment", "Reviewer"], height=30)
    first = 6
    sec_row = {}
    for i, it in enumerate(items):
        r = first + i
        req = it["requirement"]
        if it.get("context"):
            req = f"[{it['context'][:220]}{'...' if len(it['context']) > 220 else ''}] {req}"
        if len(req) > MAXLEN:
            req = req[:MAXLEN] + " ... (see the full paragraph in the Standard)"
        cl.cell(r, 1, it["id"]).font = F_NORMAL
        cl.cell(r, 2, str(it["section"])).font = F_NORMAL
        cl.cell(r, 3, it["ref"]).font = F_BOLD
        cl.cell(r, 4, it["topic"]).font = F_NORMAL
        cl.cell(r, 5, it["area"]).font = F_NORMAL
        c = cl.cell(r, 6, req)
        c.font, c.alignment = F_NORMAL, WRAP
        for col in (1, 2, 3, 4, 5):
            cl.cell(r, col).alignment = Alignment(vertical="top", wrap_text=col == 4)
        sec_row.setdefault(it["section"], []).append(r)
        c = cl.cell(r, 7)   # applicability: formula from the section answer; the preparer may overwrite with Yes/No
        c.font, c.alignment = F_LINK, Alignment(horizontal="center", vertical="top")
        c = cl.cell(r, 8)
        c.font, c.alignment = F_INPUT, Alignment(horizontal="center", vertical="top")
        c = cl.cell(r, 9, where.get(it["section"], ""))
        c.font, c.alignment = F_INPUT, WRAP
        for col in (10, 11):
            cl.cell(r, col).font, cl.cell(r, col).alignment = F_INPUT, WRAP
        cl.row_dimensions[r].height = min(409, 13 * max(1, -(-len(req) // 110)) + 2)
    last = first + len(items) - 1
    dv = DataValidation(type="list", formula1='"Yes,No,Partly,N/A"', allow_blank=True)
    dv.add(f"H{first}:H{last}")
    dv2 = DataValidation(type="list", formula1='"Yes,No"', allow_blank=False)
    dv2.add(f"G{first}:G{last}")
    cl.add_data_validation(dv)
    cl.add_data_validation(dv2)
    for val, fill in (("Yes", GREEN), ("Partly", AMBER), ("No", RED)):
        cl.conditional_formatting.add(f"H{first}:H{last}", CellIsRule(operator="equal", formula=[f'"{val}"'], fill=fill))
    cl.conditional_formatting.add(f"A{first}:K{last}", FormulaRule(formula=[f'$G{first}="No"'], fill=GREYF))
    cl.conditional_formatting.add(f"H{first}:H{last}", FormulaRule(formula=[f'AND($G{first}="Yes",$H{first}="")'], fill=AMBER))
    cl.freeze_panes = "D6"
    cl.auto_filter.ref = f"A5:K{last}"

    # ---------------- summary sheet
    sheet_title(sm, "Compliance checklist - summary by section",
                "Answer the question for each section (blue). 'Yes' makes its items applicable on the Compliance sheet. "
                "Pre-answered from the ledger where possible - confirm every answer.")
    if entity_ref:
        put(sm, "A3", f'={entity_ref}&" - "&{period_ref}' if period_ref else f"={entity_ref}", F_BOLD)
    widths(sm, [9, 44, 70, 11, 9, 11, 9, 9, 9, 12, 12])
    header_row(sm, 5, ["Section", "Title", "Applicability question", "Applies?", "Items", "Applicable", "Yes",
                       "Partly", "No", "Not yet answered", "% complete"], height=30)
    srow = {}
    for k, (sec, title) in enumerate(sections):
        r = 6 + k
        srow[sec] = r
        put(sm, f"A{r}", str(sec), F_BOLD, align=CENTER)
        put(sm, f"B{r}", title, align=WRAP)
        put(sm, f"C{r}", QUESTIONS.get(sec, ""), F_GREY, align=WRAP)
        ans = applies.get(sec, "Yes" if sec in ALWAYS else "?")
        c = put(sm, f"D{r}", ans, F_INPUT, align=CENTER)
        if ans == "?":
            c.fill = FILL_INPUT
        rng = lambda col: f"Compliance!${col}${first}:${col}${last}"
        put(sm, f"E{r}", f'=COUNTIF({rng("B")},A{r})', align=CENTER)
        put(sm, f"F{r}", f'=COUNTIFS({rng("B")},A{r},{rng("G")},"Yes")', align=CENTER)
        for col, val in (("G", "Yes"), ("H", "Partly"), ("I", "No")):
            put(sm, f"{col}{r}", f'=COUNTIFS({rng("B")},A{r},{rng("G")},"Yes",{rng("H")},"{val}")', align=CENTER)
        put(sm, f"J{r}", f'=COUNTIFS({rng("B")},A{r},{rng("G")},"Yes",{rng("H")},"")', align=CENTER)
        put(sm, f"K{r}", f"=IF(F{r}=0,\"n/a\",(F{r}-J{r}-I{r}-H{r}*0.5)/F{r})", fmt="0%", align=CENTER)
        sm.row_dimensions[r].height = 28
        for rr in sec_row[sec]:
            cl.cell(rr, 7).value = f'=IF(Compliance_Summary!$D${r}="Yes","Yes","No")'
    lastr = 5 + len(sections)
    t = lastr + 1
    put(sm, f"B{t}", "Total", F_BOLD, border=B_TOP)
    for col in "EFGHIJ":
        put(sm, f"{col}{t}", f"=SUM({col}6:{col}{lastr})", F_BOLD, align=CENTER, border=B_TOP)
    put(sm, f"K{t}", f"=IF(F{t}=0,\"n/a\",(F{t}-J{t}-I{t}-H{t}*0.5)/F{t})", F_BOLD, "0%", CENTER, B_TOP)
    dv3 = DataValidation(type="list", formula1='"Yes,No"', allow_blank=False)
    dv3.add(f"D6:D{lastr}")
    sm.add_data_validation(dv3)
    sm.conditional_formatting.add(f"D6:D{lastr}", CellIsRule(operator="equal", formula=['"?"'], fill=AMBER))
    sm.conditional_formatting.add(f"I6:I{lastr}", CellIsRule(operator="greaterThan", formula=["0"], fill=RED))
    sm.conditional_formatting.add(f"K6:K{t}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                                                         color="1BAF7A"))
    # by area
    a0 = t + 3
    put(sm, f"B{a0 - 1}", "By area", F_SUB)
    header_row(sm, a0, ["", "Area", "", "", "Items", "Applicable", "Yes", "Partly", "No", "Not yet answered", "% complete"])
    for k, area in enumerate(["Presentation", "Disclosure", "Recognition and measurement", "Transition", "Scope"]):
        r = a0 + 1 + k
        put(sm, f"B{r}", area)
        rng = lambda col: f"Compliance!${col}${first}:${col}${last}"
        put(sm, f"E{r}", f'=COUNTIF({rng("E")},B{r})', align=CENTER)
        put(sm, f"F{r}", f'=COUNTIFS({rng("E")},B{r},{rng("G")},"Yes")', align=CENTER)
        for col, val in (("G", "Yes"), ("H", "Partly"), ("I", "No")):
            put(sm, f"{col}{r}", f'=COUNTIFS({rng("E")},B{r},{rng("G")},"Yes",{rng("H")},"{val}")', align=CENTER)
        put(sm, f"J{r}", f'=COUNTIFS({rng("E")},B{r},{rng("G")},"Yes",{rng("H")},"")', align=CENTER)
        put(sm, f"K{r}", f"=IF(F{r}=0,\"n/a\",(F{r}-J{r}-I{r}-H{r}*0.5)/F{r})", fmt="0%", align=CENTER)
    note = a0 + 8
    put(sm, f"B{note}", "Status meanings: Yes = complied / disclosed; Partly = incomplete; No = not complied (explain in the comment and "
                        "raise a finding); N/A = requirement does not apply in the entity's circumstances (counts as answered). "
                        "% complete = answered (Yes or N/A) plus half of Partly, over applicable items; No counts as not complete.", F_GREY)
    sm.freeze_panes = "D6"
    sm.summary_total_row = t
    return sm, cl
