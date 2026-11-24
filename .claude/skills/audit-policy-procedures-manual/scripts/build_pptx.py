#!/usr/bin/env python3
"""Build the staff briefing / training deck (.pptx) from a content JSON file.

Usage: python build_pptx.py <content.json> <output.pptx>
"""
import json
import sys

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

NAVY = RGBColor(0x1F, 0x3A, 0x5F)
TEAL = RGBColor(0x1A, 0x7F, 0x86)
LIGHT = RGBColor(0xEE, 0xF3, 0xF8)
GREY = RGBColor(0x59, 0x59, 0x59)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
RATING = {"High": RGBColor(0xC0, 0x39, 0x2B), "Medium": RGBColor(0xD6, 0x8A, 0x00), "Low": RGBColor(0x2E, 0x7D, 0x32)}
W, H = Inches(13.333), Inches(7.5)


class Deck:
    def __init__(self, meta):
        self.p = Presentation()
        self.p.slide_width, self.p.slide_height = W, H
        self.meta = meta
        self.n = 0

    def _blank(self):
        self.n += 1
        return self.p.slides.add_slide(self.p.slide_layouts[6])

    def box(self, s, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE):
        sh = s.shapes.add_shape(shape, x, y, w, h)
        if fill is None:
            sh.fill.background()
        else:
            sh.fill.solid()
            sh.fill.fore_color.rgb = fill
        if line is None:
            sh.line.fill.background()
        else:
            sh.line.color.rgb = line
        sh.shadow.inherit = False
        return sh

    def text(self, s, x, y, w, h, txt, size=16, bold=False, color=GREY, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, shape=None):
        tb = shape or s.shapes.add_textbox(x, y, w, h)
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        for m in ("margin_left", "margin_right"):
            setattr(tf, m, Inches(0.1))
        lines = txt if isinstance(txt, list) else [txt]
        for i, ln in enumerate(lines):
            para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            para.alignment = align
            r = para.add_run()
            r.text = ln
            r.font.size, r.font.bold, r.font.color.rgb, r.font.name = Pt(size), bold, color, "Calibri"
        return tb

    def bullets(self, s, x, y, w, h, items, size=16, color=GREY):
        tb = s.shapes.add_textbox(x, y, w, h)
        tf = tb.text_frame
        tf.word_wrap = True
        for i, it in enumerate(items):
            para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            para.space_after = Pt(8)
            r = para.add_run()
            r.text = "■  "
            r.font.size, r.font.color.rgb = Pt(size - 4), TEAL
            r2 = para.add_run()
            r2.text = it
            r2.font.size, r2.font.color.rgb, r2.font.name = Pt(size), color, "Calibri"
        return tb

    def slide(self, title, kicker=None):
        s = self._blank()
        self.box(s, 0, 0, Inches(0.25), H, fill=NAVY)
        if kicker:
            self.text(s, Inches(0.6), Inches(0.3), Inches(11), Inches(0.4), kicker.upper(), 12, True, TEAL)
        self.text(s, Inches(0.6), Inches(0.6), Inches(12), Inches(0.9), title, 30, True, NAVY)
        self.box(s, Inches(0.7), Inches(1.45), Inches(1.2), Emu(38000), fill=TEAL)
        self.text(s, Inches(0.6), Inches(7.0), Inches(9), Inches(0.35), f"{self.meta['org']}  |  {self.meta['title']}  |  {self.meta['doc_ref']}", 10, color=GREY)
        self.text(s, Inches(12.0), Inches(7.0), Inches(1.0), Inches(0.35), str(self.n), 10, color=GREY, align=PP_ALIGN.RIGHT)
        return s

    def table(self, s, x, y, w, headers, rows, col_w=None, size=12, row_h=0.4, color_col=None):
        shp = s.shapes.add_table(len(rows) + 1, len(headers), x, y, w, Inches(row_h * (len(rows) + 1)))
        t = shp.table
        if col_w:
            for i, cw in enumerate(col_w):
                t.columns[i].width = Inches(cw)
        for i, h in enumerate(headers):
            c = t.cell(0, i)
            c.fill.solid()
            c.fill.fore_color.rgb = NAVY
            c.text = str(h)
            p = c.text_frame.paragraphs[0]
            p.runs[0].font.size, p.runs[0].font.bold, p.runs[0].font.color.rgb = Pt(size), True, WHITE
        for r, row in enumerate(rows, 1):
            for i, v in enumerate(row):
                c = t.cell(r, i)
                c.fill.solid()
                c.fill.fore_color.rgb = LIGHT if r % 2 else WHITE
                c.text = str(v)
                run = c.text_frame.paragraphs[0].runs[0]
                run.font.size, run.font.color.rgb = Pt(size), GREY
                if color_col is not None and i == color_col and v in RATING:
                    run.font.bold, run.font.color.rgb = True, RATING[v]
        return t

    def cards(self, s, items, top=1.8, cols=3, h=2.2, size=13, title_size=16):
        gap = 0.25
        cw = (12.2 - gap * (cols - 1)) / cols
        for i, (head, body) in enumerate(items):
            r, c = divmod(i, cols)
            x, y = Inches(0.6 + c * (cw + gap)), Inches(top + r * (h + gap))
            self.box(s, x, y, Inches(cw), Inches(h), fill=LIGHT)
            self.box(s, x, y, Inches(cw), Emu(60000), fill=TEAL)
            self.text(s, x + Inches(0.1), y + Inches(0.15), Inches(cw - 0.2), Inches(0.5), head, title_size, True, NAVY)
            self.text(s, x + Inches(0.1), y + Inches(0.65), Inches(cw - 0.2), Inches(h - 0.75), body, size, color=GREY)


def build(content, out):
    m, a, pols = content["meta"], content["analysis"], content["policies"]
    d = Deck(m)

    # 1 Title
    s = d._blank()
    d.box(s, 0, 0, W, H, fill=NAVY)
    d.box(s, Inches(0.8), Inches(3.55), Inches(2), Emu(50000), fill=TEAL)
    d.text(s, Inches(0.8), Inches(1.5), Inches(11), Inches(0.5), m["org"].upper(), 16, True, RGBColor(0x9F, 0xD8, 0xDC))
    d.text(s, Inches(0.8), Inches(2.0), Inches(11.5), Inches(1.5), m["title"], 44, True, WHITE)
    d.text(s, Inches(0.8), Inches(3.8), Inches(11.5), Inches(1.0), m["subtitle"], 20, color=WHITE)
    d.text(s, Inches(0.8), Inches(5.6), Inches(11.5), Inches(1.0), [f"Staff briefing  |  {m['version']}  |  Effective {m['effective_date']}", f"Based on: {m['source']}"], 13, color=RGBColor(0xC8, 0xD3, 0xE0))

    # 2 Agenda
    s = d.slide("Agenda", "Briefing overview")
    d.cards(s, [("1  Why this manual", "From APM guidance to firm policy"), ("2  Review findings", "Strengths and gaps in the source"),
                ("3  Policy framework", "16 policies across the audit loop"), ("4  Key requirements", "Acceptance, risk, fraud, materiality, sampling, documentation"),
                ("5  Completion and reporting", "Misstatements, going concern, opinion"), ("6  Implementation", "Dec 2026 standards, roadmap and toolkit")], h=2.3, size=15)

    # 3 Why
    s = d.slide("From guidance to firm policy", "Why this manual")
    d.bullets(s, Inches(0.6), Inches(1.8), Inches(6.4), Inches(5), [
        "The APM is guidance. It is not regulation and does not sit above the ISAs.",
        "Following it does not by itself make a file compliant. Departing from it is not by itself a breach.",
        "Inspection findings cluster on documentation, not procedures.",
        "This manual adds what the APM leaves to the firm: mandatory 'shall' requirements, owners, timing and evidence."], 17)
    d.box(s, Inches(7.4), Inches(1.9), Inches(5.3), Inches(4.4), fill=NAVY)
    d.text(s, Inches(7.6), Inches(2.1), Inches(4.9), Inches(4.0), ["THE OBJECTIVE", "", "Plan, execute, conclude and document an ISA-compliant audit.", "",
                                                                "Defend every significant judgement to a regulator, successor auditor, quality reviewer or court."], 18, False, WHITE)

    # 4 Audit loop
    s = d.slide("One chain of reasoning: the audit loop", "Framework")
    steps = [("Risk", "ISA 315"), ("Response", "ISA 330"), ("Evidence", "ISA 500"), ("Conclusion", "ISA 450"), ("Opinion", "ISA 700/705")]
    for i, (t, isa) in enumerate(steps):
        x = Inches(0.6 + i * 2.48)
        d.box(s, x, Inches(2.4), Inches(2.4), Inches(1.6), fill=NAVY if i % 2 == 0 else TEAL, shape=MSO_SHAPE.CHEVRON)
        d.text(s, x + Inches(0.3), Inches(2.4), Inches(1.8), Inches(1.6), [t, isa], 15, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    d.text(s, Inches(0.8), Inches(4.6), Inches(11.8), Inches(1.5),
           "“Audit quality depends less on the procedures performed than on the logical link between risk, response, evidence and opinion.”", 20, False, NAVY, PP_ALIGN.CENTER)

    # 5 Review: strengths vs gaps
    s = d.slide("Review of the source: what works, what is missing", "Review findings")
    d.box(s, Inches(0.6), Inches(1.8), Inches(6.0), Inches(0.5), fill=TEAL)
    d.text(s, Inches(0.7), Inches(1.83), Inches(5.8), Inches(0.45), "STRENGTHS", 15, True, WHITE)
    d.bullets(s, Inches(0.6), Inches(2.4), Inches(6.0), Inches(4.5), a["strengths"][:5], 13)
    d.box(s, Inches(6.9), Inches(1.8), Inches(5.8), Inches(0.5), fill=RATING["High"])
    d.text(s, Inches(7.0), Inches(1.83), Inches(5.6), Inches(0.45), "GAPS THIS MANUAL CLOSES", 15, True, WHITE)
    d.bullets(s, Inches(6.9), Inches(2.4), Inches(5.8), Inches(4.5), [f"{g['area']}: {g['recommendation'].split('(')[0].strip()}" for g in a["gaps"][:6]], 12)

    # 6 Gap table
    s = d.slide("Gap analysis and rating", "Review findings")
    d.table(s, Inches(0.6), Inches(1.75), Inches(12.1), ["#", "Area", "Rating", "Recommendation"],
            [[i + 1, g["area"], g["risk"], g["recommendation"]] for i, g in enumerate(a["gaps"])], [0.5, 2.6, 1.1, 7.9], 11, 0.46, color_col=2)

    # 7 Policy map
    s = d.slide("The policy framework: 16 policies", "Policy framework")
    cols = 4
    for i, p in enumerate(pols):
        r, c = divmod(i, cols)
        x, y = Inches(0.6 + c * 3.08), Inches(1.75 + r * 1.27)
        d.box(s, x, y, Inches(2.95), Inches(1.15), fill=LIGHT)
        d.box(s, x, y, Emu(60000), Inches(1.15), fill=TEAL)
        d.text(s, x + Inches(0.12), y + Inches(0.05), Inches(2.8), Inches(0.35), f"{p['id']}  ·  {p['phase']}", 10, True, TEAL)
        d.text(s, x + Inches(0.12), y + Inches(0.38), Inches(2.8), Inches(0.75), p["title"], 12, True, NAVY)

    # 8 Acceptance
    s = d.slide("Acceptance: three tests, all must pass", "POL-02  Client acceptance and continuance")
    d.cards(s, [("INTEGRITY", "Will they lie to me, and would I detect it?\nOwners, reason the last auditor left, regulator history, accounting aggressiveness."),
                ("CAPABILITY", "Can we do this job properly?\nIndustry knowledge, staff, hours to deadline, fee against hours."),
                ("INDEPENDENCE", "Is there a threat, and does a safeguard truly remove it?\nFinancial interest, family, non-audit services, fee dependence.")], top=1.8, h=2.2, size=13)
    d.text(s, Inches(0.6), Inches(4.3), Inches(12), Inches(0.4), "Continuance is a fresh decision if any of these six changed:", 15, True, NAVY)
    d.text(s, Inches(0.6), Inches(4.75), Inches(12), Inches(1.4), "  •  ".join(t[0] for t in content["continuance_triggers"]), 14, color=GREY)
    d.box(s, Inches(0.6), Inches(5.8), Inches(12.1), Inches(0.8), fill=NAVY)
    d.text(s, Inches(0.7), Inches(5.85), Inches(11.9), Inches(0.7), "The EP personally signs Ac2 BEFORE fieldwork. A memo signed after fieldwork started is not an acceptance decision.", 15, True, WHITE, anchor=MSO_ANCHOR.MIDDLE)

    # 9 Risk
    s = d.slide("A heading is not a risk", "POL-05  Risk assessment")
    d.box(s, Inches(0.6), Inches(1.8), Inches(5.9), Inches(1.5), fill=RGBColor(0xF4, 0xCC, 0xCC))
    d.text(s, Inches(0.75), Inches(1.9), Inches(5.6), Inches(1.3), ["✕  A HEADING", "\"Revenue may be overstated.\" Risk rating: High."], 15, False, NAVY)
    d.box(s, Inches(6.8), Inches(1.8), Inches(5.9), Inches(1.5), fill=RGBColor(0xD9, 0xEA, 0xD3))
    d.text(s, Inches(6.95), Inches(1.9), Inches(5.6), Inches(1.3), ["✓  A RISK", "Export revenue may be recognised before control transfers, because Incoterms vary and sales staff are bonused on shipped volume."], 13, False, NAVY)
    d.cards(s, [("WHAT", "The assertion that could be misstated (e.g. Occurrence, Cut-off)"), ("WHERE", "The class, balance or disclosure"),
                ("WHY", "The inherent risk factor: uncertainty, subjectivity, complexity, change, bias or fraud")], top=3.6, h=1.5, size=13)
    d.text(s, Inches(0.6), Inches(5.4), Inches(12), Inches(1.2), ["RMM = Inherent risk × Control risk. Assess them separately at the assertion level.",
                                                               "If controls are not tested, CR is HIGH and RMM equals IR. Complete the assessment before fieldwork."], 15, True, NAVY)

    # 10 Fraud
    s = d.slide("Fraud: a lens throughout, not a checklist", "POL-06  ISA 240 (Revised)")
    d.bullets(s, Inches(0.6), Inches(1.8), Inches(7.2), Inches(5), [
        "Identify fraud risk inside the ISA 315 process: pressure, opportunity, rationalisation.",
        "Revenue recognition: a presumed significant risk. Rebut only in narrow circumstances.",
        "Management override: significant on every audit. Test journals, review estimates for bias, add unpredictability.",
        "Judge and record document authenticity; corroborate outside the client where incentives exist.",
        "Communicate fraud matters throughout the engagement, not once at the briefing."], 15)
    d.box(s, Inches(8.2), Inches(1.9), Inches(4.5), Inches(4.2), fill=NAVY)
    d.text(s, Inches(8.4), Inches(2.1), Inches(4.1), Inches(3.9), ["EFFECTIVE", "Periods beginning on or after 15 December 2026", "", "For a 31 December year end: the FY2027 audit, reports signed in early 2028."], 16, False, WHITE)

    # 11 Materiality
    s = d.slide("Materiality: the benchmark is the judgement", "POL-07  ISA 320 / ISA 450")
    d.table(s, Inches(0.6), Inches(1.8), Inches(12.1), ["Benchmark", "Typical range", "Best used for"],
            [[b[0], f"{b[1]*100:g}% - {b[2]*100:g}%", b[3]] for b in content["materiality_benchmarks"]], [3.3, 2.3, 6.5], 14, 0.5)
    d.bullets(s, Inches(0.6), Inches(4.6), Inches(12), Inches(2.2), [
        "Justify BOTH the benchmark and the percentage. Remove unusual items first.",
        "PBT negative, volatile or nominal? Change the benchmark and say why.",
        "Reassess against actual results BEFORE evaluating misstatements. If materiality falls, revisit the sufficiency of the work done."], 15)

    # 12 Sampling
    s = d.slide("Sampling: choose the approach first", "POL-08  ISA 530 / Ac15")
    d.cards(s, [("I", "Immaterial: no sampling"), ("P", "Proof in total: no sampling"), ("%", "Test 100%: no sampling"),
                ("T", "Transaction testing: table"), ("S", "Substantive: formula")], top=1.8, cols=5, h=1.25, size=12, title_size=22)
    d.table(s, Inches(0.6), Inches(3.35), Inches(6.6), ["Population", "High", "Medium", "Low"],
            [[t[0], f"{t[3]}+", f"{t[4]}+", t[5]] for t in content["transaction_samples"]], [2.4, 1.4, 1.4, 1.4], 12, 0.42)
    ex = content["substantive_example"]
    d.box(s, Inches(7.5), Inches(3.35), Inches(5.2), Inches(3.3), fill=LIGHT)
    d.text(s, Inches(7.65), Inches(3.45), Inches(4.9), Inches(3.1), ["SUBSTANTIVE FORMULA", "Residual population ÷ overall materiality × risk factor", "",
                                                                    f"Example: {ex['population']:,} less {ex['large_items']:,} (> PM) less {ex['key_items']:,} (related party) = {ex['population']-ex['large_items']-ex['key_items']:,}",
                                                                    f"÷ {ex['overall_materiality']:,} × {ex['risk_factor']} = {ex['expected_sample']} items"], 13, False, NAVY)

    # 13 Documentation
    s = d.slide("The working paper test", "POL-10  ISA 230")
    d.cards(s, [("1  PURPOSE", "Which risk and assertion does this respond to? Reference it."), ("2  WORK DONE", "What was done and how items were selected, in enough detail to repeat."),
                ("3  RESULTS", "What was found, especially the exceptions."), ("4  CONCLUSION", "What it means for the assertion. 'No exceptions noted' is not a conclusion.")], top=1.8, cols=4, h=2.3, size=13)
    d.text(s, Inches(0.6), Inches(4.4), Inches(12), Inches(0.4), "File structure", 16, True, NAVY)
    d.table(s, Inches(0.6), Inches(4.85), Inches(12.1), ["A / B", "Ac1-Ac19", "C - O", "Ab", "Aa"],
            [["Permanent file and systems", "Planning and control (before fieldwork)", "Detailed sections (during fieldwork)", "Disclosures and critical review", "Finalisation, opinion, points forward"]], None, 12, 0.6)

    # 14 Completion
    s = d.slide("Completion: size is not the only test", "POL-12 / POL-13  ISA 450 / ISA 570")
    d.cards(s, content["material_by_nature"], top=1.75, cols=3, h=1.3, size=12, title_size=14)
    d.table(s, Inches(0.6), Inches(4.65), Inches(12.1), ["Going concern situation", "Reporting outcome"], content["going_concern"], [5.5, 6.6], 12, 0.42)

    # 15 Reporting
    s = d.slide("Reporting: the material-and-pervasive grid", "POL-14  ISA 700 / 705 / 706")
    d.table(s, Inches(0.6), Inches(1.9), Inches(12.1), ["Nature of matter", "Material, not pervasive", "Material AND pervasive"], content["opinion_grid"], [4.5, 3.8, 3.8], 15, 0.7)
    d.bullets(s, Inches(0.6), Inches(4.3), Inches(12), Inches(2.5), [
        "Emphasis of Matter highlights a matter that is properly presented. It never fixes a problem.",
        "Read the other information (ISA 720). Communicate with governance (ISA 260) and report significant deficiencies (ISA 265).",
        "Date the report only after the evidence, approval of the financial statements, the representation letter and (where required) the EQ review are complete."], 15)

    # 16 Quality
    s = d.slide("Quality management and monitoring", "POL-15  ISQM 1 / ISQM 2 / ISA 220")
    d.cards(s, [("EP responsibility", "Overall responsibility for quality, with sufficient involvement throughout"), ("EQ review", "Criteria-based; completed before the report is dated"),
                ("Inspection", "Annual file inspections covering every EP each cycle"), ("Root cause", "Diagnose every finding to its root cause, back to acceptance if needed"),
                ("Remediation", "Tracked actions reported to the Managing Partner"), ("Capacity", "Hours and staffing are quality objectives; escalate pressure")], h=2.1, size=14)

    # 17 Dec 2026
    s = d.slide("What changes on 15 December 2026", "POL-16  Methodology updates")
    d.table(s, Inches(0.6), Inches(1.8), Inches(12.1), ["Pronouncement", "What we change"], [
        ["ISA 240 (Revised)", "Fraud lens inside ISA 315 templates; revenue rebuttal narrowed; authenticity judgements recorded"],
        ["ISA 570 (Revised 2024)", "12-month assessment measured from the date of approval of the financial statements"],
        ["'Publicly traded entity'", "Replaces 'listed entity'; re-test every client and record the classification"],
        ["External experts", "Updated evaluation of experts' work"],
        ["ISSA 5000", "Sustainability assurance: same date, not an ISA, outside this manual"]], [3.5, 8.6], 14, 0.55)

    # 18 Roadmap
    s = d.slide("Implementation roadmap", "Next steps")
    rm = content["roadmap"]
    seg = 12.1 / len(rm)
    d.box(s, Inches(0.6), Inches(3.05), Inches(12.1), Emu(50000), fill=NAVY)
    for i, (when, what) in enumerate(rm):
        x = Inches(0.6 + i * seg)
        d.box(s, x + Inches(seg / 2 - 0.15), Inches(2.92), Inches(0.3), Inches(0.3), fill=TEAL, shape=MSO_SHAPE.OVAL)
        d.text(s, x, Inches(2.2), Inches(seg), Inches(0.6), when, 14, True, NAVY, PP_ALIGN.CENTER)
        d.text(s, x + Inches(0.05), Inches(3.4), Inches(seg - 0.1), Inches(2.5), what, 12, False, GREY, PP_ALIGN.CENTER)

    # 19 Toolkit
    s = d.slide("Your toolkit", "Companion files")
    d.cards(s, [("WORD  —  The Manual", "Document control, review of the source, 16 policies (purpose, scope, 'shall' statements, procedures with responsibility, WP ref and timing), appendices."),
                ("EXCEL  —  The Workbook", "Policy and procedures registers, RACI, WP index, acceptance checklist, materiality and sample size calculators, risk register, ISA 450 misstatements, gap and implementation trackers."),
                ("POWERPOINT  —  This Deck", "Staff briefing and training on the requirements and the Dec 2026 changes.")], h=3.4, size=14)

    # 20 Close
    s = d._blank()
    d.box(s, 0, 0, W, H, fill=NAVY)
    d.text(s, Inches(0.8), Inches(2.4), Inches(11.5), Inches(1.2), "Defensibility is the objective.", 40, True, WHITE)
    d.text(s, Inches(0.8), Inches(3.6), Inches(11.5), Inches(1.0), "The manual is the vehicle. Questions and discussion.", 22, color=RGBColor(0x9F, 0xD8, 0xDC))
    d.text(s, Inches(0.8), Inches(5.8), Inches(11.5), Inches(0.6), f"Contact: {m['owner']}  |  {m['org']}", 14, color=RGBColor(0xC8, 0xD3, 0xE0))

    d.p.save(out)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    with open(sys.argv[1], encoding="utf-8") as f:
        build(json.load(f), sys.argv[2])
    print("wrote", sys.argv[2])
