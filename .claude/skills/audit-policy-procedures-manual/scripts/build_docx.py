#!/usr/bin/env python3
"""Build the Policies & Procedures Manual (.docx) from a content JSON file.

Usage: python build_docx.py <content.json> <output.docx>
"""
import json
import sys

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

NAVY = RGBColor(0x1F, 0x3A, 0x5F)
TEAL = RGBColor(0x1A, 0x7F, 0x86)
GREY = RGBColor(0x59, 0x59, 0x59)
HDR_FILL = "1F3A5F"
ALT_FILL = "EEF3F8"
RISK_FILL = {"High": "F4CCCC", "Medium": "FFF2CC", "Low": "D9EAD3"}


def shade(cell, hex_fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tcPr.append(shd)


def repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    trPr.append(el)


def add_field(paragraph, instr):
    run = paragraph.add_run()
    for tag, text in (("begin", None), (None, instr), ("separate", None), (None, "Right-click and choose Update Field to build the table of contents."), ("end", None)):
        if tag:
            fc = OxmlElement("w:fldChar")
            fc.set(qn("w:fldCharType"), tag)
            run._r.append(fc)
        elif text == instr:
            it = OxmlElement("w:instrText")
            it.set(qn("xml:space"), "preserve")
            it.text = instr
            run._r.append(it)
        else:
            t = OxmlElement("w:t")
            t.text = text
            run._r.append(t)


def table(doc, headers, rows, widths=None, risk_col=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = t.rows[0]
    repeat_header(hdr)
    for i, h in enumerate(headers):
        c = hdr.cells[i]
        c.text = ""
        r = c.paragraphs[0].add_run(str(h))
        r.bold = True
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        r.font.size = Pt(9.5)
        shade(c, HDR_FILL)
    for n, row in enumerate(rows):
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ""
            run = cells[i].paragraphs[0].add_run(str(v))
            run.font.size = Pt(9.5)
            if risk_col is not None and i == risk_col and v in RISK_FILL:
                shade(cells[i], RISK_FILL[v])
                run.bold = True
            elif n % 2 == 1:
                shade(cells[i], ALT_FILL)
    if widths:
        t.autofit = False
        for i, w in enumerate(widths):
            t.columns[i].width = Cm(w)  # sets tblGrid/gridCol (honoured by Word and LibreOffice)
        for row in t.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Cm(w)
    doc.add_paragraph()
    return t


def bullets(doc, items, style="List Bullet"):
    for it in items:
        doc.add_paragraph(it, style=style)


def label(doc, text):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold = True
    r.font.color.rgb = TEAL
    r.font.size = Pt(11)
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)


def setup_styles(doc):
    st = doc.styles["Normal"]
    st.font.name = "Calibri"
    st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    for name, size in (("Heading 1", 18), ("Heading 2", 14), ("Heading 3", 12)):
        h = doc.styles[name]
        h.font.name = "Calibri"
        h.font.size = Pt(size)
        h.font.color.rgb = NAVY
        h.font.bold = True


def header_footer(section, meta):
    hp = section.header.paragraphs[0]
    hp.text = f"{meta['org']}  |  {meta['title']}  |  {meta['doc_ref']}"
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    for r in hp.runs:
        r.font.size = Pt(8)
        r.font.color.rgb = GREY
    fp = section.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = fp.add_run(f"Version {meta['version']}  |  Effective {meta['effective_date']}  |  Page ")
    r.font.size = Pt(8)
    run = fp.add_run()
    for tag, txt in (("begin", None), (None, "PAGE"), ("end", None)):
        if tag:
            fc = OxmlElement("w:fldChar")
            fc.set(qn("w:fldCharType"), tag)
            run._r.append(fc)
        else:
            it = OxmlElement("w:instrText")
            it.text = txt
            run._r.append(it)
    run.font.size = Pt(8)


def build(content, out):
    m = content["meta"]
    a = content["analysis"]
    doc = Document()
    setup_styles(doc)
    sec = doc.sections[0]
    sec.page_height, sec.page_width = Cm(29.7), Cm(21.0)
    for side in ("left_margin", "right_margin"):
        setattr(sec, side, Cm(2.2))
    sec.top_margin = sec.bottom_margin = Cm(2.0)
    sec.different_first_page_header_footer = True
    header_footer(sec, m)

    # ---- Cover
    for _ in range(5):
        doc.add_paragraph()
    p = doc.add_paragraph()
    r = p.add_run(m["org"].upper())
    r.font.size, r.font.color.rgb, r.bold = Pt(13), TEAL, True
    p = doc.add_paragraph()
    r = p.add_run(m["title"])
    r.font.size, r.font.color.rgb, r.bold = Pt(32), NAVY, True
    p = doc.add_paragraph()
    r = p.add_run(m["subtitle"])
    r.font.size, r.font.color.rgb = Pt(15), GREY
    doc.add_paragraph()
    for k, v in (("Document reference", m["doc_ref"]), ("Version", m["version"]), ("Effective date", m["effective_date"]), ("Owner", m["owner"]), ("Approved by", m["approved_by"])):
        p = doc.add_paragraph()
        r1 = p.add_run(f"{k}: ")
        r1.bold = True
        p.add_run(v)
    doc.add_paragraph()
    p = doc.add_paragraph()
    r = p.add_run(f"Developed from: {m['source']}")
    r.italic, r.font.size, r.font.color.rgb = True, Pt(9), GREY
    doc.add_page_break()

    # ---- Document control
    doc.add_heading("Document Control", 1)
    table(doc, ["Item", "Detail"], [
        ["Document reference", m["doc_ref"]], ["Version", m["version"]], ["Effective date", m["effective_date"]],
        ["Owner", m["owner"]], ["Approved by", m["approved_by"]], ["Review cycle", m["review_cycle"]],
        ["Framework", m["framework"]], ["Source document", m["source"]]], widths=[4.5, 12])
    doc.add_heading("Revision History", 2)
    table(doc, ["Version", "Date", "Description", "Author", "Approved"], [
        [m["version"], m["effective_date"], "Initial issue, based on review of the APM", m["owner"], "Pending"],
        ["", "", "", "", ""]], widths=[2.2, 3, 6.5, 3, 2.5])
    doc.add_heading("Approval", 2)
    table(doc, ["Role", "Name", "Signature", "Date"], [
        [m["owner"], "", "", ""], ["Quality Management Partner", "", "", ""], [m["approved_by"], "", "", ""]], widths=[5, 4, 4, 3])
    doc.add_page_break()

    doc.add_heading("Contents", 1)
    add_field(doc.add_paragraph(), 'TOC \\o "1-2" \\h \\z \\u')
    doc.add_page_break()

    # ---- Section 1: Introduction
    doc.add_heading("1. Introduction", 1)
    doc.add_heading("1.1 Purpose", 2)
    doc.add_paragraph(
        f"This manual sets out the policies and procedures {m['org']} applies to statutory audit engagements. "
        "It adopts the Audit Practice Manual (APM) as the firm's working method and adds mandatory requirements, "
        "named responsibilities, timing and documentation standards, so that every audit file can be defended "
        "to a regulator, a successor auditor, an engagement quality reviewer or a court.")
    doc.add_heading("1.2 The Audit Loop", 2)
    doc.add_paragraph("Every policy in this manual serves one chain of reasoning. Audit quality depends less on the "
                      "procedures performed than on the logical link between risk, response, evidence and opinion.")
    table(doc, ["Step", "ISA", "What must be shown"], [
        ["1. Risk", "ISA 315 (Revised 2019) para 28, 31", "Risks at the assertion level; IR and CR assessed separately"],
        ["2. Response", "ISA 330 para 6-7, 21", "Procedures responsive to each assessed risk, with clear linkage"],
        ["3. Evidence", "ISA 500 para 6; ISA 330 para 18", "Sufficient appropriate evidence; substantive work on every material area"],
        ["4. Conclusion", "ISA 450 para 10-11", "Misstatements evaluated individually and in aggregate, by amount and nature"],
        ["5. Opinion", "ISA 700 (Revised); ISA 705 (Revised)", "Opinion formed; modification per the material-and-pervasive grid"]], widths=[2.8, 5.5, 8.5])
    doc.add_heading("1.3 How to Read this Manual", 2)
    bullets(doc, [
        "Section 2 summarises the review and analysis of the source APM, including gaps this manual closes.",
        "Section 3 contains the policies, in audit phase order: acceptance, planning, execution, completion, reporting, quality.",
        "Each policy states its purpose, scope, policy statements ('shall' = mandatory), a procedures table (responsibility, working paper reference, timing), required documentation and red flags for reviewers.",
        "The appendices hold reference tables. The companion Excel workbook holds the calculators, registers and checklists, and the PowerPoint deck is used for staff training."])
    doc.add_heading("1.4 Roles", 2)
    table(doc, ["Code", "Role"], content["roles"], widths=[3, 13])

    # ---- Section 2: Review & analysis
    doc.add_page_break()
    doc.add_heading("2. Review and Analysis of the Source Manual", 1)
    doc.add_heading("2.1 Summary", 2)
    doc.add_paragraph(a["summary"])
    doc.add_heading("2.2 Structure of the Source", 2)
    table(doc, ["Part", "Content"], a["structure"], widths=[3.2, 13.5])
    doc.add_heading("2.3 Strengths", 2)
    bullets(doc, a["strengths"])
    doc.add_heading("2.4 Gaps and Recommendations", 2)
    doc.add_paragraph("Each gap is rated by its effect on file defensibility and points to the policy that closes it.")
    table(doc, ["#", "Area", "Finding", "Rating", "Recommendation"],
          [[i + 1, g["area"], g["finding"], g["risk"], g["recommendation"]] for i, g in enumerate(a["gaps"])],
          widths=[0.8, 2.6, 5.6, 1.6, 6.2], risk_col=3)

    # ---- Section 3: Policies
    doc.add_page_break()
    doc.add_heading("3. Policies and Procedures", 1)
    table(doc, ["Policy", "Title", "Phase", "Primary standards"],
          [[p["id"], p["title"], p["phase"], p["isa"]] for p in content["policies"]], widths=[1.8, 6, 2.8, 6.2])
    for p in content["policies"]:
        doc.add_page_break()
        doc.add_heading(f"{p['id']}  {p['title']}", 2)
        table(doc,["Phase", "Standards"], [[p["phase"], p["isa"]]], widths=[4, 12.7])
        label(doc, "Purpose")
        doc.add_paragraph(p["purpose"])
        label(doc, "Scope")
        doc.add_paragraph(p["scope"])
        label(doc, "Policy")
        for i, s in enumerate(p["statements"], 1):
            para = doc.add_paragraph()
            r = para.add_run(f"{p['id'][-2:]}.{i}  ")
            r.bold = True
            para.add_run(s)
            para.paragraph_format.left_indent = Cm(0.4)
        label(doc, "Procedures")
        table(doc, ["#", "Procedure", "Resp.", "WP ref", "Timing"],
              [[i, *row] for i, row in enumerate(p["procedures"], 1)], widths=[0.8, 9.2, 1.8, 2.2, 2.8])
        label(doc, "Documentation required on file")
        bullets(doc, p["documentation"])
        label(doc, "Red flags for reviewers")
        bullets(doc, p["red_flags"])

    # ---- Appendices
    doc.add_page_break()
    doc.add_heading("Appendix A  Working Paper Index", 1)
    table(doc, ["Ref", "Working paper", "Phase", "Standard"], content["working_papers"], widths=[2, 8.5, 3, 3.3])

    doc.add_heading("Appendix B  Minimum Substantive Procedures by Area", 1)
    table(doc, ["Sect.", "Area", "Main risks", "Key tests"], content["area_procedures"], widths=[1.3, 2.8, 6, 6.6])

    doc.add_page_break()
    doc.add_heading("Appendix C  Materiality and Sampling Reference", 1)
    doc.add_heading("C.1 Materiality benchmarks", 2)
    table(doc, ["Benchmark", "Range", "Best used for"],
          [[b[0], f"{b[1]*100:g}% - {b[2]*100:g}%", b[3]] for b in content["materiality_benchmarks"]], widths=[4, 3, 9.7])
    doc.add_paragraph("These are common practical ranges. The benchmark and percentage are judgements, and the file must justify both.")
    doc.add_heading("C.2 Transaction testing sample sizes (minimum)", 2)
    table(doc, ["Population", "High risk", "Medium risk", "Low risk"],
          [[t[0], f"{t[3]}+", f"{t[4]}+", t[5]] for t in content["transaction_samples"]], widths=[5, 3.5, 3.5, 3.5])
    doc.add_heading("C.3 Substantive testing sample size", 2)
    ex = content["substantive_example"]
    resid = ex["population"] - ex["large_items"] - ex["key_items"]
    doc.add_paragraph("Residual sample size = (population - items above PM - key items) / overall materiality x risk factor")
    doc.add_paragraph(
        f"Worked example: receivables {ex['population']:,}; overall materiality {ex['overall_materiality']:,}; "
        f"PM {ex['performance_materiality']:,}. Remove large items {ex['large_items']:,} and key item {ex['key_items']:,}: "
        f"residual {resid:,}. {resid:,} / {ex['overall_materiality']:,} x {ex['risk_factor']} = "
        f"{resid/ex['overall_materiality']*ex['risk_factor']:.1f}, rounded to {ex['expected_sample']} items. Items removed are tested in full.")
    doc.add_paragraph("Risk factors for medium and high risk shall be taken from the firm's adopted table (see Gap 8). "
                      "The Excel calculator treats the factor as a controlled input.")
    doc.add_heading("C.4 Tests of controls sample sizes", 2)
    table(doc, ["Control frequency", "Sample", "Note"], content["controls_samples"], widths=[5, 2, 9.7])
    doc.add_paragraph("One deviation normally ends reliance. Applies only where controls reliance is planned or required (ISA 330 para 8).")

    doc.add_page_break()
    doc.add_heading("Appendix D  Acceptance and Risk Reference", 1)
    doc.add_heading("D.1 Six continuance triggers", 2)
    table(doc, ["Trigger", "Why it resets the decision"], content["continuance_triggers"], widths=[5, 11.7])
    doc.add_heading("D.2 Five inherent risk factors (ISA 315)", 2)
    table(doc, ["Factor", "Meaning"], content["inherent_risk_factors"], widths=[4, 12.7])
    doc.add_heading("D.3 Material by nature: six triggers (ISA 450)", 2)
    table(doc, ["Trigger", "Why"], content["material_by_nature"], widths=[5, 11.7])

    doc.add_heading("Appendix E  Reporting Reference", 1)
    doc.add_heading("E.1 Material-and-pervasive grid (ISA 705)", 2)
    table(doc, ["Nature of matter", "Material, not pervasive", "Material AND pervasive"], content["opinion_grid"], widths=[6, 5.3, 5.3])
    doc.add_heading("E.2 Going concern outcomes (ISA 570)", 2)
    table(doc, ["Situation", "Reporting outcome"], content["going_concern"], widths=[7, 9.7])

    doc.add_heading("Appendix F  Implementation Roadmap", 1)
    table(doc, ["When", "Action"], content["roadmap"], widths=[3.5, 13.2])

    doc.add_heading("Appendix G  Glossary", 1)
    table(doc, ["Term", "Meaning"], content["glossary"], widths=[3.5, 13.2])

    # Ask Word to refresh fields (TOC, page numbers) when the file is opened
    upd = OxmlElement("w:updateFields")
    upd.set(qn("w:val"), "true")
    doc.settings.element.append(upd)
    doc.save(out)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    with open(sys.argv[1], encoding="utf-8") as f:
        build(json.load(f), sys.argv[2])
    print("wrote", sys.argv[2])
