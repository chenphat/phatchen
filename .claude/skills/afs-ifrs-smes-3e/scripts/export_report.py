#!/usr/bin/env python3
"""Export the financial statements report (Word .docx) from a workbook built by build_afs.py.

The workbook must be recalculated first (open and save it in Excel, or let build_afs.py verify it with
LibreOffice) so that formula values are stored. The report uses the presentation-currency columns only.

Usage:
  python export_report.py AFS_2026.xlsx --out AFS_2026.docx [--banner "DRAFT - for discussion"]
"""
import argparse
import re
import warnings
from datetime import date, datetime

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from openpyxl import load_workbook

FONT = "Arial"
GREY = RGBColor(0x59, 0x59, 0x59)
PLACEHOLDER = re.compile(r"\[[^\]]+\]")


# ------------------------------------------------------------------ helpers
def fmt(value, number_format):
    if value is None or value == "":
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%d %B %Y")
    if not isinstance(value, (int, float)):
        return str(value)
    nf = number_format or ""
    if nf == "0":
        return f"{value:.0f}"
    if "%" in nf:
        s = f"{abs(value) * 100:.1f}%"
    elif "0.000" in nf:
        s = f"{abs(value):,.3f}"
    elif "0.00" in nf:
        s = f"{abs(value):,.2f}"
    else:
        if round(value) == 0:
            return "-"
        s = f"{abs(round(value)):,.0f}"
    return f"({s})" if value < 0 else s


def set_cell_border(cell, top=None, bottom=None):
    tc = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge, style in (("top", top), ("bottom", bottom)):
        if style:
            el = OxmlElement(f"w:{edge}")
            el.set(qn("w:val"), style)
            el.set(qn("w:sz"), "6")
            el.set(qn("w:color"), "000000")
            borders.append(el)
    tc.append(borders)


def run(par, text, bold=False, italic=False, size=10, color=None):
    r = par.add_run(text)
    r.font.name, r.font.size, r.bold, r.italic = FONT, Pt(size), bold, italic
    r._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    if color:
        r.font.color.rgb = color
    return r


def para(doc, text="", bold=False, italic=False, size=10, align=None, space_after=4, color=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(0)
    if align:
        p.alignment = align
    if text:
        # placeholders in square brackets are shown highlighted so the preparer sees what to complete
        pos = 0
        for m in PLACEHOLDER.finditer(text):
            if m.start() > pos:
                run(p, text[pos:m.start()], bold, italic, size, color)
            r = run(p, m.group(0), bold, True, size, color)
            r.font.highlight_color = 7  # yellow
            pos = m.end()
        if pos < len(text):
            run(p, text[pos:], bold, italic, size, color)
    return p


def heading(doc, text, size=13, page_break=False):
    if page_break:
        doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    p = para(doc, text, bold=True, size=size, space_after=6)
    p.paragraph_format.keep_with_next = True
    return p


def statement_header(doc, wb, title):
    E = entity(wb)
    heading(doc, E["name"], 12, page_break=True)
    para(doc, title, bold=True, size=11, space_after=2)
    return E


def entity(wb):
    ws = wb["Entity"]
    vals = {}
    for row in ws.iter_rows(min_row=5, values_only=True):
        if row[1]:
            vals[row[1]] = row[2]
    get = lambda prefix: next((v for k, v in vals.items() if str(k).startswith(prefix)), "")
    return {"name": get("Entity name"), "date": get("Reporting date"), "currency": get("Functional currency"),
            "pres": get("Presentation currency label"), "authorised": get("Date financial statements authorised")}


def table(doc, rows, widths, bold_rows=(), total_rows=(), double_rows=(), header_rows=(0,), size=9):
    t = doc.add_table(rows=len(rows), cols=len(widths))
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    t.autofit = False
    # fixed layout + grid widths: Word and LibreOffice both honour these (cell widths alone are ignored by some)
    tblPr = t._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)
    for j, w in enumerate(widths):
        t.columns[j].width = w
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = t.cell(i, j)
            cell.width = widths[j]
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if j > 0:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            if j == 1 and len(widths) > 3 and widths[1] < Cm(1.6):
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run(p, val, bold=(i in bold_rows or i in header_rows), size=size)
            if len(rows) <= 25 and i < len(rows) - 1:
                p.paragraph_format.keep_with_next = True   # keep short tables on one page
            if j > 0 and val not in ("", None) and (i in total_rows or i in double_rows):
                set_cell_border(cell, top="single", bottom="double" if i in double_rows else None)
            if i in header_rows and j > 0 and i == max(header_rows):
                set_cell_border(cell, bottom="single")
    return t


def visible_rows(ws, first, last=None):
    last = last or ws.max_row
    for r in range(first, last + 1):
        if not ws.row_dimensions[r].hidden:
            yield r


# ------------------------------------------------------------------ sections
def cover(doc, wb, banner):
    E = entity(wb)
    for _ in range(6):
        para(doc)
    para(doc, E["name"], bold=True, size=20, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=12)
    para(doc, "Financial statements", bold=True, size=16, align=WD_ALIGN_PARAGRAPH.CENTER)
    d = E["date"]
    para(doc, f"for the year ended {d.strftime('%d %B %Y') if hasattr(d, 'strftime') else d}", size=13,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=18)
    para(doc, "Prepared in accordance with the IFRS for SMEs Accounting Standard (third edition)", size=11,
         align=WD_ALIGN_PARAGRAPH.CENTER)
    para(doc, f"Presentation currency: {E['pres']}    Functional currency: {E['currency']}", size=11,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=30)
    if banner:
        p = para(doc, banner, bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER, color=RGBColor(0xC0, 0, 0))
    heading(doc, "Contents", 13, page_break=True)
    for i, item in enumerate(["Report of the Board of Directors", "Statement of financial position",
                              "Statement of comprehensive income", "Statement of changes in equity",
                              "Statement of cash flows", "Notes to the financial statements"], start=1):
        para(doc, f"{i}.  {item}", size=11)


def directors_report(doc, wb):
    ws = wb["Directors_Report"]
    heading(doc, ws["B1"].value or "", 12, page_break=True)
    para(doc, ws["B2"].value or "", bold=True, size=11, space_after=2)
    para(doc, ws["B3"].value or "", space_after=10)
    r = 6
    while r <= ws.max_row:
        b = ws[f"B{r}"].value
        nums = [ws[f"{c}{r}"].value for c in "CD"]
        if b is None and all(v is None for v in nums):
            r += 1
            continue
        if "____" in str(b) or "____" in str(nums[0] or ""):
            para(doc, "_" * 28 + " " * 14 + "_" * 28, space_after=2)
            r += 1
            continue
        if str(b or "").startswith("[Name]"):
            para(doc, f"{b}{' ' * 30}{nums[0] or ''}", space_after=2)
            r += 1
            continue
        if any(v is not None for v in nums) or (b is None and ws[f"C{r}"].value is not None):
            block = []
            while r <= ws.max_row and (ws[f"C{r}"].value is not None or ws[f"B{r}"].value not in (None,) and
                                       ws[f"C{r}"].value is not None):
                block.append(r)
                r += 1
            rows = [[str(ws[f"B{x}"].value or "")] + [fmt(ws[f"{c}{x}"].value, ws[f"{c}{x}"].number_format) for c in "CD"]
                    for x in block]
            bold = [i for i, x in enumerate(block) if ws[f"B{x}"].font.bold]
            table(doc, rows, [Cm(9.5), Cm(3.2), Cm(3.2)], bold_rows=bold, header_rows=(0, 1))
            para(doc)
            continue
        if ws[f"B{r}"].font.bold:
            p = para(doc, str(b), bold=True, space_after=2)
            p.paragraph_format.keep_with_next = True
        else:
            para(doc, str(b), space_after=6)
        r += 1


def statement(doc, wb, sheet, last_marker=None):
    ws = wb[sheet]
    statement_header(doc, wb, ws["A2"].value)
    para(doc, ws["A3"].value or "", space_after=8)
    rows, bold, totals, doubles = [], [], [], []
    rows.append(["", "Note", fmt(ws["C5"].value, None), fmt(ws["D5"].value, None)])
    rows.append(["", "", fmt(ws["C6"].value, "0"), fmt(ws["D6"].value, "0")])
    trailer = []
    for r in visible_rows(ws, 7):
        a = ws[f"A{r}"].value
        c, d = ws[f"C{r}"].value, ws[f"D{r}"].value
        if a is None and c is None and d is None:
            if rows[-1] != ["", "", "", ""]:
                rows.append(["", "", "", ""])
            continue
        if c is None and d is None and ws[f"E{r}"].value is None:
            if rows[-1] == ["", "", "", ""] and len(rows) > 3 and not ws[f"A{r}"].font.bold:
                trailer.append(str(a))
                continue
        i = len(rows)
        rows.append([str(a or ""), str(ws[f"B{r}"].value or ""), fmt(c, ws[f"C{r}"].number_format),
                     fmt(d, ws[f"D{r}"].number_format)])
        if ws[f"A{r}"].font.bold:
            bold.append(i)
        border = ws[f"C{r}"].border
        if border.bottom and border.bottom.style == "double":
            doubles.append(i)
        elif border.top and border.top.style == "thin":
            totals.append(i)
    while rows and rows[-1] == ["", "", "", ""]:
        rows.pop()
    table(doc, rows, [Cm(10.2), Cm(1.3), Cm(2.75), Cm(2.75)], bold_rows=bold, total_rows=totals, double_rows=doubles,
          header_rows=(0, 1))
    for t in trailer:
        para(doc, t, italic=True, size=8, color=GREY) if "____" not in t else para(doc, t)
    para(doc, "The accompanying notes form an integral part of these financial statements.", italic=True, size=8,
         color=GREY)


def soce(doc, wb):
    ws = wb["SoCE"]
    statement_header(doc, wb, ws["A2"].value)
    para(doc, ws["A3"].value or "", space_after=8)
    start = next(r for r in range(5, ws.max_row) if str(ws[f"A{r}"].value or "").startswith("Presentation currency"))
    end = next((r for r in range(start + 1, ws.max_row) if str(ws[f"A{r}"].value or "").startswith("Functional currency")),
               ws.max_row)
    para(doc, ws[f"A{start}"].value, bold=True, size=9, space_after=2)
    hdr = start + 1
    rows = [[""] + [str(ws.cell(hdr, c).value or "") for c in range(2, 8)]]
    bold, totals, doubles = [0], [], []
    for r in range(hdr + 1, end):
        a = ws[f"A{r}"].value
        vals = [ws.cell(r, c).value for c in range(2, 8)]
        if a is None and all(v is None for v in vals):
            continue
        i = len(rows)
        rows.append([str(a or "")] + [fmt(v, ws.cell(r, c).number_format) for v, c in zip(vals, range(2, 8))])
        if ws[f"A{r}"].font.bold:
            bold.append(i)
        b = ws.cell(r, 2).border
        if b.bottom and b.bottom.style == "double":
            doubles.append(i)
        elif b.top and b.top.style == "thin":
            totals.append(i)
    table(doc, rows, [Cm(6.2)] + [Cm(2.05)] * 6, bold_rows=bold, total_rows=totals, double_rows=doubles, size=7.5)
    para(doc, "The accompanying notes form an integral part of these financial statements.", italic=True, size=8,
         color=GREY)


def notes(doc, wb):
    ws = wb["Notes"]
    heading(doc, ws["A1"].value, 12, page_break=True)
    para(doc, ws["A2"].value, bold=True, size=11, space_after=2)
    para(doc, ws["A3"].value, space_after=2)
    para(doc, str(ws["A4"].value or "").split(". Blue")[0] + ".", italic=True, size=8, color=GREY, space_after=10)
    cols = "CDEFGH"
    rows = list(visible_rows(ws, 6))
    i = 0
    while i < len(rows):
        r = rows[i]
        a, b = ws[f"A{r}"].value, ws[f"B{r}"].value
        vals = [ws[f"{c}{r}"].value for c in cols]
        has_vals = any(v not in (None, "") for v in vals)
        if a not in (None, "") and b:
            p = para(doc, f"{a}   {b}", bold=True, size=10.5, space_after=4)
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.keep_with_next = True
            i += 1
            continue
        if not has_vals:
            if b and "not material and is not disclosed" in str(b):
                i += 1
                continue
            if b:
                nxt = rows[i + 1] if i + 1 < len(rows) else None
                if ws[f"B{r}"].font.bold:
                    # a policy sub-heading followed by "not material" text is dropped together with it
                    if nxt and "not material and is not disclosed" in str(ws[f"B{nxt}"].value or ""):
                        i += 2
                        continue
                    p = para(doc, str(b), bold=True, space_after=2)
                    p.paragraph_format.keep_with_next = True
                else:
                    para(doc, str(b), space_after=6)
            i += 1
            continue
        # a block of table rows
        block = []
        while i < len(rows):
            rr = rows[i]
            v = [ws[f"{c}{rr}"].value for c in cols]
            if not any(x not in (None, "") for x in v):
                break
            block.append(rr)
            i += 1
        ncols = max(max((k for k, c in enumerate(cols) if ws[f"{c}{x}"].value not in (None, "")), default=0)
                    for x in block) + 1
        out, bold, totals, doubles, header = [], [], [], [], []
        for k, x in enumerate(block):
            label = ws[f"B{x}"].value
            row_vals = [fmt(ws[f"{c}{x}"].value, ws[f"{c}{x}"].number_format) for c in cols[:ncols]]
            out.append([str(label or "")] + row_vals)
            if label in (None, "") and k <= 1:
                header.append(k)
            if ws[f"B{x}"].font.bold:
                bold.append(k)
            bd = ws[f"C{x}"].border
            if bd.bottom and bd.bottom.style == "double":
                doubles.append(k)
            elif bd.top and bd.top.style == "thin":
                totals.append(k)
        label_w = Cm(17 - 2.5 * ncols) if ncols <= 4 else Cm(17 - 1.95 * ncols)
        val_w = Cm(2.5) if ncols <= 4 else Cm(1.95)
        table(doc, out, [label_w] + [val_w] * ncols, bold_rows=bold, total_rows=totals, double_rows=doubles,
              header_rows=tuple(header) or (), size=8 if ncols > 4 else 9)
        para(doc, space_after=4)


def page_numbers(doc, label):
    for section in doc.sections:
        p = section.footer.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run(p, label + "   |   Page ", size=8, color=GREY)
        r = p.add_run()
        for tag, text in (("begin", None), (None, "PAGE"), ("end", None)):
            if tag:
                el = OxmlElement("w:fldChar")
                el.set(qn("w:fldCharType"), tag)
            else:
                el = OxmlElement("w:instrText")
                el.set(qn("xml:space"), "preserve")
                el.text = text
            r._r.append(el)
        r.font.size, r.font.name = Pt(8), FONT


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("workbook")
    ap.add_argument("--out", required=True)
    ap.add_argument("--banner", default="")
    a = ap.parse_args()
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = load_workbook(a.workbook, data_only=True)
    if wb["SoFP"]["C8"].value is None and wb["SoFP"]["C10"].value is None:
        raise SystemExit("The workbook has no stored formula values. Recalculate it (open and save in Excel, or "
                         "run build_afs.py with LibreOffice installed) and try again.")
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(2)
    sec.top_margin = sec.bottom_margin = Cm(2)
    st = doc.styles["Normal"]
    st.font.name, st.font.size = FONT, Pt(10)
    cover(doc, wb, a.banner)
    directors_report(doc, wb)
    statement(doc, wb, "SoFP")
    statement(doc, wb, "SoCI")
    soce(doc, wb)
    statement(doc, wb, "SoCF")
    notes(doc, wb)
    E = entity(wb)
    page_numbers(doc, f"{E['name']} - financial statements {E['date'].year if hasattr(E['date'], 'year') else ''}")
    doc.save(a.out)
    print("Report:", a.out)


if __name__ == "__main__":
    main()
