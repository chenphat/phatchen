"""Shared helpers for the ISSB sustainability-reporting working-paper workbook."""
import datetime as dt
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell, xl_col_to_name

CLIENT = "Tonle Threads Manufacturing Plc. (TTM) - ILLUSTRATIVE SAMPLE CLIENT"
PERIOD = "Year ended 31 December 2025 (FY2025)"
C = dict(navy='#0B2545', blue='#1F5FBF', sky='#DCE9FB', teal='#13A89E', green='#2E9E5B',
         amber='#F2A900', orange='#F28C28', red='#D7263D', purple='#6C4AB6', grey='#F3F5F8',
         mid='#C9D1DC', white='#FFFFFF', ink='#1B2430', inp='#0000FF', link='#008000',
         yellow='#FFFF00', sec='#E3ECF7')


def A(r, c, absolute=True):
    return xl_rowcol_to_cell(r, c, absolute, absolute)


def ref(sheet, r, c):
    return f"'{sheet}'!{A(r, c)}"


def rng(sheet, c, r1, r2):
    return f"'{sheet}'!{A(r1, c)}:{A(r2, c)}"


class Book:
    def __init__(self, path):
        self.wb = xlsxwriter.Workbook(path)
        self._f = {}
        self.wps = []     # index records
        self.names = {}
        self.ad = {}      # free-form address store shared between parts

    def f(self, **k):
        key = tuple(sorted(k.items()))
        if key not in self._f:
            d = {'font_name': 'Arial', 'font_size': 10, 'valign': 'top'}
            d.update(k)
            self._f[key] = self.wb.add_format(d)
        return self._f[key]

    def name(self, nm, sheet, r, c):
        a = ref(sheet, r, c)
        self.wb.define_name(nm, '=' + a)
        self.names[nm] = a
        return a

    # ---- generic cell writer: blue = input, black = formula, green = cross-sheet link
    def put(self, ws, r, c, v, **k):
        base = {'text_wrap': True, 'border': 1, 'border_color': C['mid']}
        if isinstance(v, str) and v.startswith('='):
            base['font_color'] = C['link'] if '!' in v else '#000000'
            base.setdefault('num_format', '#,##0')
            base.update(k)
            ws.write_formula(r, c, v, self.f(**base))
        elif isinstance(v, bool):
            base.update(k)
            ws.write(r, c, v, self.f(**base))
        elif isinstance(v, (int, float)):
            base['font_color'] = C['inp']
            base['num_format'] = '#,##0' if isinstance(v, int) else '#,##0.00'
            base.update(k)
            ws.write_number(r, c, v, self.f(**base))
        elif isinstance(v, (dt.date, dt.datetime)):
            base['font_color'] = C['inp']
            base['num_format'] = 'dd-mmm-yyyy'
            base['align'] = 'left'
            base.update(k)
            ws.write_datetime(r, c, dt.datetime(v.year, v.month, v.day), self.f(**base))
        else:
            base.update(k)
            ws.write(r, c, '' if v is None else v, self.f(**base))

    def hdr(self, ws, r, headers, bg=None, c0=0):
        fm = self.f(bold=True, font_color='#FFFFFF', bg_color=bg or C['navy'], text_wrap=True,
                    border=1, border_color=C['mid'], valign='vcenter')
        for j, h in enumerate(headers):
            ws.write(r, c0 + j, h, fm)
        ws.set_row(r, 30)

    def table(self, ws, r, headers, rows, nf=None, inp=(), bg=None, c0=0, extra=None):
        """Write header + rows. Values may be callables f(excel_row_number)->formula.
        nf: {col: num_format}; inp: columns whose text is user input (blue)."""
        nf = nf or {}
        extra = extra or {}
        self.hdr(ws, r, headers, bg, c0)
        for i, row in enumerate(rows):
            rr = r + 1 + i
            for j, v in enumerate(row):
                if callable(v):
                    v = v(rr + 1)
                k = dict(extra.get(j, {}))
                if j in nf:
                    k['num_format'] = nf[j]
                if j in inp and isinstance(v, str) and not v.startswith('='):
                    k['font_color'] = C['inp']
                self.put(ws, rr, c0 + j, v, **k)
        return r + 1, r + len(rows)

    def sec(self, ws, r, text, ncols, c0=0):
        fm = self.f(bold=True, font_color=C['navy'], bg_color=C['sec'], border=1,
                    border_color=C['mid'], font_size=11, valign='vcenter')
        ws.merge_range(r, c0, r, c0 + ncols - 1, text, fm)
        ws.set_row(r, 20)

    def note(self, ws, r, text, ncols, c0=0, h=None):
        fm = self.f(italic=True, text_wrap=True, font_color='#44546A')
        ws.merge_range(r, c0, r, c0 + ncols - 1, text, fm)
        if h:
            ws.set_row(r, h)

    def sheet(self, name, wp, title, objective, issb, guide, widths, status='Reviewed',
              prep='Sok D.', rev='Keo V. (CFO)', tab=None, obj_h=32):
        ws = self.wb.add_worksheet(name)
        ws.hide_gridlines(2)
        n = len(widths)
        for j, w in enumerate(widths):
            ws.set_column(j, j, w)
        band = self.f(bold=True, font_color='#FFFFFF', bg_color=C['navy'], font_size=12, valign='vcenter')
        ws.merge_range(0, 0, 0, n - 1, CLIENT + "   |   Sustainability-reporting working papers (IFRS S1 & IFRS S2)", band)
        ws.set_row(0, 24)
        lab = self.f(bold=True, bg_color=C['grey'], border=1, border_color=C['mid'], valign='vcenter')
        val = self.f(text_wrap=True, border=1, border_color=C['mid'], valign='vcenter')
        big = self.f(bold=True, font_size=12, border=1, border_color=C['mid'], font_color=C['navy'], valign='vcenter')
        k, acc = 0, 0          # label area = as many leading columns as needed to reach ~16 chars
        while acc < 16 and k < n - 1:
            acc += widths[k]
            k += 1

        def label(r, text, fm):
            if k > 1:
                ws.merge_range(r, 0, r, k - 1, text, fm)
            else:
                ws.write(r, 0, text, fm)

        def value(r, text, fm):
            if n - k > 1:
                ws.merge_range(r, k, r, n - 1, text, fm)
            else:
                ws.write(r, k, text, fm)
        label(1, 'WP ref / title', lab)
        value(1, f"{wp}  |  {title}", big)
        label(2, 'Period / sign-off', lab)
        value(2, f"{PERIOD}   |   Prepared by: {prep}   |   Reviewed by: {rev}   |   WP status: {status}", val)
        label(3, 'ISSB reference', lab)
        value(3, issb, val)
        label(4, 'Guidance used', lab)
        value(4, guide, val)
        label(5, 'Objective', lab)
        value(5, objective, val)
        ws.set_row(5, obj_h)
        lnk = self.f(font_color=C['blue'], underline=1, bold=True)
        label(6, '', lnk)
        ws.write_formula(6, 0, '=HYPERLINK("#\'Index\'!A1","<< WP Index")', lnk)
        ws.write_formula(6, k, '=HYPERLINK("#\'Dashboard\'!A1","<< Dashboard")', lnk)
        if n - k > 2:
            ws.merge_range(6, k + 1, 6, n - 1,
                           'Legend:  blue font = input / hard-coded data (edit here)   |   black = formula   |   green = link from another working paper   |   yellow fill = key assumption',
                           self.f(italic=True, font_color='#44546A'))
        if tab:
            ws.set_tab_color(tab)
        ws.set_landscape()
        ws.fit_to_pages(1, 0)
        self.wps.append(dict(sheet=name, wp=wp, title=title, issb=issb, guide=guide, prep=prep, rev=rev, status=status))
        return ws, 8

    def dv(self, ws, r1, r2, c, items):
        ws.data_validation(r1, c, r2, c, {'validate': 'list', 'source': items})

    def cf_status(self, ws, r1, r2, c):
        """Traffic-light conditional format for common status words."""
        pal = [('Complete', '#C6EFCE', '#006100'), ('Reviewed', '#C6EFCE', '#006100'), ('Closed', '#C6EFCE', '#006100'),
               ('Yes', '#C6EFCE', '#006100'), ('On track', '#C6EFCE', '#006100'), ('Achieved', '#C6EFCE', '#006100'),
               ('In progress', '#FFEB9C', '#7F6000'), ('In review', '#FFEB9C', '#7F6000'), ('Partial', '#FFEB9C', '#7F6000'),
               ('Medium', '#FFEB9C', '#7F6000'), ('Behind', '#FFD9B3', '#9C4A00'), ('High', '#FFD9B3', '#9C4A00'),
               ('Open', '#FFD9B3', '#9C4A00'), ('Not started', '#FFC7CE', '#9C0006'), ('Overdue', '#FFC7CE', '#9C0006'),
               ('Critical', '#FFC7CE', '#9C0006'), ('No', '#FFC7CE', '#9C0006'), ('Low', '#DDEBF7', '#1F4E78'),
               ('N/A', '#EDEDED', '#595959')]
        for word, bg, fg in pal:
            ws.conditional_format(r1, c, r2, c, {'type': 'cell', 'criteria': '==', 'value': f'"{word}"',
                                                 'format': self.wb.add_format({'bg_color': bg, 'font_color': fg})})
