"""Shared Excel styling and formula helpers.

All statement figures are live formulas:
    GL / TB  ->  Mapping  ->  Lead (one row per AFS component)  ->  statements and notes
"""
import re

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

FONT = "Arial"
NUM = '#,##0;(#,##0);"-"'
NUM2 = '#,##0.00;(#,##0.00);"-"'
PCT = '0.0%;(0.0%);"-"'
DATE = "dd-mmm-yyyy"

F_NORMAL = Font(name=FONT, size=10)
F_BOLD = Font(name=FONT, size=10, bold=True)
F_ITALIC = Font(name=FONT, size=10, italic=True)
F_TITLE = Font(name=FONT, size=14, bold=True)
F_SUB = Font(name=FONT, size=11, bold=True)
F_INPUT = Font(name=FONT, size=10, color="0000FF")
F_LINK = Font(name=FONT, size=10, color="008000")
F_HEAD = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_GREY = Font(name=FONT, size=9, color="7F7F7F", italic=True)
FILL_HEAD = PatternFill("solid", fgColor="1F3864")
FILL_INPUT = PatternFill("solid", fgColor="FFFF00")
FILL_BAND = PatternFill("solid", fgColor="D9E1F2")
THIN, MED = Side(style="thin"), Side(style="medium")
B_TOP = Border(top=THIN)
B_TOTAL = Border(top=THIN, bottom=Side(style="double"))
B_BOTTOM = Border(bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")
RIGHT = Alignment(horizontal="right")
CENTER = Alignment(horizontal="center")

CY, PY = "CY", "PY"
LEADCOL = {("close", CY): "F", ("close", PY): "G", ("open", CY): "H", ("open", PY): "I",
           ("mvt", CY): "J", ("mvt", PY): "K"}
TBCOL = {CY: "O", PY: "I"}


def lead(kind, fy, leaves):
    """Sum of Lead balances (debit +) over AFS component codes."""
    c = LEADCOL[(kind, fy)]
    return "(" + "+".join(f'SUMIFS(Lead!${c}:${c},Lead!$A:$A,"{l}")' for l in leaves) + ")"


def face(kind, fy, code):
    c = LEADCOL[(kind, fy)]
    return f'SUMIFS(Lead!${c}:${c},Lead!$C:$C,"{code}")'


def glm(fy, leaves, types):
    """Sum of GL net movements (debit +) for the AFS components and movement types given."""
    return "(" + "+".join(
        f'SUMIFS(GL!$J:$J,GL!$K:$K,"{l}",GL!$L:$L,"{t}",GL!$A:$A,"{fy}")' for l in leaves for t in types) + ")"


def nat(fy, natures, funcs=None):
    """Sum of TB balances by expense nature (and optionally by function)."""
    c = TBCOL[fy]
    parts = []
    for n in natures:
        if funcs:
            parts += [f'SUMIFS(ATB!${c}:${c},ATB!$E:$E,"{n}",ATB!$D:$D,"{f}")' for f in funcs]
        else:
            parts.append(f'SUMIFS(ATB!${c}:${c},ATB!$E:$E,"{n}")')
    return "(" + "+".join(parts) + ")"


def widths(ws, spec):
    for i, w in enumerate(spec, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def header_row(ws, row, labels, start=1, height=30):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=start + i, value=lab)
        c.font, c.fill = F_HEAD, FILL_HEAD
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    ws.row_dimensions[row].height = height


def put(ws, ref, value, font=F_NORMAL, fmt=None, align=None, border=None, fill=None):
    c = ws[ref]
    c.value, c.font = value, font
    if fmt:
        c.number_format = fmt
    if align:
        c.alignment = align
    if border:
        c.border = border
    if fill:
        c.fill = fill
    return c


def sheet_title(ws, title, subtitle=None):
    put(ws, "A1", title, F_TITLE)
    if subtitle:
        put(ws, "A2", subtitle, F_GREY)


class Book:
    """Holds the workbook, the Entity input cell addresses and the row registry used for cross-references."""

    def __init__(self, wb):
        self.wb = wb
        self.ent = {}        # entity key -> absolute cell reference, e.g. Entity!$C$6
        self.rows = {}       # (sheet, key) -> row number
        self.notes = {}      # note key -> note number
        self.active = set()  # AFS components with balances
        self.checks = []     # (description, formula) collected while building

    def reg(self, sheet, key, row):
        self.rows[(sheet, key)] = row
        return row

    def cell(self, sheet, key, fy, cur="F"):
        """Statement cell. cur 'F' = functional currency columns (E, F); 'P' = presentation currency columns (C, D)."""
        cols = ("E", "F") if cur == "F" else ("C", "D")
        return f"{sheet}!{cols[0] if fy == CY else cols[1]}{self.rows[(sheet, key)]}"

    def rate(self, kind, fy):
        """Exchange rate cell: kind C closing, A average, O opening (the previous closing rate)."""
        key = {("C", CY): "fx_close_cy", ("C", PY): "fx_close_py", ("O", CY): "fx_close_py", ("O", PY): "fx_close_ppy",
               ("A", CY): "fx_avg_cy", ("A", PY): "fx_avg_py"}[(kind, fy)]
        return self.ent[key]

    def conv(self, kind, fy):
        """Multiplier that turns a functional-currency amount into presentation-currency units."""
        return f"*{self.rate(kind, fy)}/{self.ent['pres_scale']}"

    def to_pres(self, formula, force=None):
        """Translate every ledger term of a formula into the presentation currency (Section 30.18):
        balances at the closing rate, opening balances at the previous closing rate, income, expenses and
        movements at the average rate."""
        from .taxonomy import FACE, is_pl

        def lead_term(m):
            col, code = m.group(1), m.group(3)
            fy = CY if col in "FHJ" else PY
            if col in "HI":
                kind = "O"
            elif col in "JK":
                kind = "A"
            else:
                flow = is_pl(code) or code == "RE.DIV" or (code in FACE and FACE[code][2] == "SoCI")
                kind = "A" if flow else "C"
            return f"({m.group(0)}{self.conv(force or kind, fy)})"

        f = re.sub(r'SUMIFS\(Lead!\$([F-K]):\$\1,Lead!\$([AC]):\$\2,"([^"]+)"\)', lead_term, formula)
        f = re.sub(r'SUMIFS\(GL!\$J:\$J,GL!\$K:\$K,"[^"]+",GL!\$L:\$L,"[^"]+",GL!\$A:\$A,"(CY|PY)"\)',
                   lambda m: f"({m.group(0)}{self.conv(force or 'A', m.group(1))})", f)
        f = re.sub(r'SUMIFS\(ATB!\$([OI]):\$\1,[^()]*\)',
                   lambda m: f"({m.group(0)}{self.conv(force or 'A', CY if m.group(1) == 'O' else PY)})", f)
        return f

    def year(self, fy):
        return f"=YEAR({self.ent['reporting_date']})" + ("" if fy == CY else "-1")

    def check(self, description, formula):
        self.checks.append((description, formula))
