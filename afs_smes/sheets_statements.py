"""Primary statements: SoFP, SoCI, SoCE, SoCF. Every amount is a formula on Lead / TB / GL."""
from .taxonomy import (BORR, CASH_LEAVES, FACE, FACES, PPE_COST, PPE_DEP, SECTION_TITLES, leaves_of)
from .xl import (B_BOTTOM, B_TOP, B_TOTAL, CENTER, CY, F_BOLD, F_GREY, F_INPUT, F_ITALIC, F_NORMAL, F_SUB, F_TITLE, NUM, PY,
                 RIGHT, face, glm, lead, nat, put, widths)

NOTE_OF_FACE = {
    "PPE": "ppe", "INVPROP": "invprop", "INTANG": "intang", "INVASSOC": "invassoc", "OFA_NC": "ofa", "OFA_C": "ofa",
    "DTA": "dtax", "DTL": "dtax", "INV": "inv", "TOR": "tor", "CONTA": "revenue", "CONTL": "revenue", "CTA": "tax",
    "CTL": "tax", "CASH": "cash", "SC": "equity", "SP": "equity", "RESV": "equity", "BORR_NC": "borr",
    "BORR_C": "borr", "EBO_NC": "ebo", "PROV_NC": "prov", "PROV_C": "prov", "TOP": "top", "REV": "revenue",
    "COS": "expenses", "DIST": "expenses", "ADMIN": "expenses", "OTHEXP": "expenses", "OI": "otherinc",
    "FININC": "finance", "FINCOST": "finance", "TAX": "tax",
}
DT = ["DTL", "DTA"]
RESV = ["RESV.REVAL", "RESV.OTHER"]


def _heading(book, ws, title, period_text, last_col="D"):
    put(ws, "A1", f"={book.ent['entity_name']}", F_TITLE)
    put(ws, "A2", title, F_SUB)
    put(ws, "A3", period_text, F_NORMAL)
    put(ws, "A4", f'="Expressed in "&{book.ent["currency"]}', F_GREY)
    ws.sheet_view.showGridLines = False


def _year_heads(book, ws, row, note=True):
    if note:
        put(ws, f"B{row}", "Note", F_BOLD, align=CENTER, border=B_BOTTOM)
    put(ws, f"C{row}", book.year(CY), F_BOLD, "0", RIGHT, B_BOTTOM)
    put(ws, f"D{row}", book.year(PY), F_BOLD, "0", RIGHT, B_BOTTOM)
    put(ws, f"A{row}", None, border=B_BOTTOM)


def _as_at(book):
    return f'="As at "&TEXT({book.ent["reporting_date"]},"d mmmm yyyy")'


def _for_year(book):
    return f'="For the year ended "&TEXT({book.ent["reporting_date"]},"d mmmm yyyy")'


class Stmt:
    def __init__(self, book, ws, name):
        self.book, self.ws, self.name, self.r = book, ws, name, 6

    def line(self, key, caption, f_cy, f_py, note=None, bold=False, border=None, hidden=False, indent=0):
        r = self.r
        font = F_BOLD if bold else F_NORMAL
        c = put(self.ws, f"A{r}", caption, font)
        if indent:
            from openpyxl.styles import Alignment
            c.alignment = Alignment(indent=indent)
        if note and note in self.book.notes:
            put(self.ws, f"B{r}", self.book.notes[note], F_NORMAL, align=CENTER)
        put(self.ws, f"C{r}", "=" + f_cy, font, NUM, border=border)
        put(self.ws, f"D{r}", "=" + f_py, font, NUM, border=border)
        if hidden:
            self.ws.row_dimensions[r].hidden = True
        self.book.reg(self.name, key, r)
        self.r += 1
        return r

    def calc(self, key, caption, expr, **kw):
        """expr uses {c} as the column placeholder, e.g. '{c}7+{c}8'."""
        return self.line(key, caption, expr.format(c="C"), expr.format(c="D"), **kw)

    def text(self, caption, font=F_BOLD):
        put(self.ws, f"A{self.r}", caption, font)
        self.r += 1

    def gap(self):
        self.r += 1

    def row(self, key):
        return self.book.rows[(self.name, key)]


def build_soci(book, ws):
    _heading(book, ws, "Statement of comprehensive income", _for_year(book))
    widths(ws, [66, 7, 16, 16])
    _year_heads(book, ws, 5)
    s = Stmt(book, ws, "SoCI")
    act = book.active_faces

    def pl(code, hidden=None, bold=False):
        return s.line(code, FACE[code][1], "-" + face("close", CY, code), "-" + face("close", PY, code),
                      note=NOTE_OF_FACE.get(code), hidden=(code not in act) if hidden is None else hidden, bold=bold)

    rev, cos = pl("REV", hidden=False), pl("COS", hidden=False)
    gp = s.calc("GP", "Gross profit", f"{{c}}{rev}+{{c}}{cos}", bold=True, border=B_TOP)
    first = pl("OI")
    pl("DIST"); pl("ADMIN")
    last = pl("OTHEXP")
    op = s.calc("OP", "Operating profit", f"{{c}}{gp}+SUM({{c}}{first}:{{c}}{last})", bold=True, border=B_TOP)
    f1 = pl("FININC"); pl("FINCOST")
    f2 = pl("SHAREASSOC")
    pbt = s.calc("PBT", "Profit before tax", f"{{c}}{op}+SUM({{c}}{f1}:{{c}}{f2})", bold=True, border=B_TOP)
    tax = pl("TAX", hidden=False)
    profit = s.calc("PROFIT", "Profit for the year", f"{{c}}{pbt}+{{c}}{tax}", bold=True, border=B_TOTAL)
    s.gap()
    s.text("Other comprehensive income")
    s.text("Items that will not be reclassified subsequently to profit or loss", F_ITALIC)
    o1 = s.line("OCI_REVAL", "Gain on revaluation of property, plant and equipment",
                f"-{glm(CY, ['RESV.REVAL'], ['OCI'])}-{glm(CY, DT, ['OCI'])}",
                f"-{glm(PY, ['RESV.REVAL'], ['OCI'])}-{glm(PY, DT, ['OCI'])}", note="ppe")
    s.line("OCI_OTHER", "Other items recognised in other comprehensive income",
           "-" + glm(CY, ["RESV.OTHER"], ["OCI"]), "-" + glm(PY, ["RESV.OTHER"], ["OCI"]),
           hidden="RESV.OTHER" not in book.active)
    o2 = s.line("OCI_TAX", "Income tax relating to components of other comprehensive income",
                glm(CY, DT, ["OCI"]), glm(PY, DT, ["OCI"]), note="dtax")
    oci = s.calc("OCI", "Other comprehensive income for the year, net of tax", f"SUM({{c}}{o1}:{{c}}{o2})",
                 bold=True, border=B_TOP)
    s.gap()
    s.calc("TCI", "Total comprehensive income for the year", f"{{c}}{profit}+{{c}}{oci}", bold=True, border=B_TOTAL)
    s.gap()
    put(ws, f"A{s.r}", "Expenses are analysed by function; the analysis by nature is given in the notes. "
                       "All amounts are attributable to the owners of the entity; there are no discontinued operations.", F_GREY)
    ws.freeze_panes = "A6"


def build_sofp(book, ws):
    _heading(book, ws, "Statement of financial position", _as_at(book))
    widths(ws, [66, 7, 16, 16])
    _year_heads(book, ws, 5)
    s = Stmt(book, ws, "SoFP")
    act = book.active_faces
    tot = {}

    def section(sec, sign):
        s.text(SECTION_TITLES[sec])
        start = s.r
        for code, caption, stmt, sc in FACES:
            if stmt != "SoFP" or sc != sec:
                continue
            if code == "RE":
                f = lambda fy: (f"-{lead('close', fy, ['RE', 'RE.DIV'])}+{book.cell('SoCI', 'PROFIT', fy)}")
                s.line(code, caption, f(CY), f(PY), hidden=False)
            else:
                s.line(code, caption, sign + face("close", CY, code), sign + face("close", PY, code),
                       note=NOTE_OF_FACE.get(code), hidden=code not in act)
        tot[sec] = s.calc("T_" + sec, "Total " + SECTION_TITLES[sec].lower(),
                          f"SUM({{c}}{start}:{{c}}{s.r - 1})", bold=True, border=B_TOP)
        s.gap()

    s.text("ASSETS", F_SUB)
    section("NCA", "")
    section("CA", "")
    ta = s.calc("TA", "Total assets", f"{{c}}{tot['NCA']}+{{c}}{tot['CA']}", bold=True, border=B_TOTAL)
    s.gap()
    s.text("EQUITY AND LIABILITIES", F_SUB)
    section("EQ", "-")
    section("NCL", "-")
    section("CL", "-")
    tl = s.calc("TL", "Total liabilities", f"{{c}}{tot['NCL']}+{{c}}{tot['CL']}", bold=True, border=B_TOP)
    tel = s.calc("TEL", "Total equity and liabilities", f"{{c}}{tot['EQ']}+{{c}}{tl}", bold=True, border=B_TOTAL)
    s.gap()
    put(ws, f"A{s.r}", f'="The financial statements were authorised for issue by "&{book.ent["authorised_by"]}'
                       f'&" on "&TEXT({book.ent["authorised_date"]},"d mmmm yyyy")&" and signed on its behalf by:"', F_NORMAL)
    put(ws, f"A{s.r + 3}", "______________________________          ______________________________", F_NORMAL)
    put(ws, f"A{s.r + 4}", "Director                                                          Director", F_NORMAL)
    for fy in (CY, PY):
        book.check(f"SoFP {fy}: total assets = total equity and liabilities",
                   f"{book.cell('SoFP', 'TA', fy)}-{book.cell('SoFP', 'TEL', fy)}")
    ws.freeze_panes = "A6"


def build_soce(book, ws):
    _heading(book, ws, "Statement of changes in equity", _for_year(book))
    widths(ws, [58, 16, 16, 16, 16, 16])
    for col, lab in zip("BCDEF", ["Share capital", "Share premium", "Other reserves", "Retained earnings", "Total equity"]):
        put(ws, f"{col}5", lab, F_BOLD, align=RIGHT, border=B_BOTTOM)
    put(ws, "A5", None, border=B_BOTTOM)
    put(ws, "A6", "Note", F_GREY)
    for col in "BCD":
        put(ws, f"{col}6", book.notes.get("equity"), F_GREY, align=RIGHT)
    r = 7
    prev_close = None

    def row(caption, vals, bold=False, border=None, key=None):
        nonlocal r
        put(ws, f"A{r}", caption, F_BOLD if bold else F_NORMAL)
        for col, v in zip("BCDE", vals):
            put(ws, f"{col}{r}", ("=" + v) if isinstance(v, str) else v, F_BOLD if bold else F_NORMAL, NUM, border=border)
        put(ws, f"F{r}", f"=SUM(B{r}:E{r})", F_BOLD if bold else F_NORMAL, NUM, border=border)
        if key:
            book.reg("SoCE", key, r)
        r += 1
        return r - 1

    for fy in (PY, CY):
        yr = f'="Balance at 1 January "&TEXT(YEAR({book.ent["reporting_date"]}){"-1" if fy == PY else ""},"0")'
        if prev_close is None:
            o = row(None, [f"-{lead('open', fy, ['SC'])}", f"-{lead('open', fy, ['SP'])}",
                           f"-{lead('open', fy, RESV)}", f"-{lead('open', fy, ['RE', 'RE.DIV'])}"], bold=True,
                    key="OPEN_" + fy)
        else:
            o = row(None, [f"{c}{prev_close}" for c in "BCDE"], bold=True, key="OPEN_" + fy)
        put(ws, f"A{o}", yr, F_BOLD)
        adj = row("Changes in accounting policies and corrections of prior period errors", [0, 0, 0, 0])
        for col in "BCDE":
            ws[f"{col}{adj}"].font = F_INPUT
        p = row("Profit for the year", ["0", "0", "0", book.cell("SoCI", "PROFIT", fy)])
        oc = row("Other comprehensive income for the year", ["0", "0", book.cell("SoCI", "OCI", fy), "0"])
        t = row("Total comprehensive income for the year", [f"{c}{p}+{c}{oc}" for c in "BCDE"], bold=True, border=B_TOP)
        i = row("Issue of shares", [f"-{lead('mvt', fy, ['SC'])}", f"-{lead('mvt', fy, ['SP'])}", "0", "0"])
        d = row("Dividends", ["0", "0", "0", f"-({lead('close', fy, ['RE.DIV'])}+{glm(fy, ['RE'], ['DIVIDEND'])})"])
        x = row("Transfers between reserves", ["0", "0", f"-{glm(fy, RESV, ['TRANSFER'])}",
                                                f"-{glm(fy, ['RE'], ['TRANSFER', 'OTHER'])}"])
        c_ = row(None, [f"{c}{o}+{c}{adj}+{c}{t}+{c}{i}+{c}{d}+{c}{x}" for c in "BCDE"], bold=True, border=B_TOTAL,
                 key="CLOSE_" + fy)
        put(ws, f"A{c_}", f'="Balance at 31 December "&TEXT(YEAR({book.ent["reporting_date"]}){"-1" if fy == PY else ""},"0")', F_BOLD)
        prev_close = c_
        r += 1
        book.check(f"SoCE {fy}: closing equity = total equity in the SoFP",
                   f"SoCE!F{c_}-{book.cell('SoFP', 'T_EQ', fy)}")
        book.check(f"SoCE {fy}: closing reserves and retained earnings agree to the SoFP by component",
                   f"ABS(SoCE!D{c_}-{book.cell('SoFP', 'RESV', fy)})+ABS(SoCE!E{c_}-{book.cell('SoFP', 'RE', fy)})")
    book.check("SoCE CY: opening equity (= PY closing) agrees to the current-year opening trial balance",
               f"SoCE!F{book.rows[('SoCE', 'OPEN_CY')]}+{lead('open', CY, ['SC', 'SP', 'RE', 'RE.DIV'] + RESV)}")
    ws.freeze_panes = "A7"


def build_socf(book, ws):
    _heading(book, ws, "Statement of cash flows", _for_year(book))
    widths(ws, [66, 7, 16, 16])
    _year_heads(book, ws, 5)
    s = Stmt(book, ws, "SoCF")
    A = book.active
    M = lambda fy, lv: lead("mvt", fy, lv)
    P = lambda fy, lv: lead("close", fy, lv)

    def ln(key, caption, fn, note=None, hidden=False):
        return s.line(key, caption, fn(CY), fn(PY), note=note, hidden=hidden)

    def has(*leaves):
        return any(l in A for l in leaves)

    s.text("Cash flows from operating activities")
    pbt = ln("PBT", "Profit before tax", lambda fy: book.cell("SoCI", "PBT", fy))
    s.text("Adjustments for:", F_ITALIC)
    a1 = ln("DEP", "Depreciation of property, plant and equipment", lambda fy: nat(fy, ["DEP"]), note="ppe")
    ln("AMORT", "Amortisation of intangible assets", lambda fy: nat(fy, ["AMORT"]), note="intang",
       hidden=not has("INTANG.C"))
    ln("FINCOST", "Finance costs", lambda fy: P(fy, ["FINCOST"]), note="finance")
    ln("FININC", "Finance income", lambda fy: P(fy, ["FININC"]), note="finance", hidden=not has("FININC"))
    ln("DISP", "(Gain)/loss on disposal of property, plant and equipment", lambda fy: P(fy, ["OI.DISP"]),
       hidden=not has("OI.DISP"))
    ln("FV", "Fair value (gain)/loss on investment property", lambda fy: P(fy, ["OI.FV"]), hidden=not has("OI.FV"))
    ln("ASSOC", "Share of profit of associates and jointly controlled entities", lambda fy: P(fy, ["SHAREASSOC"]),
       hidden=not has("SHAREASSOC"))
    a2 = ln("PROV", "Increase/(decrease) in provisions and employee benefit obligations",
            lambda fy: "-" + M(fy, ["PROV_NC", "PROV_C", "EBO_NC"]), note="prov")
    sub = s.calc("OPWC", "Operating cash flows before movements in working capital",
                 f"{{c}}{pbt}+SUM({{c}}{a1}:{{c}}{a2})", bold=True, border=B_TOP)
    w1 = ln("INV", "(Increase)/decrease in inventories", lambda fy: "-" + M(fy, leaves_of("INV")))
    ln("REC", "(Increase)/decrease in trade and other receivables",
       lambda fy: "-" + M(fy, ["TOR.TRADE", "TOR.ALLOW", "TOR.OTHER", "TOR.PREPAY"]))
    ln("CONTA", "(Increase)/decrease in contract assets", lambda fy: "-" + M(fy, ["CONTA"]), hidden=not has("CONTA"))
    ln("PAY", "Increase/(decrease) in trade and other payables",
       lambda fy: "-" + M(fy, ["TOP.TRADE", "TOP.ACCR", "TOP.OTHER"]))
    w2 = ln("CONTL", "Increase/(decrease) in contract liabilities", lambda fy: "-" + M(fy, ["CONTL"]),
            hidden=not has("CONTL"))
    cgo = s.calc("CGO", "Cash generated from operations", f"{{c}}{sub}+SUM({{c}}{w1}:{{c}}{w2})", bold=True, border=B_TOP)
    ip = ln("INTPAID", "Interest paid", lambda fy: f"-{P(fy, ['FINCOST'])}-{M(fy, ['TOP.INT'])}")
    tp = ln("TAXPAID", "Income tax paid",
            lambda fy: f"-{P(fy, ['TAX.CUR', 'TAX.DEF'])}-{M(fy, ['CTL', 'CTA'] + DT)}+{glm(fy, DT, ['OCI'])}")
    op = s.calc("OPER", "Net cash from operating activities", f"{{c}}{cgo}+{{c}}{ip}+{{c}}{tp}", bold=True, border=B_TOP)
    s.gap()
    s.text("Cash flows from investing activities")
    i1 = ln("PPE_BUY", "Purchase of property, plant and equipment", lambda fy: "-" + glm(fy, PPE_COST, ["ADD"]), note="ppe")
    ln("PPE_SELL", "Proceeds from disposal of property, plant and equipment",
       lambda fy: f"-{glm(fy, PPE_COST + PPE_DEP, ['DISP'])}-{P(fy, ['OI.DISP'])}", hidden=not has("OI.DISP"))
    ln("INT_BUY", "Purchase of intangible assets", lambda fy: "-" + glm(fy, ["INTANG.C"], ["ADD"]), note="intang",
       hidden=not has("INTANG.C"))
    ln("IP", "Net cash flows on investment property", lambda fy: f"-{M(fy, ['INVPROP'])}-{P(fy, ['OI.FV'])}",
       hidden=not has("INVPROP"))
    ln("ASSOC_CF", "Net cash flows on investments in associates (acquisitions less dividends received)",
       lambda fy: f"-{M(fy, ['INVASSOC'])}-{P(fy, ['SHAREASSOC'])}", hidden=not has("INVASSOC"))
    ln("OFA", "Net (purchase)/disposal of other financial assets", lambda fy: "-" + M(fy, ["OFA_NC", "OFA_C"]),
       hidden=not has("OFA_NC", "OFA_C"))
    i2 = ln("INTREC", "Interest received", lambda fy: f"-{P(fy, ['FININC'])}-{M(fy, ['TOR.INT'])}",
            hidden=not has("FININC"))
    inv = s.calc("INVEST", "Net cash used in investing activities", f"SUM({{c}}{i1}:{{c}}{i2})", bold=True, border=B_TOP)
    s.gap()
    s.text("Cash flows from financing activities")
    f1 = ln("SHARES", "Proceeds from issue of shares", lambda fy: "-" + M(fy, ["SC", "SP"]), note="equity")
    ln("BORR_IN", "Proceeds from borrowings", lambda fy: "-" + glm(fy, BORR, ["PROCEEDS"]), note="borr")
    ln("BORR_OUT", "Repayment of borrowings", lambda fy: "-" + glm(fy, BORR, ["REPAY"]), note="borr")
    f2 = ln("DIVPAID", "Dividends paid",
            lambda fy: f"-{P(fy, ['RE.DIV'])}-{glm(fy, ['RE'], ['DIVIDEND'])}-{M(fy, ['TOP.DIV'])}", note="dividends")
    fin = s.calc("FIN", "Net cash from/(used in) financing activities", f"SUM({{c}}{f1}:{{c}}{f2})", bold=True, border=B_TOP)
    s.gap()
    net = s.calc("NET", "Net increase/(decrease) in cash and cash equivalents", f"{{c}}{op}+{{c}}{inv}+{{c}}{fin}", bold=True)
    opn = ln("OPEN", "Cash and cash equivalents at beginning of year", lambda fy: lead("open", fy, CASH_LEAVES))
    s.calc("CLOSE", "Cash and cash equivalents at end of year", f"{{c}}{net}+{{c}}{opn}", bold=True, border=B_TOTAL,
           note="cash")
    s.gap()
    put(ws, f"A{s.r}", "Indirect method. Cash and cash equivalents comprise cash on hand, bank balances and short-term deposits, "
                       "less bank overdrafts repayable on demand.", F_GREY)
    for fy in (CY, PY):
        book.check(f"SoCF {fy}: closing cash per cash flow statement = cash and cash equivalents per ledger (net of overdrafts)",
                   f"{book.cell('SoCF', 'CLOSE', fy)}-{lead('close', fy, CASH_LEAVES)}")
    ws.freeze_panes = "A6"
