"""Primary statements: SoFP, SoCI, SoCE, SoCF. Every amount is a formula on Lead / TB / GL."""
from .taxonomy import (BORR, CASH_LEAVES, EXPENSE_FUNCTIONS, FACE, FACES, NATURES, PPE_COST, PPE_DEP, SECTION_TITLES,
                       leaves_of)
from openpyxl.styles import Font

from .xl import (B_BOTTOM, B_TOP, B_TOTAL, CENTER, CY, FONT, F_BOLD, F_GREY, F_INPUT, F_ITALIC, F_NORMAL, F_SUB, F_TITLE, NUM, PY,
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


def _heading(book, ws, title, period_text):
    put(ws, "A1", f"={book.ent['entity_name']}", F_TITLE)
    put(ws, "A2", title, F_SUB)
    put(ws, "A3", period_text, F_NORMAL)
    put(ws, "A4", f'="Presentation currency: "&{book.ent["pres_label"]}&". Functional currency: "&{book.ent["currency"]}'
                  f'&" (shown for information)."', F_GREY)
    ws.sheet_view.showGridLines = False


def _year_heads(book, ws, note=True):
    """Rows 5-6: currency labels and years. C:D presentation currency, E:F functional currency."""
    put(ws, "C5", f"={book.ent['pres_label']}", F_BOLD, align=RIGHT)
    put(ws, "D5", f"={book.ent['pres_label']}", F_BOLD, align=RIGHT)
    put(ws, "E5", f"={book.ent['currency']}", F_GREY, align=RIGHT)
    put(ws, "F5", f"={book.ent['currency']}", F_GREY, align=RIGHT)
    if note:
        put(ws, "B6", "Note", F_BOLD, align=CENTER, border=B_BOTTOM)
    put(ws, "A6", None, border=B_BOTTOM)
    for col, fy, font in (("C", CY, F_BOLD), ("D", PY, F_BOLD), ("E", CY, F_GREY), ("F", PY, F_GREY)):
        put(ws, f"{col}6", book.year(fy), font, "0", RIGHT, B_BOTTOM)
    widths(ws, [64, 7, 17, 17, 15, 15])
    ws.freeze_panes = "A7"


def _as_at(book):
    return f'="As at "&TEXT({book.ent["reporting_date"]},"d mmmm yyyy")'


def _for_year(book):
    return f'="For the year ended "&TEXT({book.ent["reporting_date"]},"d mmmm yyyy")'


class Stmt:
    """Statement writer. Functional-currency formulas go to E:F; C:D translate them at the row's rate
    (C closing, A average, O opening) unless explicit presentation-currency formulas are given."""

    def __init__(self, book, ws, name):
        self.book, self.ws, self.name, self.r = book, ws, name, 7

    def line(self, key, caption, f_cy, f_py, rate="C", note=None, bold=False, border=None, hidden=False, pres=None):
        r, b = self.r, self.book
        font = F_BOLD if bold else F_NORMAL
        put(self.ws, f"A{r}", caption, font)
        if note and note in b.notes:
            put(self.ws, f"B{r}", b.notes[note], F_NORMAL, align=CENTER)
        grey = Font(name=FONT, size=10, bold=bold, color="595959")
        put(self.ws, f"E{r}", ("=" + f_cy) if f_cy is not None else 0, grey, NUM, border=border)
        put(self.ws, f"F{r}", ("=" + f_py) if f_py is not None else 0, grey, NUM, border=border)
        if pres is None:
            pres = (f"E{r}{b.conv(rate, CY)}", f"F{r}{b.conv(rate, PY)}")
        put(self.ws, f"C{r}", "=" + pres[0], font, NUM, border=border)
        put(self.ws, f"D{r}", "=" + pres[1], font, NUM, border=border)
        if hidden:
            self.ws.row_dimensions[r].hidden = True
        b.reg(self.name, key, r)
        self.r += 1
        return r

    def calc(self, key, caption, expr, **kw):
        """Subtotal in all four columns; expr uses {c} as the column placeholder."""
        return self.line(key, caption, expr.format(c="E"), expr.format(c="F"),
                         pres=(expr.format(c="C"), expr.format(c="D")), **kw)

    def text(self, caption, font=F_BOLD):
        put(self.ws, f"A{self.r}", caption, font)
        self.r += 1

    def gap(self):
        self.r += 1


def build_soci(book, ws):
    _heading(book, ws, "Statement of comprehensive income", _for_year(book))
    _year_heads(book, ws)
    s = Stmt(book, ws, "SoCI")
    act = book.active_faces

    def pl(code, hidden=None):
        return s.line(code, FACE[code][1], "-" + face("close", CY, code), "-" + face("close", PY, code), rate="A",
                      note=NOTE_OF_FACE.get(code), hidden=(code not in act) if hidden is None else hidden)

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
                f"-{glm(PY, ['RESV.REVAL'], ['OCI'])}-{glm(PY, DT, ['OCI'])}", rate="A", note="ppe")
    s.line("OCI_OTHER", "Other items recognised in other comprehensive income",
           "-" + glm(CY, ["RESV.OTHER"], ["OCI"]), "-" + glm(PY, ["RESV.OTHER"], ["OCI"]), rate="A",
           hidden="RESV.OTHER" not in book.active)
    s.line("OCI_TAX", "Income tax relating to components of other comprehensive income",
           glm(CY, DT, ["OCI"]), glm(PY, DT, ["OCI"]), rate="A", note="dtax")
    # presentation-currency formulas are filled in by build_soce once the translation reserve rows exist
    o3 = s.line("OCI_FX", "Exchange differences on translation to the presentation currency", None, None,
                pres=("0", "0"), note="basis")
    oci = s.calc("OCI", "Other comprehensive income for the year, net of tax", f"SUM({{c}}{o1}:{{c}}{o3})",
                 bold=True, border=B_TOP)
    s.gap()
    s.calc("TCI", "Total comprehensive income for the year", f"{{c}}{profit}+{{c}}{oci}", bold=True, border=B_TOTAL)
    s.gap()
    put(ws, f"A{s.r}", "Expenses are analysed by function; the analysis by nature is given in the notes. "
                       "All amounts are attributable to the owners of the entity; there are no discontinued operations.", F_GREY)


def build_sofp(book, ws):
    _heading(book, ws, "Statement of financial position", _as_at(book))
    _year_heads(book, ws)
    s = Stmt(book, ws, "SoFP")
    act, E = book.active_faces, book.ent
    scale = E["pres_scale"]
    tot = {}

    def equity_line(code, caption):
        leaves = ["RE", "RE.DIV"] if code == "RE" else leaves_of(code)
        hist = E["fx_hist_capital"] if code in ("SC", "SP") else E["fx_hist_re"]
        usd = lambda fy: f"-{lead('close', fy, leaves)}" + (f"+{book.cell('SoCI', 'PROFIT', fy)}" if code == "RE" else "")
        r = s.r
        profit = lambda fy: f"+{book.cell('SoCI', 'PROFIT', fy, 'P')}" if code == "RE" else ""
        # historical-rate roll forward: opening at the historical rate, each year's movement at that year's average rate
        p_py = (f"(-{lead('open', PY, leaves)}*{hist}-{lead('mvt', PY, leaves)}*{E['fx_avg_py']})/{scale}" + profit(PY))
        p_cy = f"D{r}-{lead('mvt', CY, leaves)}*{E['fx_avg_cy']}/{scale}" + profit(CY)
        s.line(code, caption, usd(CY), usd(PY), pres=(p_cy, p_py), note=NOTE_OF_FACE.get(code),
               hidden=code not in act and code != "RE")

    def section(sec, sign):
        s.text(SECTION_TITLES[sec])
        start = s.r
        for code, caption, stmt, sc in FACES:
            if stmt != "SoFP" or sc != sec:
                continue
            if sec == "EQ":
                equity_line(code, caption)
                if code == "RESV":
                    s.line("TR", "Foreign currency translation reserve", None, None, pres=("0", "0"), note="equity")
            else:
                s.line(code, caption, sign + face("close", CY, code), sign + face("close", PY, code), rate="C",
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
    s.calc("TEL", "Total equity and liabilities", f"{{c}}{tot['EQ']}+{{c}}{tl}", bold=True, border=B_TOTAL)
    # translation reserve = net assets at the closing rate less the other components of equity (paragraph 30.18(c))
    tr = book.rows[("SoFP", "TR")]
    others = "+".join(f"{{c}}{book.rows[('SoFP', k)]}" for k in ("SC", "SP", "RESV", "RE"))
    for col in "CD":
        ws[f"{col}{tr}"].value = f"={col}{ta}-{col}{tl}-(" + others.format(c=col) + ")"
    s.gap()
    put(ws, f"A{s.r}", f'="The financial statements were authorised for issue by "&{E["authorised_by"]}'
                       f'&" on "&TEXT({E["authorised_date"]},"d mmmm yyyy")&" and signed on its behalf by:"', F_NORMAL)
    put(ws, f"A{s.r + 3}", "______________________________          ______________________________", F_NORMAL)
    put(ws, f"A{s.r + 4}", "Director                                                          Director", F_NORMAL)
    for fy in (CY, PY):
        book.check(f"SoFP {fy}: total assets = total equity and liabilities (functional currency)",
                   f"{book.cell('SoFP', 'TA', fy)}-{book.cell('SoFP', 'TEL', fy)}")


def build_soce(book, ws):
    _heading(book, ws, "Statement of changes in equity", _for_year(book))
    widths(ws, [56, 16, 16, 16, 16, 16, 16])
    ws.freeze_panes = "B5"
    E = book.ent
    scale = E["pres_scale"]
    ALLEQ = ["SC", "SP", "RE", "RE.DIV"] + RESV
    state = {"r": 5}

    def row(caption, vals, bold=False, border=None, key=None, inp=False):
        r = state["r"]
        font = F_BOLD if bold else F_NORMAL
        put(ws, f"A{r}", caption, font)
        for col, v in zip("BCDEF", vals):
            put(ws, f"{col}{r}", ("=" + v) if isinstance(v, str) else v, F_INPUT if inp else font, NUM, border=border)
        put(ws, f"G{r}", f"=SUM(B{r}:F{r})", font, NUM, border=border)
        if key:
            book.reg("SoCE", key, r)
        state["r"] += 1
        return r

    def block(cur):
        P = cur == "P"
        r = state["r"]
        put(ws, f"A{r}", (f'="Presentation currency - "&{E["pres_label"]}' if P else
                          f'="Functional currency - "&{E["currency"]}&" (for information)"'), F_SUB)
        r += 1
        for col, lab in zip("BCDEFG", ["Share capital", "Share premium", "Other reserves", "Translation reserve",
                                       "Retained earnings", "Total equity"]):
            put(ws, f"{col}{r}", lab, F_BOLD, align=RIGHT, border=B_BOTTOM)
        put(ws, f"A{r}", None, border=B_BOTTOM)
        state["r"] = r + 1
        prev = None
        for fy in (PY, CY):
            avg = f"*{E['fx_avg_cy' if fy == CY else 'fx_avg_py']}/{scale}" if P else ""
            sc = lambda key: book.cell("SoCI", key, fy, cur)
            yr = f'TEXT(YEAR({E["reporting_date"]}){"-1" if fy == PY else ""},"0")'
            if prev is None:
                if P:
                    o = state["r"]
                    vals = [f"-{lead('open', fy, ['SC'])}*{E['fx_hist_capital']}/{scale}",
                            f"-{lead('open', fy, ['SP'])}*{E['fx_hist_capital']}/{scale}",
                            f"-{lead('open', fy, RESV)}*{E['fx_hist_re']}/{scale}",
                            f"-{lead('open', fy, ALLEQ)}*{E['fx_close_ppy']}/{scale}-(B{o}+C{o}+D{o}+F{o})",
                            f"-{lead('open', fy, ['RE', 'RE.DIV'])}*{E['fx_hist_re']}/{scale}"]
                else:
                    vals = [f"-{lead('open', fy, ['SC'])}", f"-{lead('open', fy, ['SP'])}", f"-{lead('open', fy, RESV)}",
                            "0", f"-{lead('open', fy, ['RE', 'RE.DIV'])}"]
                o = row(None, vals, bold=True, key=f"OPEN_{fy}_{cur}")
            else:
                o = row(None, [f"{c}{prev}" for c in "BCDEF"], bold=True, key=f"OPEN_{fy}_{cur}")
            put(ws, f"A{o}", f'="Balance at 1 January "&{yr}', F_BOLD)
            adj = row("Changes in accounting policies and corrections of prior period errors", [0, 0, 0, 0, 0], inp=True)
            p = row("Profit for the year", ["0", "0", "0", "0", sc("PROFIT")])
            oc = row("Other comprehensive income for the year",
                     ["0", "0", f"{sc('OCI')}-{sc('OCI_FX')}", sc("OCI_FX"), "0"])
            t = row("Total comprehensive income for the year", [f"{c}{p}+{c}{oc}" for c in "BCDEF"], bold=True, border=B_TOP)
            i = row("Issue of shares", [f"-{lead('mvt', fy, ['SC'])}{avg}", f"-{lead('mvt', fy, ['SP'])}{avg}", "0", "0", "0"])
            d = row("Dividends", ["0", "0", "0", "0",
                                  f"-({lead('close', fy, ['RE.DIV'])}+{glm(fy, ['RE'], ['DIVIDEND'])}){avg}"])
            x = row("Transfers between reserves", ["0", "0", f"-{glm(fy, RESV, ['TRANSFER'])}{avg}", "0",
                                                    f"-{glm(fy, ['RE'], ['TRANSFER', 'OTHER'])}{avg}"])
            c_ = row(None, [f"{c}{o}+{c}{adj}+{c}{t}+{c}{i}+{c}{d}+{c}{x}" for c in "BCDEF"], bold=True, border=B_TOTAL,
                     key=f"CLOSE_{fy}_{cur}")
            put(ws, f"A{c_}", f'="Balance at 31 December "&{yr}', F_BOLD)
            prev = c_
            state["r"] += 1
            name = "presentation currency" if P else "functional currency"
            book.check(f"SoCE {fy} ({name}): closing equity = total equity in the SoFP",
                       f"SoCE!G{c_}-{book.cell('SoFP', 'T_EQ', fy, cur)}")
            book.check(f"SoCE {fy} ({name}): reserves and retained earnings agree to the SoFP by component",
                       f"ABS(SoCE!D{c_}-{book.cell('SoFP', 'RESV', fy, cur)})+ABS(SoCE!E{c_}-{book.cell('SoFP', 'TR', fy, cur)})"
                       f"+ABS(SoCE!F{c_}-{book.cell('SoFP', 'RE', fy, cur)})")
        state["r"] += 1

    block("P")
    block("F")
    book.check("SoCE CY: opening equity (= PY closing) agrees to the current-year opening trial balance",
               f"SoCE!G{book.rows[('SoCE', 'OPEN_CY_F')]}+{lead('open', CY, ALLEQ)}")
    # exchange differences on translation in OCI = movement in the translation reserve
    soci = book.wb["SoCI"]
    fx, tr = book.rows[("SoCI", "OCI_FX")], book.rows[("SoFP", "TR")]
    soci[f"C{fx}"].value = f"=SoFP!C{tr}-SoFP!D{tr}"
    soci[f"D{fx}"].value = f"=SoFP!D{tr}-SoCE!E{book.rows[('SoCE', 'OPEN_PY_P')]}"
    # proof (paragraph 30.20): opening net assets x change in closing rate + movement in net assets x (closing - average)
    na = lambda fy: book.cell("SoFP", "T_EQ", fy)
    book.check("Translation CY: exchange difference in OCI agrees to the proof by rates (paragraph 30.20)",
               f"SoCI!C{fx}-({na(PY)}*({E['fx_close_cy']}-{E['fx_close_py']})"
               f"+({na(CY)}-{na(PY)})*({E['fx_close_cy']}-{E['fx_avg_cy']}))/{scale}")
    opn = f"SoCE!G{book.rows[('SoCE', 'OPEN_PY_F')]}"
    book.check("Translation PY: exchange difference in OCI agrees to the proof by rates (paragraph 30.20)",
               f"SoCI!D{fx}-({opn}*({E['fx_close_py']}-{E['fx_close_ppy']})"
               f"+({na(PY)}-{opn})*({E['fx_close_py']}-{E['fx_avg_py']}))/{scale}")


def build_socf(book, ws):
    _heading(book, ws, "Statement of cash flows", _for_year(book))
    _year_heads(book, ws)
    s = Stmt(book, ws, "SoCF")
    A = book.active
    M = lambda fy, lv: lead("mvt", fy, lv)
    P = lambda fy, lv: lead("close", fy, lv)

    def ln(key, caption, fn, note=None, hidden=False, rate="A"):
        return s.line(key, caption, fn(CY), fn(PY), rate=rate, note=note, hidden=hidden)

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
       lambda fy: "-" + M(fy, ["TOR.TRADE", "TOR.ALLOW", "TOR.RP", "TOR.OTHER", "TOR.PREPAY"]))
    ln("CONTA", "(Increase)/decrease in contract assets", lambda fy: "-" + M(fy, ["CONTA"]), hidden=not has("CONTA"))
    ln("PAY", "Increase/(decrease) in trade and other payables",
       lambda fy: "-" + M(fy, ["TOP.TRADE", "TOP.RP", "TOP.ACCR", "TOP.OTHER"]))
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
    opn = ln("OPEN", "Cash and cash equivalents at beginning of year", lambda fy: lead("open", fy, CASH_LEAVES), rate="O")
    r = s.r   # effect of exchange rate changes (paragraph 7.13): closing balance less opening balance less cash flows
    s.line("FXEFF", "Effect of exchange rate changes on cash and cash equivalents", None, None,
           pres=(f"C{r + 1}-C{net}-C{opn}", f"D{r + 1}-D{net}-D{opn}"))
    s.line("CLOSE", "Cash and cash equivalents at end of year", f"E{net}+E{opn}", f"F{net}+F{opn}", rate="C", bold=True,
           border=B_TOTAL, note="cash")
    s.gap()
    put(ws, f"A{s.r}", "Indirect method. Cash and cash equivalents comprise cash on hand, bank balances and short-term deposits, "
                       "less bank overdrafts repayable on demand. Cash flows are translated at the average rate for the year.", F_GREY)
    for fy in (CY, PY):
        book.check(f"SoCF {fy}: closing cash per cash flow statement = cash and cash equivalents per ledger (net of overdrafts)",
                   f"{book.cell('SoCF', 'CLOSE', fy)}-{lead('close', fy, CASH_LEAVES)}")


def build_soci_nature(book, ws):
    """Alternative format: expenses classified by nature. Same profit as the SoCI (by function)."""
    _heading(book, ws, "Statement of comprehensive income - expenses classified by nature (alternative format)", _for_year(book))
    _year_heads(book, ws)
    s = Stmt(book, ws, "SoCI_Nature")
    act = book.active_faces
    fl = lambda key, cap, code, **kw: s.line(key, cap, "-" + face("close", CY, code), "-" + face("close", PY, code),
                                             rate="A", **kw)
    rev = fl("REV", "Revenue", "REV", note="revenue")
    fl("OI", "Other income", "OI", note="otherinc", hidden="OI" not in act)
    first = s.r
    for code, caption in NATURES:
        if code == "OTHER":
            continue
        s.line("N_" + code, caption, "-" + nat(CY, [code], EXPENSE_FUNCTIONS), "-" + nat(PY, [code], EXPENSE_FUNCTIONS),
               rate="A", hidden=code not in book.active_natures)
    last = s.r - 1
    tot = lambda fy: "+".join(face("close", fy, f) for f in EXPENSE_FUNCTIONS)
    oth = s.line("N_OTHER", "Other expenses", f"-({tot(CY)})-SUM(E{first}:E{last})", f"-({tot(PY)})-SUM(F{first}:F{last})",
                 rate="A")
    op = s.calc("OP", "Operating profit", f"SUM({{c}}{rev}:{{c}}{oth})", bold=True, border=B_TOP)
    f1 = fl("FININC", "Finance income", "FININC", note="finance", hidden="FININC" not in act)
    fl("FINCOST", "Finance costs", "FINCOST", note="finance")
    f2 = fl("SHAREASSOC", FACE["SHAREASSOC"][1], "SHAREASSOC", hidden="SHAREASSOC" not in act)
    pbt = s.calc("PBT", "Profit before tax", f"{{c}}{op}+SUM({{c}}{f1}:{{c}}{f2})", bold=True, border=B_TOP)
    tax = fl("TAX", "Income tax expense", "TAX", note="tax")
    profit = s.calc("PROFIT", "Profit for the year", f"{{c}}{pbt}+{{c}}{tax}", bold=True, border=B_TOTAL)
    oci = s.line("OCI", "Other comprehensive income for the year, net of tax (see SoCI)", book.cell("SoCI", "OCI", CY),
                 book.cell("SoCI", "OCI", PY),
                 pres=(book.cell("SoCI", "OCI", CY, "P"), book.cell("SoCI", "OCI", PY, "P")))
    s.calc("TCI", "Total comprehensive income for the year", f"{{c}}{profit}+{{c}}{oci}", bold=True, border=B_TOTAL)
    s.gap()
    put(ws, f"A{s.r}", "An entity presents ONE analysis of expenses - by nature or by function, whichever is reliable and more relevant. "
                       "Issue either this sheet or SoCI, not both.", F_GREY)
    for fy in (CY, PY):
        book.check(f"SoCI by nature {fy}: profit for the year equals the SoCI by function",
                   f"{book.cell('SoCI_Nature', 'PROFIT', fy)}-{book.cell('SoCI', 'PROFIT', fy)}")
