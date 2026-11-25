#!/usr/bin/env python3
"""Translate summary-level USD (functional currency) statements into a presentation currency (e.g. KHR)
under IFRS for SMEs (third edition) 30.17-30.20, as a live Excel workbook.

Use this when the user gives statement-level figures rather than a trial balance and ledger
(for a TB/GL, use build_afs.py instead).

Input CSV columns: item,type,cy,py[,rate_cy,rate_py]
  type   asset | liability | income | expense | oci | dividend | share_issue
  cy/py  POSITIVE amounts in the functional currency (closing balances for asset/liability rows;
         amounts for the year for the other types; expense and dividend positive, not negative)
  rate_* optional: translation rate for that row instead of the year's average (e.g. a share issue
         or dividend at its transaction-date rate)

Rates/equity JSON:
  {"close_cy": 4020, "avg_cy": 4050, "close_py": 4080, "avg_py": 4100, "close_ppy": 4060,
   "scale": 1000, "functional": "USD", "presentation": "KHR",
   "opening": {"share_capital_usd": 500000, "share_capital_rate": 4050,
               "other_reserves_usd": 0, "other_reserves_pres": 0,
               "retained_earnings_usd": 180000, "retained_earnings_pres": 734400000}}
  rates are presentation-currency units per 1 functional unit; *_pres opening amounts are in single units;
  "opening" is at the START of the prior year.

Usage: python translate_summary.py summary.csv rates.json --out translated.xlsx
"""
import argparse
import csv
import json
import sys

from openpyxl import Workbook
from openpyxl.styles import Border, Font, PatternFill, Side

F = "Arial"
NUM = '#,##0;(#,##0);"-"'
BOLD, NORM, BLUE = Font(name=F, bold=True), Font(name=F), Font(name=F, color="0000FF")
TITLE = Font(name=F, bold=True, size=13)
TOP, TOTAL = Border(top=Side("thin")), Border(top=Side("thin"), bottom=Side("double"))
TYPES = ("asset", "liability", "income", "expense", "oci", "dividend", "share_issue")


def load(csv_path, json_path):
    rows = []
    for r in csv.DictReader(open(csv_path, newline="", encoding="utf-8-sig")):
        t = r["type"].strip().lower()
        if t not in TYPES:
            sys.exit(f"Unknown type '{t}' for item '{r['item']}'. Use one of {TYPES}.")
        rows.append({"item": r["item"].strip(), "type": t, "cy": float(r["cy"] or 0), "py": float(r["py"] or 0),
                     "rate_cy": float(r["rate_cy"]) if r.get("rate_cy") else None,
                     "rate_py": float(r["rate_py"]) if r.get("rate_py") else None})
    return rows, json.load(open(json_path))


def check_consistency(rows, cfg):
    """USD: closing net assets must equal opening equity + profit + OCI + share issues - dividends, each year."""
    o = cfg["opening"]
    eq = o["share_capital_usd"] + o.get("other_reserves_usd", 0) + o["retained_earnings_usd"]
    out = []
    for y in ("py", "cy"):
        s = lambda t: sum(r[y] for r in rows if r["type"] == t)
        eq = eq + s("income") - s("expense") + s("oci") + s("share_issue") - s("dividend")
        na = s("asset") - s("liability")
        out.append((y, na, eq))
        if abs(na - eq) > 0.5:
            sys.exit(f"Inconsistent input for {y}: net assets {na:,.2f} but equity roll-forward gives {eq:,.2f}.")
    return out


def build(rows, cfg, out):
    wb = Workbook()
    inp = wb.active
    inp.title = "Inputs"
    fc, pc, sc = cfg.get("functional", "USD"), cfg.get("presentation", "KHR"), cfg.get("scale", 1000)
    lab = f"{pc}'000" if sc == 1000 else pc
    inp["A1"], inp["A1"].font = "Inputs - functional currency figures and rates (blue = input)", TITLE
    rates = [("Closing rate, current year end", "close_cy"), ("Average rate, current year", "avg_cy"),
             ("Closing rate, prior year end", "close_py"), ("Average rate, prior year", "avg_py"),
             ("Closing rate, start of prior year", "close_ppy"), ("Presentation units shown per figure", "scale")]
    R = {}
    for i, (name, key) in enumerate(rates):
        r = 3 + i
        inp[f"A{r}"], inp[f"B{r}"] = name, cfg[key] if key != "scale" else sc
        inp[f"B{r}"].font = BLUE
        R[key] = f"Inputs!$B${r}"
    o = cfg["opening"]
    opening = [("Opening share capital and premium (" + fc + ")", o["share_capital_usd"]),
               ("Historical rate for opening share capital", o["share_capital_rate"]),
               ("Opening other reserves (" + fc + ")", o.get("other_reserves_usd", 0)),
               ("Opening other reserves (" + pc + ", single units)", o.get("other_reserves_pres", 0)),
               ("Opening retained earnings (" + fc + ")", o["retained_earnings_usd"]),
               ("Opening retained earnings (" + pc + ", single units)", o["retained_earnings_pres"])]
    O = {}
    for i, (name, v) in enumerate(opening):
        r = 10 + i
        inp[f"A{r}"], inp[f"B{r}"] = name, v
        inp[f"B{r}"].font = BLUE
        O[i] = f"Inputs!$B${r}"
    h = 17
    for c, v in zip("ABCDEFG", ["Item", "Type", f"CY {fc}", f"PY {fc}", "CY rate", "PY rate", ""]):
        inp[f"{c}{h}"], inp[f"{c}{h}"].font = v, BOLD
    first = h + 1
    for i, r in enumerate(rows):
        n = first + i
        inp[f"A{n}"], inp[f"B{n}"], inp[f"C{n}"], inp[f"D{n}"] = r["item"], r["type"], r["cy"], r["py"]
        for c in "CD":
            inp[f"{c}{n}"].font, inp[f"{c}{n}"].number_format = BLUE, "#,##0.00"
        flow = r["type"] not in ("asset", "liability")
        for c, y, key in (("E", "cy", "rate_cy"), ("F", "py", "rate_py")):
            if r[key]:
                inp[f"{c}{n}"], inp[f"{c}{n}"].font = r[key], BLUE
            else:
                close, avg = (R["close_cy"], R["avg_cy"]) if y == "cy" else (R["close_py"], R["avg_py"])
                inp[f"{c}{n}"] = f"={avg if flow else close}"
    last = first + len(rows) - 1
    rng = lambda col: f"Inputs!${col}${first}:${col}${last}"
    usd = lambda t, y: f'SUMIFS({rng("C" if y == "cy" else "D")},{rng("B")},"{t}")'
    pres = lambda t, y: (f'SUMPRODUCT(({rng("B")}="{t}")*{rng("C" if y == "cy" else "D")}'
                         f'*{rng("E" if y == "cy" else "F")})/{R["scale"]}')
    inp.column_dimensions["A"].width = 52
    for c in "BCDEF":
        inp.column_dimensions[c].width = 16

    def sheet(name, title):
        ws = wb.create_sheet(name)
        ws["A1"], ws["A1"].font = title, TITLE
        for c, v in zip("BCDE", [f"CY {lab}", f"PY {lab}", f"CY {fc}", f"PY {fc}"]):
            ws[f"{c}3"], ws[f"{c}3"].font = v, BOLD
        ws.column_dimensions["A"].width = 58
        for c in "BCDE":
            ws.column_dimensions[c].width = 17
        return ws

    def put(ws, r, label, b, c, d, e, font=NORM, border=None):
        ws[f"A{r}"], ws[f"A{r}"].font = label, font
        for col, v in zip("BCDE", (b, c, d, e)):
            ws[f"{col}{r}"] = "=" + v
            ws[f"{col}{r}"].font, ws[f"{col}{r}"].number_format = font, NUM
            if border:
                ws[f"{col}{r}"].border = border

    # ---- SoCI
    ci = sheet("SoCI", "Statement of comprehensive income")
    r = 4
    item_rows = {}
    for i, row in enumerate(rows):
        if row["type"] in ("income", "expense"):
            n = first + i
            sgn = "-" if row["type"] == "expense" else ""
            item_rows.setdefault(row["type"], []).append(r)
            put(ci, r, row["item"], f"{sgn}Inputs!C{n}*Inputs!E{n}/{R['scale']}", f"{sgn}Inputs!D{n}*Inputs!F{n}/{R['scale']}",
                f"{sgn}Inputs!C{n}", f"{sgn}Inputs!D{n}")
            r += 1
    pl = item_rows.get("income", []) + item_rows.get("expense", [])
    profit = r
    put(ci, r, "Profit for the year", *[("+".join(f"{c}{x}" for x in pl) or "0") for c in "BCDE"], font=BOLD, border=TOTAL)
    r += 2
    ci[f"A{r}"], ci[f"A{r}"].font = "Other comprehensive income - items that will not be reclassified to profit or loss", BOLD
    r += 1
    oci_first = r
    for i, row in enumerate(rows):
        if row["type"] == "oci":
            n = first + i
            put(ci, r, row["item"], f"Inputs!C{n}*Inputs!E{n}/{R['scale']}", f"Inputs!D{n}*Inputs!F{n}/{R['scale']}",
                f"Inputs!C{n}", f"Inputs!D{n}")
            r += 1
    fx = r
    put(ci, r, f"Exchange differences on translation to the presentation currency (30.18(c))", "0", "0", "0", "0")
    r += 1
    oci_t = r
    put(ci, r, "Other comprehensive income for the year", *[f"SUM({c}{oci_first}:{c}{fx})" for c in "BCDE"], font=BOLD, border=TOP)
    r += 2
    put(ci, r, "Total comprehensive income for the year", *[f"{c}{profit}+{c}{oci_t}" for c in "BCDE"], font=BOLD, border=TOTAL)

    # ---- SoCE (presentation currency), columns: share capital, other reserves, translation reserve, RE, total
    ce = wb.create_sheet("SoCE")
    ce["A1"], ce["A1"].font = f"Statement of changes in equity ({lab})", TITLE
    heads = ["Share capital", "Other reserves", "Translation reserve", "Retained earnings", "Total"]
    for c, v in zip("BCDEF", heads):
        ce[f"{c}3"], ce[f"{c}3"].font = v, BOLD
        ce.column_dimensions[c].width = 17
    ce.column_dimensions["A"].width = 46
    S = R["scale"]
    rows_ce = {}
    rr = 4

    def ce_row(label, vals, font=NORM, border=None, key=None):
        nonlocal rr
        ce[f"A{rr}"], ce[f"A{rr}"].font = label, font
        for c, v in zip("BCDE", vals):
            ce[f"{c}{rr}"] = "=" + v
            ce[f"{c}{rr}"].font, ce[f"{c}{rr}"].number_format = font, NUM
        ce[f"F{rr}"] = f"=SUM(B{rr}:E{rr})"
        ce[f"F{rr}"].font, ce[f"F{rr}"].number_format = font, NUM
        for c in "BCDEF":
            if border:
                ce[f"{c}{rr}"].border = border
        if key:
            rows_ce[key] = rr
        rr += 1
        return rr - 1

    open_na = f"({O[0]}+{O[2]}+{O[4]})"
    o_row = ce_row("Balance at the start of the prior year",
                   [f"{O[0]}*{O[1]}/{S}", f"{O[3]}/{S}",
                    f"{open_na}*{R['close_ppy']}/{S}-({O[0]}*{O[1]}+{O[3]}+{O[5]})/{S}", f"{O[5]}/{S}"], BOLD)
    for y, col, close in (("py", "C", R["close_py"]), ("cy", "B", R["close_cy"])):
        p = ce_row("Profit for the year", ["0", "0", "0", f"SoCI!{col}{profit}"])
        oci_items = f"SUM(SoCI!{col}{oci_first}:SoCI!{col}{fx - 1})" if fx > oci_first else "0"
        oc = ce_row("Other comprehensive income", ["0", oci_items, f"SoCI!{col}{fx}", "0"])
        si = ce_row("Issue of shares", [pres("share_issue", y), "0", "0", "0"])
        dv = ce_row("Dividends", ["0", "0", "0", "-" + pres("dividend", y)])
        cl = ce_row("Balance at the end of the " + ("prior" if y == "py" else "current") + " year",
                    [f"{c}{o_row}+{c}{p}+{c}{oc}+{c}{si}+{c}{dv}" for c in "BCDE"], BOLD, TOTAL, key=y)
        # translation reserve closing = net assets at closing rate less the other components
        tr = f"D{cl}"
        ce[tr] = (f"=({usd('asset', y)}-{usd('liability', y)})*{close}/{S}-(B{cl}+C{cl}+E{cl})")
        # OCI translation difference = movement in the translation reserve
        ci[f"{col}{fx}"] = f"=SoCE!D{cl}-SoCE!D{o_row}"
        rr += 1
        o_row = cl
    for y, col in (("py", "D"), ("cy", "E")):
        ci[f"{col}{fx}"] = "=0"

    # ---- SoFP
    fp = sheet("SoFP", "Statement of financial position")
    r = 4
    blocks = {}
    for t, title in (("asset", "Assets"), ("liability", "Liabilities")):
        fp[f"A{r}"], fp[f"A{r}"].font = title, BOLD
        r += 1
        s0 = r
        for i, row in enumerate(rows):
            if row["type"] == t:
                n = first + i
                put(fp, r, row["item"], f"Inputs!C{n}*{R['close_cy']}/{S}", f"Inputs!D{n}*{R['close_py']}/{S}",
                    f"Inputs!C{n}", f"Inputs!D{n}")
                r += 1
        blocks[t] = r
        put(fp, r, "Total " + title.lower(), *[f"SUM({c}{s0}:{c}{r - 1})" for c in "BCDE"], font=BOLD, border=TOP)
        r += 2
    fp[f"A{r}"], fp[f"A{r}"].font = "Equity", BOLD
    r += 1
    e0 = r
    eq_lines = [("Share capital", "B"), ("Other reserves", "C"), ("Foreign currency translation reserve", "D"),
                ("Retained earnings", "E")]
    for name, c in eq_lines:
        fp[f"A{r}"] = name
        fp[f"B{r}"], fp[f"C{r}"] = f"=SoCE!{c}{rows_ce['cy']}", f"=SoCE!{c}{rows_ce['py']}"
        for cc in "BC":
            fp[f"{cc}{r}"].number_format = NUM
        r += 1
    put(fp, r, "Total equity", f"SUM(B{e0}:B{r - 1})", f"SUM(C{e0}:C{r - 1})",
        f"D{blocks['asset']}-D{blocks['liability']}", f"E{blocks['asset']}-E{blocks['liability']}", BOLD, TOP)
    teq = r
    r += 1
    put(fp, r, "Total equity and liabilities", *[f"{c}{teq}+{c}{blocks['liability']}" for c in "BCDE"], BOLD, TOTAL)

    # ---- Proof (30.20) and checks
    pr = wb.create_sheet("Proof")
    pr["A1"], pr["A1"].font = f"Proof of the translation difference by rates (30.20), {lab}", TITLE
    pr.column_dimensions["A"].width = 70
    for c, v in zip("BC", ["CY", "PY"]):
        pr[f"{c}3"], pr[f"{c}3"].font = v, BOLD
        pr.column_dimensions[c].width = 17
    na_open = {"py": open_na, "cy": f"({usd('asset', 'py')}-{usd('liability', 'py')})"}
    prev = {"py": R["close_ppy"], "cy": R["close_py"]}
    close = {"py": R["close_py"], "cy": R["close_cy"]}
    lines = [("Opening net assets x (closing rate - previous closing rate)",
              lambda y: f"{na_open[y]}*({close[y]}-{prev[y]})/{S}")]
    for t, sign, cap in (("income", 1, "Income"), ("expense", -1, "Expenses"), ("oci", 1, "Other comprehensive income"),
                         ("share_issue", 1, "Share issues"), ("dividend", -1, "Dividends")):
        lines.append((f"{cap} x (closing rate - rate used)",
                      lambda y, t=t, sign=sign: (f"{'' if sign > 0 else '-'}({usd(t, y)}*{close[y]}/{S}-{pres(t, y)})")))
    r = 4
    for cap, fn in lines:
        pr[f"A{r}"] = cap
        pr[f"B{r}"], pr[f"C{r}"] = "=" + fn("cy"), "=" + fn("py")
        r += 1
    pr[f"A{r}"], pr[f"A{r}"].font = "Translation difference by rates", BOLD
    pr[f"B{r}"], pr[f"C{r}"] = f"=SUM(B4:B{r - 1})", f"=SUM(C4:C{r - 1})"
    tot = r
    r += 1
    pr[f"A{r}"] = "Translation difference in OCI (SoCI)"
    pr[f"B{r}"], pr[f"C{r}"] = f"=SoCI!B{fx}", f"=SoCI!C{fx}"
    r += 1
    pr[f"A{r}"], pr[f"A{r}"].font = "Difference (must be nil)", BOLD
    pr[f"B{r}"], pr[f"C{r}"] = f"=ROUND(B{tot}-B{r - 1},0)", f"=ROUND(C{tot}-C{r - 1},0)"
    for row in pr.iter_rows(min_row=4, max_row=r, min_col=2, max_col=3):
        for c in row:
            c.number_format = NUM
    r += 2
    checks = [("SoFP balances - presentation currency, CY", f"ROUND(SoFP!B{teq}-(SoFP!B{blocks['asset']}-SoFP!B{blocks['liability']}),0)"),
              ("SoFP balances - presentation currency, PY", f"ROUND(SoFP!C{teq}-(SoFP!C{blocks['asset']}-SoFP!C{blocks['liability']}),0)"),
              ("Proof agrees to OCI, CY and PY", f"B{tot - 0 + 2}+C{tot + 2}")]
    pr[f"A{r}"], pr[f"A{r}"].font = "Checks (all must be nil)", BOLD
    for i, (cap, f) in enumerate(checks, start=1):
        pr[f"A{r + i}"], pr[f"B{r + i}"] = cap, "=" + f
    wb._sheets = [wb[n] for n in ("SoFP", "SoCI", "SoCE", "Proof", "Inputs")]
    wb.save(out)
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("csv")
    ap.add_argument("rates")
    ap.add_argument("--out", default="translated.xlsx")
    a = ap.parse_args()
    rows, cfg = load(a.csv, a.rates)
    for y, na, eq in check_consistency(rows, cfg):
        print(f"{y.upper()} functional-currency net assets {na:,.2f} = equity roll-forward {eq:,.2f}")
    print("Workbook:", build(rows, cfg, a.out))


if __name__ == "__main__":
    main()
