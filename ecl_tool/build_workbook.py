"""Write IFRS9_ECL_Tool.xlsx: regression output from Python + live Excel formulas.

The coefficients are estimated in Python (ecl_models.py). Everything downstream
(PD, LGD, EAD, staging, ECL, journal entry) is an Excel formula, so the workbook
recalculates when the macro scenarios, weights, staging rules or loans change.
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

import ecl_models as em

OUT = Path(__file__).parent / "IFRS9_ECL_Tool.xlsx"

NAVY = "1F3864"
F_BASE = Font(name="Arial", size=10)
F_BOLD = Font(name="Arial", size=10, bold=True)
F_HEAD = Font(name="Arial", size=10, bold=True, color="FFFFFF")
F_TITLE = Font(name="Arial", size=14, bold=True, color=NAVY)
F_INPUT = Font(name="Arial", size=10, color="0000FF")
F_LINK = Font(name="Arial", size=10, color="008000")
F_NOTE = Font(name="Arial", size=9, italic=True, color="595959")
FILL_HEAD = PatternFill("solid", fgColor=NAVY)
FILL_SUB = PatternFill("solid", fgColor="D9E1F2")
FILL_KEY = PatternFill("solid", fgColor="FFFF00")
FILL_SC = {"Base": "D9E1F2", "Upside": "E2EFDA", "Downside": "FCE4D6"}
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

PCT, PCT2, NUM, DEC4 = "0.0%;(0.0%);-", "0.00%", "#,##0;(#,##0);-", "0.0000"

LABELS = {
    "const": "Intercept", "credit_score": "Credit score (points)",
    "dti": "Debt-to-income ratio (fraction)", "ltv": "Loan-to-value (fraction; 0 if unsecured)",
    "delinq": "Past due 1-89 days (1/0)", "consumer": "Product = Consumer (1/0)",
    "sme": "Product = SME (1/0)", "gdp_growth": "GDP growth (%)",
    "unemployment": "Unemployment rate (%)", "lending_rate": "Lending interest rate (%)",
    "hpi_growth": "House price growth (%)", "utilisation": "Utilisation = drawn / limit",
}
MACRO_LABELS = [("gdp_growth", "GDP growth (% y/y)"), ("unemployment", "Unemployment rate (%)"),
                ("lending_rate", "Lending interest rate (%)"), ("hpi_growth", "House price growth (% y/y)")]


def put(ws, ref, value, font=F_BASE, fmt=None, fill=None, border=False, align=None):
    c = ws[ref]
    c.value, c.font = value, font
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    if border:
        c.border = BOX
    if align:
        c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    return c


def header(ws, row, labels, col=1, fill=FILL_HEAD, font=F_HEAD):
    for i, text in enumerate(labels):
        put(ws, f"{L(col + i)}{row}", text, font, fill=fill, border=True, align="center")


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


def style_chart(ch, title, x=None, y=None, w=16, h=8):
    ch.title, ch.width, ch.height = title, w, h
    ch.roundedCorners = False
    ch.title.overlay = False
    if ch.legend is not None:
        ch.legend.overlay = False
    ch.x_axis.delete = ch.y_axis.delete = False
    if x:
        ch.x_axis.title = x
        ch.x_axis.title.overlay = False
    if y:
        ch.y_axis.title = y
        ch.y_axis.title.overlay = False


# ----------------------------------------------------------------- Guide
def sheet_guide(wb):
    ws = wb.active
    ws.title = "Guide"
    ws.sheet_view.showGridLines = False
    widths(ws, {"A": 26, "B": 110})
    put(ws, "A1", "IFRS 9 Expected Credit Loss Tool - logistic regression PD, LGD and EAD "
                  "with forward-looking macro scenarios", F_TITLE)
    put(ws, "A2", "ILLUSTRATIVE TEACHING MODEL. All loan and macro data are synthetic "
                  "(generate_data.py); replace the CSV files with real data before drawing conclusions.",
        Font(name="Arial", size=10, bold=True, color="C00000"))
    rows = [
        ("HOW IT WORKS", None),
        ("1. Estimate (Python)", "ecl_models.py fits three logistic-link regressions on loan history joined "
                                 "to macro history: PD (binary logit), LGD and CCF (fractional logit)."),
        ("2. Forecast (Excel)", "Scenarios sheet holds 5-year forecasts of GDP growth, unemployment, lending "
                                "rate and house prices for Base / Upside / Downside, each with a probability."),
        ("3. Predict (Excel)", "ECL_Engine applies 1/(1+EXP(-(borrower score + macro contribution))) for each "
                               "loan, scenario and year, giving PD, LGD and EAD term structures."),
        ("4. Stage (Excel)", "Stage 3 if past due >= 90 days; Stage 2 if past due >= 30 days or the weighted "
                             "12-month PD has risen by the SICR thresholds since origination; otherwise Stage 1."),
        ("5. Measure (Excel)", "ECL = sum over years of survival x PD x LGD x EAD, discounted at the loan EIR. "
                               "Stage 1 takes year 1 only, Stage 2 the whole remaining life, Stage 3 LGD x EAD. "
                               "Results are probability-weighted across scenarios (IFRS 9 para 5.5.17)."),
        ("SHEETS", None),
        ("Scenarios", "INPUT. Scenario weights, staging rules, macro forecasts. Change these to see ECL move."),
        ("Coefficients", "Python regression output (coefficient, std error, z, p-value, odds ratio)."),
        ("Portfolio", "INPUT. Loan book at the reporting date; add or overwrite loans in the blue columns."),
        ("ECL_Engine", "Formulas only. One row per loan; PD/LGD/EAD/loss by scenario and year."),
        ("Summary", "ECL by stage, product and scenario; forward-looking overlay; journal entry; charts."),
        ("Validation", "Python model diagnostics: AUC, Gini, KS, calibration by decile; reconciliation check."),
        ("Macro_History", "Quarterly macro series and observed default rate used for estimation."),
        ("COLOUR CODE", None),
        ("Blue text", "Hard-coded input you may change."),
        ("Yellow fill", "Key assumption (scenario weights, staging thresholds, opening allowance)."),
        ("Black text", "Formula - do not overwrite."),
        ("Green text", "Formula that links to another sheet."),
        ("UPDATING", None),
        ("New data", "Replace the CSVs in ecl_tool/data and run: python run_all.py. Coefficients, portfolio "
                     "and scenarios are rewritten; all formulas are regenerated."),
        ("More loans", "Rows are generated per loan; re-run the script rather than copying rows by hand."),
        ("SIMPLIFICATIONS", None),
        ("Horizon", "Lifetime is capped at 5 annual periods; longer loans need more projection years."),
        ("SICR test", "Compares 12-month PD now with 12-month PD at origination as a proxy for the lifetime "
                      "PD comparison (IFRS 9 para B5.5.13 permits this where it is a reasonable approximation)."),
        ("EAD", "Drawn balance amortises straight-line over the remaining term; undrawn limit is converted "
                "by the modelled CCF. Prepayment is ignored."),
        ("Stage 3", "PD = 100%; ECL = weighted LGD x EAD. Interest revenue then accrues on amortised cost "
                    "(gross carrying amount less loss allowance), IFRS 9 para 5.4.1(b)."),
        ("Macro lag", "Macro variables enter contemporaneously; in practice test lags and stationarity."),
    ]
    r = 4
    for a, b in rows:
        if b is None:
            r += 1
            put(ws, f"A{r}", a, F_HEAD, fill=FILL_HEAD)
            ws[f"B{r}"].fill = FILL_HEAD
        else:
            put(ws, f"A{r}", a, F_BOLD)
            put(ws, f"B{r}", b, align="left")
            if a == "Blue text":
                ws[f"A{r}"].font = Font(name="Arial", size=10, bold=True, color="0000FF")
            if a == "Green text":
                ws[f"A{r}"].font = Font(name="Arial", size=10, bold=True, color="008000")
            if a == "Yellow fill":
                ws[f"A{r}"].fill = FILL_KEY
        r += 1


# ---------------------------------------------------------- Coefficients
def sheet_coefficients(wb, models):
    """Returns {model name: {variable: absolute cell reference}}."""
    ws = wb.create_sheet("Coefficients")
    widths(ws, {"A": 16, "B": 40, "C": 13, "D": 11, "E": 10, "F": 10, "G": 12, "H": 34})
    put(ws, "A1", "Regression coefficients (Python output - statsmodels)", F_TITLE)
    put(ws, "A2", "Source: ecl_models.py fitted on data/loan_history.csv.gz (synthetic). "
                  "Re-run run_all.py to refresh. Odds ratio = EXP(coefficient).", F_NOTE)
    titles = {"PD": "PD model - binary logistic regression; target = default within 12 months",
              "LGD": "LGD model - fractional logistic regression; target = loss rate on defaulted loans",
              "CCF": "EAD model - fractional logistic regression of CCF; EAD = drawn + CCF x undrawn"}
    refs, r = {}, 4
    for m in models:
        put(ws, f"A{r}", titles[m.name], F_BOLD, fill=FILL_SUB)
        for c in "BCDEFGH":
            ws[f"{c}{r}"].fill = FILL_SUB
        header(ws, r + 1, ["Variable", "Description", "Coefficient", "Std error", "z", "p-value",
                           "Odds ratio", "Reading"])
        r += 2
        refs[m.name] = {}
        for var, row in m.table.iterrows():
            put(ws, f"A{r}", var, border=True)
            put(ws, f"B{r}", LABELS[var], border=True)
            put(ws, f"C{r}", float(row["coef"]), F_INPUT, DEC4, border=True)
            put(ws, f"D{r}", round(float(row["se"]), 6), F_INPUT, DEC4, border=True)
            put(ws, f"E{r}", round(float(row["z"]), 3), F_INPUT, "0.00", border=True)
            put(ws, f"F{r}", round(float(row["p"]), 5), F_INPUT, "0.0000", border=True)
            put(ws, f"G{r}", f"=EXP(C{r})", fmt="0.000", border=True)
            put(ws, f"H{r}", f'=IF(F{r}<0.01,"significant at 1%",IF(F{r}<0.05,"significant at 5%",'
                             f'IF(F{r}<0.1,"significant at 10%","not significant")))', border=True)
            refs[m.name][var] = f"Coefficients!$C${r}"
            r += 1
        r += 2
    put(ws, f"A{r}", "Sign check: PD rises with unemployment and lending rate, falls with GDP and house "
                     "price growth; LGD rises with LTV and unemployment, falls with house price growth.", F_NOTE)
    return refs


# -------------------------------------------------------------- Scenarios
def sheet_scenarios(wb, scenarios, refs):
    """Returns addresses for weights, staging rules and macro contributions."""
    ws = wb.create_sheet("Scenarios")
    widths(ws, {"A": 34, "B": 12, "C": 11, "D": 11, "E": 11, "F": 11, "G": 11, "H": 3, "I": 60})
    put(ws, "A1", "Forward-looking macroeconomic scenarios and staging rules", F_TITLE)
    put(ws, "A2", "Blue cells are inputs. Illustrative forecasts, not official projections - "
                  "replace with central bank / statistics office / management forecasts.", F_NOTE)
    names = list(scenarios["scenario"].unique())
    addr = {"w": {}, "m": {}}

    put(ws, "A3", "Scenario probability weights", F_BOLD)
    header(ws, 4, ["Scenario", "Weight"])
    put(ws, "I4", "Narrative", F_BOLD)
    story = {"Base": "Central forecast: steady growth, stable labour market.",
             "Upside": "Stronger growth, falling unemployment, rising house prices.",
             "Downside": "Recession in years 1-2: rate shock, job losses, house price fall, then recovery."}
    for i, n in enumerate(names):
        r = 5 + i
        put(ws, f"A{r}", n, F_BOLD, fill=PatternFill("solid", fgColor=FILL_SC[n]), border=True)
        w = float(scenarios.loc[scenarios.scenario == n, "weight"].iloc[0])
        put(ws, f"B{r}", w, F_INPUT, PCT, FILL_KEY, True)
        put(ws, f"I{r}", story.get(n, ""))
        addr["w"][n] = f"Scenarios!$B${r}"
    end = 4 + len(names)
    put(ws, f"A{end + 1}", "Total (must be 100%)", F_BOLD, border=True)
    put(ws, f"B{end + 1}", f"=SUM(B5:B{end})", F_BOLD, PCT, border=True)
    put(ws, f"C{end + 1}", f'=IF(ABS(B{end + 1}-1)<0.0001,"OK","Weights must sum to 100%")', F_BOLD)

    r = end + 3
    put(ws, f"A{r}", "Staging rules", F_BOLD)
    rules = [("sicr_mult", "SICR: 12m PD now / PD at origination >=", em.SICR_MULTIPLE, "0.0x",
              "Relative test. Both relative and absolute tests must be met."),
             ("sicr_abs", "SICR: and increase in 12m PD of at least", em.SICR_ABS, PCT2,
              "Absolute test (stops trivial moves in very low PDs triggering Stage 2)."),
             ("dpd2", "Stage 2 backstop: days past due >=", em.DPD_STAGE2, "0",
              "IFRS 9 para 5.5.11 rebuttable presumption (30 days)."),
             ("dpd3", "Stage 3 (credit-impaired): days past due >=", em.DPD_STAGE3, "0",
              "IFRS 9 para B5.5.37 rebuttable presumption (90 days).")]
    for i, (key, label, val, fmt, note) in enumerate(rules):
        rr = r + 1 + i
        put(ws, f"A{rr}", label, border=True)
        put(ws, f"B{rr}", val, F_INPUT, fmt, FILL_KEY, True)
        put(ws, f"I{rr}", note, F_NOTE)
        addr[key] = f"Scenarios!$B${rr}"

    r = r + len(rules) + 3
    for n in names:
        sc = scenarios[scenarios.scenario == n].sort_values("year")
        fill = PatternFill("solid", fgColor=FILL_SC[n])
        header(ws, r, [f"{n} scenario", "Unit"] + [f"Year {t}" for t in range(1, em.HORIZON + 1)])
        for k, (var, label) in enumerate(MACRO_LABELS):
            rr = r + 1 + k
            put(ws, f"A{rr}", label, border=True, fill=fill)
            put(ws, f"B{rr}", "% points", F_NOTE, border=True)
            for t, v in enumerate(sc[var].tolist()):
                put(ws, f"{L(3 + t)}{rr}", float(v), F_INPUT, "0.0", border=True)
        mrow = {var: r + 1 + k for k, (var, _) in enumerate(MACRO_LABELS)}
        for j, m in enumerate(("PD", "LGD", "CCF")):
            rr = r + 5 + j
            put(ws, f"A{rr}", f"Macro contribution to {m} log-odds", F_BOLD, border=True)
            put(ws, f"B{rr}", "log-odds", F_NOTE, border=True)
            for t in range(em.HORIZON):
                col = L(3 + t)
                terms = [f"{refs[m][v]}*{col}{mrow[v]}" for v, _ in MACRO_LABELS if v in refs[m]]
                put(ws, f"{col}{rr}", "=" + "+".join(terms), F_LINK, "0.000", border=True)
                addr["m"][(n, m, t + 1)] = f"Scenarios!${col}${rr}"
        put(ws, f"I{r + 5}", "Sum of coefficient x forecast macro value (links to Coefficients sheet)", F_NOTE)
        r += 10
    return addr


# -------------------------------------------------------------- Portfolio
def sheet_portfolio(wb, portfolio):
    ws = wb.create_sheet("Portfolio")
    n = len(portfolio)
    first, last = 5, 4 + n
    put(ws, "A1", "Loan portfolio at the reporting date", F_TITLE)
    put(ws, "A2", "Blue columns are inputs (synthetic data from data/portfolio.csv). Amounts in currency "
                  "units (CU). Row 5 shows the expected format: rates and ratios as fractions, term in whole "
                  "years (1-5), LTV = 0 for unsecured loans.", F_NOTE)
    heads = ["Loan ID", "Product", "Gross carrying amount (CU)", "Undrawn limit (CU)", "EIR",
             "Remaining term (yrs)", "Credit score", "DTI", "LTV", "Days past due", "12m PD at origination",
             "Past due 1-89d (1/0)", "Consumer (1/0)", "SME (1/0)", "Utilisation",
             "Stage", "ECL (CU)", "Coverage ratio"]
    header(ws, 4, heads)
    ws.row_dimensions[4].height = 42
    fmts = [None, None, NUM, NUM, PCT2, "0", "0", "0.000", "0.000", "0", PCT2]
    cols = ["loan_id", "product", "drawn", "undrawn", "eir", "term_years", "credit_score", "dti", "ltv",
            "dpd", "orig_pd_12m"]
    for i, rec in enumerate(portfolio[cols].itertuples(index=False)):
        r = first + i
        for j, v in enumerate(rec):
            v = v if isinstance(v, str) else (int(v) if cols[j] in ("term_years", "dpd", "credit_score")
                                              else float(v))
            put(ws, f"{L(j + 1)}{r}", v, F_INPUT, fmts[j])
        put(ws, f"L{r}", f"=IF(AND(J{r}>=1,J{r}<{ADDR['dpd3']}),1,0)", F_LINK, "0")
        put(ws, f"M{r}", f'=IF(B{r}="Consumer",1,0)', fmt="0")
        put(ws, f"N{r}", f'=IF(B{r}="SME",1,0)', fmt="0")
        put(ws, f"O{r}", f"=IF(C{r}+D{r}>0,C{r}/(C{r}+D{r}),0)", fmt="0.000")
    widths(ws, {"A": 10, "B": 11, "C": 16, "D": 14, "E": 9, "F": 11, "G": 9, "H": 8, "I": 8, "J": 9,
                "K": 12, "L": 11, "M": 10, "N": 8, "O": 10, "P": 8, "Q": 14, "R": 10})
    ws.freeze_panes = "C5"
    return first, last


# ------------------------------------------------------------- ECL engine
def sheet_engine(wb, names, refs, first, last):
    """Returns column letters used by the Summary and Portfolio sheets."""
    ws = wb.create_sheet("ECL_Engine")
    put(ws, "A1", "ECL engine - PD, LGD, EAD and expected loss by scenario and year", F_TITLE)
    put(ws, "A2", "PD/LGD/CCF = 1/(1+EXP(-(borrower score + macro contribution))). "
                  "EAD(t) = GCA x MAX(0,1-(t-1)/term) + CCF(t) x undrawn. "
                  "Loss(t) = survival to t-1 x PD(t) x LGD(t) x EAD(t) / (1+EIR)^t, zero after maturity.", F_NOTE)
    H = em.HORIZON
    P = "Portfolio!$"
    feat = {"credit_score": "G", "dti": "H", "ltv": "I", "delinq": "L", "consumer": "M", "sme": "N",
            "utilisation": "O"}
    cols = {"id": "A", "xb": {"PD": "B", "LGD": "C", "CCF": "D"}, "sc": {}}
    header(ws, 4, ["Loan ID", "Borrower score PD", "Borrower score LGD", "Borrower score CCF"])
    put(ws, "B3", "Borrower part of log-odds", F_BOLD, fill=FILL_SUB, align="center")
    ws.merge_cells("B3:D3")

    c = 5
    for n in names:
        blk = {}
        for kind in ("PD", "LGD", "EAD", "Loss"):
            blk[kind] = [L(c + t) for t in range(H)]
            header(ws, 4, [f"{kind} Y{t + 1}" for t in range(H)], col=c)
            c += H
        cols["sc"][n] = blk
        s, e = blk["PD"][0], blk["Loss"][-1]
        put(ws, f"{s}3", f"{n} scenario", F_BOLD, fill=PatternFill("solid", fgColor=FILL_SC[n]), align="center")
        ws.merge_cells(f"{s}3:{e}3")
    for n in names:
        cols["sc"][n]["e12"], cols["sc"][n]["elt"] = L(c), L(c + 1)
        header(ws, 4, [f"12m ECL {n}", f"Lifetime ECL {n}"], col=c)
        c += 2
    out_names = ["w_pd", "w_e12", "w_elt", "w_s3", "stage", "ecl"]
    for k, nm in enumerate(out_names):
        cols[nm] = L(c + k)
    header(ws, 4, ["Weighted 12m PD", "Weighted 12m ECL", "Weighted lifetime ECL",
                   "Weighted LGD x EAD (Stage 3)", "Stage", "ECL recognised"], col=c)
    put(ws, f"{cols['w_pd']}3", "Probability-weighted results", F_BOLD, fill=FILL_KEY, align="center")
    ws.merge_cells(f"{cols['w_pd']}3:{cols['ecl']}3")
    ws.row_dimensions[4].height = 30

    def logistic(xb, macro):
        return f"=1/(1+EXP(-({xb}+{macro})))"

    for r in range(first, last + 1):
        put(ws, f"A{r}", f"=Portfolio!A{r}", F_LINK)
        for m in ("PD", "LGD", "CCF"):
            terms = [refs[m]["const"]] + [f"{refs[m][v]}*{P}{feat[v]}{r}" for v in refs[m] if v in feat]
            put(ws, f"{cols['xb'][m]}{r}", "=" + "+".join(terms), F_LINK, "0.000")
        gca, und, eir, term = f"{P}C{r}", f"{P}D{r}", f"{P}E{r}", f"{P}F{r}"
        for n in names:
            b = cols["sc"][n]
            for t in range(1, H + 1):
                pd_c, lgd_c, ead_c, el_c = (b[k][t - 1] for k in ("PD", "LGD", "EAD", "Loss"))
                put(ws, f"{pd_c}{r}", logistic(f"$B{r}", ADDR["m"][(n, "PD", t)]), fmt=PCT2)
                put(ws, f"{lgd_c}{r}", logistic(f"$C{r}", ADDR["m"][(n, "LGD", t)]), fmt=PCT)
                put(ws, f"{ead_c}{r}",
                    f"={gca}*MAX(0,1-{t - 1}/{term})+{und}/(1+EXP(-($D{r}+{ADDR['m'][(n, 'CCF', t)]})))",
                    fmt=NUM)
                surv = "*".join(f"(1-{b['PD'][k]}{r})" for k in range(t - 1)) or "1"
                put(ws, f"{el_c}{r}",
                    f"=IF({term}>={t},{surv}*{pd_c}{r}*{lgd_c}{r}*{ead_c}{r}/(1+{eir})^{t},0)", fmt=NUM)
            put(ws, f"{b['e12']}{r}", f"={b['Loss'][0]}{r}", fmt=NUM)
            put(ws, f"{b['elt']}{r}", f"=SUM({b['Loss'][0]}{r}:{b['Loss'][-1]}{r})", fmt=NUM)

        def wsum(expr):
            return "=" + "+".join(f"{ADDR['w'][n]}*{expr(cols['sc'][n])}" for n in names)
        put(ws, f"{cols['w_pd']}{r}", wsum(lambda b: f"{b['PD'][0]}{r}"), fmt=PCT2)
        put(ws, f"{cols['w_e12']}{r}", wsum(lambda b: f"{b['e12']}{r}"), fmt=NUM)
        put(ws, f"{cols['w_elt']}{r}", wsum(lambda b: f"{b['elt']}{r}"), fmt=NUM)
        put(ws, f"{cols['w_s3']}{r}", wsum(lambda b: f"{b['LGD'][0]}{r}*{b['EAD'][0]}{r}"), fmt=NUM)
        wp, op, dpd = f"{cols['w_pd']}{r}", f"{P}K{r}", f"{P}J{r}"
        put(ws, f"{cols['stage']}{r}",
            f"=IF({dpd}>={ADDR['dpd3']},3,IF(OR({dpd}>={ADDR['dpd2']},"
            f"AND({wp}>={ADDR['sicr_mult']}*{op},{wp}-{op}>={ADDR['sicr_abs']})),2,1))", F_BOLD, "0")
        st = f"{cols['stage']}{r}"
        put(ws, f"{cols['ecl']}{r}",
            f"=IF({st}=3,{cols['w_s3']}{r},IF({st}=2,{cols['w_elt']}{r},{cols['w_e12']}{r}))", F_BOLD, NUM)

    for i in range(1, c + len(out_names)):
        ws.column_dimensions[L(i)].width = 11
    for nm in ("w_e12", "w_elt", "w_s3", "ecl"):
        ws.column_dimensions[cols[nm]].width = 15
    ws.freeze_panes = "B5"

    # bring stage and ECL back to the Portfolio sheet
    wp = wb["Portfolio"]
    for r in range(first, last + 1):
        put(wp, f"P{r}", f"=ECL_Engine!{cols['stage']}{r}", F_LINK, "0")
        put(wp, f"Q{r}", f"=ECL_Engine!{cols['ecl']}{r}", F_LINK, NUM)
        put(wp, f"R{r}", f"=IF(C{r}>0,Q{r}/C{r},0)", fmt=PCT)
    return cols


# ---------------------------------------------------------------- Summary
def sheet_summary(wb, names, cols, first, last, python_total):
    ws = wb.create_sheet("Summary", 1)
    ws.sheet_view.showGridLines = False
    widths(ws, {"A": 50, "B": 16, "C": 18, "D": 18, "E": 16, "F": 16, "G": 3, "H": 14})
    put(ws, "A1", "ECL summary - probability-weighted, forward-looking", F_TITLE)
    put(ws, "A2", "All figures are formulas on ECL_Engine / Portfolio. Amounts in CU.", F_NOTE)

    def rng(sheet, col):
        return f"{sheet}!${col}${first}:${col}${last}"
    stage, ecl = rng("ECL_Engine", cols["stage"]), rng("ECL_Engine", cols["ecl"])
    gca, und, prod = rng("Portfolio", "C"), rng("Portfolio", "D"), rng("Portfolio", "B")

    put(ws, "A4", "1. Loss allowance by stage", F_BOLD)
    header(ws, 5, ["Stage", "Number of loans", "Gross carrying amount", "Undrawn commitments",
                   "ECL allowance", "Coverage ratio"])
    desc = {1: "Stage 1 - 12-month ECL", 2: "Stage 2 - lifetime ECL (SICR)",
            3: "Stage 3 - lifetime ECL (credit-impaired)"}
    for s in (1, 2, 3):
        r = 5 + s
        put(ws, f"A{r}", desc[s], border=True)
        put(ws, f"B{r}", f"=COUNTIF({stage},{s})", F_LINK, "0", border=True)
        put(ws, f"C{r}", f"=SUMIFS({gca},{stage},{s})", F_LINK, NUM, border=True)
        put(ws, f"D{r}", f"=SUMIFS({und},{stage},{s})", F_LINK, NUM, border=True)
        put(ws, f"E{r}", f"=SUMIFS({ecl},{stage},{s})", F_LINK, NUM, border=True)
        put(ws, f"F{r}", f"=IF(C{r}>0,E{r}/C{r},0)", fmt=PCT, border=True)
    put(ws, "A9", "Total", F_BOLD, border=True)
    for c in "BCDE":
        put(ws, f"{c}9", f"=SUM({c}6:{c}8)", F_BOLD, "0" if c == "B" else NUM, border=True)
    put(ws, "F9", "=IF(C9>0,E9/C9,0)", F_BOLD, PCT, border=True)

    put(ws, "A11", "2. ECL allowance by product and stage", F_BOLD)
    header(ws, 12, ["Product", "Stage 1", "Stage 2", "Stage 3", "Total", "Coverage ratio"])
    for i, p in enumerate(("Housing", "Consumer", "SME")):
        r = 13 + i
        put(ws, f"A{r}", p, border=True)
        for s, c in zip((1, 2, 3), "BCD"):
            put(ws, f"{c}{r}", f"=SUMIFS({ecl},{prod},$A{r},{stage},{s})", F_LINK, NUM, border=True)
        put(ws, f"E{r}", f"=SUM(B{r}:D{r})", fmt=NUM, border=True)
        put(ws, f"F{r}", f"=IF(SUMIFS({gca},{prod},$A{r})>0,E{r}/SUMIFS({gca},{prod},$A{r}),0)",
            F_LINK, PCT, border=True)
    put(ws, "A16", "Total", F_BOLD, border=True)
    for c in "BCDE":
        put(ws, f"{c}16", f"=SUM({c}13:{c}15)", F_BOLD, NUM, border=True)
    put(ws, "F16", "=IF(C9>0,E16/C9,0)", F_BOLD, PCT, border=True)

    put(ws, "A18", "3. Scenario analysis (current staging held constant)", F_BOLD)
    header(ws, 19, ["Scenario", "Weight", "ECL if scenario had 100% weight", "Weighted contribution",
                    "vs Base"])
    for i, n in enumerate(names):
        r = 20 + i
        b = cols["sc"][n]
        e12, elt = rng("ECL_Engine", b["e12"]), rng("ECL_Engine", b["elt"])
        lgd1, ead1 = rng("ECL_Engine", b["LGD"][0]), rng("ECL_Engine", b["EAD"][0])
        put(ws, f"A{r}", n, border=True, fill=PatternFill("solid", fgColor=FILL_SC[n]))
        put(ws, f"B{r}", f"={ADDR['w'][n]}", F_LINK, PCT, border=True)
        put(ws, f"C{r}", f"=SUMPRODUCT(({stage}=1)*{e12})+SUMPRODUCT(({stage}=2)*{elt})"
                         f"+SUMPRODUCT(({stage}=3)*{lgd1}*{ead1})", F_LINK, NUM, border=True)
        put(ws, f"D{r}", f"=B{r}*C{r}", fmt=NUM, border=True)
        put(ws, f"E{r}", f"=IF($C$20<>0,C{r}/$C$20-1,0)", fmt="0.0%;(0.0%);-", border=True)
    end = 19 + len(names)
    put(ws, f"A{end + 1}", "Probability-weighted ECL", F_BOLD, border=True)
    put(ws, f"B{end + 1}", f"=SUM(B20:B{end})", F_BOLD, PCT, border=True)
    put(ws, f"D{end + 1}", f"=SUM(D20:D{end})", F_BOLD, NUM, border=True)
    put(ws, f"A{end + 2}", "Forward-looking overlay vs Base-only ECL", border=True)
    put(ws, f"D{end + 2}", f"=D{end + 1}-C20", fmt="#,##0;(#,##0);-", border=True)
    put(ws, f"A{end + 3}", "The weighted ECL exceeds the Base ECL because losses are non-linear in the "
                           "macro variables: the downside adds more than the upside removes.", F_NOTE)

    r = end + 5
    put(ws, f"A{r}", "4. Journal entry at the reporting date", F_BOLD)
    put(ws, f"A{r + 1}", "Opening loss allowance (enter prior-period balance)", border=True)
    put(ws, f"B{r + 1}", 400000, F_INPUT, NUM, FILL_KEY, True)
    put(ws, f"C{r + 1}", "Illustrative figure - replace with the ledger balance.", F_NOTE)
    put(ws, f"A{r + 2}", "Required closing loss allowance", border=True)
    put(ws, f"B{r + 2}", "=E9", fmt=NUM, border=True)
    put(ws, f"A{r + 3}", "Movement to recognise in profit or loss", F_BOLD, border=True)
    put(ws, f"B{r + 3}", f"=B{r + 2}-B{r + 1}", F_BOLD, "#,##0;(#,##0);-", border=True)
    header(ws, r + 5, ["Account", "Debit", "Credit"])
    mv = f"B{r + 3}"
    put(ws, f"A{r + 6}", f'=IF({mv}>=0,"Dr Impairment loss on financial assets (P&L)",'
                         f'"Dr Loss allowance (statement of financial position)")', border=True)
    put(ws, f"B{r + 6}", f"=ABS({mv})", fmt=NUM, border=True)
    put(ws, f"C{r + 6}", None, border=True)
    put(ws, f"A{r + 7}", f'=IF({mv}>=0,"   Cr Loss allowance (statement of financial position)",'
                         f'"   Cr Impairment gain / reversal (P&L)")', border=True)
    put(ws, f"B{r + 7}", None, border=True)
    put(ws, f"C{r + 7}", f"=ABS({mv})", fmt=NUM, border=True)
    put(ws, f"A{r + 8}", "IFRS 9 para 5.5.8: the amount needed to adjust the loss allowance to the required "
                         "level is an impairment gain or loss in profit or loss.", F_NOTE)

    put(ws, f"A{r + 10}", "5. Reconciliation to the Python engine", F_BOLD)
    put(ws, f"A{r + 11}", "Total ECL computed in Python (ecl_models.compute_ecl)", border=True)
    put(ws, f"B{r + 11}", round(python_total, 2), F_INPUT, NUM, border=True)
    put(ws, f"A{r + 12}", "Difference Excel - Python", border=True)
    put(ws, f"B{r + 12}", f"=E9-B{r + 11}", fmt="#,##0.00;(#,##0.00);-", border=True)
    put(ws, f"C{r + 12}", f'=IF(ABS(B{r + 12})<1,"Reconciled","Inputs changed since last Python run")')
    put(ws, f"C{r + 11}", "Static value from the last run_all.py; differs once you edit inputs.", F_NOTE)

    ch = BarChart()
    ch.add_data(Reference(ws, min_col=5, min_row=5, max_row=8), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=6, max_row=8))
    style_chart(ch, "ECL allowance by stage (CU)", w=15, h=7.5)
    ch.legend = None
    ws.add_chart(ch, "H4")
    ch2 = BarChart()
    ch2.add_data(Reference(ws, min_col=3, min_row=19, max_row=end), titles_from_data=True)
    ch2.set_categories(Reference(ws, min_col=1, min_row=20, max_row=end))
    style_chart(ch2, "Total ECL under each scenario (CU)", w=15, h=7.5)
    ch2.legend = None
    ws.add_chart(ch2, "H20")


# ------------------------------------------------- Validation and history
def sheet_validation(wb, models, calib):
    ws = wb.create_sheet("Validation")
    widths(ws, {"A": 44, "B": 16, "C": 16, "D": 16, "E": 14})
    put(ws, "A1", "Model validation (Python output)", F_TITLE)
    put(ws, "A2", "Static results from ecl_models.py; PD metrics are on a 30% hold-out sample.", F_NOTE)
    r = 4
    for m in models:
        header(ws, r, [f"{m.name} model", "Value"])
        for k, v in m.metrics.items():
            r += 1
            put(ws, f"A{r}", k, border=True)
            is_rate = any(w in k for w in ("rate", "Mean"))
            put(ws, f"B{r}", round(float(v), 6), F_INPUT,
                "#,##0" if isinstance(v, int) else (PCT2 if is_rate else "0.000"), border=True)
        r += 2
    put(ws, f"A{r}", "Rule of thumb: AUC 0.70-0.80 acceptable, above 0.80 strong; "
                     "predicted and observed rates should track each other across deciles.", F_NOTE)
    r += 2
    put(ws, f"A{r}", "PD calibration by predicted-PD decile (hold-out)", F_BOLD)
    header(ws, r + 1, ["Decile (1 = lowest risk)", "Loans", "Mean predicted PD", "Observed default rate",
                       "Difference"])
    top = r + 2
    for i, row in enumerate(calib.itertuples(index=False)):
        rr = top + i
        put(ws, f"A{rr}", int(row.decile), F_INPUT, "0", border=True)
        put(ws, f"B{rr}", int(row.loans), F_INPUT, "#,##0", border=True)
        put(ws, f"C{rr}", round(float(row.predicted_pd), 6), F_INPUT, PCT2, border=True)
        put(ws, f"D{rr}", round(float(row.observed_dr), 6), F_INPUT, PCT2, border=True)
        put(ws, f"E{rr}", f"=D{rr}-C{rr}", fmt="0.00%;(0.00%);-", border=True)
    ch = LineChart()
    ch.add_data(Reference(ws, min_col=3, max_col=4, min_row=top - 1, max_row=top + 9), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=top, max_row=top + 9))
    style_chart(ch, "PD calibration: predicted vs observed", "Predicted-PD decile", "Default rate")
    ws.add_chart(ch, "G4")


def sheet_history(wb, macro, history):
    ws = wb.create_sheet("Macro_History")
    widths(ws, {"A": 11, "B": 13, "C": 15, "D": 14, "E": 15, "F": 17})
    put(ws, "A1", "Macro history used for estimation (synthetic)", F_TITLE)
    put(ws, "A2", "Source: data/macro_history.csv and data/loan_history.csv.gz. "
                  "Default rate = share of loans observed in the quarter that defaulted within 12 months.", F_NOTE)
    header(ws, 4, ["Quarter", "GDP growth (%)", "Unemployment (%)", "Lending rate (%)",
                   "House price growth (%)", "Observed default rate (%)"])
    dr = history.groupby("quarter")["default_12m"].mean() * 100
    for i, row in enumerate(macro.itertuples(index=False)):
        r = 5 + i
        put(ws, f"A{r}", row.quarter, F_INPUT)
        for c, v in zip("BCDE", (row.gdp_growth, row.unemployment, row.lending_rate, row.hpi_growth)):
            put(ws, f"{c}{r}", float(v), F_INPUT, "0.0")
        put(ws, f"F{r}", round(float(dr[row.quarter]), 3), F_INPUT, "0.00")
    last = 4 + len(macro)
    ch = LineChart()
    for col in (2, 3, 6):
        ch.add_data(Reference(ws, min_col=col, min_row=4, max_row=last), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=5, max_row=last))
    style_chart(ch, "Default rate moves with the economic cycle", "Quarter", "%", w=22, h=10)
    ws.add_chart(ch, "H4")
    ws.freeze_panes = "A5"


ADDR = {}


def build():
    history, macro, portfolio, scenarios = em.load_data()
    pd_m, lgd_m, ccf_m, calib = em.fit_all(history)
    models = [pd_m, lgd_m, ccf_m]
    python_ecl = em.compute_ecl(portfolio, scenarios, pd_m, lgd_m, ccf_m)
    names = list(scenarios["scenario"].unique())

    wb = Workbook()
    sheet_guide(wb)
    refs = sheet_coefficients(wb, models)
    ADDR.update(sheet_scenarios(wb, scenarios, refs))
    first, last = sheet_portfolio(wb, portfolio)
    cols = sheet_engine(wb, names, refs, first, last)
    sheet_summary(wb, names, cols, first, last, float(python_ecl["ecl"].sum()))
    sheet_validation(wb, models, calib)
    sheet_history(wb, macro, history)
    wb._sheets = [wb[n] for n in ("Guide", "Summary", "Scenarios", "Coefficients", "Portfolio",
                                  "ECL_Engine", "Validation", "Macro_History")]
    wb.save(OUT)
    print(f"Saved {OUT.name}; Python total ECL = {python_ecl['ecl'].sum():,.0f}")
    return python_ecl


if __name__ == "__main__":
    build()
