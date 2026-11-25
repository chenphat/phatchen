"""Completeness, accuracy and reliability checks on the cleaned TB and GL.

Each check returns a row for the DQ_Log sheet:
  id, area, check, result (PASS / FAIL / WARN / INFO / N/A), exceptions, detail
"""
import pandas as pd

from .taxonomy import LEAF, FACE, is_pl

TOL = 0.005


def _fmt(items, n=6):
    items = list(items)
    return "; ".join(str(i) for i in items[:n]) + (f"; ... (+{len(items) - n} more)" if len(items) > n else "")


def run_checks(tb, gl, mapping, period, label):
    """tb/gl: cleaned frames for one year. period: (start, end). label: 'CY' or 'PY'."""
    res = []

    def add(cid, area, check, ok, count=0, detail="", level="FAIL"):
        res.append({"id": f"{label}-{cid}", "area": area, "check": check,
                    "result": "PASS" if ok else level, "exceptions": int(count), "detail": detail})

    # ---- trial balance ----
    s_open, s_close = round(tb["opening"].sum(), 2), round(tb["closing"].sum(), 2)
    add("TB01", "Trial balance", "Opening balances: total debits equal total credits", abs(s_open) < TOL, 0,
        f"Net difference {s_open:,.2f}")
    add("TB02", "Trial balance", "Closing balances: total debits equal total credits", abs(s_close) < TOL, 0,
        f"Net difference {s_close:,.2f}")
    cast = tb[(tb["opening"] + tb["debit"] - tb["credit"] - tb["closing"]).abs() > TOL]
    add("TB03", "Trial balance", "Arithmetic: opening + debits - credits = closing for every account",
        cast.empty, len(cast), _fmt(cast["account"]))
    dup = tb[tb["account"].duplicated(keep=False)]
    add("TB04", "Trial balance", "No duplicated account codes", dup.empty, dup["account"].nunique(),
        _fmt(dup["account"].unique()))
    noname = tb[tb["account_name"] == ""]
    add("TB05", "Trial balance", "Every account has a name", noname.empty, len(noname), _fmt(noname["account"]),
        "WARN")

    leaf_of = dict(zip(mapping["account"], mapping["leaf"]))
    pl_open = tb[tb["account"].map(lambda a: is_pl(leaf_of.get(a, ""))) & (tb["opening"].abs() > TOL)]
    add("TB06", "Trial balance", "Income and expense accounts open at zero (prior year was closed)",
        pl_open.empty, len(pl_open), _fmt(pl_open["account"]))

    # abnormal balances
    normal = {"A": 1, "C": 1, "X": 1, "T": 1, "L": -1, "E": -1, "I": -1}
    contra = {l for l, v in LEAF.items() if l.endswith(".D") or l in ("TOR.ALLOW", "RE.DIV", "OI.DISP", "OI.FV", "RE")}
    cls_of = dict(zip(mapping["account"], mapping["cls"]))
    abn = [f"{r.account} {r.account_name} ({r.closing:,.2f})" for r in tb.itertuples()
           if leaf_of.get(r.account) not in contra and abs(r.closing) > TOL
           and r.closing * normal.get(cls_of.get(r.account, "X"), 1) < 0]
    add("TB07", "Trial balance", "No balances with a sign opposite to the account's nature", not abn, len(abn),
        _fmt(abn, 4), "WARN")

    # ---- general ledger ----
    if gl is None or gl.empty:
        add("GL00", "General ledger", "General ledger supplied for this year", False, 0,
            "No GL: roll-forward notes and cash flow lines based on GL movement types will be blank", "WARN")
        return res
    td, tc = gl["debit"].sum(), gl["credit"].sum()
    add("GL01", "General ledger", "Total debits equal total credits", abs(td - tc) < TOL, 0,
        f"Debits {td:,.2f}; credits {tc:,.2f}; difference {td - tc:,.2f}")
    jb = (gl.groupby("journal")["debit"].sum() - gl.groupby("journal")["credit"].sum()).round(2)
    unbal = jb[jb.abs() > TOL]
    add("GL02", "General ledger", "Every journal entry balances", unbal.empty, len(unbal),
        _fmt(f"{j} ({v:,.2f})" for j, v in unbal.items()))
    for cid, col, text in (("GL03", "account", "account code"), ("GL04", "journal", "journal number")):
        miss = gl[gl[col] == ""]
        add(cid, "General ledger", f"No line without a(n) {text}", miss.empty, len(miss),
            _fmt(f"row {r}" for r in miss["row"]))
    nodate = gl[gl["date"].isna()]
    add("GL05", "General ledger", "Every line has a valid posting date", nodate.empty, len(nodate),
        _fmt(f"row {r}" for r in nodate["row"]))
    dated = gl[gl["date"].notna()]
    outside = dated[(dated["date"] < period[0]) | (dated["date"] > period[1])]
    add("GL06", "General ledger", f"All postings fall within the period {period[0]} to {period[1]} (cut-off)",
        outside.empty, len(outside),
        _fmt(f"row {r.row}: {r.date} journal {r.journal}" for r in outside.itertuples()), "WARN")
    zero = gl[(gl["debit"].abs() < TOL) & (gl["credit"].abs() < TOL)]
    add("GL07", "General ledger", "No zero-value lines", zero.empty, len(zero),
        _fmt(f"row {r}" for r in zero["row"]), "WARN")
    both = gl[(gl["debit"].abs() > TOL) & (gl["credit"].abs() > TOL)]
    add("GL08", "General ledger", "No line with both a debit and a credit amount", both.empty, len(both),
        _fmt(f"row {r}" for r in both["row"]), "WARN")
    key = ["date", "journal", "account", "description", "reference", "debit", "credit"]
    dups = gl[gl.duplicated(subset=key, keep=False)]
    add("GL09", "General ledger", "No remaining exact duplicate lines", dups.empty, len(dups),
        _fmt(f"row {r}" for r in dups["row"]), "WARN")
    multi = dated.groupby("journal")["date"].nunique()
    multi = multi[multi > 1]
    add("GL10", "General ledger", "Each journal carries a single posting date", multi.empty, len(multi),
        _fmt(multi.index), "WARN")

    # ---- GL to TB reconciliation ----
    glm = (gl.groupby("account")["debit"].sum() - gl.groupby("account")["credit"].sum()).round(2)
    tbm = (tb.groupby("account")["debit"].sum() - tb.groupby("account")["credit"].sum()).round(2)
    rec = pd.concat([glm.rename("gl"), tbm.rename("tb")], axis=1).fillna(0.0)
    diff = rec[(rec["gl"] - rec["tb"]).abs() > TOL]
    add("RC01", "GL to TB", "GL net movement agrees to TB movement for every account", diff.empty, len(diff),
        _fmt(f"{a} (GL {r.gl:,.2f} v TB {r.tb:,.2f})" for a, r in diff.iterrows()))
    not_tb = sorted(set(gl["account"]) - set(tb["account"]) - {""})
    add("RC02", "GL to TB", "Every GL account exists in the trial balance", not not_tb, len(not_tb), _fmt(not_tb))

    # ---- analytical / journal-entry indicators (information only) ----
    wk = dated[dated["date"].map(lambda d: d.weekday() >= 5)]
    add("JE01", "Journal review", "Postings dated on a weekend", wk.empty, wk["journal"].nunique(),
        _fmt(wk["journal"].unique(), 8), "INFO")
    amt = gl["debit"] + gl["credit"]
    big = gl[(amt >= 10000) & (amt % 1000 == 0)]
    add("JE02", "Journal review", "Large round-sum lines (>= 10,000 and a multiple of 1,000)", big.empty,
        len(big), _fmt(big["journal"].unique(), 8), "INFO")
    nodesc = gl[gl["description"] == ""]
    add("JE03", "Journal review", "Lines without a description", nodesc.empty, len(nodesc),
        _fmt(nodesc["journal"].unique(), 8), "INFO")
    return res


def continuity_checks(tb_cy, tb_py, mapping):
    """Opening balances of the current year agree to the closing balances of the prior year."""
    res = []
    if tb_py is None or tb_py.empty:
        return [{"id": "CN00", "area": "Continuity", "check": "Prior-year trial balance supplied",
                 "result": "WARN", "exceptions": 0, "detail": "No prior-year TB: comparatives will be blank"}]
    leaf_of = dict(zip(mapping["account"], mapping["leaf"]))
    closes_to_re = lambda a: is_pl(leaf_of.get(a, "")) or leaf_of.get(a) in ("RE", "RE.DIV")
    cy, py = tb_cy.set_index("account")["opening"], tb_py.set_index("account")["closing"]
    both = pd.concat([cy.rename("cy"), py.rename("py")], axis=1).fillna(0.0)
    bs = both[~both.index.map(closes_to_re)]
    diff = bs[(bs["cy"] - bs["py"]).abs() > TOL]
    res.append({"id": "CN01", "area": "Continuity",
                "check": "Current-year opening balance = prior-year closing balance (balance sheet accounts)",
                "result": "PASS" if diff.empty else "FAIL", "exceptions": len(diff),
                "detail": _fmt(f"{a} (open {r.cy:,.2f} v PY close {r.py:,.2f})" for a, r in diff.iterrows())})
    re_ = both[both.index.map(closes_to_re)]
    d = round(re_["cy"].sum() - re_["py"].sum(), 2)
    res.append({"id": "CN02", "area": "Continuity",
                "check": "Opening retained earnings = prior-year retained earnings + profit - dividends",
                "result": "PASS" if abs(d) < TOL else "FAIL", "exceptions": 0, "detail": f"Difference {d:,.2f}"})
    return res


def mapping_checks(mapping):
    bad = mapping[~mapping["leaf"].isin(LEAF)]
    review = mapping[mapping["basis"].str.contains("review", case=False)]
    exp_nonat = mapping[mapping["leaf"].isin(["COS", "DIST", "ADMIN", "OTHEXP"]) & (mapping["nature"] == "")]
    return [
        {"id": "MP01", "area": "Mapping", "check": "Every account is mapped to an AFS component",
         "result": "PASS" if bad.empty else "FAIL", "exceptions": len(bad), "detail": _fmt(bad["account"])},
        {"id": "MP02", "area": "Mapping", "check": "No account mapped by fallback (needs reviewer confirmation)",
         "result": "PASS" if review.empty else "WARN", "exceptions": len(review),
         "detail": _fmt(f"{r.account} {r.account_name}" for r in review.itertuples())},
        {"id": "MP03", "area": "Mapping", "check": "Every expense account has a nature for the expenses note",
         "result": "PASS" if exp_nonat.empty else "WARN", "exceptions": len(exp_nonat),
         "detail": _fmt(exp_nonat["account"])},
    ]
