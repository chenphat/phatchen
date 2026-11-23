#!/usr/bin/env python3
"""Month-end close calculation checker.

Recomputes every tie-out from the six input CSVs, then checks that the
agents' outputs (journal_entries.csv, draft_income_statement.csv and
draft_balance_sheet.csv) agree with those results.

Usage:
    python3 checker/check_close.py --data practice_case/data --outputs reference_outputs
    python3 checker/check_close.py --data data --outputs outputs --period-end 2026-11-30 --materiality 1000

Standard library only. Exit code 0 = all checks passed, 1 = at least one FAIL.
"""
import argparse
import csv
import os
import sys
from datetime import date

TOL = 0.01
ALLOWANCE_RATES = {"Current": 0.01, "31-60": 0.05, "61-90": 0.20, "90+": 0.50}

results = []


def check(name, ok, detail=""):
    results.append(ok)
    print(f"  [{'PASS' if ok else 'FAIL'}] {name}" + (f"  ({detail})" if detail else ""))
    return ok


def eq(a, b):
    return abs(a - b) <= TOL


def num(v):
    v = (v or "").strip()
    return float(v) if v else 0.0


def d(v):
    v = (v or "").strip()
    return date.fromisoformat(v) if v else None


def load(folder, name, required=True):
    path = os.path.join(folder, name)
    if not os.path.exists(path):
        if required:
            print(f"  [FAIL] missing file: {path}")
            results.append(False)
        return None
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def money(x):
    return f"{x:,.2f}"


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--data", required=True, help="folder with the six input CSVs")
    ap.add_argument("--outputs", help="folder with journal_entries.csv and draft statements")
    ap.add_argument("--period-end", default=None, help="YYYY-MM-DD (default: latest TB/cash month-end)")
    ap.add_argument("--materiality", type=float, default=1000.0)
    a = ap.parse_args()

    tb = load(a.data, "trial_balance.csv")
    cash = load(a.data, "cash_transactions.csv")
    ar = load(a.data, "ar_aging.csv")
    apx = load(a.data, "ap_invoices.csv")
    fa = load(a.data, "fixed_assets.csv")
    pr = load(a.data, "payroll_register.csv")
    if None in (tb, cash, ar, apx, fa, pr):
        return finish()

    if a.period_end:
        pe = date.fromisoformat(a.period_end)
    else:
        pe = max(d(r["date"]) for r in cash if r["reference"] != "OPEN")
    period = pe.strftime("%Y-%m")
    mat = a.materiality
    print(f"Period end {pe}   materiality {money(mat)}\n")

    bal = {r["account"]: num(r["debit"]) - num(r["credit"]) for r in tb}  # debit-positive
    acct = {r["account"]: r for r in tb}

    print("02 Trial balance")
    dr = sum(num(r["debit"]) for r in tb)
    cr = sum(num(r["credit"]) for r in tb)
    check("Debits = credits", eq(dr, cr), f"{money(dr)} vs {money(cr)}")
    rev = -sum(bal[k] for k, r in acct.items() if r["type"] == "Revenue")
    exp = sum(bal[k] for k, r in acct.items() if r["type"] == "Expense")
    posted_op = rev - exp
    print(f"         posted operating profit = {money(posted_op)}")

    print("03 Cash & bank")
    book = sum(num(r["book_amount"]) for r in cash)
    check("Book cash = TB 1000", eq(book, bal.get("1000", 0)), money(book))
    bank = sum(num(r["bank_amount"]) for r in cash if r["bank_amount"].strip() and d(r["bank_date"]) <= pe)
    dit = sum(num(r["book_amount"]) for r in cash
              if num(r["book_amount"]) > 0 and r["reference"] != "OPEN"
              and (not r["bank_amount"].strip() or d(r["bank_date"]) > pe))
    osp = -sum(num(r["book_amount"]) for r in cash
               if num(r["book_amount"]) < 0 and (not r["bank_amount"].strip() or d(r["bank_date"]) > pe))
    bank_only = sum(num(r["bank_amount"]) for r in cash
                    if not r["book_amount"].strip() and r["bank_amount"].strip() and d(r["bank_date"]) <= pe)
    adj_bank = bank + dit - osp
    adj_book = book + bank_only
    print(f"         bank {money(bank)} + DIT {money(dit)} - outstanding {money(osp)} = {money(adj_bank)}")
    check("Adjusted bank = adjusted book", eq(adj_bank, adj_book), f"book {money(adj_book)}")

    print("04 Receivables")
    ar_tot = sum(num(r["amount"]) for r in ar)
    check("AR aging = TB 1100", eq(ar_tot, bal.get("1100", 0)), money(ar_tot))
    bucket_ok = True
    for r in ar:
        days = (pe - d(r["invoice_date"])).days
        b = "Current" if days <= 30 else "31-60" if days <= 60 else "61-90" if days <= 90 else "90+"
        if b != r["aging_bucket"] or int(num(r["days_outstanding"])) != days:
            bucket_ok = False
            print(f"         {r['invoice_no']}: {days} days -> {b}, file says {r['aging_bucket']}")
    check("Aging buckets re-aged correctly", bucket_ok)
    req = sum(num(r["amount"]) * ALLOWANCE_RATES[r["aging_bucket"]] for r in ar)
    booked = -bal.get("1150", 0)
    diff = req - booked
    print(f"         required allowance {money(req)} vs booked {money(booked)} -> difference {money(diff)}")
    exp_entries = []  # (account, debit_positive_amount, source)
    if abs(diff) >= mat:
        exp_entries += [("5700", diff, "allowance"), ("1150", -diff, "allowance")]

    print("05 Payables")
    ap_open = sum(num(r["amount"]) for r in apx if r["status"] == "Open" and r["posted_period"] <= period)
    check("Open AP = TB 2000", eq(ap_open, -bal.get("2000", 0)), money(ap_open))

    print("06 Revenue cutoff")
    deferred_should = 0.0
    for r in ar:
        s, e, amt = d(r["service_start"]), d(r["service_end"]), num(r["amount"])
        total_days = (e - s).days + 1
        after = max(0, (e - max(s, pe)).days) if e > pe else 0
        if s > pe:
            after = total_days
        deferred_should += amt * after / total_days
    check("Deferred revenue = unearned invoices", eq(deferred_should, -bal.get("2300", 0)),
          f"{money(deferred_should)} vs TB {money(-bal.get('2300', 0))}")

    print("07 Expense cutoff (search for unrecorded liabilities)")
    accrual = 0.0
    for r in apx:
        if r["posted_period"].strip():
            continue
        s, e, amt = d(r["service_start"]), d(r["service_end"]), num(r["amount"])
        if s > pe:
            print(f"         {r['vendor']} {r['invoice_no']}: service starts {s}, correctly excluded")
            continue
        days = (e - s).days + 1
        in_period = amt * ((min(e, pe) - s).days + 1) / days
        print(f"         {r['vendor']} {r['invoice_no']}: in-period amount {money(in_period)}")
        if in_period >= mat:
            accrual += in_period
            exp_entries += [(r["gl_account"], in_period, r["invoice_no"]), ("2100", -in_period, r["invoice_no"])]
    print(f"         expected accrual {money(accrual)}")

    print("08 Fixed assets")
    for r in fa:
        m = (num(r["cost"]) - num(r["salvage_value"])) / num(r["useful_life_months"])
        acc = min(m * num(r["months_depreciated_to_period_end"]), num(r["cost"]) - num(r["salvage_value"]))
        check(f"{r['asset_id']} depreciation recomputed",
              eq(m, num(r["period_depreciation"])) and eq(acc, num(r["accumulated_depreciation"])),
              f"monthly {money(m)}, accumulated {money(acc)}")
    check("Register cost = TB 1500", eq(sum(num(r["cost"]) for r in fa), bal.get("1500", 0)))
    check("Register acc. dep. = TB 1510", eq(sum(num(r["accumulated_depreciation"]) for r in fa), -bal.get("1510", 0)))
    check("Register period dep. = TB 5400", eq(sum(num(r["period_depreciation"]) for r in fa), bal.get("5400", 0)))

    print("09 Payroll & tax")
    check("Gross + employer taxes = TB 5000",
          eq(sum(num(r["gross_wages"]) + num(r["employer_taxes"]) for r in pr), bal.get("5000", 0)))
    check("Net = gross - withholdings",
          all(eq(num(r["net_pay"]), num(r["gross_wages"]) - num(r["employee_withholdings"])) for r in pr))
    unremitted = sum(num(r["employee_withholdings"]) + num(r["employer_taxes"]) for r in pr
                     if not r["withholdings_remitted_date"].strip() or d(r["withholdings_remitted_date"]) > pe)
    check("Unremitted withholdings + employer taxes = TB 2200", eq(unremitted, -bal.get("2200", 0)), money(unremitted))
    last_end = max(d(r["pay_period"].split(" to ")[1]) for r in pr)
    check("Pay periods cover the whole month", last_end >= pe, f"last period ends {last_end}")

    # Expected draft figures
    adj = {}
    for k, v, _ in exp_entries:
        adj[k] = adj.get(k, 0) + v
    exp_adj_exp = sum(v for k, v in adj.items() if k in acct and acct[k]["type"] == "Expense" or k == "5700")
    exp_adj_rev = -sum(v for k, v in adj.items() if k in acct and acct[k]["type"] == "Revenue")
    draft_op = posted_op + exp_adj_rev - exp_adj_exp
    print(f"\nExpected: posted OP {money(posted_op)} -> draft OP {money(draft_op)}; "
          f"book cash {money(book)} -> {money(book + adj.get('1000', 0))}")

    if not a.outputs:
        return finish()

    print(f"\n10 Journal entries ({a.outputs})")
    je = load(a.outputs, "journal_entries.csv")
    if je is not None:
        by_id = {}
        for r in je:
            by_id.setdefault(r["je_id"], []).append(r)
        for jid, rows in by_id.items():
            check(f"{jid} balances", eq(sum(num(r["debit"]) for r in rows), sum(num(r["credit"]) for r in rows)))
            check(f"{jid} accounts valid", all(r["account"] in acct for r in rows))
            check(f"{jid} not approved/posted by an agent",
                  all(r["status"] == "Proposed" and not r["approved_by"].strip() for r in rows))
        je_adj = {}
        for r in je:
            je_adj[r["account"]] = je_adj.get(r["account"], 0) + num(r["debit"]) - num(r["credit"])
        keys = set(je_adj) | set(adj)
        check("Entries match the expected adjustments",
              all(eq(je_adj.get(k, 0), adj.get(k, 0)) for k in keys),
              ", ".join(f"{k}: {money(je_adj.get(k, 0))}" for k in sorted(je_adj)))

    print("11 Draft income statement")
    is_ = load(a.outputs, "draft_income_statement.csv")
    if is_ is not None:
        op = next((r for r in is_ if r["account"] == "OP"), None)
        check("Posted operating profit", op is not None and eq(num(op["posted"]), posted_op))
        check("Draft operating profit", op is not None and eq(num(op["draft"]), draft_op),
              f"expected {money(draft_op)}")
        check("Every line: posted + adjustments = draft",
              all(eq(num(r["posted"]) + num(r["adjustments"]), num(r["draft"])) for r in is_))

    print("12 Draft balance sheet")
    bs = load(a.outputs, "draft_balance_sheet.csv")
    if bs is not None:
        rows = {r["account"]: r for r in bs}
        for col in ("posted", "draft"):
            ta = num(rows.get("TA", {}).get(col))
            tle = num(rows.get("TL", {}).get(col)) + num(rows.get("TE", {}).get(col))
            check(f"{col.title()} balance sheet balances", eq(ta, tle), f"{money(ta)} vs {money(tle)}")
        c = rows.get("1000")
        check("Draft cash = expected", c is not None and eq(num(c["draft"]), book + adj.get("1000", 0)),
              "accruals do not move cash")
    return finish()


def finish():
    n_fail = results.count(False)
    print(f"\n{len(results) - n_fail} passed, {n_fail} failed")
    sys.exit(1 if n_fail else 0)


if __name__ == "__main__":
    main()
