"""Generate a two-year teaching data set (general ledger + trial balance) with deliberate data-quality defects.

The entity is fictitious. The ledgers are internally consistent; the exported
files are then "dirtied" (title rows, text numbers, mixed date formats, a
duplicated line, a subtotal row, name variants ...) so the cleaning and
checking steps have something to find.
"""
import json
import random
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook

COA = {
    "1010": "Cash on hand", "1020": "Bank - current account", "1030": "Short-term bank deposits",
    "1100": "Trade receivables", "1110": "Allowance for impairment of trade receivables",
    "1120": "Other receivables", "1130": "Prepayments", "1140": "Interest receivable",
    "1200": "Inventories - merchandise", "1500": "Land", "1510": "Buildings",
    "1515": "Accumulated depreciation - buildings", "1520": "Warehouse equipment",
    "1525": "Accumulated depreciation - warehouse equipment", "1530": "Motor vehicles",
    "1535": "Accumulated depreciation - motor vehicles", "1540": "Furniture and office equipment",
    "1545": "Accumulated depreciation - furniture and office equipment", "1600": "Computer software",
    "1605": "Accumulated amortisation - software",
    "2010": "Trade payables", "2020": "Accrued expenses", "2030": "VAT and other taxes payable",
    "2040": "Interest payable", "2050": "Dividends payable", "2060": "Customer advances",
    "2100": "Income tax payable", "2200": "Bank loan - current portion", "2210": "Provision for warranties",
    "2500": "Bank loan - non-current portion", "2600": "Deferred tax liability",
    "3010": "Share capital", "3020": "Share premium", "3100": "Revaluation reserve",
    "3200": "Retained earnings", "3300": "Dividends declared",
    "4010": "Sales of goods", "4020": "Service revenue", "4100": "Other income",
    "4110": "Gain on disposal of property, plant and equipment", "4200": "Interest income",
    "5010": "Cost of goods sold", "5020": "Freight inwards", "5030": "Depreciation - warehouse",
    "6010": "Salaries and wages - administrative", "6020": "Salaries - sales staff", "6030": "Office rent",
    "6040": "Utilities", "6050": "Advertising and promotion", "6060": "Delivery expenses",
    "6070": "Professional fees", "6080": "Depreciation - administrative",
    "6085": "Depreciation - distribution vehicles", "6090": "Amortisation of software",
    "6100": "Impairment loss on trade receivables", "6110": "Warranty expense",
    "6120": "Repairs and maintenance", "6130": "Entertainment", "6140": "Bank charges",
    "6150": "Foreign exchange loss", "7010": "Interest expense on bank loans",
    "8010": "Income tax expense - current", "8020": "Deferred tax expense",
}
OPENING = {
    "1010": 2000, "1020": 80000, "1100": 120000, "1110": -2880, "1130": 6000, "1200": 150000,
    "1500": 200000, "1510": 300000, "1515": -60000, "1520": 80000, "1525": -24000, "1530": 90000,
    "1535": -36000, "1540": 40000, "1545": -16000, "1600": 30000, "1605": -12000,
    "2010": -95000, "2020": -12000, "2030": -8000, "2100": -15000, "2200": -40000, "2210": -5000,
    "2500": -200000, "2600": -6000, "3010": -400000, "3020": -50000,
}
TAX_RATE = 0.20
PL_ACCOUNTS = [a for a in COA if a[0] in "45678"]
# overdue bands as (share of trade receivables, share of that band impaired on objective evidence, 11.21-11.22)
AGEING = [(0.08, 0.0), (0.04, 0.1), (0.02, 0.5), (0.01, 1.0)]
ALLOWANCE_RATE = sum(p * a for p, a in AGEING)


class Ledger:
    def __init__(self, year, opening):
        self.year, self.lines, self.n, self.ajes = year, [], 0, []
        self.bal = dict(opening)

    def post(self, d, desc, entries, ref=""):
        """entries: [(account, amount)] with debit positive; must sum to zero."""
        entries = [(a, round(v, 2)) for a, v in entries if round(v, 2) != 0]
        assert abs(sum(v for _, v in entries)) < 0.005, (desc, entries)
        self.n += 1
        jno = f"JV{str(self.year)[2:]}-{self.n:04d}"
        for a, v in entries:
            self.lines.append({"date": d, "journal": jno, "account": a, "name": COA[a], "desc": desc,
                               "ref": ref, "debit": v if v > 0 else 0.0, "credit": -v if v < 0 else 0.0})
            self.bal[a] = round(self.bal.get(a, 0.0) + v, 2)
        return jno

    def adjust(self, desc, entries, ref=""):
        """Year-end adjustment kept out of the ledger: it is delivered as an adjusting journal entry (AJE)."""
        self.aje_no = getattr(self, "aje_no", 0) + 1
        for a, v in entries:
            v = round(v, 2)
            if v:
                self.ajes.append({"no": self.aje_no, "account": a, "description": desc,
                                  "debit": v if v > 0 else 0, "credit": -v if v < 0 else 0})
                self.bal[a] = round(self.bal.get(a, 0.0) + v, 2)

    def movement(self, account):
        return round(sum(l["debit"] - l["credit"] for l in self.lines if l["account"] == account), 2)


def _workday(y, m, d):
    dt = date(y, m, min(d, 28))
    while dt.weekday() >= 5:
        dt += timedelta(days=1)
    return dt


def build_year(year, opening, rng, current):
    g = 1.12 if current else 1.0
    L = Ledger(year, opening)
    eom = lambda m: (date(year, m + 1, 1) - timedelta(days=1)) if m < 12 else date(year, 12, 31)
    wd = lambda m, d: _workday(year, m, d)

    # January: settle opening accruals, taxes and release the opening prepayment
    L.post(wd(1, 10), "Payment of prior-year accrued expenses", [("2020", -opening.get("2020", 0)), ("1020", opening.get("2020", 0))], "PV")
    L.post(wd(1, 15), "Payment of VAT and other taxes", [("2030", -opening.get("2030", 0)), ("1020", opening.get("2030", 0))], "PV")
    L.post(wd(1, 31), "Release of prepaid rent", [("6030", opening.get("1130", 0)), ("1130", -opening.get("1130", 0))], "GJ")
    if opening.get("2060"):
        L.post(wd(1, 20), "Goods delivered against customer advances", [("2060", -opening["2060"]), ("4010", opening["2060"])], "SI")
        L.post(wd(1, 20), "Cost of goods delivered", [("5010", -opening["2060"] * 0.62), ("1200", opening["2060"] * 0.62)], "SI")
    if opening.get("1140"):
        L.post(wd(1, 5), "Interest received on deposits", [("1020", opening["1140"]), ("1140", -opening["1140"])], "RV")
    if opening.get("2050"):
        L.post(wd(1, 25), "Payment of prior-year dividend", [("2050", -opening["2050"]), ("1020", opening["2050"])], "PV")
    L.post(wd(3, 20), "Settlement of prior-year income tax", [("2100", -opening.get("2100", 0)), ("1020", opening.get("2100", 0))], "PV")

    loan_cur, loan_nc = -opening.get("2200", 0), -opening.get("2500", 0)
    cost = {a: opening.get(a, 0) for a in ("1510", "1520", "1530", "1540", "1600")}
    for m in range(1, 13):
        season = 1 + 0.15 * (m in (11, 12)) - 0.10 * (m in (1, 2))
        # sales on credit (several invoices) with perpetual cost of sales
        month_sales = 0.0
        for k in range(6):
            s = round(95000 * g * season / 6 * rng.uniform(0.85, 1.15), 2)
            month_sales += s
            d = wd(m, 3 + 4 * k)
            L.post(d, f"Credit sales invoice batch {m:02d}/{k + 1}", [("1100", s), ("4010", -s)], f"SI-{m:02d}{k + 1}")
            L.post(d, f"Cost of goods sold batch {m:02d}/{k + 1}", [("5010", s * 0.62), ("1200", -s * 0.62)], f"SI-{m:02d}{k + 1}")
        L.post(wd(m, 12), "Service fees received", [("1020", 8000 * g), ("4020", -8000 * g)], "RV")
        L.post(wd(m, 14), "Sundry income - sale of packaging scrap", [("1010", 500), ("4100", -500)], "RV")
        for k in range(3):
            c = round(month_sales * 0.975 / 3, 2)
            L.post(wd(m, 8 + 7 * k), "Receipts from customers", [("1020", c), ("1100", -c)], "RV")
        for k in range(3):
            p = round(month_sales * 0.62 * 1.03 / 3, 2)
            L.post(wd(m, 5 + 7 * k), "Purchase of merchandise on credit", [("1200", p), ("2010", -p)], f"PI-{m:02d}{k + 1}")
            L.post(wd(m, 9 + 7 * k), "Payments to suppliers", [("2010", p * 0.985), ("1020", -p * 0.985)], "PV")
        L.post(wd(m, 6), "Freight on purchases", [("5020", 1800 * g), ("1020", -1800 * g)], "PV")
        L.post(eom(m), "Payroll - administration", [("6010", 9000 * g), ("1020", -9000 * g)], "PR")
        L.post(eom(m), "Payroll - sales staff", [("6020", 6000 * g), ("1020", -6000 * g)], "PR")
        L.post(wd(m, 2), "Office rent", [("6030", 3000), ("1020", -3000)], "PV")
        L.post(wd(m, 18), "Electricity, water and internet", [("6040", 1500 * g), ("1020", -1500 * g)], "PV")
        L.post(wd(m, 16), "Advertising", [("6050", 2500 * g), ("1020", -2500 * g)], "PV")
        L.post(wd(m, 22), "Delivery to customers", [("6060", 2000 * g), ("1020", -2000 * g)], "PV")
        L.post(wd(m, 24), "Repairs and maintenance", [("6120", 800), ("1010", -400), ("1020", -400)], "PV")
        L.post(wd(m, 27), "Client entertainment", [("6130", 400), ("1020", -400)], "PV")
        L.post(eom(m), "Bank charges", [("6140", 150), ("1020", -150)], "BK")
        L.post(wd(m, 20), "Income tax instalment", [("2100", 1500 * g), ("1020", -1500 * g)], "PV")
        if m % 3 == 0:
            L.post(wd(m, 26), "Professional fees", [("6070", 4000), ("1020", -4000)], "PV")
            L.post(wd(m, 26), "Warranty claims settled", [("2210", 1100 * g), ("1020", -1100 * g)], "PV")
        if m in (4, 9):
            L.post(wd(m, 19), "Exchange loss on settlement of foreign supplier", [("6150", 650), ("1020", -650)], "BK")

        # one-off events
        if not current and m == 2:
            L.post(wd(2, 10), "Placement of 3-month deposit", [("1030", 30000), ("1020", -30000)], "BK")
        if not current and m == 5:
            L.post(wd(5, 15), "Purchase of office equipment", [("1540", 12000), ("1020", -12000)], "FA")
            cost["1540"] += 12000
        if not current and m == 6:
            L.post(wd(6, 20), "Dividend declared", [("3300", 20000), ("2050", -20000)], "GJ")
        if not current and m == 8:
            L.post(wd(8, 15), "Dividend paid", [("2050", 20000), ("1020", -20000)], "PV")
        if current and m == 2:
            L.post(wd(2, 12), "Purchase of ERP software licence", [("1600", 8000), ("1020", -8000)], "FA")
            cost["1600"] += 8000
        if current and m == 3:
            L.post(wd(3, 10), "Purchase of warehouse racking and forklift", [("1520", 45000), ("1020", -45000)], "FA")
            cost["1520"] += 45000
        if current and m == 4:
            L.post(wd(4, 1), "Drawdown of new 5-year bank loan", [("1020", 100000), ("2500", -100000)], "BK")
            loan_nc += 100000
        if current and m == 6:
            L.post(wd(6, 16), "Issue of 50,000 ordinary shares at 1.20", [("1020", 60000), ("3010", -50000), ("3020", -10000)], "GJ")
        if current and m == 7:
            L.post(wd(7, 14), "Disposal of delivery van", [("1020", 15000), ("1535", 18000), ("1530", -30000), ("4110", -3000)], "FA")
            cost["1530"] -= 30000
        if current and m == 8:
            L.post(wd(8, 5), "Purchase of delivery truck", [("1530", 38000), ("1020", -38000)], "FA")
            cost["1530"] += 38000
        if current and m == 11:
            L.post(wd(11, 18), "Dividend declared", [("3300", 30000), ("2050", -30000)], "GJ")
        if current and m == 12:
            L.post(wd(12, 10), "Part payment of dividend", [("2050", 18000), ("1020", -18000)], "PV")

        # depreciation and amortisation
        L.post(eom(m), "Depreciation - buildings and office equipment",
               [("6080", cost["1510"] * 0.02 / 12 + cost["1540"] * 0.10 / 12),
                ("1515", -cost["1510"] * 0.02 / 12), ("1545", -cost["1540"] * 0.10 / 12)], "GJ")
        L.post(eom(m), "Depreciation - warehouse equipment", [("5030", cost["1520"] * 0.10 / 12), ("1525", -cost["1520"] * 0.10 / 12)], "GJ")
        L.post(eom(m), "Depreciation - vehicles", [("6085", cost["1530"] * 0.20 / 12), ("1535", -cost["1530"] * 0.20 / 12)], "GJ")
        L.post(eom(m), "Amortisation of software", [("6090", cost["1600"] * 0.20 / 12), ("1605", -cost["1600"] * 0.20 / 12)], "GJ")

        # bank loan: monthly interest accrual, quarterly interest payment, half-yearly principal
        interest = round((loan_cur + loan_nc) * 0.08 / 12, 2)
        L.post(eom(m), "Interest accrued on bank loans", [("7010", interest), ("2040", -interest)], "GJ")
        if m % 3 == 0:
            due = round(-L.bal.get("2040", 0) - interest, 2)
            L.post(eom(m), "Interest paid on bank loans", [("2040", due), ("1020", -due)], "BK")
        if m in (6, 12):
            L.post(eom(m), "Loan principal instalment", [("2200", loan_cur / (2 if m == 6 else 1)), ("1020", -loan_cur / (2 if m == 6 else 1))], "BK")
            loan_cur -= loan_cur / (2 if m == 6 else 1)

    # ---- year-end adjustments ----
    ye = date(year, 12, 31)
    adj = L.adjust if current else L.post     # current year: closing adjustments are proposed as AJEs, not yet in the ledger
    next_cur = 60000 if current else 40000
    L.post(ye, "Reclassification of loan instalments due within 12 months", [("2500", next_cur), ("2200", -next_cur)], "GJ")
    L.post(ye, "Rent paid in advance for next year", [("1130", 7000 * g), ("1020", -7000 * g)], "PV")
    adj(*([ye] if not current else []), "Accrual of bonuses and professional fees", [("6010", 9500 * g), ("6070", 4000), ("2020", -9500 * g - 4000)], "GJ")
    L.post(ye, "Payroll taxes withheld, payable next month", [("6010", 8500 * g), ("2030", -8500 * g)], "GJ")
    L.post(ye, "Advances received from customers for January deliveries", [("1020", 9000 * g), ("2060", -9000 * g)], "RV")
    L.post(ye, "Write-off of irrecoverable receivable", [("1110", 2000), ("1100", -2000)], "GJ")
    target = round(L.bal["1100"] * ALLOWANCE_RATE, 2)
    adj(*([ye] if not current else []), "Impairment allowance for overdue receivables with objective evidence of impairment", [("6100", target + L.bal["1110"]), ("1110", -(target + L.bal["1110"]))], "GJ")
    L.post(ye, "Warranty provision for the year's sales", [("6110", 6000 * g), ("2210", -6000 * g)], "GJ")
    dep_int = round(L.bal.get("1030", 0) * 0.04 * (10 / 12 if not current else 1), 2)
    L.post(ye, "Interest accrued on short-term deposits", [("1140", dep_int), ("4200", -dep_int)], "GJ")
    if current:
        L.post(ye, "Revaluation of land to fair value (independent valuer)", [("1500", 50000), ("3100", -40000), ("2600", -10000)], "GJ")
    pbt = -sum(L.bal.get(a, 0) for a in PL_ACCOUNTS)
    entertainment, temp_diff = L.bal["6130"], 5000
    current_tax = round(TAX_RATE * (pbt + entertainment - temp_diff), 2)
    adj(*([ye] if not current else []), "Current income tax for the year", [("8010", current_tax), ("2100", -current_tax)], "GJ")
    adj(*([ye] if not current else []), "Deferred tax on accelerated tax depreciation", [("8020", TAX_RATE * temp_diff), ("2600", -TAX_RATE * temp_diff)], "GJ")
    return L, entertainment


def close_year(bal):
    """Post-closing balances carried forward as next year's opening balances."""
    nxt = {a: v for a, v in bal.items() if a not in PL_ACCOUNTS and a != "3300" and round(v, 2) != 0}
    nxt["3200"] = round(bal.get("3200", 0) + sum(bal.get(a, 0) for a in PL_ACCOUNTS) + bal.get("3300", 0), 2)
    return nxt


def _write_tb(path, entity, year, opening, L, rng, dirty):
    wb = Workbook()
    ws = wb.active
    ws.title = "Trial Balance"
    ws.append([entity])
    ws.append([f"Trial balance as at 31 December {year}"])
    ws.append([])
    ws.append(["Account No.", "Account Description", "Opening Balance Dr", "Opening Balance Cr", "Debit", "Credit",
               "Closing Dr", "Closing Cr"])
    tot = [0.0] * 6
    for n, a in enumerate(sorted(COA)):
        o = opening.get(a, 0.0)
        dr = round(sum(l["debit"] for l in L.lines if l["account"] == a), 2)
        cr = round(sum(l["credit"] for l in L.lines if l["account"] == a), 2)
        c = round(o + dr - cr, 2)
        if o == 0 and dr == 0 and cr == 0:
            continue
        vals = [max(o, 0), max(-o, 0), dr, cr, max(c, 0), max(-c, 0)]
        tot = [round(t + v, 2) for t, v in zip(tot, vals)]
        code = int(a)
        if dirty and a in ("1100", "2010"):
            code = f" {a}  "                                   # stray spaces
        if dirty and a in ("1200", "4010", "6010"):
            vals = [f"{v:,.2f}" for v in vals]                 # numbers stored as text
        ws.append([code, COA[a]] + vals)
        if dirty and n in (8, 30):
            ws.append([])
    ws.append(["", "TOTAL"] + tot)
    wb.save(path)


def _write_gl(path, entity, year, L, rng, dirty, cut=False):
    wb = Workbook()
    ws = wb.active
    ws.title = "GL"
    ws.append([entity])
    ws.append([f"General ledger - all journals - year ended 31 December {year}"])
    ws.append(["Posting Date", "JV No.", "Acc Code", "Account Name", "Narration", "Ref", "Debit Amount",
               "Credit Amount"])
    n = len(L.lines)
    dup_at, cut_j = int(n * 0.37), L.lines[-9]["journal"]
    neg_at = next(i for i in range(int(n * 0.61), n) if L.lines[i]["credit"] and i % 17)
    for i, l in enumerate(L.lines):
        d, name, dr, cr = l["date"], l["name"], l["debit"], l["credit"]
        if dirty:
            if i % 11 == 0:
                d = d.strftime("%d/%m/%Y")                     # dates stored as text, day first
            if i % 53 == 0:
                name = name.upper()                            # inconsistent account names
            if i % 17 == 0:
                dr, cr = (f"{dr:,.2f}" if dr else "", f"{cr:,.2f}" if cr else "")
            if i == neg_at:
                dr, cr = -cr, None                             # credit keyed as a negative debit
            if l["journal"] == cut_j and cut:
                d = date(year + 1, 1, 2)                       # cut-off: dated after year end
        row = [d, l["journal"], int(l["account"]), name, l["desc"], l["ref"], dr or None, cr or None]
        ws.append(row)
        if dirty and i == dup_at:
            ws.append(row)                                     # duplicated line
        if dirty and i in (int(n * 0.25), int(n * 0.75)):
            ws.append([])
        if dirty and i == int(n * 0.5):
            ws.append([None, None, None, None, "Sub-total", None, 999999.99, 999999.99])
    ws.append([None, None, None, None, "Grand Total", None,
               round(sum(l["debit"] for l in L.lines), 2), round(sum(l["credit"] for l in L.lines), 2)])
    wb.save(path)


def generate(out_dir, cy=2026, dirty=True, seed=7):
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    rng = random.Random(seed)
    entity = "Mekong Trading Co., Ltd."
    open_py = dict(OPENING)
    open_py["3200"] = -round(sum(open_py.values()), 2)
    L_py, _ = build_year(cy - 1, open_py, rng, current=False)
    open_cy = close_year(L_py.bal)
    L_cy, entertainment = build_year(cy, open_cy, rng, current=True)
    _write_tb(out / f"TB_{cy - 1}.xlsx", entity, cy - 1, open_py, L_py, rng, dirty)
    _write_tb(out / f"TB_{cy}.xlsx", entity, cy, open_cy, L_cy, rng, dirty)
    _write_gl(out / f"GL_{cy - 1}.xlsx", entity, cy - 1, L_py, rng, dirty)
    _write_gl(out / f"GL_{cy}.xlsx", entity, cy, L_cy, rng, dirty, cut=True)
    config = {
        "entity_name": entity,
        "legal_form": "private limited liability company",
        "country": "the Kingdom of Cambodia",
        "registration_no": "[registration number]",
        "directors": "[names and positions of the directors]",
        "auditors": "[name of audit firm]",
        "registered_office": "[Registered office address]",
        "principal_activities": "the wholesale distribution of consumer goods and related after-sales services",
        "reporting_date": f"{cy}-12-31",
        "currency": "USD",
        "pres_currency": "KHR", "pres_scale": 1000,
        "pres_reason": "presentation in Khmer Riel is required for statutory reporting purposes in Cambodia",
        # ILLUSTRATIVE exchange rates (KHR per USD) - replace with the official published rates
        "fx_close_cy": 4005, "fx_avg_cy": 4012, "fx_close_py": 4010, "fx_avg_py": 4018, "fx_close_ppy": 4025,
        "fx_hist_capital": 4000, "fx_hist_re": 4060,
        "tax_rate": TAX_RATE,
        "authorised_date": f"{cy + 1}-03-25",
        "authorised_by": "the Board of Directors",
        "shares_issued_cy": 450000, "shares_issued_py": 400000, "par_value": 1.0,
        "employees_cy": 24, "employees_py": 22,
        "early_application": True,
        "shares_authorised": 600000, "dividend_proposed": 0,
        "notes_inputs": {
            "tax_nondeductible_cy": round(entertainment * TAX_RATE, 2),
            "tax_nondeductible_py": round(L_py.bal["6130"] * TAX_RATE, 2),
            # illustrative: overdue bands [gross CY, allowance CY, gross PY, allowance PY]
            "ageing": [[round(r * p, 2), round(r * p * a, 2), round(q * p, 2), round(q * p * a, 2)]
                       for r, q in [(L_cy.bal["1100"], L_py.bal["1100"])]
                       for p, a in AGEING],
            # illustrative: contractual undiscounted loan cash flows incl. interest [CY, PY] per time band
            "maturity_borrowings": [[80800, 56000], [240000, 185600], [0, 0]],
        },
        "adjustments": L_cy.ajes,
        "dayfirst": True,
        "account_classes": {"1": "A", "2": "L", "3": "E", "4": "I", "5": "C", "6": "X", "7": "X", "8": "T"},
    }
    (out / "config.json").write_text(json.dumps(config, indent=2))
    return out
