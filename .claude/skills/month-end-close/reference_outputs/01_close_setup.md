# 01 · Close Setup: Khangai Data Services LLC, November 2026

## Close parameters
| Parameter | Value |
|---|---|
| Period | 1–30 November 2026 (period end **2026-11-30**) |
| Materiality for proposing entries | $1,000 |
| Close deadline | Business day 5: 7 Dec 2026 |
| Entry IDs | AJE-2026-11-NN |
| Approver | Controller (human) |

## Input file check
| File | Rows | Header OK | Issues |
|---|---|---|---|
| trial_balance.csv | 20 | ✔ | None |
| cash_transactions.csv | 11 (incl. OPEN) | ✔ | CHK-1047 has a blank bank_amount, which is expected (uncleared) |
| ar_aging.csv | 6 | ✔ | None |
| ap_invoices.csv | 7 | ✔ | 2 rows have a blank posted_period, which is expected (received after period end) |
| fixed_assets.csv | 3 | ✔ | None |
| payroll_register.csv | 2 | ✔ | 30 Nov run has a blank remitted date, which is expected (liability) |

## Chart of accounts
1000 Cash · 1100 AR · 1150 Allowance · 1200 Prepaids · 1500 Equipment · 1510 Accumulated depreciation · 2000 AP · 2100 Accrued liabilities · 2200 Payroll liabilities · 2300 Deferred revenue · 3000 Share capital · 3100 Retained earnings · 4000 Service revenue · 5000 Salaries · 5100 Contractors · 5200 Rent · 5300 Software · 5400 Depreciation · 5500 Utilities & office · 5600 Marketing

## Close checklist
| # | Agent | Status |
|---|---|---|
| 02 | Trial Balance | Not started |
| 03 | Cash & Bank | Not started |
| 04 | Receivables | Not started |
| 05 | Payables | Not started |
| 06 | Revenue Cutoff | Not started |
| 07 | Expense Cutoff | Not started |
| 08 | Fixed Assets | Not started |
| 09 | Payroll & Tax | Not started |
| 10 | Journal Review | Not started |
| 11–14 | Reporting | Not started |
| 15 | Controller Review | Not started |

## Known support
- Accrued liabilities detail ($5,600: Harbor & Co $3,600, Sable Law $2,000) → 07
- Blue Ridge completion e-mail, 4 Dec → 07
- Prepaid insurance schedule ($780/month) → 12
- Deferred revenue detail (INV-1188) → 06
- Allowance policy rates → 04
