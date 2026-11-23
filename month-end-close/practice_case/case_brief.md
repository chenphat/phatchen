# Practice Case: Khangai Data Services LLC, November 2026 close

**Business:** a data analytics consultancy with 24 staff. It bills clients monthly for services.
**Reporting currency:** USD. **Period:** 1–30 November 2026. **Fiscal year:** starts 1 November, so every income statement balance in the trial balance is November activity.
**Close calendar:** books close on business day 5 (7 December 2026). It is now 4 December.
**Materiality for proposing entries:** $1,000.

## Data files (`data/`)

| File | Contents |
|---|---|
| `trial_balance.csv` | Posted trial balance at 30 Nov 2026, before any close adjustments |
| `cash_transactions.csv` | November cash activity, book side and bank side. A blank `bank_amount` means the item hasn't cleared. |
| `ar_aging.csv` | Open customer invoices at 30 Nov, with service periods |
| `ap_invoices.csv` | Vendor invoices: open, paid, and received after period end but not yet posted |
| `fixed_assets.csv` | Fixed asset register |
| `payroll_register.csv` | November pay runs |

## Supporting information (not in the CSVs)

**Accrued liabilities, account 2100: detail of the $5,600 balance**

| Item | Amount | Basis |
|---|---|---|
| Harbor & Co: audit fee, work performed to date | 3,600 | Engagement letter, 30% complete |
| Sable Law: November legal advice | 2,000 | Partner's estimate; invoice expected mid-December |
| **Total** | **5,600** | |

Neither item relates to Blue Ridge Engineering.

**Prepaid expenses, account 1200:** a 12-month insurance policy paid on 1 October 2026 for $9,360, amortised at $780 per month to *Utilities and office*. Balance at 30 Nov: 10 months × $780 = **$7,800**.

**Deferred revenue, account 2300:** INV-1188 to Khan Logistics. Billed on 25 Nov for December–February services at $4,000 per month. The full $12,000 is deferred.

**Allowance policy:** Current 1% · 31–60 days 5% · 61–90 days 20% · over 90 days 50%.

**Payroll:** salaries and wages expense = gross wages + employer taxes. Withholdings and employer taxes are remitted by the 20th of the following month for the 16th–end-of-month run, and within 5 days for the 1st–15th run.

**Blue Ridge Engineering** (from the operations manager's e-mail, 4 Dec): *"Blue Ridge finished the network cabling and testing on the 28th. We signed the completion sheet that day. Their invoice only came in this morning."*

## Expected headline results

| Measure | Posted | Draft (with proposed entries) |
|---|---|---|
| Operating profit | 46,700 | 34,300 |
| Book cash | 82,600 | 82,600 |
| Bank tie-out | Tied | Tied |
| Proposed entries | | 1 entry: $12,400 accrual |
| Controller sign-off | | **Pending** |
