# 07 · Expense Cutoff: search for unrecorded liabilities, 30 Nov 2026

## Cutoff decisions
| Vendor | Invoice | Service dates | Received | In period? | Existing liability? | Decision | Amount |
|---|---|---|---|---|---|---|---|
| Blue Ridge Engineering | BR-2291 | 28 Nov (completed) | 4 Dec | **Yes** | **No** (see checks) | **Accrue** | **12,400** |
| Steppe Facility Services | SF-118 | 1–31 Dec | 3 Dec | No | n/a | Exclude: December expense | 0 |

## Existing-liability checks for BR-2291
1. **AP (ap_invoices.csv, status Open):** no Blue Ridge row. ✔
2. **Accrued liabilities detail (2100 = $5,600):** Harbor & Co audit $3,600 and Sable Law $2,000. Nothing for Blue Ridge. ✔
3. **Other proposals in outputs/:** none. ✔
4. **Materiality:** $12,400 ≥ $1,000. ✔

## Delivery evidence
- ap_invoices.csv › BR-2291: service_start = service_end = 2026-11-28
- Operations manager e-mail, 4 Dec: *"Blue Ridge finished the network cabling and testing on the 28th. We signed the completion sheet that day."*

## Proposed entry (sent to 10)
| Account | Debit | Credit |
|---|---|---|
| 5100 Contractor expense | 12,400 | |
| 2100 Accrued liabilities | | 12,400 |

Memo: Accrue Blue Ridge Engineering BR-2291. Network cabling completed 28 Nov; invoice received 4 Dec.
**Reversal:** auto-reverse on 1 Dec 2026. Post the invoice to AP in December when it's approved.

## Excluded item
SF-118 Steppe Facility Services, $3,000: services run 1–31 December, so it's a December expense. Accruing it would overstate November expenses.
