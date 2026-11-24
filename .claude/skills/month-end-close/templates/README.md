# CSV templates: data dictionary

Copy these six files into a `data/` folder and fill one row per item. Keep the header row exactly as it is, because the agents and the checker read columns by name. Write dates as `YYYY-MM-DD` and amounts as plain numbers (no currency symbols or thousands separators).

### trial_balance.csv
| Column | Meaning |
|---|---|
| account | GL account number |
| account_name | Account description |
| type | `Asset`, `Liability`, `Equity`, `Revenue` or `Expense` (contra-assets use `Asset`) |
| statement | `BS` (balance sheet) or `IS` (income statement) |
| debit / credit | Closing balance, entered in one column only |

### cash_transactions.csv
| Column | Meaning |
|---|---|
| date | Book date |
| reference | Deposit, cheque or ACH reference. Use `OPEN` for the reconciled opening balance row. |
| description, counterparty | Free text |
| book_amount | Amount in the cash ledger (+ receipt / − payment) |
| bank_amount | Amount on the bank statement. **Leave blank if it hasn't cleared by the statement date.** |
| bank_date | Date the item cleared the bank. A date after period end means it's in transit. |

### ar_aging.csv
| Column | Meaning |
|---|---|
| customer, invoice_no, invoice_date | Invoice identifiers |
| service_start / service_end | Period the service covers. Used for revenue cutoff. |
| amount | Open balance |
| days_outstanding | Days from invoice date to period end |
| aging_bucket | `Current`, `31-60`, `61-90` or `90+` |
| revenue_treatment | `Recognized` or `Deferred revenue` |

### ap_invoices.csv
| Column | Meaning |
|---|---|
| vendor, invoice_no, invoice_date, received_date | Invoice identifiers |
| service_start / service_end | When goods or services were received. **This drives expense cutoff.** |
| amount, gl_account | Amount and expense account |
| posted_period | `YYYY-MM` in which it was posted. Blank if not posted. |
| status | `Open`, `Paid` or `Unposted` |

### fixed_assets.csv
| Column | Meaning |
|---|---|
| asset_id, description, in_service_date | Asset identifiers |
| cost, salvage_value, useful_life_months, method | Depreciation inputs |
| months_depreciated_to_period_end | Months charged, including the current period |
| accumulated_depreciation, period_depreciation | Register balances |

### payroll_register.csv
| Column | Meaning |
|---|---|
| pay_date, pay_period | Pay run identifiers |
| gross_wages, employer_taxes, employee_withholdings, net_pay | Pay run amounts (net = gross − withholdings) |
| net_paid_date | Date net pay left the bank |
| withholdings_remitted_date | Date withholdings + employer taxes were paid. Blank means still a liability. |
