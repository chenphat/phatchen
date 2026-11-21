# IFRS for SMEs (third edition) – financial statements platform

Reads a general ledger and trial balance, cleans and verifies them, maps every account to a
financial-statement component, and drafts a complete set of annual financial statements in one
Excel workbook: **SoFP, SoCI, SoCE, SoCF and Notes**, all as live formulas.

```
GL + TB exports ─► read ─► clean (logged) ─► check ─► map ─► AFS workbook (.xlsx)
```

## Quick start

```bash
pip install pandas openpyxl            # LibreOffice is optional (used to verify the live checks)

python -m afs_smes sample --out data/sample          # two-year teaching data set with planted defects
python -m afs_smes build \
    --tb-cy data/sample/TB_2025.xlsx --tb-py data/sample/TB_2024.xlsx \
    --gl-cy data/sample/GL_2025.xlsx --gl-py data/sample/GL_2024.xlsx \
    --config data/sample/config.json \
    --out output/AFS_Mekong_Trading_2025.xlsx
```

The command prints the cleaning and check summary and exits non-zero if any check fails.
A built example is in `output/AFS_Mekong_Trading_2025.xlsx` (fictitious entity).

## Inputs

| File | Needed columns (any order, common header spellings are recognised) |
|---|---|
| Trial balance (xlsx/csv), one per year | account code, account name, opening balance (or Dr/Cr), debits, credits, closing balance (or Dr/Cr) |
| General ledger (xlsx/csv), one per year | date, journal number, account code, description, debit, credit (or a signed amount) |
| `config.json` (optional) | entity name, reporting date, currency, tax rate, account-class prefixes … see `data/sample/config.json` |
| `--mapping file.csv` (optional) | `account, afs_code, nature` – a reviewed mapping that overrides the automatic one |

Trial balances must be **pre-closing** (income, expenses and dividends declared still in their own
accounts). The prior-year files are optional; without them comparatives are blank.

## What the tool does

1. **Read and clean** (`ingest.py`) – finds the header row, skips title/blank/total rows, converts text
   numbers and text dates, trims account codes, moves negative debits/credits to the correct side,
   removes a duplicated line only when that restores the journal's balance, standardises account
   names. Every action is listed in the `DQ_Log` sheet.
2. **Check** (`checks.py`) – TB balances and casts; every journal balances; GL movement agrees to the TB
   by account; cut-off dates; missing codes/dates; duplicates; abnormal balances; opening balances
   agree to the prior-year closing balances and retained earnings roll forward; journal-review
   indicators (weekend postings, large round sums).
3. **Map** (`mapping.py`, `taxonomy.py`) – each account gets an AFS component and, for expenses, a
   nature, from keyword rules constrained by account class. Fallback mappings are marked `REVIEW`.
   A `*_mapping.csv` is written next to the workbook so it can be reviewed and fed back with `--mapping`.
4. **Classify movements** (`movements.py`) – GL lines on PPE, intangibles, borrowings, provisions,
   deferred tax and equity accounts get a movement type (addition, disposal, depreciation, revaluation,
   proceeds, repayment …) inferred from the other accounts in the journal. These drive the roll-forward
   notes and the cash flow statement and can be overridden in the workbook.
5. **Draft** – the workbook.

## The workbook

| Sheet | Purpose |
|---|---|
| Guide, Entity | Workflow, colour legend, entity inputs |
| Checks | Live reconciliations (SoFP balances, SoCE and SoCF tie to the SoFP, notes tie to the statements, GL ties to TB …) |
| DQ_Log | Source-data checks and the cleaning log |
| SoFP, SoCI, SoCE, SoCF | Primary statements with comparatives and note references |
| Notes | General information, basis of preparation, material accounting policy information, judgements and estimates, and supporting notes (revenue, expenses by nature, tax reconciliation, PPE and intangibles roll-forwards, receivables, cash, equity, borrowings and financing-liabilities reconciliation, payables, provisions, deferred tax, financial instruments, commitments, contingencies, related parties, events after the reporting period, approval) |
| Analysis, Checklist | Ratios and year-on-year variances; disclosure checklist by Section |
| Mapping, Lead, TB, AJE, GL, Lists | Working papers: mapping, lead schedule, adjusted TB, adjusting journals, cleaned GL |

Presentation choices: single statement of comprehensive income, expenses by function with the
analysis by nature in the notes, cash flows by the indirect method. Ledger data is signed
debit +, credit −.

## Limits – read before relying on the output

* Single entity only: no consolidation, non-controlling interests, discontinued operations or
  foreign-operation translation.
* The accounting-policy wording and the disclosure checklist are drafting templates written at Section
  level. The text of the Standard could not be accessed while building this tool, so confirm them
  (and paragraph-level requirements) against the IFRS for SMEs Accounting Standard, third edition.
* Narrative in yellow cells is entity-specific and must be completed by the preparer.
* Movement types are inferred; review GL column L where a roll-forward shows "other movements".
* Adding accounts or journals requires a re-run; changing a mapping or posting an AJE does not.

## Tests

```bash
python -m pytest -q
```
