# Bundled scripts

`scripts/translate_summary.py` translates statement-level USD figures into KHR (see its header for the CSV
and rates format). The rest of this page covers the full engine.

## Workbook engine (`scripts/build_afs.py`)

The engine turns a trial balance and general ledger into one Excel workbook in which every statement figure is a
live formula. Use it whenever the user supplies ledger data (or wants a teaching example); write narrative and
judgements yourself on top of it.

## Run it

```bash
python <skill>/scripts/build_afs.py sample --out ./sample          # two-year teaching data set, FY2026 + FY2025
python <skill>/scripts/build_afs.py build \
  --tb-cy TB_2026.xlsx --tb-py TB_2025.xlsx \
  --gl-cy GL_2026.xlsx --gl-py GL_2025.xlsx \
  --config config.json [--mapping reviewed_mapping.csv] \
  --out AFS_2026.xlsx
```

Requires `pandas` and `openpyxl`. If LibreOffice (`soffice`) is installed the engine recalculates the workbook and
prints the live-check results; the command exits non-zero if any check fails. Without LibreOffice the formulas still
work when the file is opened in Excel.

## Inputs

**Trial balance** (xlsx or csv, one per year, pre-closing): account code, name, opening balance (single column or
Dr/Cr), period debits and credits, closing balance (single or Dr/Cr). Header row is found automatically; title rows,
blank rows, total rows, text numbers and stray spaces are cleaned and logged.

**General ledger** (one per year): date, journal number, account, description, debit and credit (or a signed amount).
Duplicate lines that unbalance a journal are removed and logged; out-of-period dates are flagged.

**Account classes**: the first digit of the account code is mapped to a class in `account_classes` in the config
(default 1 asset, 2 liability, 3 equity, 4 income, 5 cost of sales, 6–7 expenses, 8 tax). Change it for other charts
of accounts — mapping rules depend on it.

**Mapping file** (optional): csv/xlsx with `account, afs_code[, nature]`. The engine writes `<out>_mapping.csv` on
every run; review it, correct the `afs_code`, and feed it back with `--mapping`. Valid codes are the `LEAVES` list in
`scripts/afs_smes/taxonomy.py`; expense natures are `NATURES` in the same file.

**config.json** keys:

| Key | Meaning |
|---|---|
| `entity_name`, `legal_form`, `country`, `registered_office`, `registration_no`, `principal_activities` | General information note and directors' report |
| `reporting_date` | `YYYY-12-31`; prior year is derived |
| `currency` | functional currency (ledger currency), e.g. `USD` |
| `pres_currency`, `pres_scale` | presentation currency and rounding, e.g. `KHR`, `1000` → KHR'000 |
| `pres_reason` | reason for a presentation currency different from the functional currency (30.26) |
| `fx_close_cy`, `fx_avg_cy`, `fx_close_py`, `fx_avg_py`, `fx_close_ppy` | presentation currency units per 1 functional unit |
| `fx_hist_capital`, `fx_hist_re` | historical rates for opening share capital/premium and for opening reserves/retained earnings of the earliest year |
| `tax_rate` | applicable income tax rate for the tax reconciliation |
| `authorised_date`, `authorised_by`, `directors`, `auditors` | approval and directors' report |
| `shares_authorised`, `shares_issued_cy`, `shares_issued_py`, `par_value`, `dividend_proposed` | equity note (4.12, 6.6) |
| `employees_cy`, `employees_py` | average number of employees |
| `early_application` | `true` when the third edition is applied to a period beginning before 1 January 2027 (A1 disclosure) |
| `adjustments` | list of adjusting entries `{no, account, description, debit, credit}` in the functional currency |
| `notes_inputs` | items a ledger cannot give: `tax_nondeductible_cy/py`, receivables `ageing` (4 overdue bands × [gross CY, allowance CY, gross PY, allowance PY]), `maturity_borrowings` (3 bands × [CY, PY] undiscounted) |
| `dayfirst` | date parsing in the GL (`true` for dd/mm/yyyy) |

## What the workbook contains

Dashboard · Guide · Entity (inputs and rates) · Directors_Report · SoFP · SoCI · SoCE · SoCF · Notes ·
SoCI_Nature (alternative format) · Materiality · Policies (selection drives Note 3) · Analysis (two-year) ·
Checklist (by paragraph) · Checks (live reconciliations) · Findings (action tracker) · DQ_Log (source checks and
cleaning log) · GL · TB (unadjusted) · AJE · ATB (adjusted) · Mapping · Lead · Dash_Data · Lists.

Statements show the presentation currency in columns C:D and the functional currency in E:F. Translation follows
`usd-khr-translation.md`: closing rate for assets and liabilities, average rate for income, expenses, OCI, owner
transactions and cash flows, historical rates for opening equity, balancing figure to the translation reserve, proof
by rates on the Checks sheet.

## Limits — tell the user

- Single entity: no consolidation, NCI, discontinued operations, business combinations, share-based payment,
  hyperinflation, hedge accounting, biological assets or investment property beyond simple balances.
- Movement types for roll-forward notes are inferred from the other accounts in each journal; check GL column L
  where a note shows an unclassified movement.
- All flows are translated at the average rate. If the user wants transaction-date rates for one-off items (share
  issue, year-end revaluation, dividends), adjust the SoCE and SoCI lines and re-run the proof.
- Narrative in yellow cells (related parties, contingencies, commitments, judgements, events after the reporting
  period, directors' names) must be written from the entity's facts.
- Sample exchange rates are illustrative.
