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
    --tb-cy data/sample/TB_2026.xlsx --tb-py data/sample/TB_2025.xlsx \
    --gl-cy data/sample/GL_2026.xlsx --gl-py data/sample/GL_2025.xlsx \
    --config data/sample/config.json \
    --out output/AFS_Mekong_Trading_2026.xlsx
```

The command prints the cleaning and check summary and exits non-zero if any check fails.
A built example is in `output/AFS_Mekong_Trading_2026.xlsx` (fictitious entity).

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
| Dashboard | KPI tiles, six charts, key ratios and priority actions, with navigation to every working paper |
| Guide, Entity | Workflow, colour legend; entity inputs including exchange rates |
| Directors_Report | Report of the Board of Directors, figures linked to the statements |
| SoFP, SoCI, SoCE, SoCF | Primary statements, two years, in the presentation currency with the functional currency alongside |
| Notes | All notes, in the presentation currency (see list in the sheet) |
| SoCI_Nature | Alternative SoCI with expenses by nature (paragraph 5.11(a)); issue one format only |
| Materiality | Benchmarks, overall / performance materiality, line-by-line assessment by size and nature |
| Policies | Options the Standard permits, the policy selected, and whether it is material enough to disclose; drives Note 3 |
| Analysis, Checklist | Two-year ratios, movements and common-size analysis; disclosure checklist by paragraph |
| Checks, Findings, DQ_Log | Live reconciliations; action tracker; source-data checks and cleaning log |
| GL, TB, AJE, ATB | Cleaned ledger; unadjusted trial balance; adjusting entries; adjusted trial balance |
| Mapping, Lead, Dash_Data, Lists | Account mapping, lead schedule, chart data, drop-down lists |

Presentation choices: single statement of comprehensive income, expenses by function with the
analysis by nature in the notes, cash flows by the indirect method. Ledger data is signed
debit +, credit −.

### Functional and presentation currency (Section 30)

The ledger stays in the functional currency. Set `pres_currency`, `pres_scale` and the rates in the
config (or on the Entity sheet) to present in another currency:

* assets and liabilities at the closing rate of each year end;
* income, expenses and cash flows at the average rate for the year;
* share capital, share premium and reserves at historical rates;
* the difference goes to other comprehensive income and the translation reserve, with a proof
  by rates on the Checks sheet (paragraph 30.20).

The rates in `data/sample/config.json` are illustrative, not official rates.

### Adjusting entries

List them under `adjustments` in the config (account, description, debit, credit) or type them
into the AJE sheet. TB + AJE = ATB.

## Limits – read before relying on the output

* Single entity only: no consolidation, non-controlling interests, discontinued operations or
  foreign-operation translation.
* The statement structures, note list, policy wording and the disclosure checklist were checked against the
  IFRS for SMEs Accounting Standard, third edition (February 2025) and its Illustrative Financial Statements.
  The checklist cites paragraph numbers but summarises each requirement: read the paragraph before sign-off.
  Policy wording is a template to tailor to the entity. Not covered: consolidation (Section 9), business
  combinations (19), share-based payment (26), hyperinflation (31), specialised activities (34), hedge accounting.
* The age analysis of receivables, the maturity analysis of borrowings and lease commitments cannot be derived
  from a GL/TB: they are input tables (`notes_inputs` in the config, or typed into the yellow cells).
* Narrative in yellow cells is entity-specific and must be completed by the preparer.
* Movement types are inferred; review GL column L where a roll-forward shows "other movements".
* Adding accounts or journals requires a re-run; changing a mapping or posting an AJE does not.

## Tests

```bash
python -m pytest -q
```

## Claude skill: `afs-ifrs-smes-3e`

`.claude/skills/afs-ifrs-smes-3e/` packages this platform as a Claude skill for compiling IFRS for SMEs (third
edition) financial statements with USD functional and KHR presentation currency. It contains the workflow
(`SKILL.md`), reference guides distilled from the IFRS Foundation educational modules (Sections 2, 11, 12, 17, 23,
28, 29, 30, 34, 35), the third-edition changes and transition reliefs, the IASB illustrative statements structure,
a USD→KHR translation method, the Sections 3–8 presentation rules with a 54-row disclosure checklist, an IES
competence mapping for teaching, the bundled workbook engine and a summary-translation script, and test cases.
Claude Code loads it automatically in this repository.
