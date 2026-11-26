---
name: fa-coa-ob-review
description: Review a draft fixed-asset chart of accounts (COA) and Fixed Asset Register (FAR) workbook, revise the PP&E account structure where relevant (IAS 16 / IAS 41 / CIFRS for SMEs), and mark every asset and account balance as the opening balance (OB) at a cut-over date such as 1 Jan 2027. Use when the user uploads a FAR / "FA for OB" Excel file, asks to review or clean a fixed-asset COA, or asks to prepare PP&E opening balances for a new system or financial year.
---

# Fixed-Asset COA Review & Opening-Balance Preparation

Purpose: turn a draft FAR + draft PP&E chart of accounts into (1) a findings report,
(2) a revised COA with old→new mapping, (3) an asset-level register flagged
"OB as at <date>", (4) an OB trial balance and opening journal, (5) an exceptions list.

The source workbook is **never edited** (it usually holds pivots, images and external
links that openpyxl would destroy). All output goes into a new workbook.

## Expected input layout (the "FA Details" style FAR)

| Item | Where |
|---|---|
| Detail sheet | `FA Details` (override with `--sheet`) |
| Header row | row 8; data from row 11 |
| Class header rows | col A holds text like `121013 . Other Land Improvements - ELC`, col D the 3-letter class code |
| Asset rows | A No., B Status, C New COA (class code), D Tag, F Main (old) code, G Reference, H Acq. date, I Description, J Cost USD, K Method, L Useful life, R Opening AD, S..AD monthly depn, AE Ending AD, AF NBV |
| Status values | Confirmed / Unsupported / NonExistent / Damaged Assets |
| Summary sheet | `SUMMARY` (optional) — col I holds "Unrecorded Assets" amounts per class |

If the layout differs, inspect it first (`python3 -I` + openpyxl, `data_only=True`) and
adjust the `COL` map at the top of `scripts/review_far.py` — do not guess columns.

## Workflow

1. **Inspect** – list sheets, header rows, class headers, statuses; read SUMMARY formulas
   (load once with formulas, once `data_only=True`).
2. **Run the reviewer**
   ```bash
   python3 -I .claude/skills/fa-coa-ob-review/scripts/review_far.py \
       <input.xlsx> <output.xlsx> --ob-date 2027-01-01
   ```
3. **Recalculate** the output (`scripts/recalc.py` of the xlsx skill) and confirm zero errors
   and that `OB TB` gross cost / AD agree to the source detail totals.
4. **Read the Findings sheet critically** and add workbook-specific findings the script
   cannot detect (e.g. SUMMARY formula gaps, valuation notes, unexplained reconciling lines).
5. **Report** to the user: headline numbers, top findings, decisions they must make.

## Review checklist (what "reviewing the COA" means)

Structure
- One cost account per asset class, plus a matching **accumulated-depreciation contra**
  account (1219xx) and an **accumulated impairment** account – drafts often omit these.
- **Capital work-in-progress** account exists (not depreciated).
- Header accounts are not postable; no truncated or inconsistent names/separators.
- "Other fixed assets" catch-all is small; large balances are reclassed to real classes.

Standards (IAS 16 / IAS 41 / CIFRS for SMEs s.17 & s.34)
- **Bearer plants** (rubber, cashew, fruit trees) are PP&E, not "biological assets";
  immature bearer plants are carried at accumulated cost and **not depreciated**.
- **Livestock / produce** are IAS 41 biological assets → separate non-PP&E account at fair
  value less costs to sell.
- Land is not depreciated; improvements on concession (ELC) land depreciate over the
  shorter of useful life and remaining concession term.
- Useful lives are accounting estimates, not tax ("GDT") rates; buildings at 5 years or
  a "3 to 5 years" landscape policy needs justification. Fully depreciated assets still in
  use ⇒ review useful lives (IAS 16.51).

Opening-balance treatment by verification status
| Status | OB treatment |
|---|---|
| Confirmed | Recognise at cost / AD / NBV at cut-off |
| Unsupported (exists, no documents) | Recognise provisionally; flag for documentation or valuation |
| Damaged | Recognise, then impair to recoverable amount (default nil) |
| NonExistent | **Exclude** – derecognise against retained earnings before OB |
| Unrecorded (found, not in books) | Memorandum only until valuation approved |

Data integrity (per asset)
- Acquisition date missing, year-only, or **after** the OB date
- Useful life or depreciation method missing
- NBV ≠ cost − accumulated depreciation
- Tag missing, duplicated, or a placeholder text
- Lump-sum / balancing "plug" lines

## Output conventions
- Arial font; blue = hard-coded inputs, black = formulas; zeros as "-".
- Every OB figure on `OB TB` is a SUMIFS formula over `OB Register`; the opening journal
  must balance (check cell shows 0).
- Every sheet title carries "Opening balance as at <date>".
