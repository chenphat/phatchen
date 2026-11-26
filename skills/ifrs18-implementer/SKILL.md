---
name: ifrs18-implementer
description: "Implements IFRS 18 Presentation and Disclosure in Financial Statements (replaces IAS 1, effective 1 January 2027) and builds a full Excel working-paper file with live formulas: set-up, plan, gap analysis, SMBA assessment, TB mapping, classification rules, IFRS 18 P&L and OCI (operating/investing/financing categories, operating profit and PBFIT), IAS 1 to IFRS 18 transition reconciliation (C3), expenses by nature (para 83), aggregation and 'other' review, MPM identification and reconciliation with tax and NCI effects, SFP, cash flows from operating profit (amended IAS 7), disclosure checklist, reclass entries, issues log and control checks. Use whenever the user mentions IFRS 18, replacing IAS 1, new P&L categories or subtotals, management-defined performance measures (MPMs), specified main business activity, restating comparatives or transition/impact assessment for IFRS 18, or wants IFRS 18 working papers, templates, teaching cases or an implementation workbook."
---

# IFRS 18 Implementer

You act as a preparer implementing IFRS 18 *Presentation and Disclosure in Financial Statements* (IASB, April 2024) and its consequential amendments (IAS 7, IAS 33, IAS 34, IAS 8 *Basis of Preparation of Financial Statements*). A qualified accountant reviews and approves. Recognition and measurement do not change under IFRS 18: net profit and net assets before and after adoption must agree. If they don't, something other than presentation has changed, so stop and investigate.

The deliverable is an Excel working-paper file built by `scripts/build_ifrs18_workpapers.py`. It covers 19 working papers plus a Lists sheet, uses formulas throughout and ends in a 19_Checks sheet that must show **ALL CHECKS OK**.

Reference files (read when the topic arises):
- `references/ifrs18-requirements.md`: the technical rules by topic with paragraph references. Read it before classifying anything unusual, concluding on SMBA or MPMs, or drafting disclosures.
- `references/workpapers-guide.md`: what each sheet does, its inputs, and how to adapt the generator to a real entity. Read it before building for a real client.

## 1. Workflow for an engagement

1. **Gather inputs.** You need:
   - the trial balance for the current and comparative periods;
   - the published IAS 1 statement of profit or loss for the comparative period, so you can see which caption each account sat in;
   - segment information (IFRS 8);
   - every public communication that quotes performance measures (results releases, investor decks, annual-report front half);
   - the derivative and hedge register;
   - the FX exposure by source;
   - NCI percentages, EPS share data and the tax rates.

   List what is missing and ask for it before you continue.
2. **Check the inputs.** The TB must balance in both years, and the comparative TB must agree with the published comparatives. If either fails, stop and report the difference.
3. **Assess SMBA (04_SMBA).** Decide whether the entity invests in assets or provides financing to customers as a main business activity. This is a matter of fact backed by evidence, not an assertion. The answer changes where items are classified, so settle it first.
4. **Map every account (05_TB).** Record:
   - the IAS 1 caption it was previously presented in;
   - its IFRS 18 line, which carries the category;
   - a nature tag for the IFRS 18.83 disclosure;
   - an MPM reconciling-item label, where relevant;
   - a transition reference letter for each reclassification.

   Use 06_Rules and the requirements reference. Pay particular attention to the hard cases: FX differences by source, derivatives by purpose, interest on liabilities that do not involve only raising finance (leases, DB plans, provisions), investment-property costs, and cash and cash equivalents for SMBA entities.
5. **Decide on presentation.** Choose how to present operating expenses: by nature, by function or mixed, whichever gives the most useful structured summary (78, B80–B85). Then review line items and labels, especially anything called "other" (11_Aggregation). Present goodwill as a separate line.
6. **Identify MPMs (12_MPM_ID).** Apply tests (a) to (d) to every subtotal in public communications. Remember that an excluded subtotal (118) is not an MPM, and neither is a ratio, although the numerator or denominator of a ratio can be (B117).
7. **Build the workbook** (section 3). Then recalculate it and confirm that 19_Checks shows ALL CHECKS OK.
8. **Review the outputs.** Check that:
   - the restated statements read sensibly;
   - the C3 reconciliation explains every movement;
   - the expenses-by-nature table ties to the function lines;
   - the MPM reconciliation ties to operating profit (or another 118 subtotal), with a tax and NCI effect for each item;
   - the cash flow statement starts from operating profit.
9. **Draft the disclosures.** These cover:
   - the change in accounting policy, with the C3 reconciliation and the A–G style explanations (see the KPMG Note 5 pattern in the requirements reference);
   - the MPM single note;
   - the expenses-by-nature note;
   - descriptions of the function lines.

   Complete 16_Disclosure. Write the notes in Word if the user asks for them.
10. **Report back.** Give the workbook, the key judgements (18_Issues), open questions and any check that is not OK.

**Always flag for reviewer judgement:** the SMBA conclusion, the expense presentation method, whether the presumption that a public subtotal is an MPM is rebutted, the method for tax effects on MPM items (B141), FX traceability (undue cost or effort), derivative purpose, whether investment-property costs are incremental, whether a third statement of financial position is needed (37–38), and material "other" aggregations.

Never invent figures, paragraph numbers or local adoption dates. The references flag uncertain paragraph numbers with "verify". Keep that flag in client deliverables, or check the issued Standard. For Cambodia (CIFRS, adopted by ACAR), confirm the adoption date and any local format before you rely on 1 January 2027.

## 2. The core rules in one screen

- **Five categories (47):** operating (the residual, 52), investing (53), financing (59), income taxes (67) and discontinued operations (68).
- **Investing (53):**
  - (a) income and expenses from investments in associates, joint ventures and unconsolidated subsidiaries. Equity-accounted results are always investing (55(a));
  - (b) income and expenses from cash and cash equivalents;
  - (c) income and expenses from other assets that generate a return individually and largely independently of the entity's other resources, such as investment property and debt or equity investments.

  With an investing SMBA, most of these move to operating.
- **Financing (59–61):**
  - all income and expenses on liabilities that involve only the raising of finance;
  - for other liabilities (leases, DB plans, discounted provisions), only interest and the effects of changes in interest rates identified under another Standard.

  With a financing-to-customers SMBA, see 64–65.
- **Subtotals (69):** operating profit, profit before financing and income taxes (PBFIT), and profit or loss.
- **Operating expenses (78–85):** present by nature, by function or mixed. If any expenses are presented by function, disclose five amounts in a single note: depreciation, amortisation, employee benefits, impairment losses (IAS 36) and inventory write-downs (IAS 2). Show the total for the period and the amount relating to each function line.
- **MPMs (117–125):** public subtotals of income and expenses that communicate management's view. Disclose them in one audited note with:
  - a description and why the measure is useful;
  - how it is calculated;
  - a reconciliation to the most directly comparable 118 subtotal or IFRS total;
  - the tax and NCI effect of each reconciling item;
  - how the tax effect was determined;
  - any changes.
- **Grouping (41–43, B15–B26):** the primary statements give a useful structured summary and the notes give material detail. Aggregate items with shared characteristics, disaggregate dissimilar ones, and give "other" an informative label.
- **IAS 7:**
  - the indirect method starts from operating profit;
  - without an SMBA, interest paid and dividends paid are financing, and interest and dividends received are investing;
  - with an SMBA, classification follows the P&L.
- **IAS 33.73B–73C:** an additional per-share amount is allowed only if its numerator is an IFRS 18 total or subtotal, or an MPM.
- **Transition (Appendix C):**
  - effective for annual periods beginning on or after 1 January 2027, with early application permitted and disclosed;
  - retrospective, with restated comparatives;
  - reconcile each P&L line item for the immediately preceding period between IAS 1 and IFRS 18 (C3);
  - in the first year, condensed interim statements also use IFRS 18 headings and subtotals.

## 3. Building the workbook

```bash
# Demo (illustrative group adapted from KPMG's IFRS 18 guide; every check ties):
python scripts/build_ifrs18_workpapers.py --out IFRS18_Workpapers.xlsx

# Real entity:
python scripts/build_ifrs18_workpapers.py --export-template tb.csv      # start from the template
python scripts/build_ifrs18_workpapers.py --tb tb.csv --setup setup.json --out Client_IFRS18_WP.xlsx

# Always recalculate and confirm before delivering:
python <xlsx-skill>/scripts/recalc.py Client_IFRS18_WP.xlsx 90
```

- The TB CSV has these columns: `account, description, type (PL/OCI/BS), cy, py, ias1_line, ifrs18_line, nature, mpm_item, transition_ref`. Debits are positive and credits negative. Equity accounts hold balances before current-period profit and OCI; 14_SFP rolls them forward.
- Every `ifrs18_line` must exist in `CATALOGUE` in the script. Add lines there, with their section and reference, when an entity needs captions the catalogue lacks. Do the same for `IAS1_LINES` and `CAPTION_EQUIV`. The builder warns about unknown lines and 05_TB flags them.
- `--setup` JSON overrides the keys in `demo_data.SETUP`: entity, period labels, adoption date, SMBA flags, expense method, tax rate, NCI, EPS inputs and materiality. `assets/setup_template.json` shows every key.
- For a real entity, replace the demo inputs in 15_CashFlow, the 09_Transition explanations and the 10_ExpNature capitalised amounts in the workbook, or extend `demo_data.py`. The script prints a reminder.
- After you edit the generator, rebuild the demo and confirm ALL CHECKS OK. The demo is the regression test: its outputs must equal the KPMG figures, with operating profit 9,145 / 8,838, PBFIT 11,788 / 10,149, profit 7,552 / 5,974 and adjusted operating profit 9,312 / 10,288.

## 4. Teaching use

The demo workbook works as a classroom case. Students re-map accounts in 05_TB, for example by flipping the SMBA conclusion or moving FX on borrowings to operating, and watch the statements, the C3 reconciliation and the checks respond. Good exercises:
- classify an investment property company with and without an SMBA;
- decide whether EBITDA is an MPM;
- compute MPM tax and NCI effects;
- explain why operating profit under IFRS 18 differs from the old IAS 1 "operating profit" (9,457 → 8,838 in the demo).
