# Working papers: sheet guide and adaptation notes

Inputs are blue text on yellow cells. Everything else is a formula. In 05_TB debits are positive and credits negative; the statements show income as positive and expenses in brackets.

| WP | Sheet | Purpose | Inputs | Key formulas / outputs |
|---|---|---|---|---|
| 00 | 00_Index | Index with hyperlinks, status, sign-off, overall check status | Status, initials | Link to 19_Checks |
| 00a | 00a_Dashboard | One-page dashboard: 12 KPI tiles (with % change vs PY), charts of profit by category, IAS 1 vs IFRS 18 subtotals, expenses by nature and MPM bridge, readiness panel | – | Formula-fed helper data in columns Z:AB |
| 01 | 01_Setup | Entity, labels, adoption date, SMBA flags, expense method, tax rate, NCI, EPS data, materiality | All | Referenced everywhere |
| 02 | 02_Plan | Assess / Redesign / Implement / Report task plan; key dates | Owner, dates, status | Weighted % complete |
| 03 | 03_Gap | IAS 1 practice vs IFRS 18 gap analysis (22 areas) | Current practice, gap, impact, action | Count of high-impact gaps |
| 04 | 04_SMBA | Indicator checklist for both SMBAs | Y/N, evidence, weights | Indicator score, preliminary indication, link to setup conclusion |
| 05 | 05_TB | Trial balance and mapping | Account rows | Category lookup, mapping status, TB totals |
| 06 | 06_Rules | Classification rule book (no SMBA / investing SMBA / financing SMBA) | – | Reference |
| 07 | 07_Classification | Totals by category and by IFRS 18 line, with account counts | – | Five categories = profit |
| 08 | 08_SPLOCI | Statement of profit or loss and OCI, attribution, EPS | – | SUMIFS by line; OP, PBFIT, PBT, profit, OCI, TCI |
| 08b | 08b_PL_Visual | Colour-banded P&L: one band per category, closing with its subtotal (teaching / board view) | – | Links to 08_SPLOCI |
| 09 | 09_Transition | C3 reconciliation; lettered explanations; bridge matrix (IAS 1 captions × IFRS 18 lines) | Explanation text | IAS 1 vs IFRS 18 vs adjustment; profit unchanged |
| 10 | 10_ExpNature | IFRS 18.83 table: recognised, capitalised, incurred by function line × five natures | Capitalised amounts | Ties to function line totals |
| 11 | 11_Aggregation | Every statement line: "other" flag, PM flag, prompt, judgement record | Judgement columns | Prompts for B15–B26 review |
| 12 | 12_MPM_ID | Tests (a)–(d) for candidate measures | Y/N answers | MPM / Not an MPM; count |
| 13 | 13_MPM_Recon | Reconciliation from operating profit, line item presented in, tax and NCI effects, MPM per share, single-note checklist | Tax rate and NCI % per item | Adjusted MPM; tax and NCI totals; per-share amounts |
| 14 | 14_SFP | Statement of financial position by section, equity roll-forward | – | Balance check |
| 15 | 15_CashFlow | IAS 7 indirect method from operating profit; interest/dividend matrix | Cash flow lines | Closing cash ties to SFP cash less overdraft |
| 16 | 16_Disclosure | 36-item IFRS 18 checklist incl. IAS 7/33/34/8 and transition | Responses | Answered / non-compliant counts |
| 17 | 17_Reclass | Per P&L account: debit/credit caption and amount to move from the IAS 1 caption to the IFRS 18 line | – | Gross reclassified, count |
| 18 | 18_Issues | Judgements and review notes (seeded with the usual IFRS 18 judgements) | Conclusions | – |
| 19 | 19_Checks | 18 control checks with OK/ERROR status | – | "ALL CHECKS OK" |
| – | Lists | Line catalogue, IAS 1 captions, nature tags, status list | Extend with care | Drives look-ups and drop-downs |

## Control checks (19_Checks)
1. The TB balances in both periods.
2. There are no unmapped or conflicting accounts.
3. Operating profit from the lines equals the SUMIFS of the operating category.
4. Profit equals minus the sum of the PL accounts (CY and PY).
5. OCI equals minus the sum of the OCI accounts.
6. The SFP balances (CY and PY).
7. The five categories sum to profit.
8. Profit is unchanged between IAS 1 and IFRS 18 (PY).
9. The bridge matrix total equals profit.
10. Every PL account has an IAS 1 caption.
11. The MPM equals operating profit plus the reconciling items (CY and PY).
12. Closing cash in the cash flow equals SFP cash less overdraft (CY and PY).
13. Opening cash equals prior closing cash.

## Adapting to a real entity
1. Export the template: `--export-template tb.csv`. Replace the rows with the client TB and keep the column headings.
2. Fill `ias1_line` with the caption each account appeared in within the published comparative P&L. Fill `ifrs18_line` from the catalogue. If a caption is missing, add it to `CATALOGUE` (key, label, section, reference), and to `IAS1_LINES` and `CAPTION_EQUIV` where the old caption was renamed.
3. Tag `nature` for every operating expense account that contains depreciation, amortisation, employee benefits, IAS 36 impairment or IAS 2 write-downs. If accounts mix natures, split them, because the para 83 note depends on it.
4. Put the same `mpm_item` label on all accounts that form one MPM reconciling item. The label becomes the row description in 13_MPM_Recon.
5. Give each reclassification a `transition_ref` letter and write the explanations, either by editing `TRANSITION_EXPLANATIONS` in `demo_data.py` or directly in 09_Transition.
6. Equity accounts in the TB hold balances *before* current-period profit and OCI. 14_SFP adds profit to retained earnings, owners' OCI to reserves, and NCI's share of TCI to NCI. If the TB is post-closing, remove the current-period result from retained earnings first.
7. Enter the cash flow lines in 15_CashFlow. Operating profit is linked automatically.
8. Rebuild, recalculate (`recalc.py`) and clear every check before delivering.

## Known simplifications
- The share of the MPM attributable to ordinary shareholders, used for the per-share amount, is a single setup percentage. Replace it with a detailed computation if NCI effects are significant.
- The NCI share per MPM item defaults to 10%. Enter each subsidiary's actual percentage.
- The statement of changes in equity is not generated. Prepare it from the equity roll-forward if required.
- A third statement of financial position is not generated. Decide in 18_Issues whether one is needed.
- With an SMBA, review 06_Rules and map accounts accordingly. The generator does not reclassify automatically from the setup flags; it records the conclusion so that the mapping is a deliberate judgement.
