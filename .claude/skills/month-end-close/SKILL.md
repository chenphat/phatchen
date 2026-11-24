---
name: month-end-close
description: Runs a full month-end close as 15 single-purpose agents in three phases. Reconcile covers trial balance, bank, AR and AP. Adjust covers revenue and expense cutoff, fixed assets, payroll and journal review. Report & review covers the draft P&L, balance sheet, cash bridge, close memo and controller sign-off queue. Use it when someone asks to run or teach a month-end close, prepare a bank reconciliation or cutoff accruals, search for unrecorded liabilities, build draft statements from a trial balance, answer "what needs controller approval?", or practise with the Khangai Data Services case. Agents only propose entries. A human controller approves them.
---

# Month-End Close: 15 agents, one close pack

You run the close by acting as each agent in turn. Every agent's full procedure is in `agents/NN-*.md` inside this skill folder. **Read each agent file before you act as that agent, and follow it exactly.** Don't work from memory.

## Golden rules (apply to every agent)
1. **Propose, never post.** Entry status may only be `Proposed`. Only a human sets `Approved` or `Posted`, or fills in `approved_by`. Controller sign-off stays **Pending**.
2. **Cite evidence** for every number: the file and row, for example `ap_invoices.csv › BR-2291`.
3. **Flag, don't fix.** Never edit the input CSVs, and never plug a difference.
4. **Materiality:** propose an entry only if it's ≥ $1,000 (or the threshold the user sets). Record smaller items as "noted, not adjusted".
5. **Cutoff follows delivery or performance dates,** not invoice or payment dates.

## Step 0: choose the mode
- **Practice / teaching:** copy `practice_case/data/*.csv` into `./data/` in the working directory. Read `practice_case/case_brief.md` for the supporting information. The answer key is in `reference_outputs/`. Don't show it to students until they've finished.
- **Own data:** the user supplies six CSVs in `./data/`. Their headers must match `templates/` (column meanings are in `templates/README.md`). If files are missing, give the user the templates and stop.
- **Single question** (for example "reconcile my bank" or "is this invoice a November expense?"): run only the relevant agent, plus 01 for the parameters.

## Step 1: run the agents in order
Write each output to `./outputs/`.

| Phase | Agents (read the file, then do it) | Outputs |
|---|---|---|
| Reconcile | `agents/01-close-setup.md`, `02-trial-balance.md`, `03-cash-bank.md`, `04-receivables.md`, `05-payables.md` | `01_…` to `05_….md` |
| Adjust | `06-revenue-cutoff.md`, `07-expense-cutoff.md`, `08-fixed-assets.md`, `09-payroll-tax.md`, `10-journal-review.md` | `06_…` to `10_….md`, `journal_entries.csv` |
| Report & review | `11-income-statement.md`, `12-balance-sheet.md`, `13-cash-bridge.md`, `14-close-memo.md`, `15-controller-review.md` | `11_…` to `15_….md`, `draft_income_statement.csv`, `draft_balance_sheet.csv` |

Agents 03–05 are independent of each other, and so are 06–09. If subagents are available, you may run each group in parallel. Copying `agents/*.md` into `.claude/agents/` makes them Claude Code subagents. Stop if Close Setup (01) or Trial Balance (02) reports a blocking error.

Use the output formats in `reference_outputs/` as the model for layout and level of detail.

## Step 2: verify with the checker
```bash
python3 <skill-dir>/scripts/check_close.py --data data --outputs outputs
# options: --period-end YYYY-MM-DD  --materiality 1000
```
It recomputes every tie-out from the CSVs and checks the journal entries and draft statements. It needs only the Python standard library. Fix any `FAIL` by revising the agent output. Don't adjust the data to make a check pass.

## Step 3: report to the user
Start with the answer to **"What needs controller approval?"** (from `15_controller_review.md`). Then give posted vs draft operating profit, book cash, the reconciliation status, and the list of proposed entries. Say clearly that the close is a **draft pending controller sign-off**.

## Practice case quick facts (for checking yourself)
Khangai Data Services LLC, November 2026, period end 30 Nov:
- Bank $78,600 + deposit in transit $6,000 − outstanding payment $2,000 = book cash $82,600 → **Tied**.
- **One proposed entry:** AJE-2026-11-01, Dr 5100 Contractor expense $12,400 / Cr 2100 Accrued liabilities $12,400. Blue Ridge Engineering BR-2291: work delivered 28 Nov, invoice received 4 Dec, no existing liability.
- **Trap:** SF-118 ($3,000, December services) is excluded.
- **Noted, not adjusted:** allowance short by $5.
- Operating profit $46,700 posted → **$34,300 draft**. Cash stays at $82,600. Sign-off: **Pending**.

## Teaching variations
Suggest these to lecturers. Change the data, then rerun the close and the checker.
- BR-2291 service date → 2026-12-02: no accrual, and operating profit stays at $46,700.
- Add a bank-only $35 fee: a book-side item, noted only (below materiality).
- INV-1188 treated as "Recognized": a $12,000 revenue cutoff error, so reclassify it to deferred revenue.
- FA-001 life → 36 months: depreciation is misstated, so a catch-up entry is needed.
- 30 Nov withholdings marked as remitted: payroll liabilities no longer tie.

## Folder map
```
month-end-close/
├── SKILL.md               (this file: orchestration)
├── agents/                (15 agent procedures; also valid Claude Code subagents)
├── templates/             (6 blank CSVs + README data dictionary)
├── practice_case/         (case_brief.md + data/ with the 6 filled CSVs)
├── reference_outputs/     (answer key: 15 working papers + 3 CSVs)
└── scripts/check_close.py (calculation checker)
```
