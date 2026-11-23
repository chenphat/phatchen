# Month-End Close with Claude: Start Here

**15 agents. One close pack.** This kit splits a month-end close into 15 small, single-purpose AI agents. Each one does one job, writes one working paper, and hands off to the next. **Agents only propose entries. They never post them.** A human controller reviews and signs off.

It's built for teaching financial and cost accounting. Students can follow one transaction from the source document all the way to the controller's review, and they can see where every number came from.

---

## 1. What's in the pack

```
month-end-close/
├── START_HERE.md              ← you are here
├── agents/                    ← 15 editable Markdown agent files
│   ├── 01-close-setup.md … 05-payables.md            (Reconcile)
│   ├── 06-revenue-cutoff.md … 10-journal-review.md   (Adjust)
│   └── 11-income-statement.md … 15-controller-review.md (Report & review)
├── templates/                 ← 6 blank CSV templates + data dictionary
├── practice_case/             ← completed practice case (Khangai Data Services, Nov 2026)
│   ├── case_brief.md
│   └── data/                  ← the 6 CSVs, filled in
├── reference_outputs/         ← the answer key: what every agent should produce
└── checker/check_close.py     ← calculation checker (Python 3, no extra installs)
```

## 2. The close team

| Phase | # | Agent | Output (in `outputs/`) |
|---|---|---|---|
| **Reconcile** | 01 | Close Setup | `01_close_setup.md`: checklist, materiality, input check |
| | 02 | Trial Balance | `02_trial_balance_review.md` |
| | 03 | Cash & Bank | `03_bank_reconciliation.md` |
| | 04 | Receivables | `04_receivables_review.md` |
| | 05 | Payables | `05_payables_review.md` |
| **Adjust** | 06 | Revenue Cutoff | `06_revenue_cutoff.md` |
| | 07 | Expense Cutoff | `07_expense_cutoff.md` |
| | 08 | Fixed Assets | `08_fixed_assets.md` |
| | 09 | Payroll & Tax | `09_payroll_tax.md` |
| | 10 | Journal Review | `10_journal_review.md` + `journal_entries.csv` |
| **Report & review** | 11 | Income Statement | `11_income_statement.md` + `draft_income_statement.csv` |
| | 12 | Balance Sheet | `12_balance_sheet.md` + `draft_balance_sheet.csv` |
| | 13 | Cash Bridge | `13_cash_bridge.md` |
| | 14 | Close Memo | `14_close_memo.md` |
| | 15 | Controller Review | `15_controller_review.md`: the sign-off queue |

Run them in number order. Agents 03–05 are independent of each other, and so are 06–09, so you can run each group in parallel.

## 3. Three ways to use it

**A. Claude Code (recommended).** Each agent file is already a Claude Code subagent: it has YAML front-matter with `name`, `description` and `tools`.
1. Copy `agents/*.md` into `.claude/agents/` in your working folder.
2. Copy your six CSVs into `data/`. For practice, copy `practice_case/data/`.
3. Ask Claude: *"Run the month-end close for November 2026 using the close agents, in order, and stop at controller review."*
4. Then ask: *"What needs controller approval?"*

**B. Claude.ai Project.** Create a Project, upload the six CSVs and `case_brief.md`, and paste one agent file into the chat at a time, for example: *"Act as this agent: [paste 07-expense-cutoff.md]."* Save each output, then upload it before you run the next agent.

**C. Classroom or manual.** Print the agent files as procedure sheets. Students act as the agents, then compare their working papers with `reference_outputs/`.

## 4. Shared conventions (every agent follows these)

- **Period:** the period end comes from Close Setup (01). In the practice case it's **30 November 2026**.
- **Trial balance:** separate `debit` and `credit` columns. Contra-assets (allowance, accumulated depreciation) carry credit balances.
- **Materiality:** propose an entry only if it's ≥ **$1,000**, or if it's qualitatively important (fraud, related party, covenant). Record smaller differences as *"noted, not adjusted"*.
- **Evidence:** cite the source file and row (for example `ap_invoices.csv › BR-2291`) for every number.
- **Status words:** `Tied` · `Not tied` · `Proposed` · `Approved` · `Posted` · `Pending`. Agents may only use `Tied`, `Not tied`, `Proposed` and `Pending`. Only a human can set `Approved` or `Posted`.
- **Flag, don't fix:** agents never edit the input CSVs.
- **Entry IDs:** `AJE-YYYY-MM-NN`, assigned only by Journal Review (10).

## 5. The worked example: one late invoice

The practice case follows one late invoice from the source documents to the controller's review.

1. **Reconcile.** Start with the posted trial balance. A **$6,000 deposit in transit** and a **$2,000 outstanding payment** explain the difference between the bank statement ($78,600) and book cash ($82,600).
2. **Adjust.** Blue Ridge Engineering delivered the work on **28 November**. Their invoice arrived on **4 December**. The agent checks AP and accrued liabilities for an existing liability, finds none, and proposes **Dr Contractor expense $12,400 / Cr Accrued liabilities $12,400**. A second December invoice, SF-118, is a trap: those services happen in December, so it isn't accrued.
3. **Report & review.** Posted operating profit is **$46,700**. With the proposed accrual, draft operating profit is **$34,300**. Book cash stays at **$82,600** because an accrual is a non-cash entry. The controller can trace the change back to the invoice and see which entry is waiting for approval.

## 6. Check your numbers

```bash
python3 checker/check_close.py --data practice_case/data --outputs reference_outputs
python3 checker/check_close.py --data data --outputs outputs        # your own run
```
The checker recomputes every tie-out from the CSVs. It then checks that your `journal_entries.csv`, `draft_income_statement.csv` and `draft_balance_sheet.csv` agree with those results. Each check prints `PASS` or `FAIL`.

## 7. Using your own data

1. Fill in the six files in `templates/`. Column definitions are in `templates/README.md`.
2. Edit the materiality threshold and chart of accounts in `agents/01-close-setup.md`.
3. Add your own policies (allowance rates, capitalisation threshold, accrual rules) to the relevant agent files. They're plain Markdown.

## 8. Teaching variations (for lecturers)

| Change in the data | What students should find |
|---|---|
| Set BR-2291 `service_start/end` to 2026-12-02 | No accrual. Operating profit stays at $46,700. |
| Add a $35 bank fee on the bank side only | A book-side adjusting entry (below materiality, so noted only) |
| Move INV-1188 revenue treatment to "Recognized" | Revenue cutoff error of $12,000 → reclassify to deferred revenue |
| Change FA-001 life to 36 months | Depreciation overstated by $783.33 per month, so a catch-up is needed |
| Leave the 30 Nov withholdings remitted in November | Payroll liabilities no longer tie to the register |

Each change gives a new, self-checking exercise. Run the checker afterwards. It reports what moved.

---
*Illustrative case. Company, customers and vendors are fictional. The pack doesn't replace professional judgement or your organisation's close policy.*
