# phatchen

## Month-End Close with Claude
**15 agents. One close pack.** These are teaching materials for financial and cost accounting. A month-end close is split into 15 single-purpose AI agents (Reconcile → Adjust → Report & review). The agents propose entries, and a human controller signs off.

➡️ Start with [`month-end-close/START_HERE.md`](month-end-close/START_HERE.md)

- 15 editable Markdown agent files (also usable as Claude Code subagents): `month-end-close/agents/`
- 6 CSV templates and a data dictionary: `month-end-close/templates/`
- Completed practice case (Khangai Data Services, Nov 2026): `month-end-close/practice_case/`
- Reference outputs (answer key): `month-end-close/reference_outputs/`
- Calculation checker: `python3 month-end-close/checker/check_close.py --data month-end-close/practice_case/data --outputs month-end-close/reference_outputs`
