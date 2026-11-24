# phatchen

## Claude skills

| Skill | What it does |
|---|---|
| [`budget-vs-actual`](.claude/skills/budget-vs-actual/SKILL.md) | Static → flexible budget → actual variance analysis with F/U labels, materiality flags, review checks and commentary |

Each skill follows the same structure:

```
.claude/skills/<skill>/
├── SKILL.md                     instructions: inputs, method, review checks
├── examples/<data>.csv          example source data
├── examples/expected_results.*  reconciled output to compare against
└── scripts/<calculator>.py      reproduces the worked example (Python stdlib only)
```

**Claude Code:** skills in `.claude/skills/` load automatically when you work in this repo.

**Claude.ai:** zip the skill folder (`cd .claude/skills && zip -r budget-vs-actual.zip budget-vs-actual`)
and upload it under Settings → Capabilities → Skills.

Try it:

```bash
python .claude/skills/budget-vs-actual/scripts/bva_calculator.py \
  .claude/skills/budget-vs-actual/examples/budget_actual.csv \
  --budget-volume 1000 --actual-volume 1100 \
  --check .claude/skills/budget-vs-actual/examples/expected_results.csv
```
