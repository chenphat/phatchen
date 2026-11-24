---
name: budget-vs-actual
description: Budget vs Actual (BvA) variance analysis using a static budget, a flexible budget and actual results. Use when the user asks to compare budget to actual, explain a profit gap, split a static-budget variance into sales-volume and flexible-budget variances, flag material variances, or write variance commentary for management or a cost accounting class. Works on a CSV/Excel/pasted table of line items with budget and actual amounts plus budgeted and actual volume.
---

# Budget vs Actual Analysis

Turns a budget and an actual P&L into a reconciled variance report:
**static budget → flexible budget → actual**, with every variance labelled
Favourable (F) or Unfavourable (U), material items flagged, and a short
management commentary.

## Inputs

Ask for anything missing before you calculate. Do not invent numbers.

| Input | Required | Notes |
|---|---|---|
| Line items with `budget_amount` and `actual_amount` | Yes | One row per P&L line, same period, same currency/units |
| `category` per line | Yes | `revenue`, `cogs` or `opex` |
| `behavior` per line | Yes | `variable` (moves with volume) or `fixed`. If the user doesn't say, propose a classification and ask them to confirm it |
| Budgeted volume and actual volume | Yes | Units sold/produced, the cost driver for the variable lines |
| Materiality thresholds | No | Default: flag if \|flexible-budget variance\| ≥ 2,000 **or** ≥ 10 % of flexible budget |

The expected CSV layout is in `examples/budget_actual.csv`:

```
line_item,category,behavior,budget_amount,actual_amount
```

## Method

1. **Validate the data.** Same period and currency; no blank or text amounts; every
   line has a category and behaviour; budgeted volume > 0. List any problems and stop
   until they are fixed.
2. **Flexible budget.** For each line:
   - variable: `flexible = static budget × (actual volume ÷ budgeted volume)`
   - fixed: `flexible = static budget`
3. **Variances** (all as actual − budget):
   - Sales-volume variance = flexible − static
   - Flexible-budget variance = actual − flexible
   - Static-budget variance = actual − static = volume + flexible-budget variance
4. **Direction.** Read from the profit point of view: higher revenue or lower cost is
   **F**; lower revenue or higher cost is **U**. A positive variance on a cost line is U.
5. **Subtotals.** Revenue, COGS, gross profit, operating expenses, operating income,
   each with static / flexible / actual and both variances.
6. **Materiality.** Flag each line whose flexible-budget variance passes the thresholds.
7. **Commentary.** For each material line: what happened (amount, %, F/U), the likely
   driver (price, usage/efficiency, spending, timing, mix, classification), and the
   question management should ask. Keep the volume story (how much profit the extra or
   missing units should have produced) separate from the performance story
   (flexible-budget variances: how well the business controlled prices and costs at
   the volume actually achieved).

Run the calculator rather than doing the arithmetic in your head whenever you can
execute code:

```bash
python scripts/bva_calculator.py <input.csv> --budget-volume <n> --actual-volume <n> \
    [--abs-threshold 2000] [--pct-threshold 10] [--out results.csv]
```

It uses only the Python standard library and prints the report plus the review checks.

## Output format

1. One-line headline: operating income actual vs static budget, and the split between
   volume and flexible-budget variance.
2. Variance table (line items + subtotals): Static | Flexible | Actual | Volume var (F/U) |
   Flexible-budget var (F/U) | % | Material.
3. Material variances with commentary (bullets, largest first).
4. Review checks (below), each marked PASS/FAIL.
5. Assumptions you made (especially any fixed/variable classification you proposed).

## Review checks

Do not present results until all of these pass:

- [ ] On every line, volume variance + flexible-budget variance = static-budget variance
- [ ] Fixed lines have zero volume variance
- [ ] Operating income: actual − static budget = total static-budget variance
- [ ] Operating income: total volume variance + total flexible-budget variance = static-budget variance
- [ ] Every cost line with actual > budget is marked U, every revenue line with actual < budget is marked U
- [ ] Totals in the table agree with the sum of the lines

## Worked example (reproducible)

Furniture maker, one month, amounts in MNT thousand. Budgeted volume 1,000 chairs,
actual 1,100 chairs.

```bash
python scripts/bva_calculator.py examples/budget_actual.csv \
    --budget-volume 1000 --actual-volume 1100 --check examples/expected_results.csv
```

Line results must match `examples/expected_results.csv` and the commentary should
cover what is in `examples/expected_summary.md`. Headline: operating income 20,975
vs 19,500 budget (+1,475 F) = 9,250 F volume variance − 7,775 U flexible-budget
variance. Selling more chairs hid weak pricing and materials cost control.

## Extensions (when the user supplies the data)

- **Price and efficiency variances** for direct materials and labour, when standard
  quantity per unit, standard price, actual quantity and actual price are given:
  price variance = (AP − SP) × AQ; efficiency/usage variance = (AQ − SQ allowed) × SP;
  they sum to the line's flexible-budget variance.
- **Sales price variance** = (actual price − budgeted price) × actual units; with a
  single product this equals the revenue flexible-budget variance.
- **Multi-period** files: run each period, then show the year-to-date column.
