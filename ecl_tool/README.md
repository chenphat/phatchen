# IFRS 9 ECL tool: logistic regression PD, LGD and EAD with forward-looking scenarios

A teaching tool in two parts:

- **Python** estimates the three risk parameters from loan history joined to macro history.
- **Excel** (`IFRS9_ECL_Tool.xlsx`) applies the fitted equations with live formulas, stages
  each loan and measures the probability-weighted ECL. Change a macro forecast, a scenario
  weight or a staging threshold and the allowance and journal entry recalculate.

> All data are **synthetic** (`generate_data.py`). The numbers illustrate the method only.

## Run

```bash
pip install -r requirements.txt
python run_all.py              # fit models on ./data and rebuild the workbook
python run_all.py --new-data   # regenerate the synthetic data first
python ecl_models.py           # print regression tables and ECL by stage
```

## Method

| Parameter | Model | Target | Drivers |
|---|---|---|---|
| PD | Binary logistic regression | Default within 12 months (1/0) | Credit score, DTI, LTV, past-due flag, product, **GDP growth, unemployment, lending rate, house price growth** |
| LGD | Fractional logistic regression | Loss rate on defaulted loans (0-1) | LTV, product, **unemployment, house price growth** |
| EAD | Fractional logistic regression of CCF | Share of undrawn limit drawn at default | Utilisation, product, **unemployment** |

Each prediction is `1 / (1 + exp(-(borrower score + macro contribution)))`.
`EAD = drawn balance + CCF x undrawn limit`.

For every loan, scenario *s* (Base / Upside / Downside) and year *t* = 1..5:

```
loss(s,t) = survival(s,t-1) x PD(s,t) x LGD(s,t) x EAD(s,t) / (1 + EIR)^t
12-month ECL(s) = loss(s,1)        lifetime ECL(s) = sum of loss(s,t) to maturity
ECL = sum over scenarios of weight(s) x ECL(s)        (IFRS 9 para 5.5.17)
```

Staging: Stage 3 if 90+ days past due; Stage 2 if 30+ days past due or the weighted
12-month PD is at least 2x and +0.5 pp above the PD at origination; otherwise Stage 1.
Stage 1 recognises 12-month ECL, Stage 2 lifetime ECL, Stage 3 weighted LGD x EAD.

## Files

| File | Purpose |
|---|---|
| `generate_data.py` | Synthetic macro history, loan history, portfolio and scenarios |
| `ecl_models.py` | Model fitting, validation metrics, Python ECL engine |
| `build_workbook.py` | Writes the workbook with formulas |
| `run_all.py` | Runs the pipeline |
| `data/` | Input CSVs: replace with real data using the same column names |

## Using your own data

- `loan_history.csv.gz`: one row per loan observation with borrower features, the macro
  values for that period, `default_12m`, and `lgd` / `ccf` for defaulted loans.
- `macro_history.csv`: quarterly macro series (used for the history chart).
- `portfolio.csv`: the current loan book.
- `macro_scenarios.csv`: scenario, weight, year 1-5 and the four macro forecasts.

## Simplifications to discuss in class

- Lifetime is capped at five annual periods.
- The SICR test compares 12-month PDs as a proxy for lifetime PDs (IFRS 9 para B5.5.13).
- Drawn balances amortise straight-line; prepayment is ignored.
- Macro variables enter without lags; real models test lags, stationarity and multicollinearity.
- Staging is held constant in the single-scenario comparison on the Summary sheet.
