"""PD / LGD / EAD models and a forward-looking IFRS 9 ECL engine.

PD   binary logistic regression (default within 12 months: 1/0)
LGD  fractional logistic regression (loss rate between 0 and 1)
EAD  fractional logistic regression of the credit conversion factor (CCF);
     EAD = drawn balance + CCF x undrawn commitment

All three share the logistic link, so each prediction is
    1 / (1 + exp(-(b0 + b1*x1 + ... + macro terms)))
which is exactly the formula the Excel workbook evaluates.
Macro variables in the regressions are what make the ECL forward-looking:
forecast values are substituted for each scenario and year.
"""
from dataclasses import dataclass
from pathlib import Path

import numpy as np
import pandas as pd
import statsmodels.api as sm
from sklearn.metrics import roc_auc_score, roc_curve
from sklearn.model_selection import train_test_split

DATA = Path(__file__).parent / "data"

MACRO = ["gdp_growth", "unemployment", "lending_rate", "hpi_growth"]
PD_VARS = ["credit_score", "dti", "ltv", "delinq", "consumer", "sme"] + MACRO
LGD_VARS = ["ltv", "consumer", "sme", "unemployment", "hpi_growth"]
CCF_VARS = ["utilisation", "sme", "unemployment"]
HORIZON = 5  # projection years

# Staging rules (also editable in the workbook)
SICR_MULTIPLE = 2.0      # 12m PD now >= 2x PD at origination ...
SICR_ABS = 0.005         # ... and at least +0.5 percentage points
DPD_STAGE2 = 30          # IFRS 9 para 5.5.11 rebuttable presumption
DPD_STAGE3 = 90          # IFRS 9 para B5.5.37 rebuttable presumption


def expit(z):
    return 1.0 / (1.0 + np.exp(-z))


@dataclass
class Fitted:
    name: str
    variables: list
    table: pd.DataFrame      # coef, se, z, p per variable (index includes 'const')
    metrics: dict

    @property
    def coef(self):
        return self.table["coef"]

    def linear(self, df):
        return self.coef["const"] + sum(self.coef[v] * df[v] for v in self.variables)

    def predict(self, df):
        return expit(self.linear(df))


def _table(res):
    return pd.DataFrame({"coef": res.params, "se": res.bse,
                         "z": res.tvalues, "p": res.pvalues})


def fit_pd(history):
    """Binary logistic regression with a 70/30 hold-out for validation."""
    train, test = train_test_split(history, test_size=0.30, random_state=1,
                                   stratify=history["default_12m"])
    res = sm.Logit(train["default_12m"], sm.add_constant(train[PD_VARS])).fit(disp=0)
    model = Fitted("PD", PD_VARS, _table(res), {})
    p_tr, p_te = model.predict(train), model.predict(test)
    fpr, tpr, _ = roc_curve(test["default_12m"], p_te)
    auc_te = roc_auc_score(test["default_12m"], p_te)
    model.metrics = {
        "Training observations": len(train),
        "Hold-out observations": len(test),
        "Defaults in training sample": int(train["default_12m"].sum()),
        "Observed default rate (hold-out)": test["default_12m"].mean(),
        "Mean predicted PD (hold-out)": p_te.mean(),
        "AUC (training)": roc_auc_score(train["default_12m"], p_tr),
        "AUC (hold-out)": auc_te,
        "Gini (hold-out) = 2*AUC - 1": 2 * auc_te - 1,
        "KS statistic (hold-out)": float(np.max(tpr - fpr)),
        "McFadden pseudo R-squared": res.prsquared,
    }
    # calibration by predicted-PD decile on the hold-out sample
    t = test.assign(pd_hat=p_te)
    t["decile"] = pd.qcut(t["pd_hat"], 10, labels=False) + 1
    calib = t.groupby("decile").agg(loans=("pd_hat", "size"),
                                    predicted_pd=("pd_hat", "mean"),
                                    observed_dr=("default_12m", "mean")).reset_index()
    return model, calib


def _fit_fractional(name, df, target, variables):
    """Fractional logit (Papke & Wooldridge): binomial GLM, logit link, robust SE."""
    X = sm.add_constant(df[variables])
    res = sm.GLM(df[target], X, family=sm.families.Binomial()).fit(cov_type="HC1")
    model = Fitted(name, variables, _table(res), {})
    pred = model.predict(df)
    model.metrics = {
        "Observations (defaulted loans)": len(df),
        f"Mean observed {name}": df[target].mean(),
        f"Mean predicted {name}": pred.mean(),
        "Mean absolute error": float(np.abs(df[target] - pred).mean()),
        "Correlation (observed, predicted)": float(np.corrcoef(df[target], pred)[0, 1]),
    }
    return model


def fit_lgd(history):
    return _fit_fractional("LGD", history.dropna(subset=["lgd"]), "lgd", LGD_VARS)


def fit_ccf(history):
    return _fit_fractional("CCF", history.dropna(subset=["ccf"]), "ccf", CCF_VARS)


def add_features(portfolio):
    p = portfolio.copy()
    p["delinq"] = ((p["dpd"] >= 1) & (p["dpd"] < DPD_STAGE3)).astype(int)
    p["consumer"] = (p["product"] == "Consumer").astype(int)
    p["sme"] = (p["product"] == "SME").astype(int)
    lim = p["drawn"] + p["undrawn"]
    p["utilisation"] = np.where(lim > 0, p["drawn"] / lim, 0.0)
    return p


def compute_ecl(portfolio, scenarios, pd_m, lgd_m, ccf_m):
    """Probability-weighted, discounted ECL per loan (mirrors the Excel engine).

    For scenario s and year t (t = 1..5, limited to the remaining term):
        PD_t, LGD_t, CCF_t = logistic(borrower score + macro contribution(s,t))
        EAD_t  = drawn x max(0, 1-(t-1)/term) + CCF_t x undrawn
        loss_t = survival_(t-1) x PD_t x LGD_t x EAD_t / (1+EIR)^t
    12-month ECL = loss_1; lifetime ECL = sum of loss_t.
    """
    p = add_features(portfolio)
    out = p[["loan_id", "product", "drawn"]].copy()
    w12 = wlt = wpd = ws3 = 0.0
    for name, sc in scenarios.groupby("scenario", sort=False):
        w = sc["weight"].iloc[0]
        surv = np.ones(len(p))
        ecl12 = eclt = None
        for _, yr in sc.sort_values("year").iterrows():
            t = int(yr["year"])
            x = p.assign(**{k: yr[k] for k in MACRO})
            pd_t, lgd_t, ccf_t = pd_m.predict(x), lgd_m.predict(x), ccf_m.predict(x)
            alive = (p["term_years"] >= t).astype(float)
            ead_t = p["drawn"] * np.maximum(0, 1 - (t - 1) / p["term_years"]) + ccf_t * p["undrawn"]
            loss_t = alive * surv * pd_t * lgd_t * ead_t / (1 + p["eir"]) ** t
            if t == 1:
                ecl12, eclt = loss_t.copy(), loss_t.copy()
                wpd = wpd + w * pd_t
                ws3 = ws3 + w * lgd_t * ead_t
                out[f"pd12_{name}"] = pd_t
            else:
                eclt = eclt + loss_t
            surv = surv * (1 - pd_t)
        out[f"ecl12_{name}"], out[f"ecllt_{name}"] = ecl12, eclt
        w12, wlt = w12 + w * ecl12, wlt + w * eclt

    out["pd12_weighted"] = wpd
    sicr = (wpd / p["orig_pd_12m"] >= SICR_MULTIPLE) & (wpd - p["orig_pd_12m"] >= SICR_ABS)
    out["stage"] = np.where(p["dpd"] >= DPD_STAGE3, 3,
                            np.where((p["dpd"] >= DPD_STAGE2) | sicr, 2, 1))
    out["ecl"] = np.where(out["stage"] == 3, ws3, np.where(out["stage"] == 2, wlt, w12))
    return out


def load_data():
    return (pd.read_csv(DATA / "loan_history.csv.gz"),
            pd.read_csv(DATA / "macro_history.csv"),
            pd.read_csv(DATA / "portfolio.csv"),
            pd.read_csv(DATA / "macro_scenarios.csv"))


def fit_all(history):
    pd_m, calib = fit_pd(history)
    return pd_m, fit_lgd(history), fit_ccf(history), calib


if __name__ == "__main__":
    history, macro, portfolio, scenarios = load_data()
    pd_m, lgd_m, ccf_m, calib = fit_all(history)
    for m in (pd_m, lgd_m, ccf_m):
        print(f"\n=== {m.name} model ===")
        print(m.table.round(4).to_string())
        for k, v in m.metrics.items():
            print(f"  {k}: {v:,.4f}" if isinstance(v, float) else f"  {k}: {v:,}")
    ecl = compute_ecl(portfolio, scenarios, pd_m, lgd_m, ccf_m)
    print("\n=== ECL by stage ===")
    print(ecl.groupby("stage").agg(loans=("ecl", "size"), gross=("drawn", "sum"),
                                   ecl=("ecl", "sum")).round(0).to_string())
    print(f"Total ECL: {ecl['ecl'].sum():,.0f}")
