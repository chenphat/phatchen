"""Generate ILLUSTRATIVE (synthetic) data for the IFRS 9 ECL teaching tool.

Nothing here is real bank or national-statistics data. Replace the four CSV
files in ./data with your own data (same column names) and re-run run_all.py.

Outputs
-------
data/macro_history.csv    quarterly macro series 2010Q1-2025Q4
data/loan_history.csv     loan-level 12-month performance observations
data/portfolio.csv        current loan book at the reporting date
data/macro_scenarios.csv  5-year macro forecasts for Base / Upside / Downside
"""
from pathlib import Path

import numpy as np
import pandas as pd

DATA = Path(__file__).parent / "data"
SEED = 42
LOANS_PER_QUARTER = 1000
PORTFOLIO_SIZE = 200

# "True" data-generating coefficients (the regressions should recover these).
TRUE_PD = dict(const=-0.25, credit_score=-0.009, dti=1.8, ltv=0.8, delinq=1.5,
               consumer=0.9, sme=0.4, gdp_growth=-0.10, unemployment=0.12,
               lending_rate=0.06, hpi_growth=-0.04)
TRUE_LGD = dict(const=-2.2, ltv=1.6, consumer=2.4, sme=0.5,
                unemployment=0.06, hpi_growth=-0.05)
TRUE_CCF = dict(const=0.2, utilisation=-1.5, sme=0.5, unemployment=0.05)


def expit(z):
    return 1.0 / (1.0 + np.exp(-z))


def make_macro(rng):
    quarters = pd.period_range("2010Q1", "2025Q4", freq="Q")
    n = len(quarters)
    # economic cycle: two downturns (2015-16 and 2020)
    shock = np.zeros(n)
    for start, depth, length in ((20, -4.0, 8), (40, -6.0, 5)):
        shock[start:start + length] += depth * np.sin(np.linspace(0, np.pi, length))
    gdp = np.empty(n)
    g = 5.0
    for t in range(n):
        g = 5.0 + 0.6 * (g - 5.0) + rng.normal(0, 0.8)
        gdp[t] = g + shock[t]
    unemp = 7.0 - 0.45 * (gdp - 5.0) + rng.normal(0, 0.9, n)
    rate = 9.0 + 1.6 * np.sin(np.arange(n) / 6.0) + rng.normal(0, 0.9, n)
    hpi = 4.0 + 0.5 * (gdp - 5.0) + rng.normal(0, 3.0, n)
    return pd.DataFrame({
        "quarter": quarters.astype(str),
        "gdp_growth": gdp.round(2),
        "unemployment": unemp.clip(2.5).round(2),
        "lending_rate": rate.round(2),
        "hpi_growth": hpi.round(2),
    })


def make_borrowers(rng, n):
    product = rng.choice(["Housing", "Consumer", "SME"], size=n, p=[0.45, 0.35, 0.20])
    score = rng.normal(650, 60, n).clip(400, 850).round(0)
    dti = (0.10 + 0.70 * rng.beta(2.2, 4.0, n)).round(3)
    ltv = np.where(product == "Housing", rng.uniform(0.40, 0.95, n),
                   np.where(product == "SME", rng.uniform(0.30, 0.90, n), 0.0)).round(3)
    df = pd.DataFrame({"product": product, "credit_score": score, "dti": dti, "ltv": ltv})
    df["consumer"] = (df["product"] == "Consumer").astype(int)
    df["sme"] = (df["product"] == "SME").astype(int)
    return df


def pd_logit(df, coefs=TRUE_PD):
    return coefs["const"] + sum(coefs[k] * df[k] for k in coefs if k != "const")


def make_history(rng, macro):
    frames = []
    for _, m in macro.iterrows():
        df = make_borrowers(rng, LOANS_PER_QUARTER)
        df.insert(0, "quarter", m["quarter"])
        for k in ("gdp_growth", "unemployment", "lending_rate", "hpi_growth"):
            df[k] = m[k]
        p_del = expit(-2.6 - 0.008 * (df["credit_score"] - 650))
        df["delinq"] = (rng.random(len(df)) < p_del).astype(int)

        df["default_12m"] = (rng.random(len(df)) < expit(pd_logit(df))).astype(int)

        # exposure and facility data
        limit = np.where(df["product"] == "Housing", rng.lognormal(11.6, 0.5, len(df)),
                         np.where(df["product"] == "SME", rng.lognormal(11.0, 0.7, len(df)),
                                  rng.lognormal(8.8, 0.6, len(df))))
        util = np.where(df["product"] == "Housing", 1.0, rng.beta(3, 2, len(df)))
        df["limit"] = limit.round(0)
        df["drawn"] = (limit * util).round(0)
        df["utilisation"] = (df["drawn"] / df["limit"]).round(4)

        # LGD and CCF are observed only for defaulted loans
        z_lgd = TRUE_LGD["const"] + sum(TRUE_LGD[k] * df[k] for k in TRUE_LGD if k != "const")
        lgd = expit(z_lgd + rng.normal(0, 0.6, len(df)))
        z_ccf = TRUE_CCF["const"] + sum(TRUE_CCF[k] * df[k] for k in TRUE_CCF if k != "const")
        ccf = expit(z_ccf + rng.normal(0, 0.7, len(df)))
        dflt = df["default_12m"] == 1
        df["lgd"] = np.where(dflt, lgd.round(4), np.nan)
        df["ccf"] = np.where(dflt & (df["product"] != "Housing"), ccf.round(4), np.nan)
        frames.append(df)
    out = pd.concat(frames, ignore_index=True)
    out.insert(0, "obs_id", np.arange(1, len(out) + 1))
    return out


def make_portfolio(rng, macro):
    n = PORTFOLIO_SIZE
    df = make_borrowers(rng, n)
    limit = np.where(df["product"] == "Housing", rng.lognormal(11.6, 0.5, n),
                     np.where(df["product"] == "SME", rng.lognormal(11.0, 0.7, n),
                              rng.lognormal(8.8, 0.6, n)))
    util = np.where(df["product"] == "Housing", 1.0, rng.beta(3, 2, n))
    drawn = np.round(limit * util, -2)
    undrawn = np.round(limit * (1 - util), -2)
    eir = np.where(df["product"] == "Housing", rng.uniform(0.07, 0.10, n),
                   np.where(df["product"] == "SME", rng.uniform(0.10, 0.14, n),
                            rng.uniform(0.14, 0.20, n))).round(4)
    term = np.where(df["product"] == "Housing", rng.integers(3, 6, n), rng.integers(1, 5, n))
    dpd = rng.choice([0, 15, 45, 120], size=n, p=[0.85, 0.08, 0.04, 0.03])
    dpd = np.where(dpd > 0, dpd + rng.integers(-10, 11, n), 0)

    # 12-month PD at origination: earlier (usually better) score, long-run macro
    orig = df.copy()
    orig["credit_score"] = (df["credit_score"] + rng.normal(0, 25, n)).clip(400, 850)
    orig["delinq"] = 0
    for k in ("gdp_growth", "unemployment", "lending_rate", "hpi_growth"):
        orig[k] = macro[k].mean()
    orig_pd = expit(pd_logit(orig))

    return pd.DataFrame({
        "loan_id": [f"L{i:04d}" for i in range(1, n + 1)],
        "product": df["product"],
        "drawn": drawn,
        "undrawn": undrawn,
        "eir": eir,
        "term_years": term,
        "credit_score": df["credit_score"],
        "dti": df["dti"],
        "ltv": df["ltv"],
        "dpd": dpd,
        "orig_pd_12m": orig_pd.round(5),
    })


def make_scenarios():
    rows = []
    paths = {
        # scenario: (weight, gdp, unemployment, lending rate, hpi) for years 1-5
        "Base":     (0.50, [5.0, 5.2, 5.3, 5.2, 5.0], [6.8, 6.6, 6.5, 6.6, 6.8],
                     [9.5, 9.2, 9.0, 9.0, 9.0], [4.0, 4.5, 4.5, 4.0, 4.0]),
        "Upside":   (0.20, [6.5, 6.8, 6.2, 5.6, 5.2], [6.0, 5.6, 5.8, 6.2, 6.5],
                     [8.8, 8.5, 8.6, 8.8, 9.0], [7.0, 7.5, 6.0, 5.0, 4.5]),
        "Downside": (0.30, [0.5, 1.5, 3.5, 4.5, 5.0], [9.5, 10.5, 9.5, 8.2, 7.2],
                     [12.5, 12.0, 10.5, 9.5, 9.0], [-6.0, -3.0, 1.0, 3.0, 4.0]),
    }
    for name, (w, gdp, un, rate, hpi) in paths.items():
        for t in range(5):
            rows.append(dict(scenario=name, weight=w, year=t + 1, gdp_growth=gdp[t],
                             unemployment=un[t], lending_rate=rate[t], hpi_growth=hpi[t]))
    return pd.DataFrame(rows)


def main():
    rng = np.random.default_rng(SEED)
    DATA.mkdir(exist_ok=True)
    macro = make_macro(rng)
    history = make_history(rng, macro)
    portfolio = make_portfolio(rng, macro)
    scenarios = make_scenarios()
    macro.to_csv(DATA / "macro_history.csv", index=False)
    history.to_csv(DATA / "loan_history.csv.gz", index=False)
    portfolio.to_csv(DATA / "portfolio.csv", index=False)
    scenarios.to_csv(DATA / "macro_scenarios.csv", index=False)
    print(f"history: {len(history):,} obs, default rate {history.default_12m.mean():.2%}; "
          f"portfolio: {len(portfolio)} loans")


if __name__ == "__main__":
    main()
