#!/usr/bin/env python3
"""Budget vs Actual calculator (static budget -> flexible budget -> actual).

Standard library only. Reproduces the worked example and runs on your own data.

Input CSV columns (header required):
  line_item       free text
  category        revenue | cogs | opex
  behavior        variable | fixed   (variable lines flex with volume)
  budget_amount   static (master) budget amount for the period
  actual_amount   actual amount for the period

Usage:
  python bva_calculator.py INPUT.csv --budget-volume 1000 --actual-volume 1100
  python bva_calculator.py INPUT.csv --budget-volume 1000 --actual-volume 1100 \
      --out results.csv --check expected_results.csv
"""
from __future__ import annotations

import argparse
import csv
import sys
from decimal import Decimal, ROUND_HALF_UP

CATEGORIES = ("revenue", "cogs", "opex")
BEHAVIORS = ("variable", "fixed")
OUT_COLS = [
    "line_item", "category", "behavior",
    "static_budget", "flexible_budget", "actual",
    "volume_variance", "flexible_budget_variance", "static_budget_variance",
    "fb_variance_pct", "direction", "material",
]
TWO = Decimal("0.01")


def q(x: Decimal) -> Decimal:
    return x.quantize(TWO, rounding=ROUND_HALF_UP)


def direction(category: str, variance: Decimal) -> str:
    """F/U from the profit point of view: more revenue or less cost is favourable."""
    if variance == 0:
        return "-"
    good = variance > 0 if category == "revenue" else variance < 0
    return "F" if good else "U"


def load(path: str) -> list[dict]:
    rows, errors = [], []
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        need = {"line_item", "category", "behavior", "budget_amount", "actual_amount"}
        missing = need - set(reader.fieldnames or [])
        if missing:
            sys.exit(f"ERROR: missing columns: {', '.join(sorted(missing))}")
        for n, r in enumerate(reader, start=2):
            cat = r["category"].strip().lower()
            beh = r["behavior"].strip().lower()
            if cat not in CATEGORIES:
                errors.append(f"row {n}: category '{cat}' not in {CATEGORIES}")
            if beh not in BEHAVIORS:
                errors.append(f"row {n}: behavior '{beh}' not in {BEHAVIORS}")
            try:
                b = Decimal(r["budget_amount"].replace(",", "").strip())
                a = Decimal(r["actual_amount"].replace(",", "").strip())
            except Exception:
                errors.append(f"row {n}: non-numeric amount")
                continue
            rows.append({"line_item": r["line_item"].strip(), "category": cat,
                         "behavior": beh, "budget": b, "actual": a})
    if errors:
        sys.exit("ERROR:\n  " + "\n  ".join(errors))
    return rows


def analyse(rows, budget_volume: Decimal, actual_volume: Decimal,
            abs_threshold: Decimal, pct_threshold: Decimal) -> list[dict]:
    ratio = actual_volume / budget_volume
    out = []
    for r in rows:
        static, actual = r["budget"], r["actual"]
        flex = static * ratio if r["behavior"] == "variable" else static
        vol_var = flex - static
        fb_var = actual - flex
        st_var = actual - static
        pct = (fb_var / flex * 100) if flex != 0 else Decimal(0)
        material = abs(fb_var) >= abs_threshold or abs(pct) >= pct_threshold
        out.append({
            "line_item": r["line_item"], "category": r["category"], "behavior": r["behavior"],
            "static_budget": q(static), "flexible_budget": q(flex), "actual": q(actual),
            "volume_variance": q(vol_var), "flexible_budget_variance": q(fb_var),
            "static_budget_variance": q(st_var), "fb_variance_pct": q(pct),
            "direction": direction(r["category"], fb_var),
            "material": "YES" if material else "no",
        })
    return out


def totals(lines: list[dict]) -> dict:
    keys = ["static_budget", "flexible_budget", "actual",
            "volume_variance", "flexible_budget_variance", "static_budget_variance"]
    t = {}
    for cat in CATEGORIES:
        t[cat] = {k: sum((l[k] for l in lines if l["category"] == cat), Decimal(0)) for k in keys}
    t["gross_profit"] = {k: t["revenue"][k] - t["cogs"][k] for k in keys}
    t["operating_income"] = {k: t["gross_profit"][k] - t["opex"][k] for k in keys}
    return t


def checks(lines, t) -> list[tuple[str, bool]]:
    res = []
    res.append(("Each line: volume + flexible-budget variance = static-budget variance",
                all(l["volume_variance"] + l["flexible_budget_variance"] == l["static_budget_variance"]
                    for l in lines)))
    res.append(("Fixed lines have zero volume variance",
                all(l["volume_variance"] == 0 for l in lines if l["behavior"] == "fixed")))
    oi = t["operating_income"]
    res.append(("Operating income: actual - static budget = static-budget variance",
                oi["actual"] - oi["static_budget"] == oi["static_budget_variance"]))
    res.append(("Operating income: volume + flexible-budget variance = static-budget variance",
                oi["volume_variance"] + oi["flexible_budget_variance"] == oi["static_budget_variance"]))
    return res


def fmt(x: Decimal) -> str:
    return f"{x:,.2f}"


def report(lines, t, chk, bv, av) -> str:
    def fu(cat, v):  # direction for subtotal rows (profit lines behave like revenue)
        return direction("revenue" if cat in ("revenue", "gross_profit", "operating_income") else "cost", v)
    w = []
    w.append(f"Volume: budget {bv:,} units, actual {av:,} units (ratio {av / bv:.4f})\n")
    hdr = f"{'Line':34}{'Static':>13}{'Flexible':>13}{'Actual':>13}{'Volume var':>14}{'Flex var':>14}  Mat"
    w.append(hdr)
    w.append("-" * len(hdr))
    for l in lines:
        w.append(f"{l['line_item'][:33]:34}{fmt(l['static_budget']):>13}{fmt(l['flexible_budget']):>13}"
                 f"{fmt(l['actual']):>13}{fmt(l['volume_variance']):>12} {direction(l['category'], l['volume_variance']):1}"
                 f"{fmt(l['flexible_budget_variance']):>12} {l['direction']:1}  {l['material']}")
    w.append("-" * len(hdr))
    for key, label in [("revenue", "Total revenue"), ("cogs", "Total COGS"), ("gross_profit", "Gross profit"),
                       ("opex", "Total operating expenses"), ("operating_income", "OPERATING INCOME")]:
        s = t[key]
        w.append(f"{label:34}{fmt(s['static_budget']):>13}{fmt(s['flexible_budget']):>13}{fmt(s['actual']):>13}"
                 f"{fmt(s['volume_variance']):>12} {fu(key, s['volume_variance']):1}"
                 f"{fmt(s['flexible_budget_variance']):>12} {fu(key, s['flexible_budget_variance']):1}")
    w.append("\nReview checks:")
    for name, ok in chk:
        w.append(f"  [{'PASS' if ok else 'FAIL'}] {name}")
    return "\n".join(w)


def compare(lines, expected_path) -> bool:
    with open(expected_path, newline="", encoding="utf-8-sig") as f:
        exp = {r["line_item"]: r for r in csv.DictReader(f)}
    ok = True
    for l in lines:
        e = exp.get(l["line_item"])
        if e is None:
            print(f"  [FAIL] {l['line_item']}: not in expected file"); ok = False; continue
        for c in OUT_COLS:
            if str(l[c]) != e[c].strip():
                print(f"  [FAIL] {l['line_item']} / {c}: got {l[c]}, expected {e[c]}"); ok = False
    if len(exp) != len(lines):
        print(f"  [FAIL] row count: got {len(lines)}, expected {len(exp)}"); ok = False
    print(f"  [{'PASS' if ok else 'FAIL'}] Output matches {expected_path}")
    return ok


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("input")
    p.add_argument("--budget-volume", type=Decimal, required=True, help="budgeted output/sales units")
    p.add_argument("--actual-volume", type=Decimal, required=True, help="actual output/sales units")
    p.add_argument("--abs-threshold", type=Decimal, default=Decimal("2000"),
                   help="materiality: flag if |flexible-budget variance| >= this (default 2000)")
    p.add_argument("--pct-threshold", type=Decimal, default=Decimal("10"),
                   help="materiality: OR flag if |variance %% of flexible budget| >= this (default 10)")
    p.add_argument("--out", help="write line-level results CSV here")
    p.add_argument("--check", help="compare line-level results with this expected CSV")
    a = p.parse_args()
    if a.budget_volume <= 0:
        sys.exit("ERROR: --budget-volume must be > 0")

    lines = analyse(load(a.input), a.budget_volume, a.actual_volume, a.abs_threshold, a.pct_threshold)
    t = totals(lines)
    chk = checks(lines, t)
    print(report(lines, t, chk, a.budget_volume, a.actual_volume))

    if a.out:
        with open(a.out, "w", newline="", encoding="utf-8") as f:
            wr = csv.DictWriter(f, fieldnames=OUT_COLS)
            wr.writeheader()
            wr.writerows(lines)
        print(f"\nWrote {a.out}")
    ok = all(c[1] for c in chk)
    if a.check:
        print("\nExpected-results comparison:")
        ok = compare(lines, a.check) and ok
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
