# GHG methodology for IFRS S2 (with Cambodian data)

Anchors: S2.29(a), B19–B65; GHG Protocol Corporate Standard (2004), Scope 2 Guidance, Scope 3 Standard;
UNFCCC GHG Emissions Calculator 2022 v1.3; DESNZ (DEFRA) conversion factors; IPCC AR6.
All factor values below are **illustrative** — confirm current published or supplier-specific values
and record source and year in GHG-02 before client use.

## Contents
1. Core formulas
2. Six-step process
3. GWP values
4. Factors used in examples (Cambodia)
5. Data collection plan
6. Data quality hierarchy and EDQA scoring
7. Scope 3 screening
8. Consolidation and multi-site
9. Worked example (Mekong Garment)
10. Case F — control failure and misstatement
11. Common errors

## 1. Core formulas
- Per-gas factors: E(tCO2e) = (1/1000) × Σ activity × EF(gas) × GWP(gas)
- Factor already in CO2e (DESNZ, most databases): E(tCO2e) = (1/1000) × Σ activity × EF(CO2e) —
  **do not multiply by GWP again.**
- Scope 2: location-based = kWh × grid EF; market-based = kWh × contract-specific EF (or residual mix).
  IFRS S2 requires location-based plus contractual-instrument information. Cambodia has no residual mix
  — market-based needs contract-specific factors (PPA, I-RECs).
- Refrigerants: tCO2e = kg leaked (or recharged, screening) × GWP ÷ 1,000.
- T&D losses (Scope 3 Cat 3) where no local factor: kWh × grid EF × loss rate ÷ (1 − loss rate).
- Proration of bills to period: kWh in period = bill kWh × days of bill inside period ÷ days in bill.
- Diesel consumption = opening stock + purchases − closing stock (reconcile to run-hours × burn rate).

## 2. Six-step process (each with working paper)
| Step | Do | Para | WP | Common error |
|---|---|---|---|---|
| 1 Boundary | Consolidation approach, sites, exclusions; classify Scope 1/2/3; screen Scope 3 | S2.29(a)(ii)–(iii); S1.20 | GHG-01, GHG-06 | Collecting data before defining boundary |
| 2 Activity data | Litres, kWh, km, kg, t from invoices, meters, logs, GL; prorate | S2.B39–B54 | GHG-03/03a/03b | Purchases instead of consumption; cut-off |
| 3 Factors | Best-representing factor with source, year, GWP basis | S2.B21–B22 | GHG-02 | GWP applied twice; stale factors |
| 4 Calculate | AD × EF ÷ 1,000; convert gases at AR6 | S2.B19–B22 | GHG-04/05/06 | Reporting calculator AR4 output unchanged |
| 5 Sum and report | Scope 1, 2 (location), 3 separately; categories; approach | S2.29(a)(i),(iii),(v),(vi) | GHG-00 | Netting credits; market-based instead of location-based |
| 6 Review and improve | Recalculate, reconcile to GL, log exceptions, raise quality | S1.77–82; ISSA 5000 | GHG-07/07b/08/09 | No independent review |

## 3. GWP values (100-year)
| Gas | AR5 | AR6 (use for IFRS S2) |
|---|---|---|
| CO2 | 1 | 1 |
| CH4 fossil | 28 | 29.8 |
| CH4 non-fossil (paddy, livestock, waste) | 28 | 27.0 |
| N2O | 265 | 273 |
| HFC-32 (R-32) | 677 | 771 |
| HFC-134a | 1,300 | 1,530 |
| R-410A (blend) | 2,088 (AR4) | ≈2,256 |
| SF6 | 23,500 | 25,200 |
Published factors already in CO2e need not be recalculated (B22).

## 4. Factors used in the examples
| Activity | Factor | Unit | Note |
|---|---|---|---|
| Cambodia grid electricity | 0.5878 | kg CO2e/kWh | IFI 2021 harmonised factor (UNFCCC calculator). ☐ Ministry of Environment factor: ACLEDA 2025 report states it switched to the MoE grid factor and restated 2023 — obtain the official value/year |
| Diesel, 100% mineral | 2.6988 | kg CO2e/L | Use rather than UK biofuel blend (2.5578) |
| Diesel well-to-tank | 0.6287 | kg CO2e/L | Scope 3 Cat 3 |
| Petrol | 2.1619 | kg CO2e/L | |
| LPG | 1.5571 | kg CO2e/L | |
| Van ≤3.5 t diesel | 0.2316 | kg CO2e/km | |
| Water supply / treatment | 0.149 / 0.272 | kg CO2e/m³ | UK proxies |
| Waste to landfill (C&I) | 467.0 | kg CO2e/t | UK proxy |
| Local bus (worker trucks proxy) | 0.1078 | kg CO2e/pkm | |
| Purchased fabric | 9.0 | kg CO2e/kg | Placeholder — replace with supplier EPD |
Known limitations: most calculator factors are DEFRA 2022 (UK); T&D UK-only; refrigerant GWPs AR4;
hotel factors exclude Cambodia; the tool records no evidence or approvals.

## 5. Data collection plan (Cambodian entity)
| Scope/source | Unit | Source document | Owner | Frequency | Key control |
|---|---|---|---|---|---|
| S1 diesel generators | L | Fuel invoices, tank dips, run-hour meters | Facilities | Monthly | Purchases vs stock vs run-hours |
| S1 vehicles | L or km | Fuel cards, logbooks | Fleet | Monthly | Fuel/km outliers |
| S1 LPG | kg/L | Supplier invoices | Facilities | Monthly | Cylinder count |
| S1 biomass (rice husk) | t | Weighbridge tickets | Production | Monthly | Calibration; biogenic CO2 reported separately, CH4/N2O in Scope 1 |
| S1 refrigerants | kg | Service records | Facilities | Per service | Asset register of AC/chillers |
| S1 land use (rubber, cassava, livestock) | ha, t, head | Land titles, plantation maps | Operations | Annual | Area tied to land register |
| S2 electricity | kWh | EDC/licensee bills, sub-meters | Finance | Monthly | Bills to AP ledger and meter reads |
| S2 solar PPA / I-RECs | kWh, certificates | PPA invoices, redemption statements | Finance | Monthly | No double claim |
| S3 Cat 1 purchased goods | kg / USD | Purchase ledger, supplier EPDs | Procurement | Quarterly | Supplier-specific > database > spend-based |
| S3 Cat 4/9 transport | t-km | Freight invoices, B/L | Logistics | Quarterly | Route distance check |
| S3 Cat 6/7 travel, commuting | km, nights | Travel reports, survey | HR | Annual | Survey response ≥70% |
| S3 Cat 15 financed | USD × attribution | Loan system, borrower data | Credit risk | Annual | PCAF score per loan |

## 6. Data quality
Hierarchy: (1) measured — meters, invoices, weighbridge, supplier primary data; (2) specific proxy —
engineering estimates, run-hours × burn rate; (3) generic — national averages, spend-based (disclose and
plan to replace). For Scope 3 IFRS S2 prioritises direct measurement, primary over secondary data and
verified data, all else equal (B43–B54).

**EDQA scoring** (each 1–3; total /15; <10 → improvement action and estimate note in the disclosure):
| Dimension | 3 strong | 2 acceptable | 1 weak |
|---|---|---|---|
| Completeness | All sites/months | One month or site estimated | Several gaps/unknown |
| Accuracy | Meter/invoice, reconciled | Supplier/engineering estimate | Spend-based/generic |
| Timeliness | Monthly, closed with financial close | Quarterly | Annual/after year-end |
| Consistency | Same method/factor YoY | Change disclosed and restated | Unexplained change |
| Source reliability | Third-party document (EDC, PPWSA, weigh ticket) | Internal log, second-person review | Uncontrolled spreadsheet |

## 7. Scope 3 screening (GHG-06)
Screen all 15 categories: relevant? (size, influence, risk, stakeholder/buyer demand, outsourcing,
sector guidance) → method (supplier-specific, hybrid, average-data, spend-based) → data quality → included
or excluded with reason. Categories: 1 purchased goods and services; 2 capital goods; 3 fuel- and
energy-related; 4 upstream transport; 5 waste; 6 business travel; 7 commuting; 8 upstream leased;
9 downstream transport; 10 processing of sold products; 11 use of sold products; 12 end-of-life;
13 downstream leased; 14 franchises; 15 investments (financed emissions).

## 8. Consolidation and multi-site
One boundary (e.g. operational control, GHG-01). Same template and units for every site; aggregate by
scope and category; check leased buildings and shared meters are not double counted; agree site totals
to GHG-00. Disaggregate Scope 1 and 2 for consolidated group vs other investees (S2.29(a)(iv)). Show an
equity-share figure only if additionally requested.

## 9. Worked example — Mekong Garment Co., FY2026 (illustrative)
| Source | Activity | Factor | tCO2e |
|---|---|---|---|
| Diesel generators (S1) | 50,000 L | 2.6988 | 134.94 |
| Delivery vans (S1) | 60,000 km | 0.2316 | 13.90 |
| R-410A recharged (S1) | 20 kg | 2,256 (AR6) | 45.12 (AR4 calculator: 41.76) |
| LPG canteen (S1) | 3,000 L | 1.5571 | 4.67 |
| **Scope 1** | | | **≈198.6** |
| EDC electricity (S2 location) | 1,200,000 kWh | 0.5878 | 705.36 |
| Diesel WTT (Cat 3) | 50,000 L | 0.6287 | 31.44 |
| Water supply (Cat 1) / wastewater (Cat 5) | 20,000 / 18,000 m³ | 0.149 / 0.272 | 2.98 / 4.90 |
| Waste to landfill (Cat 5) | 100 t | 467 | 46.70 |
| Worker bus/truck (Cat 7) | 400,000 pkm | 0.1078 | 43.12 |
| Purchased fabric (Cat 1) | 300,000 kg | 9.0 (placeholder) | 2,700.00 |
Scope 3 ≈ 2,829 t — about 75% of the inventory: supplier data matters more than refining Scope 1.

## 10. Case F — Green Garment Factory (control failure)
EDC invoices 120,000 kWh; monthly register 115,000; manager spreadsheet 123,000; factor copied from last
year (0.50); reported 60 tCO2e; nobody signed. Procedure: don't pick a figure — reconcile all three to
bills, meter-read dates and GL; confirm boundary (exclude out-of-boundary sub-meters; move generator
energy to Scope 1); replace factor with current documented Cambodian factor; recalculate: 120,000 ×
0.5878 ÷ 1,000 = 70.5 tCO2e. Understatement 10.5 t = 17.6% of reported (15% of corrected) — a material
misstatement of the metric caused by control failures (no review, no factor register, unreconciled
sources). kWh differences alone (≤8,000 kWh ≈ 4.7 t; 6.7% of invoices) exceed a 5% review threshold →
escalate. Independent review, exception log, remediation plan to CFO and audit committee.

## 11. Common errors
GWP twice; AR4 refrigerants; UK biofuel-blend diesel; purchases ≠ consumption; billing-cycle cut-off;
kWh vs MWh; generator kWh counted as Scope 2; market-based shown instead of location-based; credits
netted; Scope 3 omitted without screening; biogenic CO2 mixed into Scope 1; prior-year factor reused.
