# Cambodia carbon pricing, ETS readiness and carbon credits

Working papers: CR-06 (carbon pricing / ETS exposure model), CR-07 (carbon credit register).
IFRS anchors: S2.10(b) transition risk (policy/legal); S2.14(a)(iv)–(v) transition plan; S2.15–21
anticipated financial effects; S2.22 and SA6 scenario analysis; S2.29(f) internal carbon price;
S2.36(e) planned use of carbon credits; S1.23 consistency with the financial statements.

## 1. Status (as understood at November 2026 — confirm each ☐ before use)
| Item | Position | ☐ |
|---|---|---|
| Domestic ETS or carbon tax | **None in force.** Cambodia has no mandatory carbon pricing instrument (Slaughter and May, *ESG in APAC 2025 — Cambodia*). | ☐ |
| Carbon market sub-decree | NDC 3.0 (August 2025) sets the milestone "by 2026: Sub-Decree approved". Approval not confirmed. | ☐ |
| Article 6 framework | Ministry of Environment Article 6 Operations Manual (reported as December 2023 / 2024). Eligible mechanisms: JCM, Gold Standard, VCS, Article 6.4 (PACM). Forestry credits counted toward the NDC are not exportable; transfers to companies via long-term (~15-year) contracts. | ☐ |
| Institution | Carbon Crediting Secretariat (CCS) leads implementation (Khmer Times, May 2026). | ☐ |
| Bilateral | JCM with Japan; Korea Article 6.2 partnership (first authorised transport project 2025); Cambodia–Singapore MoU (2023). Letters of authorisation already issued; credits supplied to CORSIA airlines. | ☐ |
| Registry | UN PACM registry and independent-mechanism registries until a national registry exists. | ☐ |
| External carbon prices | EU CBAM definitive regime from 2026 (cement, iron and steel, aluminium, fertilisers, hydrogen, electricity) — limited direct effect on garments; buyer internal carbon prices and supplier decarbonisation clauses matter more. A domestic carbon price paid may reduce CBAM liability, which is one policy argument for a Cambodian scheme. | ☐ |
Do not describe a Cambodian ETS as existing. Model it as a **scenario** (CR-06) and disclose the
assumptions.

## 2. Why it belongs in an IFRS S2 report even though no ETS exists
- A *possible* policy change that could reasonably be expected to affect cash flows is a transition risk
  (S2.10). The sub-decree milestone, NDC 3.0 (up to 55% by 2035, conditional) and regional schemes
  (Indonesia, Vietnam pilot, China, Korea) make it reasonable to assess.
- Anticipated financial effects may be given as **ranges** across scenarios (S2.15–21); the CR-06
  summary provides cumulative cost, NPV and peak annual cost per scenario.
- If management uses a price path in investment decisions, that is an **internal carbon price** —
  disclose whether and how it is applied and the price per tCO2e (S2.29(f)).
- Planned credits (Article 6 or voluntary) used to reach net targets must be disclosed — extent,
  scheme, type, integrity (S2.36(e)) — never netted against gross emissions.

## 3. CR-06 model logic
Inputs (yellow, all assumptions): earliest compliance year; coverage threshold (direct Scope 1 per
entity/facility); whether power generation is covered and the pass-through rate; free allocation in the
first year and its annual decline; offset limit and eligible credit price; emissions trajectory; FX
(presentation units per USD); discount rate; EBITDA. Linked inputs: base-year Scope 1 (GHG-00), purchased
kWh (GHG-05), grid factor (GHG-02).

Per scenario and year (2027–2035):
- Scheme start = later of the earliest year and the first year the scenario price is positive.
- Covered emissions = base Scope 1 × (1 + annual change)^(year − base year).
- In scheme = year ≥ start AND covered emissions ≥ threshold.
- Free allocation = max(0, first-year % − decline × years since start).
- Net obligation = covered × (1 − free %) × in-scheme flag; credits used = obligation × offset limit;
  allowances purchased = obligation − credits.
- Direct cost = allowances × price + credits × credit price.
- Indirect cost = kWh × grid EF ÷ 1,000 × price × pass-through × power-covered flag (from start year).
- Total in USD and presentation currency; % of EBITDA; NPV.

Illustrative price paths (USD/tCO2e, 2027→2035): Current Policies 0 until 2030, then 5→8; NDC/Delayed
5 (2029)→15; Net Zero 10→50. These are teaching assumptions, not forecasts — replace with the entity's
SA2 scenario sources (e.g. NGFS carbon price variables for Asia) and document the choice.

Mekong Garment result (≈199 tCO2e Scope 1, 1.2 GWh): below a 25,000 t threshold, so **no direct
compliance cost**; indirect power-tariff cost USD 11k–91k cumulative 2027–2035 across scenarios (NPV
KHR 22–208 million) — immaterial against EBITDA, but the qualitative risk (buyer and tariff pass-through)
is still disclosed. Lowering the threshold to 100 t brings the entity into the scheme (direct cost
≈USD 25k cumulative in Net Zero) — a useful sensitivity for teaching.

High-exposure Cambodian sectors to test with larger inputs: cement, brick kilns, power generation
(coal, diesel), steel, garment groups with coal/biomass boilers, aviation (CORSIA), and banks through
financed emissions (borrowers' compliance costs → PD/LGD in IFRS 9 ECL).

## 4. Accounting if a scheme starts (connectivity, S1.21–23)
No IFRS deals specifically with emission allowances (IFRIC 3 was withdrawn). Select and disclose an
accounting policy (IAS 8), commonly:
- Allowances: intangible assets (IAS 38) or inventory (IAS 2, if held for trading/production use),
  at cost (granted allowances often at nil) — or a government-grant approach (IAS 20) at fair value.
- Obligation: liability under IAS 37 as emissions occur, measured at the carrying amount of allowances
  held plus market price for any shortfall ("net" approach).
- Purchased carbon credits held for own voluntary use: typically intangible assets; expense on retirement.
- Use the same carbon price, FX and emissions assumptions in impairment tests (IAS 36), useful lives and
  provisions as in CR-06, or explain the difference (S1.23).

## 5. Carbon credit integrity checklist (CR-07)
Eligible mechanism; letter of authorisation from the MoE; corresponding adjustment (needed for
internationally transferred mitigation outcomes and most net-zero claims); vintage; removal vs avoidance;
nature-based vs technological; permanence and reversal buffer; additionality; leakage; double-claiming
with the NDC; registry serial numbers and retirement evidence; price and cost; intended use (S2.36(e) net
target, offset under a future ETS, voluntary claim).
