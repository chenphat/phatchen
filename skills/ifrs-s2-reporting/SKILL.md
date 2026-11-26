---
name: ifrs-s2-reporting
description: Process, procedures and Excel working-paper templates for IFRS S2 (ISSB) climate-related disclosures - climate governance, physical and transition risks and opportunities, transition plans, financial effects, scenario analysis and resilience (SA1-SA9), GHG accounting under the GHG Protocol (Scope 1, 2 location-based, 3; AR6 GWPs; GHG-00 to GHG-09 with factor register and reconciliations), S2.29 cross-industry metrics, targets, carbon credits and Cambodia carbon pricing/ETS scenarios, PCAF financed emissions for banks and MFIs, December 2025 amendments and reliefs, with Cambodian factors and data. Use whenever the user mentions IFRS S2, ISSB climate, TCFD, climate risk, scenario analysis, transition plan, GHG inventory, carbon footprint, Scope 1/2/3, emission or grid factors, financed emissions, carbon pricing, ETS or carbon tax, Article 6 credits, internal carbon price, net-zero or SBTi targets, or climate working papers, even without naming IFRS S2. Use with ifrs-s1-reporting for general requirements.
---

# IFRS S2 reporting — process, procedures and working papers

IFRS S2 requires information about climate-related risks and opportunities (CRROs) that could reasonably
be expected to affect the entity's prospects, organised in the four pillars and applying IFRS S1's general
requirements (boundary, materiality, connectivity, judgements, comparatives, compliance statement). Load
`ifrs-s1-reporting` alongside this skill for those; this skill adds the climate-specific content,
calculations and working papers.

## How to use this skill

1. **Confirm the context**: entity, industry (for Industry-based Guidance / SASB), jurisdiction, FS
   framework, year end, first year of application (reliefs S2.C3–C5; S1.E3–E6), whether the entity is in
   asset management, commercial banking or insurance (financed emissions), whether it applies the
   December 2025 amendments early (mandatory from 1 January 2027).
2. **Choose the deliverable**: GHG inventory, scenario analysis, climate risk register, transition plan
   disclosure, full S2 note, review of a draft/published report, or teaching case.
3. **Generate the workbook** when templates are wanted:
   ```bash
   python scripts/build_s2_workpapers.py --entity "Mekong Garment Co." --year-end 2026-12-31 \
       --currency "KHR million" --out S2_Working_Papers.xlsx [--bank] [--no-examples]
   ```
   `--bank` adds the PCAF financed-emissions sheet (FE-01). Recalculate formulas (xlsx skill
   `scripts/recalc.py`) before sharing. Example values are illustrative (Mekong Garment case, UNFCCC
   calculator/DEFRA 2022 factors); the factor register flags every factor to replace with a current
   published or supplier-specific value.
4. **Cite precisely** using `references/s2-disclosure-checklist.md`; cite a range when unsure of a letter.

## Reference files

| File | Read it when |
|---|---|
| `references/s2-disclosure-checklist.md` | Drafting/reviewing any S2 disclosure; paragraph map; Dec 2025 amendments; reliefs |
| `references/ghg-methodology.md` | Any GHG calculation or review: six steps, formulas, GWPs, factors, Cambodian grid factor, data quality (EDQA), common errors, Case F |
| `references/scenario-analysis.md` | Climate resilience/scenario analysis (S2.22, B1–B18): proportionality, quality ladder, SA1–SA9, financial-effect translation |
| `references/financed-emissions.md` | Banks, MFIs, asset managers, insurers: Cat 15, PCAF methods and data quality, climate-adjusted ECL |
| `references/cambodia-carbon-pricing.md` | Carbon pricing / possible Cambodian ETS, Article 6 credits, CBAM, internal carbon price, allowance accounting (CR-06, CR-07) |

## The process

### Phase 1 — Scoping and climate governance → CR-00 (index), S1 WPs
Agree boundary and period via S1-02 (same as FS; S1.20). Governance (S2.5–7): body responsible for
climate oversight, terms of reference, skills, frequency of information, consideration in strategy and
major transactions, target oversight and remuneration links (S2.6(a)); management's role and controls
(S2.6(b)). Avoid duplicating S1 governance where integrated (S2.7).

### Phase 2 — Identify climate-related risks and opportunities → CR-01
- Describe each CRRO, classify as **physical (acute/chronic)** or **transition (policy/legal,
  technology, market, reputation)**, assign horizon and define horizons (S2.10).
- Refer to and consider the Industry-based Guidance topics (S2.12, S2.32).
- Assess by likelihood × magnitude using S1-03 scales; consider own operations, upstream, downstream.
- Include nature and social aspects of climate risks (ISSB educational material, Dec 2023): e.g.
  floods → crop yields → borrower arrears.

### Phase 3 — Strategy, transition plan and financial effects → CR-02
- Business model and value-chain effects and concentration (S2.13).
- Strategy and decision-making (S2.14): changes to business model and resource allocation; direct and
  indirect mitigation and adaptation; **transition plan** (key assumptions, dependencies); how targets
  will be achieved; how plans are resourced; progress on previously disclosed plans. Label plans as
  ambitions, not results.
- Financial effects (S2.15–21): current-period effects; significant risk of material adjustment next
  year; anticipated effects on financial position (investment/disposal plans, funding) and on
  performance and cash flows. Quantitative unless not separately identifiable, uncertainty too high, or
  (anticipated) lacking skills/resources — then explain and give qualitative information (S2.19–21).
- Tie every amount to the FS through S1-09 (impairment, useful lives, provisions, IAS 41, IFRS 9 ECL).

### Phase 4 — Climate resilience and scenario analysis → SA1–SA9
Follow `references/scenario-analysis.md`. Approach commensurate with circumstances (S2.22, B1–B18):
qualitative narratives are acceptable for low capability, quantitative expected as resources allow (B17).
Cover physical and transition risks; a diverse range including a scenario aligned with the latest
international agreement on climate change (Paris); state sources (NGFS, IPCC SSP-RCP, IEA), horizons,
scope, key assumptions, results, uncertainties and capacity to adjust (S2.22(a)–(b)). Scenario analysis
updated at least with the strategic planning cycle; resilience assessed every reporting period (B18).

### Phase 5 — Risk management → S1-07 + CR-01
Processes to identify, assess, prioritise and monitor CRROs, inputs and parameters, scenario use,
integration into ERM, changes from the prior period (S2.24–26).

### Phase 6 — GHG emissions → GHG-00 to GHG-09
Six steps (see `ghg-methodology.md`):
1. **Boundary** — consolidation approach (equity share, financial or operational control), sites,
   exclusions; Scope 3 screening of all 15 categories (GHG-01, GHG-06).
2. **Collect activity data** — consumption, not purchases; prorate billing periods (GHG-03/03a/03b).
3. **Factors** — best-representing factor with source, year and GWP basis in the factor register
   (GHG-02); Cambodia grid factor for location-based Scope 2.
4. **Calculate** — AD × EF ÷ 1,000 = tCO2e; convert non-CO2 gases with 100-year AR6 GWPs (S2.B21–B22;
   jurisdictional relief under the Dec 2025 amendments); don't apply GWP twice (GHG-04/05/06).
5. **Sum and report** — Scope 1, Scope 2 location-based and Scope 3 separately, absolute gross tCO2e;
   Scope 3 categories; disaggregate Scope 1 and 2 between the consolidated group and other investees
   (associates, JVs, unconsolidated subsidiaries); contractual-instrument information for Scope 2;
   measurement approach, inputs, assumptions and changes (S2.29(a)) (GHG-00).
6. **Review, evidence, improve** — reconcile to GL, recalculate, log exceptions, raise data quality
   (GHG-07, 07b, 08, 09).

### Phase 7 — Other cross-industry and industry metrics → CR-03, CR-05
S2.29(b)–(g): amount and percentage of assets or business activities vulnerable to transition risks and to
physical risks; aligned with climate-related opportunities; capital expenditure, financing or investment
deployed towards CRROs; internal carbon price (whether and how used, price per tCO2e); climate-linked
executive remuneration (whether and how, and % of remuneration recognised in the period). Industry-based
metrics: refer to and consider the Industry-based Guidance (S2.32).

### Phase 8 — Targets and carbon credits → CR-04
For each target (S2.33–36): metric, objective (mitigation, adaptation, sector initiative), part of entity,
period, base period, milestones/interim targets, absolute or intensity, how the latest international
agreement informed it, third-party validation (e.g. SBTi), review process, revisions, performance and
trends. GHG targets: gases, scopes, **gross vs net (if net, also gross)**, sectoral decarbonisation
approach, planned use of carbon credits — extent of reliance, scheme, type (nature-based/technological;
removal/avoidance), credibility factors. Never net credits against the inventory.

### Phase 8a — Carbon pricing, ETS exposure and credits → CR-06, CR-07
No Cambodian ETS or carbon tax is in force (Nov 2026); a carbon-market sub-decree is an NDC 3.0
milestone and Article 6 rules exist. Model a possible ETS as a **scenario**, not a fact
(`references/cambodia-carbon-pricing.md`): direct compliance cost above the coverage threshold
(free allocation, offset limit), indirect cost through power-tariff pass-through, by scenario price
path; feed the ranges into S2.15–21 and SA6, and the price into S2.29(f) if used internally. Record
Article 6/voluntary credits in CR-07 with authorisation and corresponding-adjustment status for
S2.36(e). If a scheme starts, set an allowance accounting policy and keep assumptions consistent with
the FS (S1.23).

### Phase 9 — Financed emissions (banks, MFIs, asset managers, insurers) → FE-01
See `references/financed-emissions.md`. S2.29(a)(vi)(2) and B58–B63 with the December 2025 amendments
(limits on derivatives, facilitated and insurance-associated emissions; alternative industry
classification to GICS).

### Phase 10 — Compile and review → S2-CL, then S1-11 to S1-14
Complete the S2 checklist; reconcile every disclosed figure to GHG-00/CR sheets and the S1-13 evidence
index; review against the scenario quality checklist and anti-greenwashing controls; joint sign-off.

## Working-paper index (`scripts/build_s2_workpapers.py`)

| WP | Content | Anchor |
|---|---|---|
| CR-00 | Index, reliefs elected, amendments applied | S2.C1–C5; Dec 2025 amendments |
| CR-01 | Climate risk and opportunity register (physical/transition) | S2.10–12, 24–26 |
| CR-02 | Strategy, transition plan, financial effects | S2.13–21 |
| SA | Scenario analysis SA1–SA9 + quality ladder | S2.22, B1–B18 |
| CR-03 | Cross-industry metrics (b)–(g) | S2.29(b)–(g) |
| CR-04 | Targets and carbon credits | S2.33–36 |
| CR-05 | Industry-based metrics | S2.32 |
| CR-06 | Carbon pricing / Cambodia ETS exposure model (scenario-based: direct, indirect, credits, NPV) | S2.10, 15–22, 29(f) |
| CR-07 | Carbon credit register (Article 6 / voluntary; LoA, corresponding adjustment, type) | S2.36(e) |
| GHG-00 | Lead schedule (scopes, categories, prior year, intensity, checks) | S2.29(a) |
| GHG-01 | Boundary and approach memo | S2.29(a)(ii)–(iii) |
| GHG-02 | Emission factor register (source, year, GWP basis, approval) | S2.B21–B22 |
| GHG-03 / 03a / 03b | Source-data register; invoice register with cut-off proration; stock and logs | S2.B39–B54 (Scope 3 data) |
| GHG-04 / 05 / 06 | Scope 1; Scope 2 (location + contractual); Scope 3 screening and calculations | S2.29(a)(i), (v), (vi) |
| GHG-07 / 07b | Three-way reconciliation and recalculation (Case F); invoice register vs GL | S1.77–82 |
| GHG-08 | Review checklist and EDQA scoring | — |
| GHG-09 | Assurance evidence index (GHG) | ISSA 5000 |
| FE-01 | Financed emissions (PCAF), `--bank` only | S2.29(a)(vi)(2), B58–B63 |
| S2-CL | IFRS S2 disclosure checklist | whole standard |

## Quality rules

- **Report gross, by scope, in tCO2e.** A single combined total, market-based Scope 2 instead of
  location-based, or emissions net of credits fails S2.29(a).
- **Calculators are not controls.** Adjust tool output (e.g. AR4 refrigerant GWPs → AR6) and document it.
- **Every factor has a source, year, unit and GWP basis**, approved in GHG-02; stale factors are the
  most common material misstatement (Case F: 0.50 vs 0.5878 kg CO2e/kWh understated Scope 2 by 17.6%).
- **Proportionate, not optional.** Scenario analysis and anticipated financial effects may be
  qualitative at first, but the explanation of why is itself a required disclosure.
- **Consistency with the FS** for prices, FX, asset lives and forecasts (S1.23); explain differences.
- **Ambition ≠ result.** Net-zero pledges, transition plans and targets are labelled forward-looking;
  progress is measured and evidenced.
- **State the version and open items**: IFRS S2 (June 2023) with Dec 2025 amendments; local
  adoption, grid factor and carbon-market rules as items to confirm.

## Output formats
- **GHG statement**: table Scope 1 | Scope 2 (location) | Scope 2 contractual info | Scope 3 by
  category | total of disclosed scopes, current and prior year, then basis of preparation (consolidation
  approach, GHG Protocol standards, GWP source, factor sources and years, exclusions, estimates,
  restatements).
- **Scenario summary**: scenario × risk type matrix with sources, horizon, key assumptions, results
  (financial effect ranges), strategic implications, capacity to adjust, uncertainties, quality-ladder level.
- **Review report**: Requirement | Para | Finding | Effect | Recommendation | Priority.
