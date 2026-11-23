# ISSB sustainability-reporting working papers (IFRS S1 / IFRS S2)

Illustrative working-paper file and Excel dashboard for a **fictional** Cambodian garment
manufacturer, *Tonle Threads Manufacturing Plc. (TTM)*, year ended 31 December 2025.
All client data are invented for teaching purposes.

| File | Purpose |
|---|---|
| `TTM_ISSB_Sustainability_Working_Papers_FY2025.xlsx` | The deliverable: Dashboard, Index and 27 cross-referenced working papers |
| `build_workpapers.py`, `wp_lib.py`, `wp_part1-5.py` | Python (xlsxwriter) generator that rebuilds the workbook |

## Working papers

| Ref | Content | Main ISSB references |
|---|---|---|
| A1-A4 | Basis of preparation, value chain, SRO register, materiality | S1.17-24, 30, 54-73, B1-B38, E3-E6; S2.10-12, C3-C5 |
| B1 | Governance | S1.26-27; S2.5-7 |
| C1-C3 | Strategy, transition plan, financial effects, scenario analysis | S1.28-42; S2.8-23, B1-B18 |
| D1 | Risk management | S1.43-44; S2.24-26 |
| E1-E9 | Activity data, emission factors, Scope 1/2/3, GHG summary, cross-industry and industry metrics, targets | S1.45-53; S2.27-37, B19-B71 |
| E10 | ESG KPI library (74 Environmental / Social / Governance / supply-chain KPIs) mapped to ISSB: required, material-SRO metric, or monitored only | S1.45-53, B25; S2.29, 32 |
| E11 | Supply-chain ESG register: tier-1 supplier environmental and social risk, assessment coverage, concentration | S1.32(b), B2-B5; S2.13(b), B47-B56 |
| E12 | ESG theme map: 14 rating themes, KPI coverage and evidence, and which themes belong in the ISSB report (coverage, not a score) | S1.17, B25-B30, 55-58 |
| G1 | ISSA 5000 assurance readiness: engagement parameters, readiness checklist, assertions, indicative materiality | ISSA 5000 (by topic); S1.B8-B10, D21-D24 |
| F1-F5 | Disclosure checklist, judgements and uncertainty, FS connectivity, evidence tracker, action log | S1.21-24, 72-86, B39-B59 |

Colour code: blue font = input, black = formula, green = link to another working paper,
yellow fill = key assumption. The dashboard is entirely formula-driven (sheet `Calc`).

## Rebuild

```
python3 build_workpapers.py
```
Open the result in Excel (or run a LibreOffice recalculation) so formula values are cached.

## Caveats

* Emission factors in `E2 EF-GWP` section 3 (grid and Scope 3) are placeholders of realistic
  magnitude - replace them with documented current factors before any real use.
* Requirements are summarised; always read the Standards. SASB metric codes should be
  verified against the current version.
* ISSA 5000 is cited by topic only (the standard was not among the source documents); add
  paragraph references and confirm its effective date before relying on G1.
* Jurisdictional adoption of ISSB Standards in Cambodia is treated as an assumption to confirm.
