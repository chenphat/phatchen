# ISSB sustainability-reporting working papers (IFRS S1 / IFRS S2)

Illustrative working-paper file and Excel dashboard for a **fictional** Cambodian garment
manufacturer, *Tonle Threads Manufacturing Plc. (TTM)*, year ended 31 December 2025.
All client data are invented for teaching purposes.

| File | Purpose |
|---|---|
| `TTM_ISSB_Sustainability_Working_Papers_FY2025.xlsx` | The deliverable: Dashboard, Index and 23 cross-referenced working papers |
| `build_workpapers.py`, `wp_lib.py`, `wp_part1-3.py` | Python (xlsxwriter) generator that rebuilds the workbook |

## Working papers

| Ref | Content | Main ISSB references |
|---|---|---|
| A1-A4 | Basis of preparation, value chain, SRO register, materiality | S1.17-24, 30, 54-73, B1-B38, E3-E6; S2.10-12, C3-C5 |
| B1 | Governance | S1.26-27; S2.5-7 |
| C1-C3 | Strategy, transition plan, financial effects, scenario analysis | S1.28-42; S2.8-23, B1-B18 |
| D1 | Risk management | S1.43-44; S2.24-26 |
| E1-E9 | Activity data, emission factors, Scope 1/2/3, GHG summary, cross-industry and industry metrics, targets | S1.45-53; S2.27-37, B19-B71 |
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
* Jurisdictional adoption of ISSB Standards in Cambodia is treated as an assumption to confirm.
