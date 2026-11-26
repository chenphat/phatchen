# phatchen: ISSB (IFRS S1 and S2) reporting skills

These are Claude skills and Excel working-paper templates for preparing, reviewing and assuring
sustainability-related financial disclosures under **IFRS S1** and **IFRS S2**. They include a layer for
Cambodia: SERC Prakas 034, ACAR, NBC, CIFRS, EDC grid factor and local data sources.

| Path | Contents |
|---|---|
| `analysis/sustainability-reporting-requirements-analysis.md` | Review and analysis of the requirements: architecture, boundary, materiality, foundations, climate, the ACLEDA 2025 gap analysis, cross-document findings and open items |
| `skills/ifrs-s1-reporting/` | IFRS S1 skill: an 8-phase process, a paragraph-level checklist, the materiality procedure, ICSR and assurance guidance, Cambodian context, and a builder for working papers S1-00 to S1-14 |
| `skills/ifrs-s2-reporting/` | IFRS S2 skill: a 10-phase climate process, a checklist including the December 2025 amendments, GHG methodology, scenario analysis (SA1–SA9), PCAF financed emissions, and a builder for working papers CR-00 to CR-07 (incl. the Cambodia carbon pricing/ETS scenario model and the carbon credit register), SA, GHG-00 to GHG-09 and FE-01 |
| `examples/` | Generated workbooks: an illustrative Mekong Garment case (inventory totals 3,733.1 tCO2e; Case F gives 70.5 t) and blank templates |
| `dist/*.skill` | Packaged skills, ready to install |

## Regenerate the workbooks

```bash
python skills/ifrs-s1-reporting/scripts/build_s1_workpapers.py --entity "Entity" --year-end 2026-12-31 --currency "KHR million" --out S1.xlsx
python skills/ifrs-s2-reporting/scripts/build_s2_workpapers.py --entity "Entity" --year-end 2026-12-31 --currency "KHR million" --bank --out S2.xlsx
```

Add `--no-examples` to get blank templates. After generating a workbook, open it in Excel or
LibreOffice so the formulas calculate.

The emission factors in the examples are illustrative (UNFCCC calculator, DEFRA 2022, IFI 2021 grid
factor 0.5878). Replace them with current published or supplier-specific values in the `GHG-02`
register before using the workbooks for a client.
