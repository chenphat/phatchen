# Content JSON schema

One JSON file drives all three outputs. `assets/apm_content.json` is the complete example.

## meta (object, required)
| key | used in | notes |
|---|---|---|
| title, subtitle | all | Manual title and subtitle |
| org | all | Firm or organisation name. Keep a `[Firm Name]` placeholder if unknown |
| doc_ref | all, file names | e.g. `APPM-2026-01`. Used as the output filename prefix |
| version, effective_date | all | Shown on the cover and in footers |
| owner, approved_by, review_cycle | docx, xlsx | Document control |
| source | all | Title of the reviewed source document |
| framework | docx | Standards the manual relies on |

## analysis (object, required)
- `summary` (string): 4-6 sentences: what the source is, its status, and why firm policy is needed.
- `structure` (list of [part, content]): map of the source.
- `strengths` (list of strings).
- `gaps` (list of {area, finding, risk: High|Medium|Low, recommendation}). The recommendation should name the policy (POL-xx) or appendix that closes the gap.

## roles (list of [code, name])
Codes must match the `Resp.` values used in procedures, so the RACI COUNTIFS formulas can find them. A procedure may name several roles, e.g. `"MGR / EP"`.

## policies (list, required)
```json
{
  "id": "POL-02", "title": "...", "phase": "Acceptance", "isa": "ISA 210; ISQM 1",
  "purpose": "...", "scope": "...",
  "statements": ["... shall ..."],
  "procedures": [["Action", "RESP", "WP ref", "Timing"]],
  "documentation": ["..."], "red_flags": ["..."]
}
```
Statements are numbered automatically as `<last two digits of id>.<n>`.

## Reference tables (optional; audit-specific)
| key | shape | used in |
|---|---|---|
| working_papers | [ref, title, phase, standard] | docx App. A, xlsx WP_Index |
| area_procedures | [section, area, risks, tests] | docx App. B |
| materiality_benchmarks | [name, low%, high%, use] (decimals) | docx, xlsx Materiality, pptx |
| transaction_samples | [label, from, to, high, medium, low] | docx, xlsx Sample_Size, pptx |
| controls_samples | [frequency, size, note] | docx, xlsx, pptx |
| substantive_example | {population, overall_materiality, performance_materiality, large_items, key_items, risk_factor, expected_sample} | docx, xlsx, pptx |
| material_by_nature, going_concern, continuance_triggers, inherent_risk_factors, glossary | [a, b] | docx appendices, xlsx, pptx |
| opinion_grid | [matter, material, pervasive] | docx, pptx |
| roadmap | [when, action] | docx App. F, xlsx Implementation, pptx |

## Adapting for a non-audit manual (e.g. procurement, treasury, HR, IT)
1. Keep `meta`, `analysis`, `roles`, `policies`, `glossary` and `roadmap`.
2. In `build_xlsx.py`, delete the Acceptance, Materiality, Sample_Size, Risk_Register and Misstatements blocks
   (each starts with a `# n Name` comment) and their rows on the Cover sheet index. Replace them with domain
   calculators if useful, such as an approval-limits matrix or a reconciliation checklist.
3. In `build_docx.py`, drop Appendices B to E. In `build_pptx.py`, drop the slides headed POL-05 to POL-14 and write
   domain slides with the same `d.slide()`, `d.cards()`, `d.table()` and `d.bullets()` helpers.
4. Re-run `build_all.py`. The formula check still guards whatever sheets remain.
