---
name: cambodia-aml-cft-accounting
description: Build or update an AML/CFT/CPF Business Risk Assessment (BRA) Excel toolkit for accounting, audit and tax firms in Cambodia (reporting entities under the 2020 Law on AML/CFT). Use when asked for an AML/CFT risk assessment, enterprise-wide risk assessment, customer risk scoring, TF/PF/sanctions risk component, risk appetite or AML teaching workbook for a Cambodian accounting firm or DNFBP.
---

# Cambodia AML/CFT BRA toolkit for accounting firms

Builds a formula-driven Excel workbook that follows the six-step BRA method (from the
ADGM BRA guide and FATF's 2019 RBA Guidance for the Accounting Profession), re-based on
the Cambodian framework: the 2020 Law on AML/CFT, the 2020 law on countering
proliferation financing, CAFIU (STR/CTR, guidelines), ACAR/KICPAA and the National Risk Assessment.

## Generate

```bash
python .claude/skills/cambodia-aml-cft-accounting/scripts/build_workbook.py OUTPUT.xlsx --firm "ABC Accounting Co., Ltd."
python <xlsx-skill>/scripts/recalc.py OUTPUT.xlsx 60   # must report status: success, 0 errors
```

## Workbook map (15 sheets)

| Sheet | Step | Purpose |
|---|---|---|
| Cover | - | Firm details, approval/version, live headline results, legal basis (with "verify" notes) |
| Scales | 1 | Likelihood x impact 4x4 matrix, control rule (Effective = down one level; residual never above inherent) |
| BRA_Register | 2-4 | 24 Cambodia-specific risks across 8 factors + 10 spare rows; inherent -> residual -> appetite -> decision |
| TFS_TF_PF | 5 | Separate TF, PF and sanctions-evasion component with its own rating + TFS screening checklist |
| Risk_Appetite | - | Max residual per factor (feeds register), ML/TF/PF/TFS statements, "will not accept", KRIs with RAG |
| Heat_Map | - | One-page management pack: heat map, counts, top-5 residual risks, approval block |
| Control_Tests | 3 | Design + sample testing -> proposed effectiveness rating |
| CRA_Clients | 6 | Client-level weighted scoring, PEP/sanctions overrides, CDD level, review dates |
| Action_Plan | 4 | Actions for Medium-High/High residual and out-of-appetite risks; overdue tracking |
| Case_XingTianDy | - | Fictional OFAC-linked client case: KYC, 50% rule look-through, screening log, decision log, discussion Q&A |
| Case_PPHotel | - | Fictional hotel client: PEP-family BO, revenue reasonableness test, accept with EDD |
| Traceability | 6 | Risk -> CRA -> policy -> monitoring -> training -> MI |
| Red_Flags | - | 25 Cambodia accounting/audit red flags with responses |
| Triggers_Versions | - | Version log and trigger register |
| Evidence_Checklist | - | Supervisor evidence file, 30-day build plan, common mistakes |

## Rules when customising

- Rate inherent risk before controls; each rating needs a firm-specific reason (column N).
- Rate a control "Effective" only from tested evidence (Control_Tests), never from policy existence.
- Keep TF/PF/TFS as a separate, visible component.
- Inputs: blue text on yellow; formulas black; cross-sheet links green. Never hardcode results.
- Cambodian legal details (article numbers, STR timing, cash threshold, current supervisor of accountants,
  whether the 2010 CAFIU DNFBP Prakas is still in force) are flagged "verify". Do not state them as fact
  without checking the official Khmer text or cafiu.gov.kh.
- Pre-filled risks/ratings are illustrative teaching defaults. Tell the user to replace them with firm data.
