---
name: ifrs-s1-reporting
description: Process, procedures and Excel working-paper templates (S1-00 to S1-14) for preparing, reviewing or assuring IFRS S1 (ISSB) sustainability-related financial disclosures - gap analysis, reporting boundary, identifying sustainability-related risks and opportunities (SRROs), materiality, governance, strategy and financial effects, risk management, metrics and targets, connectivity to the financial statements, judgements, reliefs, compliance statement, ICSR controls and ISSA 5000 assurance readiness, with a Cambodia layer (SERC Prakas 034, ACAR, NBC, CIFRS). Use whenever the user mentions IFRS S1, ISSB, SRROs, investor or enterprise-value materiality, an ISSB-aligned ESG or sustainability report, a disclosure checklist, sustainability working papers, ICSR or assurance readiness, even without naming IFRS S1. For climate content (GHG, scenarios, transition plans) also use ifrs-s2-reporting.
---

# IFRS S1 reporting — process, procedures and working papers

IFRS S1 sets the *general requirements*: what an entity must disclose about **all** sustainability-related
risks and opportunities (SRROs) that could reasonably be expected to affect its cash flows, access to
finance or cost of capital over the short, medium or long term (its "prospects", S1.3). IFRS S2 is the
climate application of the same architecture. Treat this skill as the engagement backbone; call
`ifrs-s2-reporting` for the climate content.

The approach mirrors a financial-statement close: plan → identify → measure → disclose → review →
assure, with a numbered working paper (WP) behind every disclosed figure. The reason is practical:
sustainability information is estimate-heavy and value-chain-dependent, so it only becomes
decision-useful (and assurable) when each figure can be traced from source document to disclosure.

## How to use this skill

1. **Establish the engagement context first** (ask only for what you cannot infer):
   - Entity, industry (SASB/SICS industry), jurisdiction, financial reporting framework (IFRS, CIFRS,
     CIFRS for SMEs, local GAAP — S1.8 lets ISSB apply under any GAAP), reporting period end.
   - Role: preparer, reviewer/internal audit, external assurer, lecturer/student.
   - First year of application? (determines reliefs, S1.E3–E6 and S2.C3–C5).
   - Mandatory or voluntary? Full compliance or partial ("voluntarily applying" — then it must not
     claim compliance, S1.72).
2. **Pick the output the user needs** — a plan, a gap analysis, a WP pack, a draft disclosure note,
   a review of someone else's report, or teaching material — and follow the matching phase(s) below.
3. **Generate the working-paper workbook** when templates are wanted:
   ```bash
   python scripts/build_s1_workpapers.py --entity "Mekong Garment Co." --year-end 2026-12-31 \
       --currency "KHR million" --out S1_Working_Papers.xlsx
   ```
   Then recalculate formulas (xlsx skill `scripts/recalc.py`) before sharing. The workbook contains
   S1-00 to S1-14 (index below), each with a legend, one example row and sign-off blocks.
4. **Cite paragraphs precisely.** Use the paragraph map in `references/s1-disclosure-checklist.md`.
   When unsure of a sub-paragraph, cite the range (e.g. "S1.34–40") rather than guess — a wrong
   citation leads a preparer to apply the wrong requirement.

## Reference files (read when needed)

| File | Read it when |
|---|---|
| `references/s1-disclosure-checklist.md` | Building or reviewing a disclosure; paragraph-level checklist of every S1 requirement incl. reliefs |
| `references/materiality-procedure.md` | Identifying SRROs and deciding what information is material (S1.17–19, B1–B37); scoring scales; nature (TNFD LEAP) |
| `references/controls-and-assurance.md` | Designing ICSR, maker-checker-approver, testing, deficiency classification, ISSA 5000 readiness, maturity scorecard |
| `references/cambodia-context.md` | Entity is Cambodian or the user teaches/works in Cambodia — regulators, SERC Prakas 034 31-item mapping, data sources, open items |

## The process (eight phases)

Each phase lists the objective, the procedures, the IFRS anchor and the WP produced. Phases 3–6 are
iterative: a new risk found in strategy work goes back into the SRRO register.

### Phase 1 — Readiness, scoping and planning  → S1-00, S1-01
- Agree objective, scope (all SRROs vs climate-first relief, S1.E5), timetable aligned to the
  financial close (S1.64: same time as the financial statements, relief S1.E4 in year 1), owners
  (RACI) and assurance intentions with the CFO, audit committee and assurer.
- Gap analysis: current reporting (GRI, SDG mapping, local regulator template, CDP) vs S1/S2
  requirements. Record each gap with owner, action and date. For Cambodian listed entities also
  map to SERC Prakas 034's 31 items (see `cambodia-context.md`).
- Inventory existing data sources and systems (ERP, payroll, meters, loan system, spreadsheets).

### Phase 2 — Reporting entity, boundary and period  → S1-02
- Reporting entity = same as the financial statements (S1.20, B38): 100% of consolidated subsidiaries.
  Joint operations, leased assets, associates/JVs: follow the general principle and document the
  judgement (S1 has no specific guidance; SRG 3.2 practice).
- Value chain = own operations + upstream + downstream, including relationships, resources and
  financing (S1 Appendix A, B2–B5). Investments under fiduciary duty are **not** excluded under ISSB.
- Period = financial-statement period, including 52/53-week years and changes in period (S1.65–66).
- Time horizons: entity-specific short/medium/long, linked to strategic planning and capital
  allocation (S1.30(b)–(c)); document if different horizons are used for different SRROs.
- Acquisitions/disposals: include from/to the date consolidated, aligned to financial reporting
  (TIG June 2024); disaggregate if needed so information is not obscured (S1.B26–B27).

### Phase 3 — Identify SRROs and determine material information  → S1-03, S1-04
Follow `references/materiality-procedure.md`. In outline:
1. Understand the business model, value chain, dependencies and impacts on resources and
   relationships (S1.B2–B5).
2. Long-list using the required sources: SASB disclosure topics for the industry ("shall refer to and
   consider", S1.55(a)); may also use CDSB guidance, other investor-focused standards, peer practice
   (S1.55(b), 58). GRI/ESRS may help identify *information* but not *SRROs* (S1.BC136–137).
3. Assess each SRRO for likelihood and magnitude of effect on cash flows, access to finance and cost
   of capital, by time horizon, using all reasonable and supportable information available without
   undue cost or effort (S1.B6–B10). Consider completed mitigation, not planned actions.
4. Determine the **material information** about each SRRO: would omitting, misstating or obscuring it
   reasonably be expected to influence primary users' decisions (S1.18, B13–B37)? Consider
   qualitative and quantitative factors and aggregate exposure (S1.B23).
5. Re-assess on significant events or changes (S1.B11–B12); otherwise at least annually.
Keep financial materiality (ISSB) and impact materiality (GRI/ESRS) in separate columns — mixing them is
the most common error. Impacts matter under S1 only to the extent they give rise to SRROs.

### Phase 4 — Governance  → S1-05
Disclose the governance processes, controls and procedures used to monitor, manage and oversee SRROs
(S1.26–27): board/committee responsibilities and terms of reference, skills and competencies, how and
how often informed, how SRROs are considered in strategy, major transactions and risk management,
oversight of targets and remuneration links (27(a)); management's role, delegated positions, controls
and procedures (27(b)). Evidence: board minutes, charters, RACI, remuneration policy. Avoid
duplicating governance across topics — integrated disclosure is allowed (S1.B42(b)).

### Phase 5 — Strategy  → S1-06
- SRROs and time horizons (S1.30); effects on business model and value chain and where concentrated
  (S1.32); response in strategy and decision-making, progress against plans, trade-offs (S1.33).
- **Financial effects** (S1.34–40): current effects on financial position, performance and cash flows;
  SRROs with significant risk of material adjustment to carrying amounts within the next annual
  period; anticipated effects over time. Quantitative unless effects are not separately identifiable,
  measurement uncertainty is too high, or (anticipated effects) the entity lacks skills, capabilities
  or resources — then explain why and give qualitative information (S1.37–40).
- **Resilience** of strategy and business model (S1.41–42).
- Plain-language aid: the five zones (why/where/what/who/when) + resilience test — see S1-06.

### Phase 6 — Risk management  → S1-07
Processes and policies to identify, assess, prioritise and monitor SRROs — inputs and data sources,
whether scenario analysis informs identification, how nature/likelihood/magnitude are assessed,
prioritisation relative to other risks, monitoring, changes from prior period (S1.44(a)); processes for
opportunities (44(b)); integration into overall risk management/ERM (44(c)). Include greenwashing and
fraud risk controls (COSO/ACFE) — this is where reviewers look for them.

### Phase 7 — Metrics and targets  → S1-08 (KPI cards and target register)
- Metrics required by an applicable IFRS SDS; metrics the entity uses to measure and monitor each SRRO
  and performance, including industry metrics from SASB (S1.45–50). For entity-developed metrics,
  disclose definition, method, whether validated externally, and limitations (S1.50).
- Targets: metric, objective, scope, period, base period, milestones/interim targets, performance and
  trends, revisions (S1.51–53).
- Every metric gets a KPI card (nine attributes: definition, unit, period, boundary, owner, source
  system, method and factor, frequency, evidence reference). Redefined/replaced metrics → revise
  comparatives unless impracticable (S1.B52); new metrics → comparatives unless impracticable (B53).

### Phase 8 — Compile, connect, review and sign off  → S1-09 to S1-14
- **Connected information** (S1.21–24, B39–B44): link each SRRO across pillars and to financial
  statement items — impairment (IAS 36), useful lives (IAS 16), provisions (IAS 37), biological assets
  (IAS 41), ECL (IFRS 9), going concern. Data and assumptions must be consistent with the financial
  statements to the extent possible (S1.23); explain significant differences.
- **Judgements** (S1.74–76), **measurement uncertainty** (S1.77–82), **errors** (S1.83–86, B55–B59):
  log in S1-10.
- **Presentation**: location within general purpose financial reports (S1.60–63), cross-references
  (B45–B47), comparatives (S1.70–71, B49–B54), commercially sensitive *opportunity* information
  (S1.B34–B37), information prohibited by law (S1.73), sources of guidance used (S1.59).
- **Compliance statement**: explicit and unreserved only if all requirements are met (S1.72); if
  reliefs are used (climate-first, timing, comparatives), disclose that fact.
- **Review**: complete the disclosure checklist (S1-11), controls matrix (S1-12), evidence index
  (S1-13) and review notes/exception log (S1-14). Joint CFO–sustainability sign-off before release.

## Working-paper index (generated by `scripts/build_s1_workpapers.py`)

| WP | Title | Main IFRS anchor |
|---|---|---|
| S1-00 | Index, engagement details, sign-off and timetable | S1.64, E4 |
| S1-01 | Readiness and gap analysis (incl. SERC Prakas 034 map) | S1.72, E3–E6 |
| S1-02 | Reporting entity, boundary, period and time horizons memo | S1.20, 30, 65–66, B38 |
| S1-03 | SRRO register and assessment (likelihood × magnitude, heat map) | S1.30, B1–B12 |
| S1-04 | Material information assessment | S1.17–19, B13–B37 |
| S1-05 | Governance disclosure working paper | S1.26–27 |
| S1-06 | Strategy: five zones, financial effects, resilience | S1.28–42 |
| S1-07 | Risk management process documentation | S1.43–44 |
| S1-08 | Metrics (KPI cards) and targets register | S1.45–53 |
| S1-09 | Connectivity to the financial statements | S1.21–24, B39–B44 |
| S1-10 | Judgements, measurement uncertainty and errors log | S1.74–86 |
| S1-11 | IFRS S1 disclosure checklist with status | whole standard |
| S1-12 | ICSR controls matrix (maker-checker-approver, COSO) | S1.77–82; COSO ICSR |
| S1-13 | Assurance evidence index (figure → WP → evidence) | ISSA 5000 |
| S1-14 | Review notes and exception log; ICSR maturity scorecard | ISA 265 concepts |

Climate WPs (GHG-00 to GHG-09, SA1–SA9, CR-01…) come from `ifrs-s2-reporting`; S1-13 cross-references them.

## Quality rules that apply to every output

- **Investor lens, entity-specific.** Boilerplate ("we take sustainability seriously") is not
  material information. Each SRRO needs a stated response, even "monitor only".
- **Separate ambition from results.** Targets, plans and commitments are labelled as such; reported
  performance is measured, sourced and dated. Run every claim through the anti-greenwashing test:
  metric + baseline + evidence + WP reference.
- **Gross before net; facts before narrative.** Never net credits, offsets or mitigation into a
  reported risk exposure or metric.
- **Estimates are allowed, hidden estimates are not.** State the method, data quality and
  uncertainty (S1.77–82).
- **Proportionality has limits.** "Undue cost or effort" and "commensurate with skills, capabilities
  and resources" (S1.37, 39, B8–B10) relieve the effort, not the obligation; the more useful the
  information, the more effort is expected (S1.BC17).
- **Tools calculate; working papers control.** Calculators and dashboards do not record evidence,
  approvals or changes.
- **Flag currency of law.** Standards and local rules change; state the version relied on (IFRS S1/S2
  June 2023; S2 amendments December 2025, effective 1 January 2027) and list jurisdictional
  assumptions as open items to confirm.

## Output formats

- **Draft disclosure note**: headings by pillar (Governance, Strategy, Risk management, Metrics and
  targets), each paragraph tagged with its IFRS reference and WP reference in a review copy, e.g.
  "[S1.33(a); S1-06 row 3]". Remove tags in the publication copy.
- **Gap analysis / review report**: table with Requirement | Para | Current disclosure | Gap |
  Severity (High/Med/Low) | Action | Owner | Date — then a short summary of the top five gaps.
- **Teaching material**: link each activity to paragraph, WP and a worked Cambodian example; include
  the common error to avoid.
