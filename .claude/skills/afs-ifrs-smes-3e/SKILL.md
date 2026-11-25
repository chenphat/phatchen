---
name: afs-ifrs-smes-3e
description: Compile annual financial statements (AFS) under the IFRS for SMEs Accounting Standard, third edition (2025), for an entity whose functional currency is USD and whose presentation currency is KHR (Khmer riel) — from general ledger and trial balance through adjusting entries, mapping, materiality, accounting-policy selection, translation to KHR, the SoFP, SoCI, SoCE, SoCF, notes, the report of the Board of Directors and a disclosure checklist, normally as a live Excel workbook. Use this skill whenever someone asks to prepare, draft, compile, review or teach financial statements under IFRS for SMEs (CIFRS for SMEs, "3E", "third edition"), mentions a Cambodian company reporting in riel with a dollar ledger, asks how to translate USD statements into KHR, needs SME notes or a disclosure checklist, or wants AFS working papers — even if they only say "do the year-end accounts" or "turn this TB into financial statements".
---

# AFS compilation — IFRS for SMEs Accounting Standard, third edition (USD → KHR)

This skill turns an SME's books into a complete, internally consistent set of annual financial statements and the
working papers that support them. The ledger stays in **USD (functional currency)**; the statements are **presented in
KHR** under Section 30.17–30.19. A bundled engine builds the workbook; your job is the judgement, the narrative and
the review — the parts a script cannot do.

Two things make this harder than it looks, so keep them in view throughout:
- **Two currency steps, not one.** KHR balances inside a USD ledger are remeasured to USD with differences in profit or
  loss (30.10). Only then is the whole USD set translated into KHR, with the difference in OCI and a translation
  reserve that is never recycled (30.18(c)). Mixing these up is the most common error. Read
  `references/usd-khr-translation.md` before translating anything.
- **The third edition changed a lot.** Revenue (new five-step Section 23), financial instruments (merged Section 11),
  fair value (new Section 12), the conceptual framework (Section 2), materiality wording, supplier finance and the
  financing-liabilities reconciliation (Section 7). It is effective for periods beginning on or after 1 January 2027;
  **applying it to a 2026 year is early application and must be disclosed (A1)**. See `references/third-edition-changes.md`.

## Inputs to ask for (only what is missing)

1. Trial balance for the current and prior year (pre-closing), and the general ledger for both years if available.
2. Entity facts: name, legal form, registered office, registration number, principal activities, directors,
   auditors, date of authorisation, shares authorised/issued, par value, average employees.
3. Exchange rates KHR per USD: closing CY, average CY, closing PY, average PY, closing at the start of the prior year,
   and historical rates for share capital. Ask for the source (e.g. National Bank of Cambodia official rate) — never
   invent rates; if the user has none, use clearly labelled placeholders and say so.
4. Items a ledger cannot show: receivables ageing, loan contractual cash flows, lease commitments, related parties,
   contingencies, events after the reporting date, tax computation details.
5. Whether this is the first year under IFRS for SMEs (Section 35) or a move from the 2015 edition (Appendix A).

If the user has no data and wants an example or teaching material, generate the sample set with the engine.

**Pick the tool by the input you have:**
- Trial balance and ledger → `scripts/build_afs.py` (full workbook: working papers, statements, notes, directors'
  report, checks). See `references/workbook-engine.md`.
- Statement-level USD figures only (e.g. "here is our USD SoFP and P&L, put it into riel") →
  `scripts/translate_summary.py`: live KHR SoFP, SoCI, SoCE and the 30.20 proof from a small typed CSV and a rates
  file (usage in the script header). It refuses inputs whose USD equity does not roll forward, which catches
  transcription errors early.

## Workflow

Work through these phases in order. Each phase lists what to produce and where the detail lives.

### 1. Scope and acceptance
Confirm the entity is eligible (no public accountability, Section 1), the reporting period, the applicable edition,
transition status and the deliverables. Transition has two different dates, and mixing them up is common: a
**first-time adopter** (Section 35) restates from the start of the comparative year (date of transition, 35.6); an
entity **moving from the 2015 edition** applies the third edition from the start of the current year (date of initial
application, A2) with the A3 relief from current-period line-item disclosures. Both are covered in
`references/sections/s35-transition.md` and `references/third-edition-changes.md`. Confirm the deliverables (workbook, directors' report, Word/PDF statements). Note the preparer's
competence needs if the user is teaching — `references/competence-ies.md` maps each phase to IES 2/3/4 outcomes.

### 2. Read, clean and verify the data
Run the engine (`references/workbook-engine.md`). It finds headers, converts text numbers and dates, removes subtotal
and duplicate lines, and logs every change in the DQ_Log sheet. Then read its findings: TB and GL balance, every
journal balances, GL movement agrees to TB by account, opening balances agree to prior-year closing, cut-off,
abnormal balances, weekend and round-sum journals. Resolve or report every FAIL before going further — statements
built on an unreconciled TB are worthless.

### 3. Map accounts to statement components
Review the `<out>_mapping.csv` the engine writes. Rows marked "Fallback - review" need a decision. Map expenses to a
function (cost of sales, distribution, administrative, other) **and** a nature (for the expense-by-nature note).
Feed the corrected file back with `--mapping`.

### 4. Adjusting entries → adjusted trial balance
Identify year-end adjustments the books lack: accruals, prepayments, depreciation, impairment of receivables
(incurred-loss model, 11.21), inventory write-downs, provisions, current and deferred tax, reclassification of
current portions of loans, **remeasurement of KHR monetary items at the closing rate (30.9–30.10)** — but not of
advances paid or received in KHR, which are non-monetary and keep their original rate; the related revenue or expense
later uses that same rate (30.8A, new in the third edition). An allowance for receivables needs objective evidence
of impairment (incurred-loss model, 11.21–11.22), so tie it to the age analysis rather than a flat percentage. Put them in
`adjustments` in the config (or the AJE sheet): TB + AJE = ATB. Use the section guides for the journals.

### 5. Materiality and accounting policies
Set overall, performance and clearly-trivial materiality from a documented benchmark, and assess each line by size and
by nature (3.15A–3.16; `references/presentation-and-disclosure.md` §7). On the Policies sheet, select the policy wherever
the Standard gives a choice (cost or revaluation model by PPE class, FIFO or weighted average, actuarial gains in P&L
or OCI, investment property/biological assets at fair value or cost when undue cost or effort applies, presentation
choices). Disclose only **material** accounting policy information (8.5), but always the translation policy.

### 6. Draft the statements in USD, then translate to KHR
The engine drafts all four statements in USD and translates them. Check:
- SoFP balances in USD and in KHR; translation reserve present as a separate equity component.
- SoCI: translation difference in OCI under "will not be reclassified"; revaluation and its tax there too.
- SoCE has a translation-reserve column and agrees to the SoFP by component in both currencies.
- SoCF shows "effect of exchange rate changes on cash" separately (7.13) and closes to cash in the SoFP.
- A statement of income and retained earnings (3.18) is normally not available: the KHR translation difference is
  OCI, so a SoCI and a SoCE are needed.
- The proof by rates (30.20) agrees to the OCI translation line.

### 7. Notes
Follow the order in 8.4 and the IASB illustrative structure (`references/illustrative-fs-structure.md`):
general information; basis of preparation including **compliance statement, early application of the third
edition, functional and presentation currency with the reason and the rates**; material accounting policy
information; judgements and estimation uncertainty; then supporting notes in statement order; then other
disclosures (commitments, contingencies, related parties, events after the reporting period, approval). Pull the
section-specific requirements from the guides below. Write entity-specific narrative from facts the user gives —
leave a clearly marked placeholder rather than inventing facts.

### 8. Report of the Board of Directors
Required by local law, not by the Standard. Use the structure in `references/presentation-and-disclosure.md` §10;
every figure must agree to the statements (both currencies if shown).

### 9. Review and sign-off
Run the disclosure checklist (54 rows by paragraph), the live Checks sheet, the two-year analysis (flag movements above
performance materiality), and the Findings tracker. Report honestly what remains open.

## Section guides — read the one you need, when you need it

| When you are dealing with… | Read |
|---|---|
| Elements, recognition, measurement bases, materiality, undue cost or effort | `references/sections/s02-concepts.md` |
| Receivables, payables, loans, bank balances, investments, impairment, hedging | `references/sections/s11-financial-instruments.md` |
| Any fair value measurement (revaluation, investment property, biological assets, Part II instruments) | `references/sections/s12-fair-value.md` |
| PPE: cost/revaluation, components, depreciation, roll-forward | `references/sections/s17-ppe.md` |
| Revenue: five steps, contract balances, contract costs, disaggregation | `references/sections/s23-revenue.md` |
| Salaries, bonuses, severance/seniority, pensions | `references/sections/s28-employee-benefits.md` |
| Current and deferred tax, tax reconciliation | `references/sections/s29-income-tax.md` |
| Functional currency, FX transactions, translation, hyperinflation | `references/sections/s30-foreign-currency.md` and `references/usd-khr-translation.md` |
| Agriculture, extractive activities, service concessions | `references/sections/s34-specialised.md` |
| First-time adoption; moving from the 2015 edition | `references/sections/s35-transition.md`, `references/third-edition-changes.md` |
| Statement formats, note order, wording patterns | `references/illustrative-fs-structure.md` |
| Sections 3–8 requirements, disclosure checklist, directors' report | `references/presentation-and-disclosure.md` |
| Teaching, competence mapping, assessment design | `references/competence-ies.md` |

For a section without a guide (e.g. 13 Inventories, 20 Leases, 21 Provisions, 27 Impairment), use the disclosure
checklist rows and quote the Standard's paragraph numbers; if the user has supplied the Standard, read the paragraph.

## Standards of work

- **Every number traceable.** Statement → note → lead schedule → ATB → TB/AJE → GL. Prefer formulas to typed values
  so the set recalculates when an adjustment changes.
- **Cite paragraphs** of the Standard for every policy, judgement and disclosure you write.
- **Do not invent** facts, rates, related parties or events. Placeholders in square brackets are better than plausible fiction.
- **Comparatives** for every amount (3.14); prior-year KHR figures use prior-year rates (30.18), not this year's.
- **Rounding**: present KHR'000; per-share amounts in single currency units; let rounding differences fall in one line.
- **Be candid** in the final message: what was verified, what is placeholder, what judgement the user must confirm.

## Output

Default deliverable: the Excel workbook from the engine plus a short summary to the user covering — the checks
result; adjustments posted; materiality and key policy choices; translation rates used and their source; open
items (placeholders, findings, judgements). For the finished report (cover, contents, directors' report, the four
statements and all notes in KHR'000), run `scripts/export_report.py <workbook> --out <file>.docx [--banner "DRAFT"]`
on the recalculated workbook; it carries placeholders across highlighted in yellow and drops policies judged not
material on the Policies sheet. Convert to PDF with LibreOffice if needed, and look at the pages before sending.
