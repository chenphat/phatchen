# IFRS for SMEs Accounting Standard (third edition) — compliance checklist

> Source: the text of the Standard (February 2025), extracted by `scripts/extract_checklist.py` into `scripts/afs_smes/data/ifrs_smes_3e_checklist.json`. Requirements below are abridged; the workbook carries the full text.

## Contents
1. What the checklist is and how to complete it
2. Applicability by section
3. Third-edition items to watch
4. Items by section

## 1. What the checklist is and how to complete it

- 813 items: 308 disclosure, 282 recognition and measurement, 158 presentation, 54 transition, 11 scope. One row per lettered sub-paragraph of a disclosure or presentation requirement (e.g. 17.31(a)–(e)); one row per paragraph for recognition and measurement requirements.
- **Where it lives:** every AFS workbook built by `build_afs.py` has `Compliance_Summary` (one row per section, with the applicability question) and `Compliance` (every item). `compliance_checklist.py` makes a standalone copy.
- **Step 1 — applicability.** Answer each section's question on Compliance_Summary. The engine pre-answers from the ledger (e.g. PPE balances → Section 17 Yes; no associates → Section 14 No); cells showing `?` need a human answer. Confirm every pre-answer — a ledger cannot show a lease that is expensed under another name, a contingent liability, or a government grant netted off costs.
- **Step 2 — status per item.** `Yes` complied/disclosed; `Partly`; `No` (explain and raise a finding); `N/A` where the requirement does not arise in the entity's circumstances (e.g. 4.14 with no binding sale agreement). Fill *Where addressed* (note or statement); it is pre-filled with the note numbers usually affected.
- **Step 3 — sign-off.** The reviewer initials each applicable item. `% complete` on the summary counts Yes and N/A as answered and Partly as half; the Dashboard tile shows the overall figure.
- **Judgement, not ticking.** Materiality still applies (3.16): a disclosure that is immaterial need not be given, but record that conclusion as N/A with the reason, so the review trail shows it was considered.
- **Never mark Yes on the strength of the template alone.** The workbook drafts most notes, but entity facts (related parties, commitments, contingencies, judgements, valuation details) must be present before an item is Yes.

## 2. Applicability by section

| Section | Title | Items |
|---|---|---|
| 2 | Concepts and Pervasive Principles | 1 |
| 3 | Financial Statement Presentation | 45 |
| 4 | Statement of Financial Position | 49 |
| 5 | Statement of Comprehensive Income and Income Statement | 22 |
| 6 | Statement of Changes in Equity and Statement of Income and Retained Earnings | 13 |
| 7 | Statement of Cash Flows | 54 |
| 8 | Notes to the Financial Statements | 15 |
| 9 | Consolidated and Separate Financial Statements | 19 |
| 10 | Accounting Policies, Estimates and Errors | 19 |
| 11 | Financial Instruments | 59 |
| 12 | Fair Value Measurement | 17 |
| 13 | Inventories | 11 |
| 14 | Investments in Associates | 14 |
| 15 | Joint Arrangements | 19 |
| 16 | Investment Property | 7 |
| 17 | Property, Plant and Equipment | 27 |
| 18 | Intangible Assets other than Goodwill | 21 |
| 19 | Business Combinations and Goodwill | 40 |
| 20 | Leases | 21 |
| 21 | Provisions and Contingencies | 18 |
| 22 | Liabilities and Equity | 14 |
| 23 | Revenue from Contracts with Customers | 73 |
| 24 | Government Grants | 5 |
| 25 | Borrowing Costs | 1 |
| 26 | Share-based Payment | 23 |
| 27 | Impairment of Assets | 13 |
| 28 | Employee Benefits | 38 |
| 29 | Income Tax | 31 |
| 30 | Foreign Currency Translation | 19 |
| 31 | Hyperinflation | 3 |
| 32 | Events after the End of the Reporting Period | 4 |
| 33 | Related Party Disclosures | 9 |
| 34 | Specialised Activities | 19 |
| 35 | Transition to the IFRS for SMEs Accounting Standard | 18 |
| A | Appendix A - Effective date and transition | 52 |

## 3. Third-edition items to watch

| Paragraphs | What to check |
|---|---|
| 2.30–2.31 | Undue cost or effort: re-assess each period; disclose its use |
| 3.15A–3.16 | Materiality: 'obscuring' wording; aggregation |
| 7.19A | Reconciliation of liabilities arising from financing activities |
| 7.19C | Supplier finance arrangements |
| 8.5 | Material accounting policy information (not 'significant policies') |
| 11.43, 11.43A | Age analysis of receivables; maturity analysis of financial liabilities |
| 12.28–12.31 | Fair value hierarchy, techniques and inputs; Level 3 gains and losses |
| 23.82–23.90 | Revenue: disaggregation, contract balances, promises, judgements |
| 29.38–29.40 | Income tax disclosures incl. Pillar Two |
| 30.5A, 30.8A | Exchangeability; advance consideration |
| A1–A3 | Effective date, early application disclosure, transition relief |

## 4. Items by section

### Section 2 — Concepts and Pervasive Principles

| Ref | Area | Requirement (abridged) |
|---|---|---|
| 2.31 | Disclosure | If an entity applies an undue cost or effort exemption, the entity shall disclose that fact and the reasons why applying the requirement would involve undue cost or effort. This requirement does not apply to the undue cost or … |

### Section 3 — Financial Statement Presentation

| Ref | Area | Requirement (abridged) |
|---|---|---|
| 3.1 | Scope | This section explains fair presentation of financial statements, what compliance with the IFRS for SMEs Accounting Standard requires and what comprises a complete set of financial statements. |
| 3.2(a) | Presentation | *Financial statements shall present fairly the financial position, financial performance …* → the application of the IFRS for SMEs Accounting Standard, with additional disclosure when necessary, is presumed to result in financial statements that achieve a fair presentation of the financial position, financial performance … |
| 3.2(b) | Presentation | *Financial statements shall present fairly the financial position, financial performance …* → the application of this Standard by an entity with public accountability does not result in a fair presentation in accordance with this Standard (see paragraph 1.5). The additional disclosures referred to in (a) are necessary … |
| 3.3 | Disclosure | An entity whose financial statements comply with the IFRS for SMEs Accounting Standard shall make an explicit and unreserved statement of such compliance in the notes. Financial statements shall not be described as complying with … |
| 3.4 | Presentation | In the extremely rare circumstances when management concludes that compliance with this Standard would be so misleading that it would conflict with the objective of financial statements of SMEs set out in Section 2, the entity … |
| 3.5(a) | Disclosure | *When an entity departs from a requirement of this Standard in accordance with paragraph …* → that management has concluded that the financial statements present fairly the entity's financial position, financial performance and cash flows; |
| 3.5(b) | Disclosure | *When an entity departs from a requirement of this Standard in accordance with paragraph …* → that it has complied with the IFRS for SMEs Accounting Standard, except that it has departed from a particular requirement to achieve a fair presentation; and |
| 3.5(c) | Disclosure | *When an entity departs from a requirement of this Standard in accordance with paragraph …* → the nature of the departure, including the treatment that the IFRS for SMEs Accounting Standard would require, the reason why that treatment would be so misleading in the circumstances that it would conflict with the objective of … |
| 3.6 | Presentation | When an entity has departed from a requirement of this Standard in a prior period, and that departure affects the amounts recognised in the financial statements for the current period, it shall make the disclosures set out in … |
| 3.7(a) | Presentation | *In the extremely rare circumstances when management concludes that compliance with a …* → the nature of the requirement in this Standard and the reason why management has concluded that complying with that requirement is so misleading in the circumstances that it conflicts with the objective of financial statements … |
| 3.7(b) | Presentation | *In the extremely rare circumstances when management concludes that compliance with a …* → for each period presented, the adjustments to each item in the financial statements that management has concluded would be necessary to achieve a fair presentation. |
| 3.8 | Presentation | When preparing financial statements, the management of an entity using this Standard shall make an assessment of the entity's ability to continue as a going concern. An entity is a going concern unless management either intends … |
| 3.9 | Disclosure | When management is aware, in making its assessment, of material uncertainties related to events or conditions that cast significant doubt upon the entity's ability to continue as a going concern, the entity shall disclose those … |
| 3.10(a) | Disclosure | *An entity shall present a complete set of financial statements (including comparative …* → that fact; |
| 3.10(b) | Disclosure | *An entity shall present a complete set of financial statements (including comparative …* → the reason for using a longer or shorter period; and |
| 3.10(c) | Disclosure | *An entity shall present a complete set of financial statements (including comparative …* → the fact that comparative amounts presented in the financial statements (including the related notes) are not entirely comparable. |
| 3.11(a) | Presentation | *An entity shall retain the presentation and classification of items in the financial …* → it is apparent, following a significant change in the nature of the entity's operations or a review of its financial statements, that another presentation or classification would be more appropriate having regard to the criteria … |
| 3.11(b) | Presentation | *An entity shall retain the presentation and classification of items in the financial …* → this Standard requires a change in presentation. |
| 3.12(a) | Disclosure | *When the presentation or classification of items in the financial statements is changed, …* → the nature of the reclassification; |
| 3.12(b) | Disclosure | *When the presentation or classification of items in the financial statements is changed, …* → the amount of each item or class of items that is reclassified; and |
| 3.12(c) | Disclosure | *When the presentation or classification of items in the financial statements is changed, …* → the reason for the reclassification. |
| 3.13 | Disclosure | If it is impracticable to reclassify comparative amounts, an entity shall disclose why reclassification was not practicable. |
| 3.14 | Disclosure | Except when this Standard permits or requires otherwise, an entity shall disclose comparative information in respect of the previous comparable period for all amounts presented in the current period's financial statements. An … |
| 3.15 | Presentation | An entity shall present separately each material class of similar items. An entity shall present separately items of a dissimilar nature or function unless they are immaterial. |
| 3.15A | Presentation | When applying this Standard an entity shall decide, after taking into consideration all the relevant facts and circumstances, how it aggregates information in the financial statements, including the notes. An entity shall not … |
| 3.16 | Presentation | Information is material if omitting, misstating or obscuring it could reasonably be expected to influence decisions that the primary users of general purpose financial statements make on the basis of those financial statements, … |
| 3.16A | Presentation | An entity shall prepare its financial statements, except for cash flow information, using the accrual basis of accounting. On the accrual basis, items are recognised as assets, liabilities, equity, income or expenses when they … |
| 3.17(a) | Presentation | *A complete set of financial statements of an entity shall include all of the following:* → a statement of financial position as at the reporting date; |
| 3.17(b) | Presentation | *A complete set of financial statements of an entity shall include all of the following:* → either: (i) a single statement of comprehensive income for the reporting period displaying all items of income and expense recognised during the period including those items recognised in determining profit or loss (which is a … |
| 3.17(c) | Presentation | *A complete set of financial statements of an entity shall include all of the following:* → a statement of changes in equity for the reporting period; |
| 3.17(d) | Presentation | *A complete set of financial statements of an entity shall include all of the following:* → a statement of cash flows for the reporting period; and |
| 3.17(e) | Presentation | *A complete set of financial statements of an entity shall include all of the following:* → notes, comprising material accounting policy information and other explanatory information. |
| 3.18 | Presentation | If the only changes to equity during the periods for which financial statements are presented arise from profit or loss, payment of dividends, corrections of prior period errors, and changes in accounting policy, the entity may … |
| 3.19 | Presentation | If an entity has no items of other comprehensive income in any of the periods for which financial statements are presented, it may present only an income statement or it may present a statement of comprehensive income in which … |
| 3.20 | Presentation | Because paragraph 3.14 requires comparative amounts in respect of the previous period for all amounts presented in the financial statements, a complete set of financial statements means that an entity shall present, as a minimum, … |
| 3.21 | Presentation | In a complete set of financial statements, an entity shall present each financial statement with equal prominence. |
| 3.22 | Presentation | An entity may use titles for the financial statements other than those used in this Standard as long as they are not misleading. |
| 3.23(a) | Presentation | *An entity shall clearly identify each of the financial statements and the notes and …* → the name of the reporting entity and any change in its name since the end of the preceding reporting period; |
| 3.23(b) | Presentation | *An entity shall clearly identify each of the financial statements and the notes and …* → whether the financial statements cover the individual entity or a group of entities; |
| 3.23(c) | Presentation | *An entity shall clearly identify each of the financial statements and the notes and …* → the date of the end of the reporting period and the period covered by the financial statements; |
| 3.23(d) | Presentation | *An entity shall clearly identify each of the financial statements and the notes and …* → the presentation currency, as defined in Section 30 Foreign Currency Translation; and |
| 3.23(e) | Presentation | *An entity shall clearly identify each of the financial statements and the notes and …* → the level of rounding, if any, used in presenting amounts in the financial statements. |
| 3.24(a) | Disclosure | *An entity shall disclose the following in the notes:* → the domicile and legal form of the entity, its country of incorporation and the address of its registered office (or principal place of business, if different from the registered office); and |
| 3.24(b) | Disclosure | *An entity shall disclose the following in the notes:* → a description of the nature of the entity's operations and its principal activities. |
| 3.25 | Disclosure | This Standard does not address presentation of segment information, earnings per share, or interim financial reports by a small or medium-sized entity. An entity making such disclosures shall describe the basis for preparing and … |

### Section 4 — Statement of Financial Position

| Ref | Area | Requirement (abridged) |
|---|---|---|
| 4.1 | Scope | This section sets out the information that is to be presented in a statement of financial position and how to present it. The statement of financial position (sometimes called the balance sheet) presents an entity's assets, … |
| 4.2(a) | Presentation | *The statement of financial position shall include line items that present the following …* → cash and cash equivalents; |
| 4.2(b) | Presentation | *The statement of financial position shall include line items that present the following …* → trade and other receivables; |
| 4.2(c) | Presentation | *The statement of financial position shall include line items that present the following …* → financial assets (excluding amounts shown under (a), (b), (j) and (k)); |
| 4.2(d) | Presentation | *The statement of financial position shall include line items that present the following …* → inventories; |
| 4.2(e) | Presentation | *The statement of financial position shall include line items that present the following …* → property, plant and equipment (including bearer plants in the scope of Section 17 Property, Plant and Equipment); |
| 4.2(ea) | Presentation | *The statement of financial position shall include line items that present the following …* → investment property carried at cost less accumulated depreciation and impairment; |
| 4.2(f) | Presentation | *The statement of financial position shall include line items that present the following …* → investment property carried at fair value through profit or loss; |
| 4.2(g) | Presentation | *The statement of financial position shall include line items that present the following …* → intangible assets; |
| 4.2(h) | Presentation | *The statement of financial position shall include line items that present the following …* → biological assets in the scope of Section 34 Specialised Activities carried at cost less accumulated depreciation and impairment; |
| 4.2(i) | Presentation | *The statement of financial position shall include line items that present the following …* → biological assets in the scope of Section 34 carried at fair value through profit or loss; |
| 4.2(j) | Presentation | *The statement of financial position shall include line items that present the following …* → investments in associates; |
| 4.2(k) | Presentation | *The statement of financial position shall include line items that present the following …* → investments in jointly controlled entities; |
| 4.2(l) | Presentation | *The statement of financial position shall include line items that present the following …* → trade and other payables; |
| 4.2(m) | Presentation | *The statement of financial position shall include line items that present the following …* → financial liabilities (excluding amounts shown under (l) and (p)); |
| 4.2(n) | Presentation | *The statement of financial position shall include line items that present the following …* → liabilities and assets for current tax; |
| 4.2(o) | Presentation | *The statement of financial position shall include line items that present the following …* → deferred tax liabilities and deferred tax assets (these shall always be classified as non-current); |
| 4.2(p) | Presentation | *The statement of financial position shall include line items that present the following …* → provisions; |
| 4.2(q) | Presentation | *The statement of financial position shall include line items that present the following …* → non-controlling interest, presented within equity separately from the equity attributable to the owners of the parent; and |
| 4.2(r) | Presentation | *The statement of financial position shall include line items that present the following …* → equity attributable to the owners of the parent. |
| 4.3 | Presentation | An entity shall present additional line items (including by disaggregating the line items listed in paragraph 4.2), headings and subtotals in the statement of financial position when such presentation is relevant to an … |
| 4.4 | Presentation | An entity shall present current and non-current assets, and current and non-current liabilities, as separate classifications in its statement of financial position in accordance with paragraphs 4.5-4.8, except when a presentation … |
| 4.5(a) | Presentation | *An entity shall classify an asset as current when:* → it expects to realise the asset, or intends to sell or consume it, in the entity's normal operating cycle; |
| 4.5(b) | Presentation | *An entity shall classify an asset as current when:* → it holds the asset primarily for the purpose of trading; |
| 4.5(c) | Presentation | *An entity shall classify an asset as current when:* → it expects to realise the asset within twelve months after the reporting date; or |
| 4.5(d) | Presentation | *An entity shall classify an asset as current when:* → the asset is cash or a cash equivalent, unless it is restricted from being exchanged or used to settle a liability for at least twelve months after the reporting date. |
| 4.6 | Presentation | An entity shall classify all other assets as non-current. When the entity's normal operating cycle is not clearly identifiable, its duration is assumed to be twelve months. |
| 4.7(a) | Presentation | *An entity shall classify a liability as current when:* → it expects to settle the liability in the entity's normal operating cycle; |
| 4.7(b) | Presentation | *An entity shall classify a liability as current when:* → it holds the liability primarily for the purpose of trading; |
| 4.7(c) | Presentation | *An entity shall classify a liability as current when:* → the liability is due to be settled within twelve months after the reporting date; or |
| 4.7(d) | Presentation | *An entity shall classify a liability as current when:* → the entity does not have an unconditional right to defer settlement of the liability for at least twelve months after reporting date. |
| 4.8 | Presentation | An entity shall classify all other liabilities as non-current. |
| 4.9(a) | Presentation | *This Standard does not prescribe the sequence or format in which items are to be …* → line items are included when the size, nature or function of an item or aggregation of similar items is such that separate presentation is relevant to an understanding of the entity's financial position; and |
| 4.9(b) | Presentation | *This Standard does not prescribe the sequence or format in which items are to be …* → the descriptions used and the sequencing of items or aggregation of similar items may be amended according to the nature of the entity and its transactions, to provide information that is relevant to an understanding of the … |
| 4.10(a) | Presentation | *The judgement on whether additional items are presented separately is based on an …* → the amounts, nature and liquidity of assets; |
| 4.10(b) | Presentation | *The judgement on whether additional items are presented separately is based on an …* → the function of assets within the entity; and |
| 4.10(c) | Presentation | *The judgement on whether additional items are presented separately is based on an …* → the amounts, nature and timing of liabilities. |
| 4.11(a) | Disclosure | *An entity shall disclose, either in the statement of financial position or in the notes, …* → property, plant and equipment in classifications appropriate to the entity; |
| 4.11(b) | Disclosure | *An entity shall disclose, either in the statement of financial position or in the notes, …* → trade and other receivables showing separately amounts receivable from related parties and receivable from other parties; |
| 4.11(c) | Disclosure | *An entity shall disclose, either in the statement of financial position or in the notes, …* → inventories, showing separately amounts of inventories: (i) held for sale in the ordinary course of business; (ii) in the process of production for such sale; and (iii) in the form of materials or supplies to be consumed in the … |
| 4.11(d) | Disclosure | *An entity shall disclose, either in the statement of financial position or in the notes, …* → trade and other payables, showing separately amounts payable to trade suppliers, payable to related parties, deferred income and accruals; |
| 4.11(e) | Disclosure | *An entity shall disclose, either in the statement of financial position or in the notes, …* → provisions for employee benefits and other provisions; and |
| 4.11(f) | Disclosure | *An entity shall disclose, either in the statement of financial position or in the notes, …* → classes of equity, such as paid-in capital, share premium, retained earnings and items of income and expense that, as permitted or required by this Standard, are recognised in other comprehensive income and presented separately … |
| 4.12(a) | Disclosure | *An entity with share capital shall disclose the following, either in the statement of …* → for each class of share capital: (i) the number of shares authorised. (ii) the number of shares issued and fully paid, and issued but not fully paid. (iii) par value per share or that the shares have no par value. (iv) a … |
| 4.12(b) | Disclosure | *An entity with share capital shall disclose the following, either in the statement of …* → a description of each reserve within equity. |
| 4.13 | Disclosure | An entity without share capital, such as a partnership or trust, shall disclose information equivalent to that required by paragraph 4.12(a), showing changes during the period in each category of equity, and the rights, … |
| 4.14(a) | Disclosure | *If, at the reporting date, an entity has a binding sale agreement for a major disposal of …* → a description of the asset(s) or the group of assets and liabilities; |
| 4.14(b) | Disclosure | *If, at the reporting date, an entity has a binding sale agreement for a major disposal of …* → a description of the facts and circumstances of the sale; and |
| 4.14(c) | Disclosure | *If, at the reporting date, an entity has a binding sale agreement for a major disposal of …* → the carrying amount of the assets or, if the disposal involves a group of assets and liabilities, the carrying amounts of those assets and liabilities. |

### Section 5 — Statement of Comprehensive Income and Income Statement

| Ref | Area | Requirement (abridged) |
|---|---|---|
| 5.1 | Scope | This section requires an entity to present its total comprehensive income for a period-ie its financial performance for the period-in one or two financial statements. It sets out the information that is to be presented in those … |
| 5.2(a) | Presentation | *An entity shall present its total comprehensive income for a period either:* → in a single statement of comprehensive income, in which case the statement of comprehensive income presents all items of income and expense recognised in the period; or |
| 5.2(b) | Presentation | *An entity shall present its total comprehensive income for a period either:* → in two statements-an income statement and a statement of comprehensive income-in which case the income statement presents all items of income and expense recognised in the period except those that are recognised in total … |
| 5.3 | Presentation | A change from the single-statement approach to the two-statement approach, or vice versa, is a change in accounting policy to which Section 10 Accounting Policies, Estimates and Errors applies. Single-statement approach |
| 5.4(a) | Presentation | *Under the single-statement approach, the statement of comprehensive income shall include …* → the effects of corrections of errors and changes in accounting policies are presented as retrospective adjustments of prior periods instead of as part of profit or loss in the period in which they are identified (see Section 10); … |
| 5.4(b) | Presentation | *Under the single-statement approach, the statement of comprehensive income shall include …* → four types of other comprehensive income are recognised as part of total comprehensive income, outside of profit or loss, when they arise: (i) some gains and losses arising on translating the financial statements of a foreign … |
| 5.5(a) | Presentation | *As a minimum, an entity shall include, in the statement of comprehensive income, line …* → revenue. |
| 5.5(b) | Presentation | *As a minimum, an entity shall include, in the statement of comprehensive income, line …* → finance costs. |
| 5.5(c) | Presentation | *As a minimum, an entity shall include, in the statement of comprehensive income, line …* → share of the profit or loss of investments in associates (see Section 14 Investments in Associates) and jointly controlled entities (see Section 15 Joint Arrangements) accounted for using the equity method. |
| 5.5(d) | Presentation | *As a minimum, an entity shall include, in the statement of comprehensive income, line …* → tax expense excluding tax allocated to items (e), (g) and (h) (see paragraph 29.35). |
| 5.5(e) | Presentation | *As a minimum, an entity shall include, in the statement of comprehensive income, line …* → a single amount comprising the total of: (i) the post-tax profit or loss of a discontinued operation; and (ii) the post-tax gain or loss attributable to an impairment, or reversal of an impairment, of the assets in the … |
| 5.5(f) | Presentation | *As a minimum, an entity shall include, in the statement of comprehensive income, line …* → profit or loss (if an entity has no items of other comprehensive income, this line need not be presented). |
| 5.5(g) | Presentation | *As a minimum, an entity shall include, in the statement of comprehensive income, line …* → each item of other comprehensive income (see paragraph 5.4(b)) classified by nature (excluding amounts in (h)). Such items shall be grouped into those that, in accordance with this Standard: (i) will not be reclassified … |
| 5.5(h) | Presentation | *As a minimum, an entity shall include, in the statement of comprehensive income, line …* → share of the other comprehensive income of associates and jointly controlled entities accounted for using the equity method. |
| 5.5(i) | Presentation | *As a minimum, an entity shall include, in the statement of comprehensive income, line …* → total comprehensive income (if an entity has no items of other comprehensive income, it may use another term for this line such as profit or loss). |
| 5.6(a) | Disclosure | *An entity shall disclose separately the following items in the statement of comprehensive …* → profit or loss for the period attributable to (i) non-controlling interest; and (ii) owners of the parent. |
| 5.6(b) | Disclosure | *An entity shall disclose separately the following items in the statement of comprehensive …* → total comprehensive income for the period attributable to (i) non-controlling interest; and (ii) owners of the parent. Two-statement approach |
| 5.7 | Presentation | Under the two-statement approach, the income statement shall display, as a minimum, line items that present the amounts in paragraph 5.5(a)-5.5(f) for the period, with profit or loss as the last line. The statement of … |
| 5.8 | Presentation | Under this Standard, the effects of corrections of errors and changes in accounting policies are presented as retrospective adjustments of prior periods instead of as part of profit or loss in the period in which they are … |
| 5.9 | Presentation | An entity shall present additional line items, headings and subtotals in the statement of comprehensive income (and in the income statement, if presented), when such presentation is relevant to an understanding of the entity's … |
| 5.10 | Presentation | An entity shall not present or describe any items of income and expense as 'extraordinary items' in the statement of comprehensive income (or in the income statement, if presented) or in the notes. |
| 5.11 | Presentation | An entity shall present an analysis of expenses using a classification based on either the nature of expenses or the function of expenses within the entity, whichever provides information that is reliable and more relevant. … |

### Section 6 — Statement of Changes in Equity and Statement of Income and Retained Earnings

| Ref | Area | Requirement (abridged) |
|---|---|---|
| 6.1 | Scope | This section sets out requirements for presenting the changes in an entity's equity for a period, either in a statement of changes in equity or, if specified conditions are met and an entity chooses, in a statement of income and … |
| 6.2 | Presentation | The statement of changes in equity presents an entity's profit or loss for a reporting period, other comprehensive income for the period, the effects of changes in accounting policies and corrections of errors recognised in the … |
| 6.3(a) | Presentation | *The statement of changes in equity includes the following information:* → total comprehensive income for the period, showing separately the total amounts attributable to owners of the parent and to non-controlling interests; |
| 6.3(b) | Presentation | *The statement of changes in equity includes the following information:* → for each component of equity, the effects of retrospective application or retrospective restatement recognised in accordance with Section 10 Accounting Policies, Estimates and Errors; and |
| 6.3(c) | Presentation | *The statement of changes in equity includes the following information:* → for each component of equity, a reconciliation between the carrying amount at the beginning and the end of the period, separately disclosing changes resulting from: (i) profit or loss; (ii) other comprehensive income; and (iii) … |
| 6.4 | Presentation | The statement of income and retained earnings presents an entity's profit or loss and changes in retained earnings for a reporting period. Paragraph 3.18 permits an entity to present a statement of income and retained earnings in … |
| 6.5(a) | Presentation | *An entity shall present, in the statement of income and retained earnings, the following …* → retained earnings at the beginning of the reporting period; |
| 6.5(b) | Presentation | *An entity shall present, in the statement of income and retained earnings, the following …* → dividends declared and paid or payable during the period; |
| 6.5(c) | Presentation | *An entity shall present, in the statement of income and retained earnings, the following …* → restatements of retained earnings for corrections of prior period errors; |
| 6.5(d) | Presentation | *An entity shall present, in the statement of income and retained earnings, the following …* → restatements of retained earnings for changes in accounting policy; and |
| 6.5(e) | Presentation | *An entity shall present, in the statement of income and retained earnings, the following …* → retained earnings at the end of the reporting period. |
| 6.6(a) | Disclosure | *An entity shall disclose in the notes:* → the amount of dividends proposed (or declared) before the financial statements were authorised for issue but not recognised as a distribution to owners during the reporting period, and the amount per share; and |
| 6.6(b) | Disclosure | *An entity shall disclose in the notes:* → the amount of any cumulative preference dividends not recognised. |

### Section 7 — Statement of Cash Flows

| Ref | Area | Requirement (abridged) |
|---|---|---|
| 7.1 | Scope | This section sets out the information that is to be presented in a statement of cash flows and how to present it. The statement of cash flows provides information about the changes in cash and cash equivalents of an entity for a … |
| 7.2 | Presentation | Cash equivalents are short-term, highly liquid investments that are readily convertible to known amounts of cash and that are subject to an insignificant risk of changes in value. They are held to meet short-term cash commitments … |
| 7.3 | Presentation | An entity shall present a statement of cash flows that presents cash flows for a reporting period classified by operating activities, investing activities and financing activities. Operating activities |
| 7.4(a) | Presentation | *Operating activities are the principal revenue-producing activities of the entity. …* → cash receipts from the sale of goods and the rendering of services; |
| 7.4(b) | Presentation | *Operating activities are the principal revenue-producing activities of the entity. …* → cash receipts from royalties, fees, commissions and other revenue; |
| 7.4(c) | Presentation | *Operating activities are the principal revenue-producing activities of the entity. …* → cash payments to suppliers for goods and services; |
| 7.4(d) | Presentation | *Operating activities are the principal revenue-producing activities of the entity. …* → cash payments to and on behalf of employees; |
| 7.4(e) | Presentation | *Operating activities are the principal revenue-producing activities of the entity. …* → cash payments or refunds of income tax, unless they can be specifically identified with financing and investing activities; and |
| 7.4(f) | Presentation | *Operating activities are the principal revenue-producing activities of the entity. …* → cash receipts and payments from investments, loans and other contracts held for dealing or trading purposes, which are similar to inventory acquired specifically for resale. Some transactions, such as the sale of an item of plant … |
| 7.5(a) | Presentation | *Investing activities are the acquisition and disposal of long-term assets and other …* → cash payments to acquire property, plant and equipment (including self-constructed property, plant and equipment), intangible assets and other long-term assets; |
| 7.5(b) | Presentation | *Investing activities are the acquisition and disposal of long-term assets and other …* → cash receipts from sales of property, plant and equipment, intangibles and other long-term assets; |
| 7.5(c) | Presentation | *Investing activities are the acquisition and disposal of long-term assets and other …* → cash payments to acquire equity or debt instruments of other entities (other than payments for those instruments classified as cash equivalents or held for dealing or trading); |
| 7.5(d) | Presentation | *Investing activities are the acquisition and disposal of long-term assets and other …* → cash receipts from sales of equity or debt instruments of other entities (other than receipts for those instruments classified as cash equivalents or held for dealing or trading); |
| 7.5(e) | Presentation | *Investing activities are the acquisition and disposal of long-term assets and other …* → cash advances and loans made to other parties; |
| 7.5(f) | Presentation | *Investing activities are the acquisition and disposal of long-term assets and other …* → cash receipts from the repayment of advances and loans made to other parties; |
| 7.5(g) | Presentation | *Investing activities are the acquisition and disposal of long-term assets and other …* → cash payments for futures contracts, forward contracts, option contracts and swap contracts, except when the contracts are held for dealing or trading, or the payments are classified as financing activities; and |
| 7.5(h) | Presentation | *Investing activities are the acquisition and disposal of long-term assets and other …* → cash receipts from futures contracts, forward contracts, option contracts and swap contracts, except when the contracts are held for dealing or trading, or the receipts are classified as financing activities. When a contract is … |
| 7.6(a) | Presentation | *Financing activities are activities that result in changes in the size and composition of …* → cash proceeds from issuing shares or other equity instruments; |
| 7.6(b) | Presentation | *Financing activities are activities that result in changes in the size and composition of …* → cash payments to owners to acquire or redeem the entity's shares; |
| 7.6(c) | Presentation | *Financing activities are activities that result in changes in the size and composition of …* → cash proceeds from issuing debentures, loans, notes, bonds, mortgages and other short-term or long-term borrowings; |
| 7.6(d) | Presentation | *Financing activities are activities that result in changes in the size and composition of …* → cash repayments of amounts borrowed; and |
| 7.6(e) | Presentation | *Financing activities are activities that result in changes in the size and composition of …* → cash payments by a lessee for the reduction of the outstanding liability relating to a finance lease. |
| 7.7(a) | Presentation | *An entity shall present cash flows from operating activities using either:* → the indirect method, whereby profit or loss is adjusted for the effects of non-cash transactions, any deferrals or accruals of past or future operating cash receipts or payments and items of income or expense associated with … |
| 7.7(b) | Presentation | *An entity shall present cash flows from operating activities using either:* → the direct method, whereby major classes of gross cash receipts and gross cash payments are disclosed. Indirect method |
| 7.8(a) | Presentation | *Under the indirect method, the net cash flow from operating activities is determined by …* → changes during the period in inventories and operating receivables and payables; |
| 7.8(b) | Presentation | *Under the indirect method, the net cash flow from operating activities is determined by …* → non-cash items such as depreciation, provisions, deferred tax, accrued income (expenses) not yet received (paid) in cash, unrealised foreign currency gains and losses, and undistributed profits of associates; and |
| 7.8(c) | Presentation | *Under the indirect method, the net cash flow from operating activities is determined by …* → all other items for which the cash effects relate to investing or financing. Direct method |
| 7.9(a) | Presentation | *Under the direct method, net cash flow from operating activities is presented by …* → from the accounting records of the entity; or |
| 7.9(b) | Presentation | *Under the direct method, net cash flow from operating activities is presented by …* → by adjusting sales, cost of sales and other items in the statement of comprehensive income (or the income statement, if presented) for: (i) changes during the period in inventories and operating receivables and payables; (ii) … |
| 7.10 | Presentation | An entity shall present separately major classes of gross cash receipts and gross cash payments arising from investing and financing activities. The aggregate cash flows arising from acquisitions and from disposals of … |
| 7.11 | Presentation | An entity shall record cash flows arising from transactions in a foreign currency in the entity's functional currency by applying to the foreign currency amount the exchange rate between the functional currency and the foreign … |
| 7.12 | Presentation | The entity shall translate cash flows of a foreign subsidiary at the exchange rates between the entity's functional currency and the foreign currency at the dates of the cash flows. |
| 7.12A | Presentation | Paragraph 30.19 explains when an exchange rate that approximates the actual rate can be used. |
| 7.13 | Presentation | Unrealised gains and losses arising from changes in foreign currency exchange rates are not cash flows. However, to reconcile cash and cash equivalents at the beginning and the end of the period, the effect of exchange rate … |
| 7.14 | Presentation | An entity shall present separately cash flows from interest and dividends received and paid. The entity shall classify cash flows consistently from period to period as operating, investing or financing activities. |
| 7.15 | Presentation | An entity may classify interest paid and interest and dividends received as operating cash flows because they are included in profit or loss. Alternatively, the entity may classify interest paid and interest and dividends … |
| 7.16 | Presentation | An entity may classify dividends paid as a financing cash flow because they are a cost of obtaining financial resources. Alternatively, the entity may classify dividends paid as a component of cash flows from operating activities … |
| 7.17 | Disclosure | An entity shall present separately cash flows arising from income tax and shall classify them as cash flows from operating activities unless they can be specifically identified with financing and investing activities. When tax … |
| 7.18 | Disclosure | An entity shall exclude from the statement of cash flows investing and financing transactions that do not require the use of cash or cash equivalents. An entity shall disclose such transactions elsewhere in the financial … |
| 7.19(a) | Presentation | *Many investing and financing activities do not have a direct impact on current cash flows …* → the acquisition of assets either by assuming directly related liabilities or by means of a finance lease; |
| 7.19(b) | Presentation | *Many investing and financing activities do not have a direct impact on current cash flows …* → the acquisition of an entity by means of an equity issue; and |
| 7.19(c) | Presentation | *Many investing and financing activities do not have a direct impact on current cash flows …* → the conversion of debt to equity. |
| 7.19A(a) | Disclosure | *An entity shall disclose a reconciliation between the opening and closing balances in the …* → changes from financing cash flows; |
| 7.19A(b) | Disclosure | *An entity shall disclose a reconciliation between the opening and closing balances in the …* → changes arising from obtaining or losing control of subsidiaries or other businesses; |
| 7.19A(c) | Disclosure | *An entity shall disclose a reconciliation between the opening and closing balances in the …* → the effect of changes in foreign exchange rates; |
| 7.19A(d) | Disclosure | *An entity shall disclose a reconciliation between the opening and closing balances in the …* → changes in fair values; and |
| 7.19A(e) | Disclosure | *An entity shall disclose a reconciliation between the opening and closing balances in the …* → other changes. |
| 7.19B(a) | Presentation | *Supplier finance arrangements are characterised by one or more finance providers offering …* → arrangements that are solely credit enhancements for an entity (for example, financial guarantees including letters of credit used as guarantees); and |
| 7.19B(b) | Presentation | *Supplier finance arrangements are characterised by one or more finance providers offering …* → instruments used by an entity to settle directly with a supplier the amounts owed (for example, a situation in which an entity uses a credit card to settle the amount owed to a supplier and then has an obligation to pay the … |
| 7.19C(a) | Disclosure | *An entity shall disclose in aggregate for its supplier finance arrangements:* → the key terms and conditions of the arrangements (for example, terms and conditions related to the interest rate, fees charged, extended payment terms and security or guarantees provided). However, an entity shall disclose … |
| 7.19C(b) | Disclosure | *An entity shall disclose in aggregate for its supplier finance arrangements:* → as at the beginning and end of the reporting period: (i) the carrying amounts, and associated line items presented in the entity's statement of financial position, of the financial liabilities that are part of a supplier finance … |
| 7.19C(c) | Disclosure | *An entity shall disclose in aggregate for its supplier finance arrangements:* → the type and effect of non-cash changes in the carrying amounts of the financial liabilities disclosed in accordance with (b)(i). These non-cash changes include the effect of business combinations, exchange differences and other … |
| 7.20 | Disclosure | An entity shall disclose the components of cash and cash equivalents and shall present a reconciliation of the amounts in the statement of cash flows with the equivalent items reported in the statement of financial position. … |
| 7.21 | Disclosure | An entity shall disclose, together with a commentary by management, the amount of significant cash and cash equivalent balances held by the entity that are not available for use by the entity. Cash and cash equivalents held by an … |

### Section 8 — Notes to the Financial Statements

| Ref | Area | Requirement (abridged) |
|---|---|---|
| 8.1 | Scope | This section sets out the principles underlying information that is to be presented in the notes to the financial statements and how to present it. Notes contain information in addition to that presented in the statement of … |
| 8.2(a) | Presentation | *The notes shall:* → present information about the basis of preparation of the financial statements and the specific accounting policies used, in accordance with paragraphs 8.5-8.7; |
| 8.2(b) | Presentation | *The notes shall:* → disclose the information required by this Standard that is not presented elsewhere in the financial statements; and |
| 8.2(c) | Presentation | *The notes shall:* → provide information that is not presented elsewhere in the financial statements but is relevant to an understanding of any of them. |
| 8.3 | Presentation | An entity shall, as far as practicable, present the notes in a systematic manner. An entity shall cross-reference each item in the financial statements to any related information in the notes. |
| 8.4(a) | Presentation | *An entity normally presents the notes in the following order:* → a statement that the financial statements have been prepared in compliance with the IFRS for SMEs Accounting Standard (see paragraph 3.3); |
| 8.4(b) | Presentation | *An entity normally presents the notes in the following order:* → material accounting policy information (see paragraph 8.5); |
| 8.4(c) | Presentation | *An entity normally presents the notes in the following order:* → supporting information for items presented in the financial statements, in the sequence in which each statement and each line item is presented; and |
| 8.4(d) | Presentation | *An entity normally presents the notes in the following order:* → any other disclosures. |
| 8.5 | Disclosure | An entity shall disclose material accounting policy information. Accounting policy information is material if, when considered together with other information included in the entity's financial statements, it can reasonably be … |
| 8.6(a) | Disclosure | *An entity shall disclose, along with material accounting policy information or other …* → appropriate classes of assets and liabilities in accordance with paragraph 12.30 for which disclosures about fair value measurements are provided; |
| 8.6(b) | Disclosure | *An entity shall disclose, along with material accounting policy information or other …* → that the entity has control of another entity; and |
| 8.6(c) | Disclosure | *An entity shall disclose, along with material accounting policy information or other …* → that the entity has joint control of an arrangement or significant influence over another entity. |
| 8.7(a) | Disclosure | *An entity shall disclose in the notes information about the key assumptions concerning …* → their nature; and |
| 8.7(b) | Disclosure | *An entity shall disclose in the notes information about the key assumptions concerning …* → their carrying amount as at the end of the reporting period. |

### Section 9 — Consolidated and Separate Financial Statements

| Ref | Area | Requirement (abridged) |
|---|---|---|
| 9.2 | Presentation | Except as permitted or required by paragraphs 9.3 and 9.3C, a parent shall present consolidated financial statements. Consolidated financial statements shall include all subsidiaries of the parent, except subsidiaries to which … |
| 9.4 | R&M | A subsidiary is an entity (an investee) that is controlled by another entity (an investor). An investor, regardless of the nature of its involvement with an investee, shall determine whether it is a parent by assessing whether it … |
| 9.12 | R&M | When an investor with decision-making rights (a decision-maker) assesses whether it controls an investee, it shall determine whether it is a principal or an agent. An investor shall also determine whether another entity with … |
| 9.18B | R&M | If a parent loses control of a subsidiary, the parent shall account for all amounts previously recognised in other comprehensive income for that subsidiary, except for the cumulative amount of any exchange differences that relate … |
| 9.20 | Presentation | An entity shall present non-controlling interest in the consolidated statement of financial position within equity, separately from the equity of the owners of the parent, as required by paragraph 4.2(q). |
| 9.20A | R&M | An entity shall treat changes in a parent's controlling interest in a subsidiary that do not result in a loss of control as transactions with owners in their capacity as owners. Accordingly, the carrying amount of the … |
| 9.21 | Disclosure | An entity shall disclose non-controlling interest in the profit or loss of the group separately in the statement of comprehensive income, as required by paragraph 5.6 (or in the income statement, if presented, as required by … |
| 9.23A | Disclosure | In addition to the disclosure requirements in Section 11, a parent shall disclose the carrying amount of investments in subsidiaries that are not consolidated (see paragraphs 9.3A-9.3C) at the reporting date, in total, either in … |
| 9.23B(a) | Disclosure | *An entity shall disclose the gain or loss, if any, calculated in accordance with …* → the portion of that gain or loss attributable to measuring any interest retained in the former subsidiary at its fair value at the date control is lost; and |
| 9.23B(b) | Disclosure | *An entity shall disclose the gain or loss, if any, calculated in accordance with …* → the line items in profit or loss in which the gain or loss is recognised (if not presented separately). |
| 9.26 | R&M | When an entity prepares separate financial statements and describes them as conforming to the IFRS for SMEs Accounting Standard, those statements shall comply with all of the requirements of this Standard except as follows. The … |
| 9.27(a) | Disclosure | *When an entity prepares separate financial statements, those separate financial …* → that the statements are separate financial statements; |
| 9.27(b) | Disclosure | *When an entity prepares separate financial statements, those separate financial …* → a description of the methods used to account for the investments in subsidiaries, jointly controlled entities and associates; and |
| 9.27(c) | Disclosure | *When an entity prepares separate financial statements, those separate financial …* → either: (i) the consolidated financial statements or other financial statements to which they relate; or (ii) the name and principal place of business (and country of incorporation, if different) of the entity whose consolidated … |
| 9.30(a) | Disclosure | *The combined financial statements shall disclose the following:* → the fact that the financial statements are combined financial statements; |
| 9.30(b) | Disclosure | *The combined financial statements shall disclose the following:* → the reason why combined financial statements are prepared; |
| 9.30(c) | Disclosure | *The combined financial statements shall disclose the following:* → the basis for determining which entities are included in the combined financial statements; |
| 9.30(d) | Disclosure | *The combined financial statements shall disclose the following:* → the basis of preparation of the combined financial statements; and |
| 9.30(e) | Disclosure | *The combined financial statements shall disclose the following:* → the related party disclosures required by Section 33 Related Party Disclosures. |

### Section 10 — Accounting Policies, Estimates and Errors

| Ref | Area | Requirement (abridged) |
|---|---|---|
| 10.3 | R&M | If this Standard specifically addresses a transaction, other event or condition, an entity shall apply this Standard. However, the entity need not follow a requirement in this Standard if the effect of doing so would not be … |
| 10.4 | R&M | If this Standard does not specifically address a transaction, other event or condition, an entity's management shall use its judgement in developing and applying an accounting policy that results in information that is: (a) … |
| 10.11 | R&M | An entity shall account for: (a) a change in accounting policy resulting from a change in the requirements of this Standard in accordance with the transitional provisions, if any, specified in Appendix A of this Standard; and (b) … |
| 10.12 | R&M | When a change in accounting policy is applied retrospectively in accordance with paragraph 10.11, the entity shall apply the new accounting policy to comparative information for prior periods to the earliest date for which it is … |
| 10.13(a) | Disclosure | *When an amendment to this Standard has an effect on the current period or any prior …* → the nature of the change in accounting policy; |
| 10.13(b) | Disclosure | *When an amendment to this Standard has an effect on the current period or any prior …* → for the current period and each prior period presented, to the extent practicable, the amount of the adjustment for each financial statement line item affected; |
| 10.13(c) | Disclosure | *When an amendment to this Standard has an effect on the current period or any prior …* → the amount of the adjustment relating to periods before those presented, to the extent practicable; and |
| 10.13(d) | Disclosure | *When an amendment to this Standard has an effect on the current period or any prior …* → an explanation if it is impracticable to determine the amounts to be disclosed in (b) or (c). Financial statements of subsequent periods need not repeat these disclosures. |
| 10.14(a) | Disclosure | *When a voluntary change in accounting policy has an effect on the current period or any …* → the nature of the change in accounting policy; |
| 10.14(b) | Disclosure | *When a voluntary change in accounting policy has an effect on the current period or any …* → the reasons why applying the new accounting policy provides reliable and more relevant information; |
| 10.14(c) | Disclosure | *When a voluntary change in accounting policy has an effect on the current period or any …* → to the extent practicable, the amount of the adjustment for each financial statement line item affected, shown separately: (i) for the current period; (ii) for each prior period presented; and (iii) in the aggregate for periods … |
| 10.14(d) | Disclosure | *When a voluntary change in accounting policy has an effect on the current period or any …* → an explanation if it is impracticable to determine the amounts to be disclosed in (c). Financial statements of subsequent periods need not repeat these disclosures. |
| 10.16 | R&M | An entity shall recognise the effect of a change in an accounting estimate, other than a change to which paragraph 10.17 applies, prospectively by including it in profit or loss in: (a) the period of the change, if the change … |
| 10.17 | R&M | To the extent that a change in an accounting estimate gives rise to changes in assets and liabilities, or relates to an item of equity, the entity shall recognise it by adjusting the carrying amount of the related asset, … |
| 10.18 | Disclosure | An entity shall disclose the nature of any change in an accounting estimate and the effect of the change on assets, liabilities, income and expense for the current period. If it is practicable for the entity to estimate the … |
| 10.23(a) | Disclosure | *An entity shall disclose the following about prior period errors:* → the nature of the prior period error; |
| 10.23(b) | Disclosure | *An entity shall disclose the following about prior period errors:* → for each prior period presented, to the extent practicable, the amount of the correction for each financial statement line item affected; |
| 10.23(c) | Disclosure | *An entity shall disclose the following about prior period errors:* → to the extent practicable, the amount of the correction at the beginning of the earliest prior period presented; and |
| 10.23(d) | Disclosure | *An entity shall disclose the following about prior period errors:* → an explanation if it is not practicable to determine the amounts to be disclosed in (b) or (c). Financial statements of subsequent periods need not repeat these disclosures. |

### Section 11 — Financial Instruments

| Ref | Area | Requirement (abridged) |
|---|---|---|
| 11.8 | R&M | An entity shall account for the following financial instruments as basic financial instruments in accordance with Part I of Section 11: (a) cash; (b) a debt instrument (such as an account, note or loan receivable or payable) that … |
| 11.12 | R&M | An entity shall recognise a financial asset or a financial liability only when the entity becomes a party to the contractual provisions of the instrument. |
| 11.13 | R&M | When a financial asset or financial liability is recognised initially, an entity shall measure it at the transaction price (including transaction costs except in the initial measurement of financial assets and liabilities that … |
| 11.13A | R&M | A trade receivable shall be recognised initially at the amount determined by applying Section 23, unless the arrangement constitutes, in effect, a financing transaction and the entity does not apply the option in paragraph 23.38. … |
| 11.13B | R&M | An arrangement constitutes a financing transaction if payment is deferred beyond normal business terms, for example, providing interest-free credit to a buyer for the sale of goods, or is financed at a rate of interest that is … |
| 11.14 | R&M | At the end of each reporting period, an entity shall measure financial instruments as follows, without any deduction for transaction costs the entity may incur on sale or other disposal: (a) debt instruments that meet the … |
| 11.18 | R&M | When calculating the effective interest rate, an entity shall amortise any related fees, finance charges paid or received (such as 'points'), transaction costs and other premiums or discounts over the expected life of the … |
| 11.20 | R&M | If an entity revises its estimates of payments or receipts, the entity shall adjust the carrying amount of the financial asset or financial liability (or group of financial instruments) to reflect actual and revised estimated … |
| 11.21 | R&M | At the end of each reporting period, an entity shall assess whether there is objective evidence of impairment of any financial assets that are measured at cost or amortised cost. If there is objective evidence of impairment, the … |
| 11.25 | R&M | An entity shall measure an impairment loss on the following financial assets measured at cost or amortised cost as follows: (a) for a financial asset measured at amortised cost in accordance with paragraph 11.14(a), the … |
| 11.26 | R&M | If, in a subsequent period, the amount of an impairment loss decreases and the decrease can be related objectively to an event occurring after the impairment was recognised (such as an improvement in the debtor's credit rating), … |
| 11.33 | R&M | An entity shall derecognise a financial asset only when either: (a) the contractual rights to the cash flows from the financial asset expire or are settled; (b) the entity transfers to another party substantially all of the risks … |
| 11.34 | R&M | If a transfer does not result in derecognition because the entity has retained significant risks and rewards of ownership of the transferred asset, the entity shall continue to recognise the transferred asset in its entirety and … |
| 11.35 | R&M | If a transferor provides non-cash collateral (such as debt or equity instruments) to the transferee, the accounting for the collateral by the transferor and the transferee depends on whether the transferee has the right to sell … |
| 11.36 | R&M | An entity shall derecognise a financial liability (or a part of a financial liability) only when it is extinguished-ie when the obligation specified in the contract is discharged, is cancelled or expires. |
| 11.37 | R&M | If an existing borrower and lender exchange financial instruments with substantially different terms, the entities shall account for the transaction as an extinguishment of the original financial liability and the recognition of … |
| 11.38 | R&M | The entity shall recognise in profit or loss any difference between the carrying amount of the financial liability (or part of a financial liability) extinguished or transferred to another party and the consideration paid, … |
| 11.40 | Disclosure | In accordance with paragraph 8.5, an entity shall disclose material accounting policy information. Information about the measurement basis (or bases) for financial instruments used in preparing the financial statements is … |
| 11.41(a) | Disclosure | *An entity shall disclose the carrying amounts of each of the following categories of …* → financial assets measured at fair value through profit or loss (paragraph 11.14(c)(i) and paragraph 11.54); |
| 11.41(b) | Disclosure | *An entity shall disclose the carrying amounts of each of the following categories of …* → financial assets that are debt instruments measured at amortised cost (paragraph 11.14(a)); |
| 11.41(c) | Disclosure | *An entity shall disclose the carrying amounts of each of the following categories of …* → financial assets that are equity instruments measured at cost less impairment (paragraph 11.14(c)(ii) and paragraph 11.54); |
| 11.41(d) | Disclosure | *An entity shall disclose the carrying amounts of each of the following categories of …* → financial liabilities measured at fair value through profit or loss (paragraph 11.54); |
| 11.41(e) | Disclosure | *An entity shall disclose the carrying amounts of each of the following categories of …* → financial liabilities measured at amortised cost (paragraph 11.14(a)); and |
| 11.41(f) | Disclosure | *An entity shall disclose the carrying amounts of each of the following categories of …* → loan commitments measured at cost less impairment (paragraph 11.14(b)). |
| 11.42 | Disclosure | An entity shall disclose information that enables users of its financial statements to evaluate the significance of financial instruments for its financial position and performance. For example, for long-term debt such … |
| 11.43(a) | Disclosure | *An entity shall disclose an analysis of the age, by reference to due date, of trade …* → the amortised cost of the financial assets before adjusting for any reduction (directly or by using an allowance account) for impairment or uncollectability (see paragraph 11.15(d)); and |
| 11.43(b) | Disclosure | *An entity shall disclose an analysis of the age, by reference to due date, of trade …* → any reduction (directly or by using an allowance account) for impairment or uncollectability (see paragraph 11.15(d)). |
| 11.43A | Disclosure | An entity shall disclose a maturity analysis for financial liabilities by category (see paragraph 11.41(d)-(e)). The maturity analysis shall include the remaining contractual maturities. The contractual amounts disclosed in the … |
| 11.43B | R&M | In preparing the disclosures in paragraphs 11.43-11.43A, an entity shall use time bands it considers to be the most useful. For example, for paragraph 11.43 the time bands might be: (a) not later than one month; (b) later than … |
| 11.44 | Disclosure | If a reliable measure of fair value is no longer available, or is not available without undue cost or effort when such an exemption is provided, for any financial instruments that would otherwise be required to be measured at … |
| 11.45(a) | Disclosure | *If an entity has transferred financial assets to another party in a transaction that does …* → the nature of the assets; |
| 11.45(b) | Disclosure | *If an entity has transferred financial assets to another party in a transaction that does …* → the nature of the risks and rewards of ownership to which the entity remains exposed; and |
| 11.45(c) | Disclosure | *If an entity has transferred financial assets to another party in a transaction that does …* → the carrying amounts of the assets and of any associated liabilities that the entity continues to recognise. Collateral |
| 11.46(a) | Disclosure | *When an entity has pledged financial assets as collateral for liabilities or contingent …* → the carrying amount of the financial assets pledged as collateral; and |
| 11.46(b) | Disclosure | *When an entity has pledged financial assets as collateral for liabilities or contingent …* → the terms and conditions relating to its pledge. Defaults and breaches on loans payable |
| 11.47(a) | Disclosure | *For loans payable recognised at the reporting date for which there is a breach of terms …* → details of that breach or default; |
| 11.47(b) | Disclosure | *For loans payable recognised at the reporting date for which there is a breach of terms …* → the carrying amount of the related loans payable at the reporting date; and |
| 11.47(c) | Disclosure | *For loans payable recognised at the reporting date for which there is a breach of terms …* → whether the breach or default was remedied, or the terms of the loans payable were renegotiated, before the financial statements were authorised for issue. Items of income, expense, gains or losses |
| 11.48(a) | Disclosure | *An entity shall disclose the following items of income, expense, gains or losses:* → income, expense, gains or losses, including changes in fair value, recognised on: (i) financial assets measured at fair value through profit or loss; (ii) financial liabilities measured at fair value through profit or loss; (iii) … |
| 11.48(b) | Disclosure | *An entity shall disclose the following items of income, expense, gains or losses:* → total interest income and total interest expense (calculated using the effective interest method) for financial assets or financial liabilities that are not measured at fair value through profit or loss; and |
| 11.48(c) | Disclosure | *An entity shall disclose the following items of income, expense, gains or losses:* → the amount of any impairment loss for each class of financial asset. |
| 11.52 | R&M | An entity shall recognise a financial asset or a financial liability only when the entity becomes a party to the contractual provisions of the instrument. |
| 11.53 | R&M | When a financial asset or financial liability is recognised initially, an entity shall measure it at its fair value, which is normally the transaction price. |
| 11.54 | R&M | At the end of each reporting period, an entity shall measure all financial instruments within the scope of Part II of Section 11 at fair value and recognise changes in fair value in profit or loss, except as follows: (a) some … |
| 11.56 | R&M | If a reliable measure of fair value is no longer available without undue cost or effort for an equity instrument, or a contract linked to such an instrument that if exercised will result in the delivery of such instruments, that … |
| 11.57 | R&M | An entity shall apply the guidance on fair value in Section 12 to fair value measurements in accordance with Section 11. |
| 11.59 | R&M | An entity shall apply the guidance on impairment in paragraphs 11.21-11.26 to financial assets measured at cost less impairment in accordance with Part II of Section 11. |
| 11.60 | R&M | An entity shall apply the derecognition requirements in paragraphs 11.33-11.38 to financial assets and financial liabilities to which Section 11 applies. |
| 11.66 | R&M | If the hedged risk is the fixed interest rate risk of a debt instrument measured at amortised cost, the entity shall recognise the periodic net cash settlements on the interest rate swap that is the hedging instrument in profit … |
| 11.69 | R&M | If the conditions in paragraph 11.62 are met and the hedged risk is: (a) the variable interest rate risk in a debt instrument measured at amortised cost; (b) the foreign exchange risk in a firm commitment or a highly probable … |
| 11.70 | R&M | If the hedged risk is the variable interest rate risk in a debt instrument measured at amortised cost, the entity shall subsequently recognise in profit or loss the periodic net cash settlements from the interest rate swap that … |
| 11.73(a) | Disclosure | *An entity shall disclose the following separately for hedges of each of the four types of …* → a description of the hedge; |
| 11.73(b) | Disclosure | *An entity shall disclose the following separately for hedges of each of the four types of …* → a description of the financial instruments designated as hedging instruments and their fair values at the reporting date; and |
| 11.73(c) | Disclosure | *An entity shall disclose the following separately for hedges of each of the four types of …* → the nature of the risks being hedged, including a description of the hedged item. |
| 11.74(a) | Disclosure | *If an entity uses hedge accounting for a hedge of fixed interest rate risk or commodity …* → the amount of the change in fair value of the hedging instrument recognised in profit or loss for the period; and |
| 11.74(b) | Disclosure | *If an entity uses hedge accounting for a hedge of fixed interest rate risk or commodity …* → the amount of the change in fair value of the hedged item recognised in profit or loss for the period. |
| 11.75(a) | Disclosure | *If an entity uses hedge accounting for a hedge of variable interest rate risk, foreign …* → the periods when the cash flows are expected to occur and when they are expected to affect profit or loss; |
| 11.75(b) | Disclosure | *If an entity uses hedge accounting for a hedge of variable interest rate risk, foreign …* → a description of any forecast transaction for which hedge accounting had previously been used, but which is no longer expected to occur; |
| 11.75(c) | Disclosure | *If an entity uses hedge accounting for a hedge of variable interest rate risk, foreign …* → the amount of the change in fair value of the hedging instrument that was recognised in other comprehensive income during the period (paragraph 11.69); (d) the amount that was reclassified to profit or loss for the period … |

### Section 12 — Fair Value Measurement

| Ref | Area | Requirement (abridged) |
|---|---|---|
| 12.4 | R&M | Fair value is the price that would be received to sell an asset, or paid to transfer a liability, in an orderly transaction between market participants at the measurement date. Fair value is a market-based rather than … |
| 12.14 | R&M | If a price for an identical asset or liability is not observable, an entity shall measure fair value using another valuation technique. The entity shall use a valuation technique that is appropriate in the circumstances and for … |
| 12.15 | R&M | Three widely used valuation approaches are the market approach, the cost approach and the income approach. An entity shall use a valuation technique consistent with one or more of these approaches: (a) the market approach uses as … |
| 12.16 | R&M | An entity shall account for revisions resulting from a change in the valuation technique or a change in its application as a change in accounting estimate in Section 10. However, an entity is not required to apply the disclosure … |
| 12.17 | R&M | An asset or a liability measured at fair value might have a bid price and an ask price (for example, an input from a dealer market). If so, an entity shall measure the fair value of the asset or liability using the price within … |
| 12.21 | R&M | If a reliable measure of fair value is no longer available for an asset measured at fair value (or is not available without undue cost or effort when an undue cost or effort exemption is provided-for example, see paragraphs … |
| 12.22 | R&M | For consistency and comparability in fair value measurements and related disclosures, an entity shall apply a fair value hierarchy, categorising the inputs to valuation techniques used to measure fair value into three levels- … |
| 12.23 | R&M | Level 1 inputs are unadjusted quoted prices in active markets for identical assets or liabilities that the entity can access at the measurement date. A quoted price in an active market provides the most reliable evidence of fair … |
| 12.24 | R&M | If an entity holds a position in a single asset or liability (including a position comprising many identical assets or liabilities, such as a holding of financial instruments) and the asset or liability is traded in an active … |
| 12.28(a) | Disclosure | *An entity shall disclose for each class of assets and liabilities measured at fair value …* → the carrying amounts at the end of the reporting period; |
| 12.28(b) | Disclosure | *An entity shall disclose for each class of assets and liabilities measured at fair value …* → the level of the fair value hierarchy within which the fair value measurements are categorised in their entirety (Level 1, Level 2 or Level 3); and |
| 12.28(c) | Disclosure | *An entity shall disclose for each class of assets and liabilities measured at fair value …* → a description of the valuation technique(s) the entity used for fair value measurements categorised within Level 2 and Level 3 of the fair value hierarchy, and the inputs used in the fair value measurement. |
| 12.29(a) | Disclosure | *For recurring fair value measurements categorised within Level 3 of the fair value …* → profit or loss, and the line items in profit or loss in which those gains or losses are recognised; and |
| 12.29(b) | Disclosure | *For recurring fair value measurements categorised within Level 3 of the fair value …* → other comprehensive income, and the line items in other comprehensive income in which those gains or losses are recognised. |
| 12.30 | R&M | An entity shall determine appropriate classes of assets and liabilities on the basis of: (a) the nature, characteristics and risks of the asset or liability; and (b) the level of the fair value hierarchy within which the fair … |
| 12.31 | Disclosure | A class of assets and liabilities will often require greater disaggregation than the line items presented in the statement of financial position; however, an entity shall disclose sufficient information to permit reconciliation … |
| 12.32 | Presentation | An entity shall present the quantitative disclosures required by paragraphs 12.28-12.31 in a table unless another format would be more useful. |

### Section 13 — Inventories

| Ref | Area | Requirement (abridged) |
|---|---|---|
| 13.4 | R&M | An entity shall measure inventories at the lower of cost and estimated selling price less costs to complete and sell. |
| 13.9 | R&M | An entity shall allocate fixed production overheads to the costs of conversion on the basis of the normal capacity of the production facilities. Normal capacity is the production expected to be achieved on average over a number … |
| 13.10 | R&M | A production process may result in more than one product being produced simultaneously. This is the case, for example, when joint products are produced or when there is a main product and a by-product. When the costs of raw … |
| 13.17 | R&M | An entity shall measure the cost of inventories of items that are not ordinarily interchangeable and goods or services produced and segregated for specific projects by using specific identification of their individual costs. |
| 13.18 | R&M | An entity shall measure the cost of inventories, other than those dealt with in paragraph 13.17, by using the first-in, first-out (FIFO) or weighted average cost formula. An entity shall use the same cost formula for all … |
| 13.20 | R&M | When inventories are sold, the entity shall recognise the carrying amount of those inventories as an expense in the period in which the related revenue is recognised. |
| 13.22(a) | Disclosure | *An entity shall disclose the following:* → the accounting policies adopted in measuring inventories, including the cost formula used; |
| 13.22(b) | Disclosure | *An entity shall disclose the following:* → the total carrying amount of inventories and the carrying amount in classifications appropriate to the entity; |
| 13.22(c) | Disclosure | *An entity shall disclose the following:* → the amount of inventories recognised as an expense during the period; |
| 13.22(d) | Disclosure | *An entity shall disclose the following:* → impairment losses recognised or reversed in profit or loss in accordance with Section 27 Impairment of Assets; and |
| 13.22(e) | Disclosure | *An entity shall disclose the following:* → the total carrying amount of inventories pledged as security for liabilities. |

### Section 14 — Investments in Associates

| Ref | Area | Requirement (abridged) |
|---|---|---|
| 14.4 | R&M | An investor shall account for all of its investments in associates using one of the following: (a) the cost model in paragraph 14.5; (b) the equity method in paragraph 14.8; or (c) the fair value model in paragraph 14.9. Cost … |
| 14.5 | R&M | An investor shall measure its investments in associates, other than those for which there is a published price quotation (see paragraph 14.7) at cost less any accumulated impairment losses recognised in accordance with Section 27 … |
| 14.6 | R&M | The investor shall recognise dividends and other distributions received from the investment as income without regard to whether the distributions are from accumulated profits of the associate arising before or after the date of … |
| 14.7 | R&M | An investor shall measure its investments in associates for which there is a published price quotation using the fair value model (see paragraph 14.9). Equity method |
| 14.8 | R&M | Under the equity method of accounting, an equity investment is initially recognised at the transaction price (including transaction costs) and is subsequently adjusted to reflect the investor's share of the profit or loss and … |
| 14.9 | R&M | When an investment in an associate is recognised initially, an investor shall measure it at the transaction price. Transaction price excludes transaction costs. |
| 14.10 | R&M | At each reporting date, an investor shall measure its investments in associates at fair value, with changes in fair value recognised in profit or loss, using the fair value measurement guidance in Section 12 Fair Value … |
| 14.11 | Presentation | An investor shall classify investments in associates as non-current assets. |
| 14.12(a) | Disclosure | *An entity shall disclose the following:* → its accounting policy for investments in associates; |
| 14.12(b) | Disclosure | *An entity shall disclose the following:* → the carrying amount of investments in associates (see paragraph 4.2(j)); and |
| 14.12(c) | Disclosure | *An entity shall disclose the following:* → the fair value of its investment in an associate if a market price for the investment is quoted and the entity accounts for the associate using the equity method. |
| 14.13 | Disclosure | For investments in associates accounted for using the cost model, an investor shall disclose the amount of dividends and other distributions recognised as income. |
| 14.14 | Disclosure | For investments in associates accounted for using the equity method, an investor shall disclose separately its share of the profit or loss and its share of any discontinued operations. |
| 14.15 | Disclosure | For investments in associates accounted for using the fair value model, an investor shall make the disclosures required in Section 12. If an investor applies the undue cost or effort exemption in paragraph 14.10 for any … |

### Section 15 — Joint Arrangements

| Ref | Area | Requirement (abridged) |
|---|---|---|
| 15.5 | R&M | In respect of its interests in jointly controlled operations, a party shall recognise in its financial statements: (a) the assets that it controls and the liabilities that it incurs; and (b) the expenses that it incurs and its … |
| 15.7 | R&M | In respect of its interest in a jointly controlled asset, a party shall recognise in its financial statements: (a) its share of the jointly controlled assets, classified according to the nature of the assets; (b) any liabilities … |
| 15.9 | R&M | A party that has joint control shall account for all of its investments in jointly controlled entities using one of the following: (a) the cost model in paragraph 15.10; (b) the equity method in paragraph 15.13; or (c) the fair … |
| 15.10 | R&M | A party that has joint control shall measure its investments in jointly controlled entities, other than those for which there is a published price quotation (see paragraph 15.12) at cost less any accumulated impairment losses … |
| 15.11 | R&M | The party that has joint control shall recognise distributions received from the investment as income without regard to whether the distributions are from accumulated profits of the jointly controlled entity arising before or … |
| 15.12 | R&M | Applying the cost model in paragraph 15.10, a party that has joint control shall measure its investments in jointly controlled entities for which there is a published price quotation using the fair value model (see paragraph … |
| 15.13 | R&M | A party that has joint control shall measure its investments in jointly controlled entities using the equity method following the procedures in paragraph 14.8 (substituting 'joint control' where that paragraph refers to … |
| 15.14 | R&M | When an investment in a jointly controlled entity is recognised initially, a party that has joint control shall measure it at transaction price. Transaction price excludes transaction costs. |
| 15.15 | R&M | At each reporting date, a party that has joint control shall measure its investments in jointly controlled entities at fair value, with changes in fair value recognised in profit or loss, using the fair value measurement guidance … |
| 15.16 | R&M | When a party to the joint arrangement that has joint control contributes or sells assets to a joint arrangement, recognition of any portion of a gain or loss from the transaction shall reflect the substance of the transaction. … |
| 15.17 | R&M | When a party to the joint arrangement that has joint control purchases assets from a joint arrangement, that party shall not recognise its share of the profits of the joint arrangement from the transaction until it resells the … |
| 15.18 | R&M | A party that participates in, but does not have joint control of, a jointly controlled entity shall account for its interest in the arrangement in accordance with Section 11 Financial Instruments unless it has significant … |
| 15.18A | R&M | A party that participates in, but does not have joint control of, a jointly controlled operation shall account for its interest in the arrangement in accordance with paragraph 15.5. |
| 15.18B | R&M | A party that participates in, but does not have joint control of, jointly controlled assets shall account for its interest in the arrangement in accordance with paragraph 15.7. |
| 15.19(a) | Disclosure | *An entity shall disclose the following:* → the accounting policy it uses for recognising its investments in jointly controlled entities; |
| 15.19(b) | Disclosure | *An entity shall disclose the following:* → the carrying amount of investments in jointly controlled entities (see paragraph 4.2(k)); |
| 15.19(c) | Disclosure | *An entity shall disclose the following:* → the fair value of its investment in a jointly controlled entity, if a market price for the investment is quoted and the entity accounts for the jointly controlled entity using the equity method; and |
| 15.19(d) | Disclosure | *An entity shall disclose the following:* → the aggregate amount of its commitments relating to jointly controlled entities, including its share in the commitments that have been incurred jointly with other parties. |
| 15.21 | Disclosure | For jointly controlled entities accounted for using the fair value model, a party that has joint control shall make the disclosures required in Section 12. If a party applies the undue cost or effort exemption in paragraph 15.15 … |

### Section 16 — Investment Property

| Ref | Area | Requirement (abridged) |
|---|---|---|
| 16.3A | R&M | An entity shall use its judgement to determine whether the acquisition of investment property is the acquisition of an asset or a group of assets, or a business combination within the scope of Section 19 Business Combinations and … |
| 16.5 | R&M | An entity shall measure investment property at its cost at initial recognition. The cost of a purchased investment property comprises its purchase price and any directly attributable expenditure such as legal and brokerage fees, … |
| 16.7 | R&M | Investment property whose fair value can be measured reliably without undue cost or effort shall be measured at fair value at each reporting date with changes in fair value recognised in profit or loss. If a property interest … |
| 16.10(b) | Disclosure | *An entity shall disclose the following for all investment property accounted for at fair …* → the extent to which the fair value of investment property (as measured or disclosed in the financial statements) is based on a valuation by an independent valuer who holds a recognised and relevant professional qualification and … |
| 16.10(c) | Disclosure | *An entity shall disclose the following for all investment property accounted for at fair …* → the existence and amounts of restrictions on the realisability of investment property or the remittance of income and proceeds of disposal. |
| 16.10(d) | Disclosure | *An entity shall disclose the following for all investment property accounted for at fair …* → contractual obligations to purchase, construct or develop investment property or for repairs, maintenance or enhancements. |
| 16.10(e) | Disclosure | *An entity shall disclose the following for all investment property accounted for at fair …* → a reconciliation between the carrying amounts of investment property at the beginning and end of the period, showing separately: (i) additions, disclosing separately those additions resulting from acquisitions through business … |

### Section 17 — Property, Plant and Equipment

| Ref | Area | Requirement (abridged) |
|---|---|---|
| 17.4 | R&M | An entity shall recognise the cost of an item of property, plant and equipment as an asset if, and only if: (a) it is probable that future economic benefits associated with the item will flow to the entity; and (b) the cost of … |
| 17.6 | R&M | Parts of some items of property, plant and equipment may require replacement at regular intervals (for example, the roof of a building). An entity shall add to the carrying amount of an item of property, plant and equipment the … |
| 17.8 | R&M | Land and buildings are separable assets and an entity shall account for them separately, even when they are acquired together. |
| 17.9 | R&M | An entity shall measure an item of property, plant and equipment at initial recognition at its cost. Elements of cost |
| 17.11 | R&M | The following costs are not costs of an item of property, plant and equipment and an entity shall recognise them as an expense when they are incurred: (a) costs of opening a new facility; (b) costs of introducing a new product or … |
| 17.14 | R&M | An item of property, plant or equipment may be acquired in exchange for a non-monetary asset, or assets, or a combination of monetary and non- monetary assets. An entity shall measure the cost of the acquired asset at fair value … |
| 17.15 | R&M | An entity shall choose either the cost model in paragraph 17.15A or the revaluation model in paragraph 17.15B as its accounting policy and shall apply that policy to an entire class of property, plant and equipment. An entity … |
| 17.15A | R&M | An entity shall measure an item of property, plant and equipment after initial recognition at cost less any accumulated depreciation and any accumulated impairment losses. Revaluation model |
| 17.15B | R&M | An entity shall measure an item of property, plant and equipment whose fair value can be measured reliably at a revalued amount, being its fair value at the date of the revaluation less any subsequent accumulated depreciation and … |
| 17.16 | R&M | If the major components of an item of property, plant and equipment have significantly different patterns of consumption of economic benefits, an entity shall allocate the initial cost of the asset to its major components and … |
| 17.18 | R&M | An entity shall allocate the depreciable amount of an asset on a systematic basis over its useful life. |
| 17.19 | R&M | Factors such as a change in how an asset is used, significant unexpected wear and tear, technological advancement and changes in market prices may indicate that the residual value or useful life of an asset has changed since the … |
| 17.23 | R&M | If there is an indication that there has been a significant change since the last annual reporting date in the pattern by which an entity expects to consume an asset's future economic benefits, the entity shall review its present … |
| 17.24 | R&M | At each reporting date, an entity shall apply Section 27 Impairment of Assets to determine whether an item or group of items of property, plant and equipment is impaired and, if so, how to recognise and measure the impairment … |
| 17.27 | R&M | An entity shall derecognise an item of property, plant and equipment: (a) on disposal; or (b) when no future economic benefits are expected from its use or disposal. |
| 17.28 | R&M | An entity shall recognise the gain or loss on the derecognition of an item of property, plant and equipment in profit or loss when the item is derecognised (unless Section 20 Leases requires otherwise on a sale and leaseback). … |
| 17.30 | R&M | An entity shall determine the gain or loss arising from the derecognition of an item of property, plant and equipment as the difference between the net disposal proceeds, if any, and the carrying amount of the item. |
| 17.31(a) | Disclosure | *An entity shall disclose the following for each class of property, plant and equipment …* → the measurement bases used for determining the gross carrying amount; |
| 17.31(b) | Disclosure | *An entity shall disclose the following for each class of property, plant and equipment …* → the depreciation methods used; |
| 17.31(c) | Disclosure | *An entity shall disclose the following for each class of property, plant and equipment …* → the useful lives or the depreciation rates used; |
| 17.31(d) | Disclosure | *An entity shall disclose the following for each class of property, plant and equipment …* → the gross carrying amount and the accumulated depreciation (aggregated with accumulated impairment losses) at the beginning and end of the reporting period; and |
| 17.31(e) | Disclosure | *An entity shall disclose the following for each class of property, plant and equipment …* → a reconciliation of the carrying amount at the beginning and end of the reporting period showing separately: (i) additions; (ii) disposals; (iii) acquisitions through business combinations; (iv) increases or decreases resulting … |
| 17.32(a) | Disclosure | *An entity shall also disclose the following:* → the existence and carrying amounts of property, plant and equipment to which the entity has restricted title or that is pledged as security for liabilities; |
| 17.32(b) | Disclosure | *An entity shall also disclose the following:* → the amount of contractual commitments for the acquisition of property, plant and equipment; and |
| 17.32(c) | Disclosure | *An entity shall also disclose the following:* → if an entity has investment property whose fair value cannot be measured reliably without undue cost or effort it shall disclose that fact and the reasons why fair value measurement would involve undue cost or effort for those … |
| 17.33(a) | Disclosure | *If items of property, plant and equipment are stated at revalued amounts, an entity shall …* → the effective date of the revaluation; |
| 17.33(b) | Disclosure | *If items of property, plant and equipment are stated at revalued amounts, an entity shall …* → whether an independent valuer was involved; |

### Section 18 — Intangible Assets other than Goodwill

| Ref | Area | Requirement (abridged) |
|---|---|---|
| 18.4 | R&M | An entity shall recognise an intangible asset as an asset if, and only if: (a) it is probable that the expected future economic benefits that are attributable to the asset will flow to the entity; (b) the cost or value of the … |
| 18.9 | R&M | An entity shall measure an intangible asset initially at cost. Separate acquisition |
| 18.13 | R&M | An intangible asset may be acquired in exchange for a non-monetary asset or assets, or a combination of monetary and non-monetary assets. An entity shall measure the cost of such an intangible asset at fair value unless (a) the … |
| 18.14 | R&M | An entity shall recognise expenditure incurred internally on an intangible item, including all expenditure for both research and development activities, as an expense when it is incurred unless it forms part of the cost of … |
| 18.15 | R&M | As examples of applying the preceding paragraph, an entity shall recognise expenditure on the following items as an expense and shall not recognise such expenditure as intangible assets: (a) internally generated brands, logos, … |
| 18.18 | R&M | An entity shall measure intangible assets at cost less any accumulated amortisation and any accumulated impairment losses. The requirements for amortisation are set out in this section. The requirements for recognition of … |
| 18.21 | R&M | An entity shall allocate the depreciable amount of an intangible asset on a systematic basis over its useful life. The amortisation charge for each period shall be recognised as an expense, unless another section of this Standard … |
| 18.22 | R&M | Amortisation begins when the intangible asset is available for use, ie when it is in the location and condition necessary for it to be usable in the manner intended by management. Amortisation ceases when the asset is … |
| 18.24 | R&M | Factors such as a change in how an intangible asset is used, technological advancement and changes in market prices may indicate that the residual value or useful life of an intangible asset has changed since the most recent … |
| 18.25 | R&M | To determine whether an intangible asset is impaired, an entity shall apply Section 27. That section explains when and how an entity reviews the carrying amount of its assets, how it determines the recoverable amount, of an asset … |
| 18.26 | R&M | An entity shall derecognise an intangible asset, and shall recognise a gain or loss in profit or loss: (a) on disposal; or (b) when no future economic benefits are expected from its use or disposal. |
| 18.27(a) | Disclosure | *An entity shall disclose the following for each class of intangible assets:* → the useful lives or the amortisation rates used; |
| 18.27(b) | Disclosure | *An entity shall disclose the following for each class of intangible assets:* → the amortisation methods used; |
| 18.27(c) | Disclosure | *An entity shall disclose the following for each class of intangible assets:* → the gross carrying amount and any accumulated amortisation (aggregated with accumulated impairment losses) at the beginning and end of the reporting period; |
| 18.27(d) | Disclosure | *An entity shall disclose the following for each class of intangible assets:* → the line item(s) in the statement of comprehensive income (and in the income statement, if presented) in which any amortisation of intangible assets is included; and |
| 18.27(e) | Disclosure | *An entity shall disclose the following for each class of intangible assets:* → a reconciliation of the carrying amount at the beginning and end of the reporting period showing separately: (i) additions; (ii) disposals; (iii) acquisitions through business combinations; (iv) amortisation; (v) impairment … |
| 18.28(a) | Disclosure | *An entity shall also disclose:* → a description, the carrying amount and remaining amortisation period of any individual intangible asset that is material to the entity's financial statements; |
| 18.28(b) | Disclosure | *An entity shall also disclose:* → for intangible assets acquired by way of a government grant and initially recognised at fair value (see paragraph 18.12): (i) the fair value initially recognised for these assets; and (ii) their carrying amounts. |
| 18.28(c) | Disclosure | *An entity shall also disclose:* → the existence and carrying amounts of intangible assets to which the entity has restricted title or that are pledged as security for liabilities; and |
| 18.28(d) | Disclosure | *An entity shall also disclose:* → the amount of contractual commitments for the acquisition of intangible assets. |
| 18.29 | Disclosure | An entity shall disclose the aggregate amount of research and development expenditure recognised as an expense during the period (ie the amount of expenditure incurred internally on research and development that has not been … |

### Section 19 — Business Combinations and Goodwill

| Ref | Area | Requirement (abridged) |
|---|---|---|
| 19.3 | R&M | An entity shall determine whether a transaction or other event meets the definition of a business combination. A business combination is a transaction or other event in which an entity (the acquirer) obtains control of one or … |
| 19.6 | R&M | An entity shall account for each business combination by applying the acquisition method. |
| 19.12 | R&M | At the acquisition date, the acquirer shall recognise, separately from goodwill, the identifiable assets acquired, the liabilities assumed and any non-controlling interest in the acquiree-except as specified in paragraphs … |
| 19.13 | R&M | The acquirer shall measure the identifiable assets acquired and the liabilities assumed at their acquisition-date fair values-except as specified in paragraphs 19.16-19.21. |
| 19.14 | R&M | For each business combination, the acquirer shall measure at the acquisition date any non-controlling interests in the acquiree at the non-controlling interest's proportionate share of the recognised amounts of the acquiree's … |
| 19.16 | R&M | An acquirer shall recognise an intangible asset acquired in a business combination if the asset meets the recognition principles set out in Section 18 Intangible Assets other than Goodwill and the asset's fair value can be … |
| 19.17 | R&M | If an acquirer assumes a liability or contingent liability in a business combination and that liability or contingent liability would have been within the scope of Section 21 Provisions and Contingencies had the acquirer incurred … |
| 19.18 | R&M | If an acquirer determines, in accordance with paragraph 19.17, that a present obligation exists, the acquirer shall recognise a contingent liability at the acquisition date if: (a) the present obligation meets the definition of a … |
| 19.20 | R&M | An acquirer shall recognise and measure in accordance with Section 29 Income Tax a deferred tax asset or deferred tax liability arising from the assets acquired and liabilities assumed in a business combination. |
| 19.21 | R&M | An acquirer shall recognise and measure in accordance with Section 28 Employee Benefits a liability (or asset, if any) related to the acquiree's employee benefit arrangements. Recognising and measuring goodwill or a gain from a … |
| 19.22 | R&M | The acquirer shall recognise goodwill as of the acquisition date measured as the excess of (a) over (b): (a) the sum of: (i) the consideration transferred, measured in accordance with paragraph 19.25; (ii) the amount of any … |
| 19.23 | R&M | Occasionally, an acquirer will make a bargain purchase-a business combination in which the amount in paragraph 19.22(b) exceeds the sum of the amounts specified in paragraph 19.22(a). If an excess remains after an acquirer … |
| 19.24 | R&M | Before recognising a gain on a bargain purchase, the acquirer shall reassess whether it has correctly identified all the assets acquired and all the liabilities assumed and shall recognise any additional assets or liabilities … |
| 19.25 | R&M | An acquirer shall measure the consideration transferred in a business combination at fair value. The acquirer shall calculate the consideration transferred as the sum of the acquisition-date fair values of the assets transferred … |
| 19.26 | R&M | The consideration an acquirer transfers in exchange for an acquiree includes any asset or liability resulting from a contingent consideration arrangement (see paragraph 19.25). The acquirer shall recognise the acquisition-date … |
| 19.27 | R&M | If the acquirer cannot reliably measure the fair value of contingent consideration at the acquisition date without undue cost or effort, it shall measure the contingent consideration at the acquisition date using the most likely … |
| 19.28 | Presentation | The acquirer shall apply Section 22 Liabilities and Equity to classify an obligation to pay contingent consideration that is a financial instrument as a financial liability or as equity. An acquirer shall classify a right to the … |
| 19.30 | R&M | If a party to a joint arrangement obtains control of a business that is a jointly controlled operation or a jointly controlled asset immediately before the acquisition date, the transaction is a business combination achieved in … |
| 19.31 | R&M | If an acquirer's initial accounting for a business combination is incomplete by the end of the reporting period in which the combination occurs, the acquirer shall recognise in its financial statements provisional amounts for the … |
| 19.32 | R&M | Acquisition-related costs are costs an acquirer incurs to effect a business combination. The acquirer shall account for acquisition-related costs separately from the business combination, as expenses in the periods in which the … |
| 19.33 | R&M | After initial recognition, an acquirer shall apply the applicable sections of this Standard to assets acquired, liabilities assumed or incurred, and equity instruments issued in a business combination. After initial recognition, … |
| 19.34 | R&M | After an acquirer initially recognises goodwill acquired in a business combination, the acquirer shall measure it at cost less accumulated amortisation and accumulated impairment losses. The acquirer shall apply: (a) paragraphs … |
| 19.35 | R&M | After an acquirer initially recognises a contingent liability in a business combination and until the liability is settled, cancelled or expires, it shall measure the liability at the higher of: (a) the amount that the acquirer … |
| 19.36 | R&M | Changes resulting from events after the acquisition date (such as meeting an earnings target, reaching a specified share price or reaching a milestone on a research and development project) are not measurement-period adjustments … |
| 19.37 | R&M | If the contingent consideration is: (a) classified as equity the acquirer shall not remeasure that contingent consideration and shall account for its subsequent settlement within equity. (b) other contingent consideration: (i) … |
| 19.38(a) | Disclosure | *For each business combination that occurs during the reporting period, the acquirer shall …* → the name and a description of the acquiree; |
| 19.38(b) | Disclosure | *For each business combination that occurs during the reporting period, the acquirer shall …* → the acquisition date; |
| 19.38(c) | Disclosure | *For each business combination that occurs during the reporting period, the acquirer shall …* → the percentage of voting-equity instruments acquired; |
| 19.38(d) | Disclosure | *For each business combination that occurs during the reporting period, the acquirer shall …* → the acquisition-date fair value of the total consideration transferred and a description of the components of that consideration (such as cash, equity instruments and debt instruments); |
| 19.38(e) | Disclosure | *For each business combination that occurs during the reporting period, the acquirer shall …* → for contingent consideration arrangements: (i) the amount recognised as of the acquisition date; (ii) a description of the arrangement and the basis for determining the amount of the payment; and (iii) if applicable, the fact … |
| 19.38(f) | Disclosure | *For each business combination that occurs during the reporting period, the acquirer shall …* → the amounts recognised at the acquisition date for each class of the acquiree's assets and liabilities; |
| 19.38(g) | Disclosure | *For each business combination that occurs during the reporting period, the acquirer shall …* → for a bargain purchase, the amount of any gain recognised in profit or loss in accordance with paragraph 19.23 and the line item in the statement of comprehensive income (and in the income statement, if presented) in which the … |
| 19.38(h) | Disclosure | *For each business combination that occurs during the reporting period, the acquirer shall …* → a qualitative description of the factors that make up the goodwill recognised-for example, expected synergies from combining operations of the acquiree and the acquirer, or intangible assets not recognised in accordance with … |
| 19.38(i) | Disclosure | *For each business combination that occurs during the reporting period, the acquirer shall …* → the information required by paragraph 21.15 for each contingent liability that the acquirer does not recognise in accordance with paragraph 19.18 (because its fair value cannot be measured reliably). For all business combinations |
| 19.39(a) | Disclosure | *An acquirer shall disclose the useful lives used to amortise goodwill and a …* → additional goodwill it recognised during the reporting period arising from new business combinations; |
| 19.39(b) | Disclosure | *An acquirer shall disclose the useful lives used to amortise goodwill and a …* → impairment losses it recognised during the reporting period in accordance with Section 27; |
| 19.39(c) | Disclosure | *An acquirer shall disclose the useful lives used to amortise goodwill and a …* → goodwill relating to previously acquired businesses the acquirer derecognised during the reporting period; and |
| 19.39(d) | Disclosure | *An acquirer shall disclose the useful lives used to amortise goodwill and a …* → other changes. |
| 19.41(a) | Disclosure | *For each reporting period after the acquisition date until the entity collects, sells or …* → any changes in the recognised amounts of contingent consideration, including any differences arising upon settlement; and |
| 19.41(b) | Disclosure | *For each reporting period after the acquisition date until the entity collects, sells or …* → the valuation techniques and key model inputs the acquirer used to measure contingent consideration. |

### Section 20 — Leases

| Ref | Area | Requirement (abridged) |
|---|---|---|
| 20.9 | R&M | At the commencement of the lease term, a lessee shall recognise its rights of use and obligations under finance leases as assets and liabilities in its statement of financial position at amounts equal to the fair value of the … |
| 20.11 | R&M | A lessee shall apportion minimum lease payments between the finance charge and the reduction of the outstanding liability using the effective interest method (see paragraphs 11.15-11.20). The lessee shall allocate the finance … |
| 20.12 | R&M | A lessee shall depreciate an asset leased under a finance lease in accordance with the relevant section of this Standard for that type of asset, for example, Section 17 Property, Plant and Equipment or Section 18. If there is no … |
| 20.15 | R&M | A lessee shall recognise lease payments under operating leases (excluding costs for services such as insurance and maintenance) as an expense over the lease term on a straight-line basis unless either: (a) another systematic … |
| 20.17 | R&M | A lessor shall recognise assets held under a finance lease in its statement of financial position and present them as a receivable at an amount equal to the net investment in the lease. The net investment in a lease is the … |
| 20.23(a) | Disclosure | *A lessor shall make the following disclosures for finance leases:* → a reconciliation between the gross investment in the lease at the end of the reporting period and the present value of minimum lease payments receivable at the end of the reporting period. In addition, a lessor shall disclose the … |
| 20.23(b) | Disclosure | *A lessor shall make the following disclosures for finance leases:* → unearned finance income. |
| 20.23(c) | Disclosure | *A lessor shall make the following disclosures for finance leases:* → the unguaranteed residual values accruing to the benefit of the lessor. |
| 20.23(d) | Disclosure | *A lessor shall make the following disclosures for finance leases:* → the accumulated allowance for uncollectable minimum lease payments receivable. |
| 20.23(e) | Disclosure | *A lessor shall make the following disclosures for finance leases:* → contingent rents recognised as income in the period. |
| 20.23(f) | Disclosure | *A lessor shall make the following disclosures for finance leases:* → a general description of the lessor's significant leasing arrangements, including, for example, information about contingent rent, renewal or purchase options and escalation clauses, subleases, and restrictions imposed by lease … |
| 20.24 | Presentation | A lessor shall present assets subject to operating leases in its statement of financial position according to the nature of the asset. |
| 20.25 | R&M | A lessor shall recognise lease income from operating leases (excluding amounts for services such as insurance and maintenance) in profit or loss on a straight-line basis over the lease term, unless either: (a) another systematic … |
| 20.26 | R&M | A lessor shall recognise as an expense costs, including depreciation, incurred in earning the lease income. The depreciation policy for depreciable leased assets shall be consistent with the lessor's normal depreciation policy … |
| 20.27 | R&M | A lessor shall add to the carrying amount of the leased asset any initial direct costs it incurs in negotiating and arranging an operating lease and shall recognise such costs as an expense over the lease term on the same basis … |
| 20.28 | R&M | To determine whether a leased asset has become impaired, a lessor shall apply Section 27. |
| 20.30(a) | Disclosure | *A lessor shall disclose the following for operating leases:* → the future minimum lease payments under non-cancellable operating leases for each of the following periods: (i) not later than one year; (ii) later than one year and not later than five years; and (iii) later than five years. |
| 20.30(b) | Disclosure | *A lessor shall disclose the following for operating leases:* → total contingent rents recognised as income; and |
| 20.30(c) | Disclosure | *A lessor shall disclose the following for operating leases:* → a general description of the lessor's significant leasing arrangements, including, for example, information about contingent rent, renewal or purchase options and escalation clauses and restrictions imposed by lease arrangements. |
| 20.33 | R&M | If a sale and leaseback transaction results in a finance lease, the seller-lessee shall not recognise immediately, as income, any excess of sales proceeds over the carrying amount. Instead, the seller-lessee shall defer such … |
| 20.34 | R&M | If a sale and leaseback transaction results in an operating lease, and it is clear that the transaction is established at fair value, the seller-lessee shall recognise any profit or loss immediately. If the sale price is below … |

### Section 21 — Provisions and Contingencies

| Ref | Area | Requirement (abridged) |
|---|---|---|
| 21.4 | R&M | An entity shall recognise a provision only when: (a) the entity has an obligation at the reporting date as a result of a past event; (b) it is probable (ie more likely than not) that the entity will be required to transfer … |
| 21.5 | R&M | The entity shall recognise the provision as a liability in the statement of financial position and shall recognise the amount of the provision as an expense, unless another section of this Standard requires the cost to be … |
| 21.7 | R&M | An entity shall measure a provision at the best estimate of the amount required to settle the obligation at the reporting date. The best estimate is the amount an entity would rationally pay to settle the obligation at the end of … |
| 21.9 | R&M | When some or all of the amount required to settle a provision may be reimbursed by another party (for example, through an insurance claim), the entity shall recognise the reimbursement as a separate asset only when it is … |
| 21.12 | R&M | A contingent liability is either a possible but uncertain obligation or a present obligation that is not recognised because it fails to meet one or both of the conditions (b) and (c) in paragraph 21.4. An entity shall not … |
| 21.13 | R&M | An entity shall not recognise a contingent asset as an asset. Disclosure of a contingent asset is required by paragraph 21.16 when an inflow of economic benefits is probable. However, when the flow of future economic benefits to … |
| 21.14(a) | Disclosure | *For each class of provision, an entity shall disclose all of the following:* → a reconciliation showing: (i) the carrying amount at the beginning and end of the period; (ii) additions during the period, including adjustments that result from changes in measuring the discounted amount; (iii) amounts charged … |
| 21.14(b) | Disclosure | *For each class of provision, an entity shall disclose all of the following:* → a brief description of the nature of the obligation and the expected amount and timing of any resulting payments; |
| 21.14(c) | Disclosure | *For each class of provision, an entity shall disclose all of the following:* → an indication of the uncertainties about the amount or timing of those outflows; and |
| 21.14(d) | Disclosure | *For each class of provision, an entity shall disclose all of the following:* → the amount of any expected reimbursement, stating the amount of any asset that has been recognised for that expected reimbursement. Comparative information for prior periods is not required. Disclosures about contingent … |
| 21.15(a) | Disclosure | *Unless the possibility of any outflow of resources in settlement is remote, an entity …* → an estimate of its financial effect, measured in accordance with paragraphs 21.7-21.11; |
| 21.15(b) | Disclosure | *Unless the possibility of any outflow of resources in settlement is remote, an entity …* → an indication of the uncertainties relating to the amount or timing of any outflow; and |
| 21.15(c) | Disclosure | *Unless the possibility of any outflow of resources in settlement is remote, an entity …* → the possibility of any reimbursement. If it is impracticable to make one or more of these disclosures, that fact shall be stated. Disclosures about contingent assets |
| 21.16 | Disclosure | If an inflow of economic benefits is probable (more likely than not) but not virtually certain, an entity shall disclose a description of the nature of the contingent assets at the end of the reporting period and, unless it would … |
| 21.17 | Disclosure | In extremely rare cases, disclosure of some or all of the information required by paragraphs 21.14-21.16 can be expected to prejudice seriously the position of the entity in a dispute with other parties on the subject matter of … |
| 21.18(a) | Disclosure | *An entity shall disclose for issued financial guarantee contracts in the scope of this …* → the nature and business purpose of these contracts; |
| 21.18(b) | Disclosure | *An entity shall disclose for issued financial guarantee contracts in the scope of this …* → an indication of the uncertainties relating to the amount or timing of any outflow of resources; and |
| 21.18(c) | Disclosure | *An entity shall disclose for issued financial guarantee contracts in the scope of this …* → the maximum amount the entity could be required to pay if the guarantees are called on. |

### Section 22 — Liabilities and Equity

| Ref | Area | Requirement (abridged) |
|---|---|---|
| 22.3A | Presentation | An entity shall classify a financial instrument as a financial liability or as equity in accordance with the substance of the contractual arrangement, not merely its legal form, and in accordance with the definitions of a … |
| 22.7(a) | Presentation | *An entity shall recognise the issue of shares or other equity instruments as equity when …* → if the equity instruments are issued before the entity receives the cash or other resources, the entity shall present the amount receivable as an offset to equity in its statement of financial position, unless local laws or … |
| 22.7(b) | Presentation | *An entity shall recognise the issue of shares or other equity instruments as equity when …* → if the entity receives the cash or other resources before the equity instruments are issued, and the entity cannot be required to repay the cash or other resources received, the entity shall recognise the corresponding increase … |
| 22.7(c) | Presentation | *An entity shall recognise the issue of shares or other equity instruments as equity when …* → to the extent that the equity instruments have been subscribed for but not issued, and the entity has not yet received the cash or other resources, the entity shall not recognise an increase in equity. |
| 22.8 | R&M | An entity shall measure equity instruments, other than those issued as part of a business combination or those accounted for in accordance with paragraphs 22.13-22.15B, at the fair value of the cash or other resources received or … |
| 22.9 | R&M | An entity shall account for the transaction costs of an equity transaction as a deduction from equity. Income tax relating to the transaction costs shall be accounted for in accordance with Section 29 Income Tax. |
| 22.11 | R&M | An entity shall apply the principles in paragraphs 22.7-22.10 to equity issued by means of sales of options, rights, warrants and similar equity instruments. |
| 22.13 | R&M | On issuing convertible debt or similar compound financial instruments that contain both a liability and an equity component, an entity shall allocate the proceeds between the liability component and the equity component. To make … |
| 22.15 | R&M | In periods after the instruments were issued, the entity shall account for the liability component as follows: (a) in accordance with Part I of Section 11 Financial Instruments if the liability component meets the conditions in … |
| 22.15A | R&M | An entity may renegotiate the terms of a financial liability with a creditor of the entity with the result that the entity extinguishes the liability fully or partially by issuing equity instruments to the creditor. Issuing … |
| 22.15B | R&M | If part of the consideration paid relates to a modification of the terms of the remaining part of the liability, the entity shall allocate the consideration paid between the part of the liability extinguished and the part that … |
| 22.16 | R&M | Treasury shares are the equity instruments of an entity that have been issued and subsequently reacquired by the entity. An entity shall deduct from equity the fair value of the consideration given for the treasury shares. The … |
| 22.18 | R&M | Sometimes an entity distributes assets other than cash to its owners ('non-cash distributions'). When an entity declares such a distribution and has an obligation to distribute non-cash assets to its owners, it shall recognise a … |
| 22.20 | Disclosure | If the fair value of the assets to be distributed as described in paragraphs 22.18-22.18A cannot be measured reliably without undue cost or effort, the entity shall disclose that fact and the reasons why a reliable fair value … |

### Section 23 — Revenue from Contracts with Customers

| Ref | Area | Requirement (abridged) |
|---|---|---|
| 23.2 | Scope | A contract with a customer might be partially within the scope of this section and partially within the scope of other sections listed in paragraph 23.1. If one of the other sections specifies how to initially measure any parts … |
| 23.3 | R&M | This section describes the revenue recognition model an entity shall apply to account for revenue from contracts with customers. Applying the model requires an entity to recognise an amount of revenue that depicts the goods or … |
| 23.4 | R&M | An entity shall apply this section consistently to contracts with similar characteristics and in similar circumstances. |
| 23.7 | R&M | An entity shall apply the revenue recognition model to account for a contract with a customer that is within the scope of this section only when all of the following criteria are met: (a) the parties to the contract have approved … |
| 23.10 | R&M | When a contract with a customer does not meet the criteria in paragraph 23.7, an entity shall recognise any consideration received from the customer as a liability. If the consideration is non-refundable, the entity shall … |
| 23.11 | R&M | Some contracts with customers might have no fixed duration or might automatically renew periodically. An entity shall apply this section to the duration of the contract in which the parties to the contract have present … |
| 23.12 | R&M | If a contract is modified, an entity shall account for the contract modification by applying paragraphs 23A.2-23A.4. Combination of contracts |
| 23.15 | R&M | An entity might transfer distinct goods or services to a customer that are essentially the same, but at discrete instances over time (that is, transfer a series of distinct goods or services). The entity shall account for those … |
| 23.24 | R&M | An entity shall determine the transaction price based on the terms of the contract and its usual business practices. When determining the transaction price, an entity shall assume the goods or services will be transferred to the … |
| 23.25 | R&M | If the consideration promised by the customer is in a form other than cash, an entity shall measure the non-cash consideration by applying paragraphs 23A.21-23A.23. Variable consideration |
| 23.27 | R&M | If the consideration includes a variable amount, an entity shall measure the amount of variable consideration to include in the transaction price by: (a) estimating the amount of variable consideration (see paragraphs … |
| 23.29 | R&M | An entity shall apply one method to estimate the amount of variable consideration consistently throughout the contract. |
| 23.31 | R&M | At the end of each reporting period, an entity shall update the amount of variable consideration included in the transaction price to reflect any relevant changes in circumstances. An entity shall allocate any subsequent changes … |
| 23.32 | R&M | An entity shall apply paragraph 23A.37 to account for consideration in the form of a royalty based on sales or use that is received in exchange for a licence of intellectual property. Refund liabilities |
| 23.33 | R&M | If an entity receives consideration from a customer and expects to refund some or all of that consideration to the customer, the entity shall recognise a refund liability. |
| 23.34 | R&M | The entity shall measure a refund liability by first applying paragraphs 23.27-23.30 to measure the amount of variable consideration to include in the transaction price. The entity shall then recognise as a refund liability the … |
| 23.35 | R&M | As an exception to paragraph 23.34, an entity shall apply paragraphs 23A.23-23A.27 to account for the variable consideration and refund liability relating to a sale with a right of return. Deferred payment |
| 23.36 | R&M | If payment is deferred and the arrangement constitutes a financing transaction (see paragraph 11.13B), an entity shall discount the promised amount of consideration at a market rate of interest for a similar debt instrument as … |
| 23.37 | Presentation | An entity shall present interest revenue separately from revenue from contracts with customers. |
| 23.39 | R&M | An entity shall allocate the transaction price to each promise it has identified in a contract in accordance with paragraphs 23.40-23.48. |
| 23.41 | R&M | Step 4 does not apply if a contract contains a single promise. However, an entity shall apply paragraph 23.48 if the entity accounts for a series of distinct goods or services as a single promise in accordance with paragraph … |
| 23.42 | R&M | At the start of a contract, an entity shall determine the stand-alone selling price of the distinct good or service (or distinct bundle of goods or services) in each promise in the contract. The entity shall allocate the … |
| 23.44 | R&M | If a stand-alone selling price is not directly observable, an entity shall estimate that price. An entity shall take into account all information that is reasonably available to it, including market conditions, entity-specific … |
| 23.47 | R&M | An entity shall allocate a discount or variable amount to each promise in the contract on a relative stand-alone selling price basis. However, if allocating the discount or variable amount on this basis does not represent the … |
| 23.48 | R&M | If an entity accounts for a series of distinct goods or services as a single promise in accordance with paragraph 23.15, it shall allocate a variable amount to each distinct good or service on a relative stand-alone selling price … |
| 23.49 | R&M | An entity shall recognise revenue when (or as) it fulfils a promise. An entity fulfils a promise by transferring control of the distinct good or service in the promise to the customer. |
| 23.52 | R&M | At the start of a contract, an entity shall determine whether it transfers control of the distinct good or service in a promise to a customer either: (a) over time, in which case the entity fulfils the promise (and recognises … |
| 23.53 | R&M | If a contract with a customer includes a licence that is distinct from the other goods or services in the contract, an entity shall apply paragraphs 23A.28-23A.36 to determine whether the promise to grant the licence is fulfilled … |
| 23.58 | R&M | For each promise an entity fulfils at a point in time, the entity shall recognise revenue at the specific point in time a customer obtains control of the distinct good or service. To identify this point in time, an entity shall … |
| 23.63 | R&M | An entity shall choose a method of measuring progress that depicts the entity's performance in transferring control of goods or services promised to the customer. An entity shall apply one method of measuring progress for each … |
| 23.64 | R&M | At the end of each reporting period, an entity shall remeasure its progress towards fulfilment of a promise to reflect any changes in circumstances. The entity shall account for changes to the measure of progress as a change in … |
| 23.68 | R&M | An entity shall recognise the costs of obtaining a contract with a customer as an expense when incurred, unless another section of this Standard requires the costs to be recognised as part of the cost of an asset. Costs to fulfil … |
| 23.69 | R&M | An entity shall account for the costs incurred to fulfil a contract with a customer in accordance with the relevant section of this Standard for those costs (for example, Section 13 Inventories, Section 17 Property, Plant and … |
| 23.70 | R&M | If the costs incurred to fulfil a contract are not within the scope of another section of this Standard, an entity shall recognise those costs as an asset if: (a) the costs relate directly to a contract (including future … |
| 23.73 | R&M | An entity shall recognise costs that relate to promises that are fulfilled (or partially fulfilled) as expenses when incurred, because those costs relate to past performance. Measurement after recognition |
| 23.74 | R&M | After initial recognition, an entity shall measure assets arising from the costs to fulfil a contract in accordance with paragraph 23.70 at cost less accumulated amortisation and any accumulated impairment losses: (a) an entity … |
| 23.77(a) | Presentation | *When either party to a contract has performed, an entity shall present the contract in …* → the entity's performance in transferring the goods or services promised in the contract to the customer; and |
| 23.77(b) | Presentation | *When either party to a contract has performed, an entity shall present the contract in …* → the customer's payment. |
| 23.78 | R&M | If an entity has received consideration (or has a trade receivable) before it transfers goods or services to the customer, the entity shall recognise a contract liability when the payment is made or due, whichever is earlier. A … |
| 23.79 | Presentation | If an entity transfers goods or services to a customer before the customer has paid (or before payment is due), the entity has a right to consideration. The entity shall present that right as a contract asset, excluding any … |
| 23.80 | Presentation | A receivable is an entity's unconditional right to consideration. A right to consideration is unconditional if only the passage of time is necessary before payment of that consideration becomes due. An entity can have an … |
| 23.83(a) | Disclosure | *If not otherwise separately presented or disclosed, an entity shall disclose:* → the opening and closing balances of trade receivables and contract assets separately; and |
| 23.83(b) | Disclosure | *If not otherwise separately presented or disclosed, an entity shall disclose:* → total impairment losses recognised on trade receivables and contract assets during the reporting period in accordance with Section 11. |
| 23.84(a) | Disclosure | *An entity shall disclose:* → the opening and closing balances of contract liabilities, if they are not otherwise separately presented or disclosed; and |
| 23.84(b) | Disclosure | *An entity shall disclose:* → revenue recognised in the reporting period that was included in the contract liability balance at the beginning of the period. |
| 23.85(a) | Disclosure | *An entity shall disclose information about its promises in contracts with customers, …* → when the entity typically fulfils its promises (for example, upon shipment, upon delivery, as services are rendered or upon completion of service); |
| 23.85(b) | Disclosure | *An entity shall disclose information about its promises in contracts with customers, …* → the main payment terms (for example, when payment is typically due, whether the contract constitutes a financing transaction and whether the consideration includes a variable amount); |
| 23.85(c) | Disclosure | *An entity shall disclose information about its promises in contracts with customers, …* → the nature of the goods or services that the entity has promised to transfer to customers, highlighting any promises to arrange for another party to transfer goods or services if the entity is an agent; |
| 23.85(d) | Disclosure | *An entity shall disclose information about its promises in contracts with customers, …* → obligations for returns, refunds and other similar obligations; and |
| 23.85(e) | Disclosure | *An entity shall disclose information about its promises in contracts with customers, …* → types of warranties and related obligations. |
| 23.86 | Disclosure | An entity shall disclose the methods it used to measure its progress towards fulfilment of promises fulfilled over time (for example, a description of the output or input methods it used and how those methods were applied). |
| 23.87(a) | Disclosure | *An entity shall disclose any judgements it has made that had a significant effect on the …* → determining the transaction price; and |
| 23.87(b) | Disclosure | *An entity shall disclose any judgements it has made that had a significant effect on the …* → allocating the transaction price to the promises identified in the contract. |
| 23.88 | Disclosure | An entity shall disclose the closing balance of assets arising from the costs to fulfil a contract (in accordance with paragraph 23.70) by main category of asset (for example, pre-contract costs and set-up costs). |
| 23.89 | Disclosure | If an entity chooses to use the option in paragraph 23.38 not to account for interest revenue separately from revenue from contracts with customers, it shall disclose that fact. |
| 23.90 | Disclosure | If an entity cannot account for an option to acquire additional goods or services that meets the criteria in paragraph 23A.9(a) as a separate promise without undue cost or effort, the entity shall disclose that fact and the … |
| 23A.3 | R&M | An entity shall account for a contract modification as a separate contract if: (a) additional goods or services are promised that are distinct from those in the original contract; and (b) the price of the original contract … |
| 23A.4 | R&M | For a contract modification that is not a separate contract in accordance with paragraph 23A.3, an entity shall account for the contract modification in one of two ways. If the goods or services an entity will transfer after the … |
| 23A.6 | R&M | If a customer can choose to buy the good or service either with or without a warranty, the warranty is distinct. In these circumstances, an entity shall account for the warranty as a separate promise and allocate a portion of the … |
| 23A.7 | R&M | If the customer has no option to buy the good or service without a warranty, the warranty is not distinct. In these circumstances, an entity shall account for the warranty in accordance with Section 21 Provisions and … |
| 23A.9 | R&M | An entity shall account for an option to acquire additional goods or services as a separate promise if: (a) the option gives the customer a material right that: (i) is only available to a customer that has entered into that … |
| 23A.11 | R&M | In accordance with Step 4, an entity shall allocate the transaction price to each promise in the contract on a relative stand-alone selling price basis. If the stand-alone selling price is not directly observable, an entity shall … |
| 23A.12 | R&M | An entity might grant a customer an option to renew a contract on similar terms. If the entity accounts for the renewal option as a separate promise, it shall allocate the transaction price to the option based on the total … |
| 23A.14 | R&M | If another party is involved in providing goods or services to a customer, an entity shall determine whether its promise is to provide the specified goods or services itself (that is, the entity is a principal) or to arrange for … |
| 23A.19 | R&M | A principal shall recognise as revenue the gross amount of the consideration it expects to be entitled to in exchange for transferring the specified good or service when (or as) it transfers control of that good or service to the … |
| 23A.20 | R&M | An entity that is an agent does not control the specified good or service provided by another party before that good or service is transferred to the customer. An agent shall recognise as revenue the amount of the fee or … |
| 23A.22 | R&M | An entity shall measure the non-cash consideration at fair value unless the fair value cannot be reasonably estimated. If the entity cannot reasonably estimate the fair value of the non-cash consideration, the entity shall … |
| 23A.24 | R&M | To account for products sold with a right of return, an entity shall recognise: (a) revenue for consideration received or receivable for transferred products the entity expects not to be returned; (b) a refund liability for … |
| 23A.26 | R&M | An entity shall initially measure the returns asset at the former carrying amount of the products the entity expects to be returned minus: (a) any expected costs to recover those products; and (b) any allowances for possible … |
| 23A.27 | R&M | At the end of each reporting period, an entity shall adjust the refund liability and returns asset for changes in its expectations about products to be returned. The entity shall recognise adjustments to: (a) the refund liability … |
| 23A.29 | R&M | If a contract with a customer includes a licence (or licences) and other goods or services, an entity shall apply Step 2 to identify each of the promises in the contract. If the licence is not distinct from the other goods or … |
| 23A.36 | R&M | If a promise to grant a licence is fulfilled over time, an entity shall apply paragraphs 23.62-23.67 to choose an appropriate method to measure its progress towards fulfilment of that promise. If a promise to grant a licence is … |
| 23A.37 | R&M | An entity might grant a licence of intellectual property in exchange for a royalty based on sales or use. If the licence is the sole or main item to which the royalty relates, the entity shall recognise revenue for such a royalty … |

### Section 24 — Government Grants

| Ref | Area | Requirement (abridged) |
|---|---|---|
| 24.4 | R&M | An entity shall recognise government grants as follows: (a) a grant that does not impose specified future performance conditions on the recipient is recognised in income when the grant proceeds are receivable; (b) a grant that … |
| 24.5 | R&M | An entity shall measure grants at the fair value of the asset received or receivable. |
| 24.6(a) | Disclosure | *An entity shall disclose the following:* → the nature and amounts of government grants recognised in the financial statements; |
| 24.6(b) | Disclosure | *An entity shall disclose the following:* → unfulfilled conditions and other contingencies attaching to government grants that have not been recognised in income; and |
| 24.6(c) | Disclosure | *An entity shall disclose the following:* → an indication of other forms of government assistance from which the entity has directly benefited. |

### Section 25 — Borrowing Costs

| Ref | Area | Requirement (abridged) |
|---|---|---|
| 25.2 | R&M | An entity shall recognise all borrowing costs as an expense in profit or loss in the period in which they are incurred. |

### Section 26 — Share-based Payment

| Ref | Area | Requirement (abridged) |
|---|---|---|
| 26.1D | Scope | This section uses the term 'fair value' in a way that differs in some respects from the definition in the Glossary and in other sections of this Standard. When applying this section, an entity shall apply the definition of fair … |
| 26.3 | R&M | An entity shall recognise the goods or services received or acquired in a share- based payment transaction when it obtains the goods or as the services are received. The entity shall recognise a corresponding increase in equity … |
| 26.4 | R&M | When the goods or services received or acquired in a share-based payment transaction do not qualify for recognition as assets, the entity shall recognise them as expenses. |
| 26.5 | R&M | If the share-based payments granted to an employee or another party (the counterparty) vest immediately, the counterparty is not required to complete a specified period of service before becoming unconditionally entitled to those … |
| 26.6 | R&M | If the share-based payments do not vest until the counterparty completes a specified period of service, the entity shall presume that the services to be rendered by the counterparty as consideration for those share-based payments … |
| 26.7 | R&M | For equity-settled share-based payment transactions, an entity shall measure the goods or services received, and the corresponding increase in equity, at the fair value of the goods or services received, unless that fair value … |
| 26.9 | R&M | A grant of equity instruments might be conditional upon satisfying specified vesting conditions related to service or performance. An example of a service condition is when a grant of shares or share options to an employee is … |
| 26.10 | R&M | An entity shall measure the fair value of shares (and the related goods or services received) using the following three-tier measurement hierarchy: (a) if an observable market price is available for the equity instruments … |
| 26.11 | R&M | An entity shall measure the fair value of share options and equity-settled share appreciation rights (and the related goods or services received) using the following three-tier measurement hierarchy: (a) if an observable market … |
| 26.13 | R&M | An entity shall account for a cancellation or settlement of an equity-settled share-based payment award as an acceleration of vesting, and therefore shall recognise immediately the amount that otherwise would have been recognised … |
| 26.14 | R&M | For cash-settled share-based payment transactions, an entity shall measure the goods or services acquired and the liability incurred at the fair value of the liability. Until the liability is settled, the entity shall remeasure … |
| 26.14A | R&M | A cash-settled share-based payment transaction might be conditional upon satisfying specified vesting conditions related to service or performance. Vesting conditions and conditions that are not vesting conditions (non-vesting … |
| 26.15 | R&M | Some share-based payment transactions give either the entity or the counterparty a choice of settling the transaction in cash (or other assets) or by transfer of equity instruments. In such a case, the entity shall account for … |
| 26.15C | R&M | For a transaction that satisfies the conditions in paragraph 26.15B, an entity shall account for the payment to the tax authority as a deduction from equity for the equity instruments withheld, except to the extent that the … |
| 26.17 | R&M | If the identifiable consideration received appears to be less than the fair value of the equity instruments granted or the liability incurred, this situation typically indicates that other consideration (ie unidentifiable goods … |
| 26.18(a) | Disclosure | *An entity shall disclose the following information about the nature and extent of …* → a description of each type of share-based payment arrangement that existed at any time during the period, including the general terms and conditions of each arrangement, such as vesting requirements, the maximum term of options … |
| 26.18(b) | Disclosure | *An entity shall disclose the following information about the nature and extent of …* → the number and weighted average exercise prices of share options for each of the following groups of options: (i) outstanding at the beginning of the period; (ii) granted during the period; (iii) forfeited during the period; (iv) … |
| 26.19 | Disclosure | For equity-settled share-based payment arrangements, an entity shall disclose information about how it measured the fair value of goods or services received or the value of the equity instruments granted. If a valuation … |
| 26.20 | Disclosure | For cash-settled share-based payment arrangements, an entity shall disclose information about how the liability was measured. |
| 26.21 | Disclosure | For share-based payment arrangements that were modified during the period, an entity shall disclose an explanation of those modifications. |
| 26.22 | Disclosure | If the entity is part of a group share-based payment plan, and it measures its share-based payment expense on the basis of a reasonable allocation of the expense recognised for the group, it shall disclose that fact and the basis … |
| 26.23(a) | Disclosure | *An entity shall disclose the following information about the effect of share-based …* → the total expense recognised in profit or loss for the period; and |
| 26.23(b) | Disclosure | *An entity shall disclose the following information about the effect of share-based …* → the total carrying amount at the end of the period for liabilities arising from share-based payment transactions. |

### Section 27 — Impairment of Assets

| Ref | Area | Requirement (abridged) |
|---|---|---|
| 27.6 | R&M | An entity shall recognise an impairment loss immediately in profit or loss, unless the asset is carried at a revalued amount in accordance with the revaluation model in Section 17 Property, Plant and Equipment. Any impairment … |
| 27.14 | R&M | When measuring fair value less costs to sell, an entity shall apply the guidance on measuring fair value in Section 12 Fair Value Measurement and deduct the costs of disposal. Value in use |
| 27.29 | R&M | For all assets other than goodwill, an entity shall assess at each reporting date whether there is any indication that an impairment loss recognised in prior periods may no longer exist or may have decreased. Indications that an … |
| 27.30 | R&M | When the prior impairment loss was based on the recoverable amount of the individual impaired asset, the following requirements apply: (a) the entity shall estimate the recoverable amount of the asset at the current reporting … |
| 27.31 | R&M | When the original impairment loss was based on the recoverable amount of the cash-generating unit to which the asset belongs, the following requirements apply: (a) the entity shall estimate the recoverable amount of that cash- … |
| 27.32(a) | Disclosure | *An entity shall disclose the following for each class of assets indicated in paragraph …* → the amount of impairment losses recognised in profit or loss during the period and the line item(s) in the statement of comprehensive income (and in the income statement, if presented) in which those impairment losses are … |
| 27.32(b) | Disclosure | *An entity shall disclose the following for each class of assets indicated in paragraph …* → the amount of reversals of impairment losses recognised in profit or loss during the period and the line item(s) in the statement of comprehensive income (and in the income statement, if presented) in which those impairment … |
| 27.33(a) | Disclosure | *An entity shall disclose the information required by paragraph 27.32 for each of the …* → inventories; |
| 27.33(b) | Disclosure | *An entity shall disclose the information required by paragraph 27.32 for each of the …* → property, plant and equipment (including investment property accounted for by the cost model); |
| 27.33(c) | Disclosure | *An entity shall disclose the information required by paragraph 27.32 for each of the …* → goodwill; |
| 27.33(d) | Disclosure | *An entity shall disclose the information required by paragraph 27.32 for each of the …* → intangible assets other than goodwill; |
| 27.33(e) | Disclosure | *An entity shall disclose the information required by paragraph 27.32 for each of the …* → investments in associates; and |
| 27.33(f) | Disclosure | *An entity shall disclose the information required by paragraph 27.32 for each of the …* → investments in jointly controlled entities. |

### Section 28 — Employee Benefits

| Ref | Area | Requirement (abridged) |
|---|---|---|
| 28.2 | Scope | Employee benefits also include share-based payment transactions by which employees receive equity instruments (such as shares or share options) or cash or other assets of the entity in amounts that are based on the price of the … |
| 28.3 | R&M | An entity shall recognise the cost of all employee benefits to which its employees have become entitled as a result of service rendered to the entity during the reporting period: (a) as a liability, after deducting amounts that … |
| 28.5 | R&M | When an employee has rendered service to an entity during the reporting period, the entity shall measure the amounts recognised in accordance with paragraph 28.3 at the undiscounted amount of short-term employee benefits expected … |
| 28.6 | Presentation | An entity may compensate employees for absence for various reasons including annual vacation leave and sick leave. Some short-term compensated absences accumulate-they can be carried forward and used in future periods if the … |
| 28.7 | R&M | An entity shall recognise the cost of other (non-accumulating) compensated absences when the absences occur. The entity shall measure the cost of non-accumulating compensated absences at the undiscounted amount of salaries and … |
| 28.8 | R&M | An entity shall recognise the expected cost of profit-sharing and bonus payments only when: (a) the entity has a present legal or constructive obligation to make such payments as a result of past events (this means that the … |
| 28.9 | R&M | Post-employment benefits include, for example: (a) retirement benefits, such as pensions; and (b) other post-employment benefits, such as post-employment life insurance and post-employment medical care. Arrangements whereby an … |
| 28.11 | R&M | Multi-employer plans and state plans are classified as defined contribution plans or defined benefit plans on the basis of the terms of the plan, including any constructive obligation that goes beyond the formal terms. However, … |
| 28.13 | R&M | An entity shall recognise the contribution payable for a period: (a) as a liability, after deducting any amount already paid. If contribution payments exceed the contribution due for service before the reporting date, an entity … |
| 28.14 | R&M | In applying the general recognition principle in paragraph 28.3 to defined benefit plans, an entity shall recognise: (a) a liability for its obligations under defined benefit plans net of plan assets-its 'defined benefit … |
| 28.15 | R&M | An entity shall measure a defined benefit liability for its obligations under defined benefit plans at the net total of the following amounts: (a) the present value of its obligations under defined benefit plans (its defined … |
| 28.17 | R&M | An entity shall determine the rate used to discount the future payments by reference to market yields at the reporting date on high quality corporate bonds. For currencies for which there is no deep market in such high quality … |
| 28.21 | R&M | If a defined benefit plan has been introduced or changed in the current period, the entity shall increase or decrease its defined benefit liability to reflect the change, and shall recognise the increase (decrease) as an expense … |
| 28.22 | R&M | If the present value of the defined benefit obligation at the reporting date is less than the fair value of plan assets at that date, the plan has a surplus. An entity shall recognise a plan surplus as a defined benefit plan … |
| 28.23 | R&M | An entity shall recognise the net change in its defined benefit liability during the period, other than a change attributable to benefits paid to employees during the period or to contributions from the employer, as the cost of … |
| 28.24 | R&M | An entity is required to recognise all actuarial gains and losses in the period in which they occur. An entity shall: (a) recognise all actuarial gains and losses in profit or loss; or (b) recognise all actuarial gains and losses … |
| 28.27 | R&M | If defined benefits are reduced for amounts that will be paid to employees under government-sponsored plans, an entity shall measure its defined benefit obligations on a basis that reflects the benefits payable under the … |
| 28.28 | R&M | If an entity is virtually certain that another party will reimburse some or all of the expenditure required to settle a defined benefit obligation, the entity shall recognise its right to reimbursement as a separate asset. The … |
| 28.30 | R&M | An entity shall recognise a liability for other long-term employee benefits measured at the net total of the following amounts: (a) the present value of the benefit obligation at the reporting date; minus (b) the fair value at … |
| 28.32 | R&M | Because termination benefits do not provide an entity with future economic benefits, an entity shall recognise them as an expense in profit or loss immediately. |
| 28.34 | R&M | An entity shall recognise a liability and an expense for termination benefits at the earlier of: (a) the date at which the entity can no longer withdraw the offer of those benefits; and (b) the date at which the entity recognises … |
| 28.36 | R&M | An entity shall measure termination benefits at the best estimate of the expenditure that would be required to settle the obligation at the reporting date. In the case of an offer made to encourage voluntary redundancy, the … |
| 28.40 | Disclosure | An entity shall disclose the amount recognised in profit or loss as an expense for defined contribution plans. If an entity treats a defined benefit multi-employer or state plan as a defined contribution plan because sufficient … |
| 28.41(a) | Disclosure | *Except for any defined benefit multi-employer or state plan that is accounted for as a …* → a general description of the type of plan, including the funding policy; |
| 28.41(b) | Disclosure | *Except for any defined benefit multi-employer or state plan that is accounted for as a …* → the entity's accounting policy for recognising actuarial gains and losses (either in profit or loss or as an item of other comprehensive income) and the amount of actuarial gains and losses recognised during the period; |
| 28.41(c) | Disclosure | *Except for any defined benefit multi-employer or state plan that is accounted for as a …* → if the entity applies paragraph 28.19 in measuring its defined benefit obligation, it shall disclose: (i) that fact; (ii) the reasons why using the projected unit credit method to measure its obligation and cost under defined … |
| 28.41(d) | Disclosure | *Except for any defined benefit multi-employer or state plan that is accounted for as a …* → the date of the most recent comprehensive actuarial valuation and, if it was not as of the reporting date, a description of the adjustments that were made to measure the defined benefit obligation at the reporting date; |
| 28.41(e) | Disclosure | *Except for any defined benefit multi-employer or state plan that is accounted for as a …* → a reconciliation of opening and closing balances of the defined benefit obligation showing separately, if applicable: (i) the change in the defined benefit liability arising from employee service rendered during the reporting … |
| 28.41(f) | Disclosure | *Except for any defined benefit multi-employer or state plan that is accounted for as a …* → a reconciliation of the opening and closing balances of the plan assets and of the opening and closing balances of any reimbursement right recognised as an asset, showing separately, when applicable: (i) contributions; (ii) … |
| 28.41(h) | Disclosure | *Except for any defined benefit multi-employer or state plan that is accounted for as a …* → for each major class of plan assets, which shall include, but is not limited to, equity instruments, debt instruments, property, and all other assets, the percentage or amount that each major class of plan assets constitutes of … |
| 28.41(i) | Disclosure | *Except for any defined benefit multi-employer or state plan that is accounted for as a …* → the amounts included in the fair value of plan assets for: (i) each class of the entity's own financial instruments; and (ii) any property occupied by, or other assets used by, the entity; |
| 28.41(k) | Disclosure | *Except for any defined benefit multi-employer or state plan that is accounted for as a …* → the principal actuarial assumptions used, including, when applicable: (i) the discount rates; (ii) the expected rates of return on any plan assets for the periods presented in the financial statements; (iii) the expected rates of … |
| 28.41(l) | Disclosure | *Except for any defined benefit multi-employer or state plan that is accounted for as a …* → the expected contributions to the defined benefit plan for the next annual reporting period. |
| 28.41C(a) | Disclosure | *If an entity participates in a defined benefit plan that is a group plan, it shall …* → the contractual agreement or stated policy for charging the net defined benefit cost, or the fact that there is no such policy; |
| 28.41C(b) | Disclosure | *If an entity participates in a defined benefit plan that is a group plan, it shall …* → the policy for determining the contribution to be paid by the entity; and |
| 28.41C(c) | Disclosure | *If an entity participates in a defined benefit plan that is a group plan, it shall …* → if the entity accounts for an allocation of the net defined benefit cost as noted in paragraph 28.38, all the information about the plan as a whole required by paragraph 28.41. |
| 28.42 | Disclosure | For each category of other long-term employee benefits that an entity provides to its employees, the entity shall disclose the nature of the benefit, the amount of its obligation and the extent of funding at the reporting date. … |
| 28.43 | Disclosure | For each category of termination benefits that an entity provides to its employees, the entity shall disclose the nature of the benefit, the amount of its obligation and the extent of funding at the reporting date. |

### Section 29 — Income Tax

| Ref | Area | Requirement (abridged) |
|---|---|---|
| 29.4 | R&M | An entity shall recognise a current tax liability for tax payable on taxable profit for the current and past periods. If the amount paid for the current and past periods exceeds the amount payable for those periods, the entity … |
| 29.5 | R&M | An entity shall recognise a current tax asset for the benefit of a tax loss that can be carried back to recover tax paid in a previous period. |
| 29.6 | R&M | An entity shall measure a current tax liability (asset) at the amount it expects to pay (recover) using the tax rates and laws that have been enacted or substantively enacted by the reporting date. An entity shall regard tax … |
| 29.8 | R&M | An entity shall recognise a deferred tax asset or liability for tax recoverable or payable in future periods as a result of past transactions or events. Such tax arises from the differences between the carrying amounts of the … |
| 29.25 | R&M | An entity shall recognise a deferred tax liability for all taxable temporary differences associated with investments in subsidiaries, branches and associates, and interests in joint arrangements, except to the extent that both of … |
| 29.26 | R&M | An entity shall recognise a deferred tax asset for all deductible temporary differences arising from investments in subsidiaries, branches and associates and interests in joint arrangements, only to the extent that it is probable … |
| 29.27 | R&M | An entity shall measure a deferred tax liability (asset) using the tax rates and tax laws that have been enacted or substantively enacted by the reporting date. An entity shall regard tax rates and tax laws as substantively … |
| 29.28 | R&M | When different tax rates apply to different levels of taxable profit, an entity shall measure deferred tax liabilities (assets) using the average enacted or substantively enacted rates that it expects to be applicable to the … |
| 29.32 | R&M | An entity shall not discount current or deferred tax assets and liabilities. |
| 29.33 | R&M | In some jurisdictions, income tax is payable at a higher or lower rate if part or all of the profit or retained earnings is paid out as a dividend to shareholders of the entity. In other jurisdictions, income tax may be … |
| 29.34B | R&M | An entity shall determine whether to consider each uncertain tax treatment separately or together with one or more other uncertain tax treatments based on which approach better predicts the resolution of the uncertainty. |
| 29.34C | R&M | An entity shall assume that a taxation authority will examine amounts it has a right to examine and have full knowledge of all related information when making those examinations. If an entity concludes that it is: (a) probable … |
| 29.35 | R&M | An entity shall recognise tax expense in the same component of total comprehensive income (ie continuing operations, discontinued operations or other comprehensive income) or equity as the transaction or other event that resulted … |
| 29.38 | Disclosure | An entity shall disclose information that enables users of its financial statements to evaluate the nature and financial effect of the current and deferred tax consequences of recognised transactions and other events (including … |
| 29.39(a) | Disclosure | *An entity shall disclose separately the major components of tax expense (income). Such …* → current tax expense (income); |
| 29.39(b) | Disclosure | *An entity shall disclose separately the major components of tax expense (income). Such …* → any adjustments recognised in the period for current tax of prior periods; |
| 29.39(c) | Disclosure | *An entity shall disclose separately the major components of tax expense (income). Such …* → the amount of deferred tax expense (income) relating to the origination and reversal of temporary differences; |
| 29.39(d) | Disclosure | *An entity shall disclose separately the major components of tax expense (income). Such …* → the amount of deferred tax expense (income) relating to changes in tax rates or the imposition of new taxes; |
| 29.39(e) | Disclosure | *An entity shall disclose separately the major components of tax expense (income). Such …* → the amount of the benefit arising from a previously unrecognised tax loss, tax credit or temporary difference of a prior period that is used to reduce tax expense; |
| 29.39(f) | Disclosure | *An entity shall disclose separately the major components of tax expense (income). Such …* → adjustments to deferred tax expense (income) arising from a change in the tax status of the entity or its shareholders; |
| 29.39(g) | Disclosure | *An entity shall disclose separately the major components of tax expense (income). Such …* → deferred tax expense (income) arising from the write-down, or reversal of a previous write-down, of a deferred tax asset in accordance with paragraph 29.31; and |
| 29.39(h) | Disclosure | *An entity shall disclose separately the major components of tax expense (income). Such …* → the amount of tax expense (income) relating to those changes in accounting policies and errors that are included in profit or loss in accordance with Section 10, because they cannot be accounted for retrospectively. |
| 29.40(a) | Disclosure | *An entity shall disclose the following separately:* → the aggregate current and deferred tax relating to items that are recognised as items of other comprehensive income. |
| 29.40(b) | Disclosure | *An entity shall disclose the following separately:* → the aggregate current and deferred tax relating to items that are charged or credited directly to equity. |
| 29.40(c) | Disclosure | *An entity shall disclose the following separately:* → an explanation of any significant differences between the tax expense (income) and accounting profit multiplied by the applicable tax rate. For example such differences may arise from transactions such as revenue that are exempt … |
| 29.40(d) | Disclosure | *An entity shall disclose the following separately:* → an explanation of changes in the applicable tax rate(s) compared with the previous reporting period. |
| 29.40(e) | Disclosure | *An entity shall disclose the following separately:* → for each type of temporary difference and for each type of unused tax losses and tax credits: (i) the amount of deferred tax liabilities and deferred tax assets at the end of the reporting period; and (ii) an analysis of the … |
| 29.40(f) | Disclosure | *An entity shall disclose the following separately:* → the amount (and expiry date, if any) of deductible temporary differences, unused tax losses and unused tax credits for which no deferred tax asset is recognised in the statement of financial position. |
| 29.40(g) | Disclosure | *An entity shall disclose the following separately:* → in the circumstances described in paragraph 29.33, an explanation of the nature of the potential income tax consequences that would result from the payment of dividends to its shareholders. |
| 29.42 | Disclosure | An entity within the scope of Pillar Two legislation shall disclose that it has applied the exception to recognising and disclosing information about deferred tax assets and liabilities related to Pillar Two income taxes (see … |
| 29.43 | Disclosure | An entity shall disclose separately its current tax expense (income) related to Pillar Two income taxes. |

### Section 30 — Foreign Currency Translation

| Ref | Area | Requirement (abridged) |
|---|---|---|
| 30.8A | R&M | When an entity pays or receives consideration in advance in a foreign currency, it recognises a non-monetary asset or non-monetary liability. The exchange rate to be used on the initial recognition of the related asset, expense … |
| 30.10 | R&M | An entity shall recognise, in profit or loss in the period in which they arise, exchange differences arising on the settlement of monetary items or on translating monetary items at rates different from those at which they were … |
| 30.11 | R&M | When another section of this Standard requires a gain or loss on a non-monetary item to be recognised in other comprehensive income, an entity shall recognise any exchange component of that gain or loss in other comprehensive … |
| 30.14 | R&M | When there is a change in an entity's functional currency, the entity shall apply the translation procedures applicable to the new functional currency prospectively from the date of the change. |
| 30.22 | R&M | In incorporating the assets, liabilities, income and expenses of a foreign operation with those of the reporting entity, the entity shall follow normal consolidation procedures, such as the elimination of intragroup balances and … |
| 30.25(a) | Disclosure | *An entity shall disclose the following:* → the amount of exchange differences recognised in profit or loss during the period, except for those arising on financial instruments measured at fair value through profit or loss in accordance with Section 11 Financial … |
| 30.25(b) | Disclosure | *An entity shall disclose the following:* → the amount of exchange differences arising during the period and classified in a separate component of equity at the end of the period. |
| 30.26 | Disclosure | An entity shall disclose the currency in which the financial statements are presented. When the presentation currency is different from the functional currency, an entity shall state that fact and shall disclose the functional … |
| 30.27 | Disclosure | When there is a change in the functional currency of either the reporting entity or a significant foreign operation, the entity shall disclose that fact and the reason for the change in functional currency. |
| 30.28(a) | Disclosure | *When an entity estimates a spot exchange rate because a currency is not exchangeable into …* → the currency and a description of the restrictions that result in that currency not being exchangeable into the other currency; |
| 30.28(b) | Disclosure | *When an entity estimates a spot exchange rate because a currency is not exchangeable into …* → a description of affected transactions; |
| 30.28(c) | Disclosure | *When an entity estimates a spot exchange rate because a currency is not exchangeable into …* → the carrying amount of affected assets and liabilities; |
| 30.28(d) | Disclosure | *When an entity estimates a spot exchange rate because a currency is not exchangeable into …* → the spot exchange rates used and whether those rates are: (i) observable exchange rates without adjustment (see paragraphs 30A.13-30A.17); or (ii) spot exchange rates estimated using another estimation technique (see paragraph … |
| 30.28(e) | Disclosure | *When an entity estimates a spot exchange rate because a currency is not exchangeable into …* → information about the estimation process, including qualitative and quantitative information about the inputs and assumptions used. |
| 30.29(a) | Disclosure | *When a foreign operation's functional currency is not exchangeable into an entity's …* → the name of the foreign operation; |
| 30.29(b) | Disclosure | *When a foreign operation's functional currency is not exchangeable into an entity's …* → whether the foreign operation is a subsidiary, jointly controlled entity, associate or branch; |
| 30.29(c) | Disclosure | *When a foreign operation's functional currency is not exchangeable into an entity's …* → the foreign operation's principal place of business; |
| 30.29(d) | Disclosure | *When a foreign operation's functional currency is not exchangeable into an entity's …* → summarised financial information about the foreign operation; and |
| 30.29(e) | Disclosure | *When a foreign operation's functional currency is not exchangeable into an entity's …* → the nature and terms of any contractual arrangements that could require the entity to provide financial support to the foreign operation, including events or circumstances that could expose the entity to a loss. |

### Section 31 — Hyperinflation

| Ref | Area | Requirement (abridged) |
|---|---|---|
| 31.15(a) | Disclosure | *An entity to which this section applies shall disclose the following:* → the fact that financial statements and other prior period data have been restated for changes in the general purchasing power of the functional currency; |
| 31.15(b) | Disclosure | *An entity to which this section applies shall disclose the following:* → the identity and level of the price index at the reporting date and changes during the current reporting period and the previous reporting period; and |
| 31.15(c) | Disclosure | *An entity to which this section applies shall disclose the following:* → amount of gain or loss on monetary items. |

### Section 32 — Events after the End of the Reporting Period

| Ref | Area | Requirement (abridged) |
|---|---|---|
| 32.8 | R&M | If an entity declares dividends to holders of its equity instruments after the end of the reporting period, the entity shall not recognise those dividends as a liability at the end of the reporting period. The amount of the … |
| 32.9 | Disclosure | An entity shall disclose the date when the financial statements were authorised for issue and who gave that authorisation. If the entity's owners or others have the power to amend the financial statements after issue, the entity … |
| 32.10(a) | Disclosure | *An entity shall disclose the following for each category of non-adjusting event after the …* → the nature of the event; and |
| 32.10(b) | Disclosure | *An entity shall disclose the following for each category of non-adjusting event after the …* → an estimate of its financial effect or a statement that such an estimate cannot be made. |

### Section 33 — Related Party Disclosures

| Ref | Area | Requirement (abridged) |
|---|---|---|
| 33.5 | Disclosure | Relationships between a parent and its subsidiaries shall be disclosed irrespective of whether there have been related party transactions. An entity shall disclose the name of its parent and, if different, the ultimate … |
| 33.7 | Disclosure | An entity shall disclose key management personnel compensation in total. |
| 33.9(a) | Disclosure | *If an entity has related party transactions, it shall disclose the nature of the related …* → the amount of the transactions; |
| 33.9(b) | Disclosure | *If an entity has related party transactions, it shall disclose the nature of the related …* → the amount of outstanding balances, including commitments and: (i) their terms and conditions, including whether they are secured and the nature of the consideration to be provided in settlement; and (ii) details of any … |
| 33.9(c) | Disclosure | *If an entity has related party transactions, it shall disclose the nature of the related …* → provisions for uncollectable receivables related to the amount of outstanding balances; and |
| 33.9(d) | Disclosure | *If an entity has related party transactions, it shall disclose the nature of the related …* → the expense recognised during the period in respect of bad or doubtful debts due from related parties. Such transactions could include purchases, sales or transfers of goods or services; leases; guarantees; and settlements by the … |
| 33.15(a) | Disclosure | *If a reporting entity applies the exemption in paragraph 33.11, it shall disclose the …* → the name of the government and the nature of its relationship with the reporting entity (that is, control, joint control or significant influence). |
| 33.15(b) | Disclosure | *If a reporting entity applies the exemption in paragraph 33.11, it shall disclose the …* → the nature and amount of each individually significant transaction. |
| 33.15(c) | Disclosure | *If a reporting entity applies the exemption in paragraph 33.11, it shall disclose the …* → for transactions that are collectively, but not individually, significant, a qualitative or quantitative indication of their extent. Types of transactions include those listed in paragraph 33.12. |

### Section 34 — Specialised Activities

| Ref | Area | Requirement (abridged) |
|---|---|---|
| 34.2 | R&M | An entity applying this Standard that is engaged in agricultural activity shall determine its accounting policy for each class of its biological assets, except for bearer plants that can be measured separately from the produce on … |
| 34.3 | R&M | An entity shall recognise a biological asset or agricultural produce when, and only when: (a) the entity controls the asset as a result of past events; (b) it is probable that future economic benefits associated with the asset … |
| 34.4 | R&M | An entity shall measure a biological asset on initial recognition and at each reporting date at its fair value less costs to sell. Changes in fair value less costs to sell shall be recognised in profit or loss. |
| 34.7(a) | Disclosure | *An entity shall disclose the following with respect to its biological assets measured at …* → a description of each class of its biological assets. |
| 34.7(c) | Disclosure | *An entity shall disclose the following with respect to its biological assets measured at …* → a reconciliation of changes in the carrying amount of biological assets between the beginning and the end of the current period. The reconciliation shall include: (i) the gain or loss arising from changes in fair value less costs … |
| 34.8 | R&M | The entity shall measure at cost less any accumulated depreciation and any accumulated impairment losses those biological assets whose fair value is not readily determinable without undue cost or effort. |
| 34.9 | R&M | The entity shall measure agricultural produce harvested from its biological assets at fair value less estimated costs to sell at the point of harvest. Such measurement is the cost at that date when applying Section 13 or other … |
| 34.10(a) | Disclosure | *An entity shall disclose the following with respect to its biological assets measured …* → a description of each class of its biological assets; |
| 34.10(b) | Disclosure | *An entity shall disclose the following with respect to its biological assets measured …* → an explanation of why fair value cannot be measured reliably without undue cost or effort; |
| 34.10(c) | Disclosure | *An entity shall disclose the following with respect to its biological assets measured …* → the depreciation method used; |
| 34.10(d) | Disclosure | *An entity shall disclose the following with respect to its biological assets measured …* → the useful lives or the depreciation rates used; and |
| 34.10(e) | Disclosure | *An entity shall disclose the following with respect to its biological assets measured …* → the gross carrying amount and the accumulated depreciation (aggregated with accumulated impairment losses) at the beginning and end of the period. |
| 34.11 | R&M | An entity using this Standard that is engaged in the exploration for, or evaluation of, mineral resources shall determine an accounting policy that specifies which expenditures are recognised as exploration and evaluation assets … |
| 34.11B | R&M | Exploration and evaluation assets shall be measured on initial recognition at cost. After initial recognition, an entity shall apply Section 17 and Section 18 Intangible Assets other than Goodwill to the exploration and … |
| 34.11C | R&M | Exploration and evaluation assets shall be assessed for impairment when facts and circumstances suggest that the carrying amount of an exploration and evaluation asset may exceed its recoverable amount. An entity shall measure, … |
| 34.11F | R&M | An entity shall determine an accounting policy for allocating exploration and evaluation assets to cash-generating units or groups of cash-generating units for the purpose of assessing such assets for impairment. |
| 34.14 | R&M | The operator shall recognise a financial asset to the extent that it has an unconditional contractual right to receive cash or another financial asset from or at the direction of the grantor for the construction services. The … |
| 34.15 | R&M | The operator shall recognise an intangible asset to the extent that it receives a right (a licence) to charge users of the public service. The operator shall initially measure the intangible asset at fair value. Thereafter, it … |
| 34.16 | R&M | The operator of a service concession arrangement shall recognise, measure and disclose revenue in accordance with Section 23 Revenue from Contracts with Customers for the services it performs. |

### Section 35 — Transition to the IFRS for SMEs Accounting Standard

| Ref | Area | Requirement (abridged) |
|---|---|---|
| 35.1 | Scope | This section applies to a first-time adopter of the IFRS for SMEs Accounting Standard, regardless of whether its previous accounting framework was full IFRS Accounting Standards or another set of generally accepted accounting … |
| 35.2 | Scope | An entity that has applied the IFRS for SMEs Accounting Standard in a previous reporting period, but whose most recent previous annual financial statements did not contain an explicit and unreserved statement of compliance with … |
| 35.3 | Transition | A first-time adopter of the IFRS for SMEs Accounting Standard shall apply this section in its first financial statements that conform to this Standard. |
| 35.4 | Transition | An entity's first financial statements that conform to this Standard are the first annual financial statements in which the entity makes an explicit and unreserved statement in those financial statements of compliance with the … |
| 35.5 | Transition | Paragraph 3.17 defines a complete set of financial statements. |
| 35.6 | Transition | Paragraph 3.14 requires an entity to disclose, in a complete set of financial statements, comparative information in respect of the previous comparable period for all monetary amounts presented in the financial statements, as … |
| 35.7 | Transition | Except as provided in paragraphs 35.9-35.11, an entity shall on its date of transition to the IFRS for SMEs Accounting Standard (ie the beginning of the earliest period presented): (a) recognise all assets and liabilities whose … |
| 35.8 | Transition | The accounting policies that an entity uses on adoption of this Standard may differ from those that it used for the same date using its previous financial reporting framework. The resulting adjustments arise from transactions, … |
| 35.9 | Transition | On first-time adoption of this Standard, an entity shall not retrospectively change the accounting that it followed under its previous financial reporting framework for any of the following transactions: (a) derecognition of … |
| 35.10 | Transition | An entity may use one or more of the following exemptions in preparing its first financial statements that conform to this Standard: (a) business combinations. A first-time adopter may elect not to apply Section 19 Business … |
| 35.11 | Transition | If it is impracticable for an entity to make one or more of the adjustments required by paragraph 35.7 at the date of transition, the entity shall apply paragraphs 35.7-35.10 for such adjustments in the earliest period for which … |
| 35.12 | Disclosure | An entity shall explain how the transition from its previous financial reporting framework to this Standard affected its reported financial position, financial performance and cash flows. |
| 35.12A(a) | Disclosure | *An entity that has applied the IFRS for SMEs Accounting Standard in a previous period, as …* → the reason it stopped applying the IFRS for SMEs Accounting Standard; |
| 35.12A(b) | Disclosure | *An entity that has applied the IFRS for SMEs Accounting Standard in a previous period, as …* → the reason it is resuming the application of the IFRS for SMEs Accounting Standard; and |
| 35.12A(c) | Disclosure | *An entity that has applied the IFRS for SMEs Accounting Standard in a previous period, as …* → whether it has applied this section or has applied the IFRS for SMEs Accounting Standard retrospectively in accordance with Section 10. Reconciliations |
| 35.13 | Transition | To comply with paragraph 35.12, an entity's first financial statements prepared using this Standard shall include: (a) a description of the nature of each change in accounting policy; (b) reconciliations of its equity determined … |
| 35.14 | Transition | If an entity becomes aware of errors made under its previous financial reporting framework, the reconciliations required by paragraph 35.13(b) and (c) shall, to the extent practicable, distinguish the correction of those errors … |
| 35.15 | Disclosure | If an entity did not present financial statements for previous periods, it shall disclose that fact in its first financial statements that conform to this Standard. |

### Section A — Appendix A - Effective date and transition

| Ref | Area | Requirement (abridged) |
|---|---|---|
| A1 | Disclosure | The third edition of the IFRS for SMEs Accounting Standard issued in February 2025 amended and revised sections of the IFRS for SMEs Accounting Standard. An entity shall apply the amended and revised sections for annual periods … |
| A2 | Transition | An entity shall retrospectively apply the amended and revised sections in the third edition of the IFRS for SMEs Accounting Standard in accordance with Section 10 Accounting Policies, Estimates and Errors except as stated in … |
| A3 | Transition | An entity need not disclose the information required by paragraph 10.13(b) for the current period. Furthermore, if an entity has not restated comparative information in accordance with the requirements in this appendix, the … |
| A4 | Transition | An entity shall prospectively apply the amended requirements in paragraphs 3.15A-3.16 from the date of initial application. Control model |
| A5 | Transition | If an entity applies the amended requirements in Section 9 Consolidated and Separate Financial Statements at the date of initial application and as a result consolidates an investee not previously consolidated, the entity shall: … |
| A6 | Transition | If an entity applies the amended requirements in Section 9 at the date of initial application and as a result no longer consolidates an investee it previously consolidated, the entity shall: (a) measure its interest in the … |
| A7 | Transition | If an entity applies paragraph A6(a) and applies the cost model in accordance with paragraph 14.4 or 15.9 and it is impracticable to determine the cost at the date of purchase, the entity is permitted to measure the cost (deemed … |
| A8 | Transition | At the date of initial application an entity is not required to make adjustments for previous accounting periods for its involvement in entities that: (a) would have been consolidated before the date of initial application and … |
| A9 | Transition | An entity applying the amended requirements in Section 9 is only required to disclose the amount of the adjustment for each line item affected in accordance with paragraph 10.13(b) for the annual period immediately preceding the … |
| A10 | Transition | An entity shall not restate the carrying amount of an investment in a former subsidiary if the entity lost control before the date of initial application. In addition, an entity shall not remeasure any gain or loss on the loss of … |
| A11 | Transition | An entity that previously applied the recognition and measurement requirements in IAS 39 Financial Instruments: Recognition and Measurement shall retrospectively apply the recognition and measurement requirements in Section 11 … |
| A12 | Transition | An entity shall prospectively apply the new Section 12 Fair Value Measurement and paragraph 2.89 from the date of initial application. |
| A13 | Transition | An entity is not required to apply the disclosure requirements in the new Section 12 to comparative information provided for periods before the date of initial application. |
| A14 | Transition | The new Section 12 amended paragraphs 11.14(c), 11.43, 11.57, 13.3(a), 14.10, 14.15, 15.15, 15.21, 16.7, 17.15B, 27.9(c), 27.14, 28.15(b), 34.4 and 34.6; deleted paragraphs 11.27-11.32, 12.12, 16.10(a), 17.33(c) and 34.7(b); and … |
| A15 | Transition | An entity applying the amended requirements in Section 15 Joint Arrangements is only required to disclose the amount of the adjustment for each line item affected in accordance with paragraph 10.13(b) for the annual period … |
| A16 | Transition | An entity shall prospectively apply the requirements in paragraph 16.3A for acquisitions of investment property on or after the date of initial application. Transfers of investment property |
| A17 | Transition | An entity shall apply the amended requirements in paragraph 16.9 to changes in use that occur on or after the date of initial application. At the date of initial application, an entity shall reassess the classification of … |
| A18 | Disclosure | If an entity reclassifies property at the date of initial application in accordance with paragraph 16.9, the entity shall disclose the amounts reclassified to or from investment property as part of the reconciliation between the … |
| A19 | Transition | An entity shall prospectively apply the amended requirements in paragraphs 17.21-17.22 from the date of initial application. Amortisation of intangible assets |
| A20 | Transition | An entity shall prospectively apply the requirements in paragraph 18.22A from the date of initial application. Business combinations and goodwill |
| A21 | Transition | An entity shall prospectively apply the revised Section 19 Business Combinations and Goodwill to business combinations for which the acquisition date is on or after the date of initial application and to asset acquisitions that … |
| A22 | Transition | An entity shall leave unadjusted on the date of initial application assets and liabilities from business combinations whose acquisition dates preceded the date of initial application. |
| A23 | Transition | The revised Section 19 amended paragraph 11.49(g), added paragraph 21.1(e) and deleted paragraph 22.2(c). An entity shall prospectively apply those amendments to business combinations for which the acquisition date is on or after … |
| A24 | Transition | An entity shall not adjust on the date of initial application any contingent consideration balances arising from business combinations whose acquisition dates preceded the date of initial application. An entity shall omit these … |
| A25 | Transition | If a business combination agreement provides for an adjustment to the cost of the combination contingent on future events and the adjustment is probable and can be measured reliably, the acquirer shall include the amount of that … |
| A26 | Transition | However, if a business combination agreement provides for such an adjustment and the adjustment either is not probable or cannot be measured reliably, an entity shall not include the adjustment in the cost of the combination at … |
| A27 | Transition | An entity shall apply the revised Section 23 Revenue from Contracts with Customers consistently to all contracts with customers either: (a) retrospectively in accordance with Section 10, subject to paragraphs A28-A31; or (b) … |
| A28 | Transition | If an entity retrospectively applies the revised Section 23, it is permitted to use one or more of three exemptions: (a) for completed contracts, an entity is not required to restate contracts that: (i) begin and end within the … |
| A29 | Transition | In paragraph A28(a)-(b), a completed contract is a contract for which the entity has transferred all goods or services identified in accordance with the requirements for accounting for revenue from contracts with customers that … |
| A30 | Disclosure | If an entity uses any of the exemptions in paragraph A28, the entity shall disclose that fact and apply that exemption consistently to all periods presented. |
| A31 | Disclosure | An entity retrospectively applying the revised Section 23 shall disclose the amount of the adjustment for each line item affected for the annual period immediately preceding the date of initial application and not earlier … |
| A32 | Transition | If an entity prospectively applies the revised Section 23, the entity shall apply the revised Section 23 to contracts that begin after the date of initial application. The entity shall not change its accounting policy for any … |
| A33 | Transition | The revised Section 23 amended paragraphs 4.11(b), 11.13, 11.14(a), 17.29, 18.1, 21.1(b), 27.1(f) and 34.16; deleted paragraphs 13.2(a), 13.14 and 21A.5; and added paragraphs 11.7(g), 11.7A, 11.13A-11.13B, 11.14A, 11.49(j), 11.55 … |
| A34(a) | Disclosure | *An entity prospectively applying the revised Section 23 shall disclose:* → that it has not restated comparative information as a result of applying the revised Section 23; and |
| A34(b) | Disclosure | *An entity prospectively applying the revised Section 23 shall disclose:* → the nature of the change in accounting policy. |
| A35 | Transition | An entity prospectively applying the revised Section 23 is not required to apply the disclosure requirements in that section to comparative information provided for periods before the date of initial application. |
| A36(a) | Disclosure | *In the financial statements for periods in which contracts that were in progress at the …* → the accounting policy for the recognition of revenue for those contracts, including the methods adopted to determine the stage of completion of transactions involving the rendering of services; and |
| A36(b) | Disclosure | *In the financial statements for periods in which contracts that were in progress at the …* → the revenue recognised in the current period from those contracts. Share-based payment Scope of Section 26 Share-based Payment |
| A37 | Transition | An entity shall prospectively apply paragraph 26.1C for transactions occurring on or after the date of initial application. Definition of fair value measurement |
| A38 | Transition | An entity shall prospectively apply paragraphs 26.1D-26.1E from the date of initial application. Definition of a vesting condition |
| A39 | Transition | An entity shall prospectively apply the amendments to paragraph 26.9, the amended definitions of 'vesting condition' and 'market vesting condition' and the new definitions of 'performance condition' and 'service condition' to … |
| A40 | Transition | An entity shall prospectively apply paragraphs 26.14A-26.14B to share-based payment transactions that are unvested at the date of initial application and to share-based payment transactions with a grant date on or after the date … |
| A41 | Transition | An entity shall apply the amended requirements in paragraphs 26.15-26.15D to share-based payment transactions that are unvested (or vested but unexercised) at the date of initial application and to share-based payment … |
| A42 | Transition | An entity shall retrospectively apply the amended requirements in paragraph 28.19 in accordance with Section 10. However, an entity is not required to adjust the carrying amount of assets in the scope of other sections of this … |
| A43 | Disclosure | An entity shall retrospectively apply the amended requirements in paragraphs 29.16A, 29.19(a) and 29.19A in accordance with Section 10. However, on initial application of these amended requirements, an entity is permitted to … |
| A44 | Transition | An entity shall retrospectively apply paragraphs 29.34A-29.34D either: (a) in accordance with Section 10; or (b) if (a) is impossible without the use of hindsight or is otherwise impracticable, recognising the cumulative effect … |
| A45 | Transition | An entity shall not retrospectively apply paragraph 30.5A. Instead, if the entity: (a) reports foreign currency transactions in its functional currency and concludes at the date of initial application that the two currencies are … |
| A46 | Transition | An entity is not required to apply the disclosure requirements in paragraph 30.28 to comparative information provided for periods before the date of initial application. Foreign currency transactions and advance consideration |
| A47 | Transition | An entity is permitted to prospectively apply paragraph 30.8A to all assets, expenses and income within the scope of paragraph 30.8A that were initially recognised on or after either: (a) the date of initial application; or (b) … |
| A48 | Transition | An entity that applies paragraph A47 shall apply paragraph 30.8A on initial application to assets, expenses and income initially recognised on or after the date in paragraph A47(a) or paragraph A47(b) for which the entity has … |
| A49 | Transition | An entity is permitted to measure a bearer plant at its fair value at the beginning of the earliest period presented and use that fair value as its deemed cost at that date. The entity shall recognise any difference between the … |
| A50 | Transition | Table A1 lists the paragraphs that have been added, amended and deleted in the third edition of the Standard compared to the second edition of the Standard. Table A1-Paragraphs added, amended and deleted in the third edition of … |
