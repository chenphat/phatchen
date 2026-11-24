# Section 30 — Foreign Currency Translation (third edition)

> Sources: Educational Module 30 (third edition); IFRS for SMEs Accounting Standard (2025) paragraphs 30.1–30.29, Appendix 30A.1–30A.18, transition A45–A48; cash-flow cross-references 7.11–7.13.
> Example numbers are as printed in the module. "CU1 = LCA2.3" means multiply the functional-currency amount by 2.3, which is the same direction as "USD1 = KHR x".

## When this section matters in an SME compilation
- Every compilation: the entity must **identify** (not choose) its functional currency (30.2) and state it in the notes (30.26).
- The ledger holds balances or transactions in any currency other than the functional currency: KHR cash and bank accounts, receivables, payables, tax liabilities, loans, advances paid or received (30.6–30.11).
- The financial statements are presented in a currency other than the functional currency, for example USD ledger with KHR statements required by local law (30.17–30.19; Example 28 and Example 38 are this exact fact pattern).
- Group or branch situations: foreign subsidiaries, associates, branches, intragroup loans, goodwill on a foreign acquisition (30.12–30.13, 30.22–30.23).
- A currency cannot be exchanged (restrictions) at the measurement date (30.5A, 30A).

## What changed in the third edition
- **Lack of exchangeability (new 30.5A and Appendix 30A):** how to assess whether a currency is exchangeable, and how to estimate a spot rate when it is not. New disclosures in 30.28–30.29.
- **Advance consideration (new 30.8A):** the transaction date for an asset, expense or income linked to an advance paid or received in foreign currency is the date the non-monetary prepayment asset or deferred-income liability was first recognised.
- Transition: 30.5A is **not** applied retrospectively. Remeasure at the estimated spot rate at the date of initial application, with the effect taken to opening retained earnings (transaction case) or to the cumulative translation differences (presentation/foreign operation case). No comparative 30.28 disclosures are required (A45–A46). For 30.8A, apply it fully retrospectively or prospectively from the date of initial application or from the start of the earliest comparative period (A47–A48). The third edition is effective for periods beginning on or after 1 January 2027, with early application permitted.

## Scope and key definitions
- **Scope (30.1):** foreign currency transactions, foreign operations, and translation into a presentation currency. Out of scope: FX derivatives such as forwards, futures and options, and hedge accounting. These fall under Part II of Section 11. Section 30 still applies to all financial instruments when the statements are translated into a presentation currency (module note).
- **Functional currency:** "the currency of the primary economic environment in which the entity operates" (30.2).
- **Presentation currency:** the currency in which the financial statements are presented. The entity can pick any currency (30.17).
- **Foreign currency transaction (30.6):** a transaction denominated in, or requiring settlement in, a foreign currency. Examples are buying or selling goods and services, borrowing or lending, and acquiring or disposing of assets or incurring or settling liabilities.
- **Monetary items:** units of currency held, and assets or liabilities to be received or paid in a fixed or determinable number of units of currency. Examples: cash, trade receivables and payables, loans, **tax assets and tax liabilities**, employee benefits payable in cash, cash-settled provisions, and dividends recognised. An equity investment is **not** monetary (module note to 30.9).
- **Non-monetary items:** prepaid goods and services such as prepaid rent, goodwill, intangibles, inventories, PPE, equity investments, shareholders' equity, provisions settled by delivering a non-monetary asset, and **advance consideration** (30.8A).
- **Closing rate:** the spot rate at the end of the reporting period. A predetermined contracted rate cannot be used instead (module note).
- **Foreign operation:** a subsidiary, associate, joint arrangement or branch whose activities are based or conducted in a country or currency other than those of the reporting entity (Glossary).

## Recognition
### Functional currency determination (30.2–30.5)
| Tier | Paragraph | Indicator | Weight |
|---|---|---|---|
| Primary | 30.3(a)(i) | Currency that mainly influences sales prices. This is often the currency in which prices are denominated and settled | Most important. Look at what drives pricing more than the currency of invoicing |
| Primary | 30.3(a)(ii) | Currency of the country whose competitive forces and regulations mainly determine sales prices | Most important |
| Primary | 30.3(b) | Currency that mainly influences labour, material and other costs | Most important |
| Supporting | 30.4(a) | Currency in which financing (debt and equity) is raised | Corroborative. Less significant if the entity is funded mainly by retained earnings |
| Supporting | 30.4(b) | Currency in which operating receipts are usually retained, that is, where excess working capital is held | Corroborative |
| Foreign operation only | 30.5(a)–(d) | Whether it is an extension of the parent or autonomous. Whether parent transactions are a high or low proportion of activity. Whether its cash flows directly affect the parent and are readily remittable. Whether it can service its debt without parent funds | Used together with 30.4. More likely to point to the parent's currency for subsidiaries and branches than for associates |

Decision flow (Figure 1):
1. If 30.3 clearly identifies one currency, that is the functional currency. Stop.
2. Otherwise weigh 30.3 together with 30.4, adding 30.5 for a foreign operation, with priority to 30.3.
3. If the result is still mixed, use judgement to choose the currency that most faithfully reflects the economic effects. A clear indicator that points to only one currency carries more weight than one that points weakly to several.

Module examples (condensed):
- **Ex 1:** Domestic manufacturer in Country A. Prices are set by local supply, demand and regulation, and inputs are local. Functional currency is LCA.
- **Ex 2:** Miner in Country A. Sales are in LCZ at prices set by global demand, with Country Z about 50% of demand. About 90% of costs (expatriate salaries, imported chemicals and machinery) are in LCZ. Functional currency is **LCZ, not the local currency**.
- **Ex 3:** Exporter from Country A to Country B. Prices are in LCB and set by Country B competition, but most costs are in LCA. The 30.3 indicators are mixed. Borrowings are in LCB and excess cash is held in LCB (30.4), so the functional currency is LCB. Machinery bought in LCC is ignored because the cost is insignificant.
- **Ex 4:** Autonomous subsidiary in Country Z. All 30.3 indicators point to LCZ, so 30.4–30.5 need not be considered.
- **Ex 5:** R&D subsidiary in Country B paid an LCA fee by the parent and funded by an LCA loan. 30.3(a) is not relevant and 30.3(b) alone is insufficient. 30.4(a) and 30.5 point to the parent's currency, so the functional currency is **LCA**.
- **Case study 1:** The parent (CU) and the autonomous Subsidiary A (LCA) are clear on 30.3. Subsidiary B imports the parent's shoes priced in CU, needs parent approval for price changes and remits all excess cash monthly. 30.3 is mixed, 30.4(b) and 30.5(a)–(c) point to CU, so B's functional currency is **CU**.
- **Branches:** a division operating in a different economic environment may have its own functional currency and is then translated under 30.22.

### Exchangeability (30.5A, 30A.2–30A.18)
- A currency is exchangeable if the entity is **able** to obtain the other currency within a normal administrative delay, through a market or mechanism that creates **enforceable** rights and obligations (30A.2). The test is applied at the measurement date and for a specified purpose.
- Ability matters, not intention (30A.6). Obtaining no more than an insignificant amount means the currency is not exchangeable (30A.3, 30A.11). Exchangeability can differ by direction (30A.4).
- Purpose (30A.9): for transactions, assume the purpose is to settle the individual items. For presentation translation, assume it is to realise net assets. For a foreign operation, assume it is to realise the net investment. Assess each purpose separately (30A.10).
- If the currency is not exchangeable, estimate the spot rate as the rate for an orderly transaction between market participants under prevailing conditions (30.5A). Options: (a) an observable rate without adjustment, either a rate for another purpose (30A.14–15) or the first rate available after exchangeability is restored (30A.16–17), or (b) another estimation technique (30A.18).
- Module examples: Ex 6 (a three-day administrative delay is normal, so the currency is exchangeable). Ex 7 (resellers' transactions are not enforceable, so not exchangeable). Ex 8 (PC is available only for food and medicine imports, so not exchangeable for realising the net investment). Ex 9 (a free-floating rate updated several times a day may be used). Ex 10 (a first subsequent rate four months later in a hyperinflationary economy does not meet the objective, so the rate must be adjusted).

## Measurement
### Initial: transactions in the functional currency (30.7–30.8A)
- Translate the foreign amount at the **spot rate at the transaction date** (30.7). The date of transaction is the date the item first qualifies for recognition, for example revenue on the date the Section 23 recognition conditions are met (30.8).
- An average rate for a week or month may approximate actual rates (30.8), but it is **inappropriate if rates fluctuate significantly**. Use actual rates for large one-off transactions. A weighted average is better if activity is uneven (module note).
- **Advance consideration (30.8A):** an advance paid or received is a **non-monetary** item, translated once and never retranslated. The related asset, expense or revenue uses the rate at the advance date for the portion covered by the advance, and each advance has its own date.

### Subsequent: each reporting date (30.9–30.11)
| Item | Rate | Where the difference goes |
|---|---|---|
| Foreign currency monetary items | Closing rate (30.9(a)) | Profit or loss, both on settlement and on retranslation (30.10) |
| Non-monetary items at historical cost | Rate at the transaction date (30.9(b)). No retranslation | None |
| Non-monetary items at fair value | Rate at the date fair value was determined (30.9(c)) | Follows the gain or loss: the FX component goes to OCI if the gain is in OCI, and to P&L if the gain is in P&L (30.11) |
| Impaired assets and inventory at the lower of two amounts | Compare the carrying amount at its historical rate with the recoverable amount (or selling price less costs) translated at the rate when that amount was determined, which is usually the closing rate | An impairment can arise in the functional currency even when there is none in the foreign currency, and vice versa (Case study 2) |

**Books kept in a foreign currency (module note to 30.6):** if the records are kept in a currency other than the functional currency, every amount is first remeasured into the functional currency under 30.6–30.11. Monetary items go at the closing rate, historical-cost non-monetary items at historical rates, and fair-value items at the rate on the fair-value date. Only then are the statements translated into a presentation currency.

### Net investment in a foreign operation (30.12–30.13)
- A monetary receivable from, or payable to, a foreign operation whose settlement is **neither planned nor likely in the foreseeable future** is part of the net investment. Trade receivables and payables never qualify, even if they are long-outstanding (30.12). A long-dated loan does not qualify merely because it is long unless renewal is intended. A loan repayable on demand can qualify. Reassess periodically.
- Exchange differences on such an item go to **P&L in the separate or individual statements** and to **OCI (equity) in the consolidated statements**. They are **not recycled on disposal** (30.13).

### Change in functional currency (30.14–30.16)
- A change happens only when the underlying transactions, events and conditions change (30.15). It is not a change in accounting policy, so apply it **prospectively** from the date of change (30.14). If the original identification was an error, restate retrospectively (module note).
- Translate all items at the rate on the date of change. The translated non-monetary amounts become their **historical cost** (30.16).
- **Ex 25:** A subsidiary stops manufacturing and distributes the parent's products, so its functional currency changes from FCU to CU.
- **Ex 26:** The functional currency changes from CU to FCU on 1 January 20X3. Situation 1: the presentation currency also changes to FCU, so comparatives are retranslated, because a change in presentation currency is a change in accounting policy and is applied retrospectively. Situation 2: the entity keeps presenting in CU, so 20X3 figures are translated under 30.18 and the 20X2 comparatives stay as reported.
- **Ex 27:** A machine bought for FCU100,000 at CU1 = FCU2 is recorded at CU50,000. After 3 years of depreciation it is carried at CU35,000. On the change date the rate is CU1 = FCU1.8, so the machine becomes FCU63,000 deemed cost. Over the remaining 7 years depreciation is FCU9,000 a year, giving a carrying amount of FCU54,000 at 31 December 20X3.

### Translation to a presentation currency (30.17–30.21)
30.18 applies where the functional currency is **not hyperinflationary**:
- (a) **assets and liabilities** for each statement of financial position presented, including comparatives, are translated at the **closing rate at that statement's date**;
- (b) **income and expenses** for each period presented, including comparatives, at the **rates at the transaction dates**;
- (c) **all resulting differences go to OCI** and accumulate in a component of equity. They are **never reclassified to profit or loss**.
- 30.19: an average rate for the period may be used for income and expenses for practical reasons. It is **inappropriate if rates fluctuate significantly**, in which case use narrower periods such as monthly averages (Ex 30) or actual rates, and use actual rates for large one-off items.
- 30.20: the difference arises from (a) translating income and expenses at transaction or average rates while assets and liabilities are at the closing rate, and (b) retranslating **opening net assets** at a closing rate that differs from the previous closing rate. It also includes movements in equity such as dividends and share issues (module note). The NCI share is allocated to NCI.
- **Equity components:** the section gives no explicit rule (module, Significant judgements). Retained earnings are built from translated profits and dividends. Share capital at the **historical rate** is the module's approach in every example. If the closing rate is used, the difference is still a 30.18(c) difference, and total equity is the same either way (Ex 28 note (a)).
- Translation **does not change the measurement** of underlying items (module note to 30.18).
- **Hyperinflation (30.21):** if the functional currency is hyperinflationary, translate under Section 31 instead. Section 31 applies to an entity whose *functional* currency is hyperinflationary (31.1), and its indicators are listed in 31.2.

### Foreign operations in group statements (30.22–30.23)
- Use normal consolidation procedures plus 30.17–30.21. Intragroup monetary items cannot be eliminated without showing their FX effect. The difference stays in P&L, or goes to OCI if 30.13 applies (30.22).
- Goodwill and fair-value adjustments are assets and liabilities **of the foreign operation**. They are carried in its functional currency and translated at the closing rate (30.23).
- If a subsidiary has a different reporting date, use the rate at the subsidiary's date and adjust for significant movements (module note, 9.16).

## Key judgements and estimates
| Judgement | What the module says | Example |
|---|---|---|
| Functional currency | It is determined, not chosen. Prioritise 30.3. Mixed results are likely when there are significant transactions in several currencies or international influence on prices or costs | Ex 3: LCB wins on 30.4 evidence |
| Use of average rates (30.8, 30.19) | Consider volatility and whether transactions arise evenly. Choose a year, month or week average, or actual rates | Ex 30: monthly rates |
| Net investment (30.12) | Whether settlement is neither planned nor likely | Ex 21 qualifies, Ex 22 (repayable in 5 years) does not |
| Rates for equity items | No explicit guidance. Choose a policy, for example historical rates for share capital | Ex 28 note (a) |
| Exchangeability (30.5A) | Ability, enforceability, purpose, insignificant amounts | Ex 7–Ex 10 |

## Presentation
- Exchange differences on monetary items are presented in P&L. The module's disclosure illustration places them in "other expenses" or finance costs (Case study 4).
- Translation differences are presented in OCI in a separate equity reserve, for example "foreign exchange reserve" or "foreign currency translation reserve" (Ex 28, Ex 36).
- **Cash flow statement:** foreign cash flows are recorded at the rates at the cash-flow dates (7.11–7.12). An approximate rate is allowed under 30.19 (7.12A). The effect of exchange-rate changes on cash is shown separately from operating, investing and financing activities (7.13). Case study 3 translates operating flows at the average rate and tax paid, PPE purchases and share issues at their actual rates, then shows a reconciling FX line.

## Disclosure requirements
| Para | Requirement | Typical wording or table |
|---|---|---|
| 30.24 | In a group, "functional currency" in 30.26–30.27 means the parent's | — |
| 30.25(a) | Amount of exchange differences recognised in P&L, excluding those on financial instruments at FVTPL under Section 11 | Ex 35: "Foreign exchange losses (gains) on trade receivables and payables (included in other expenses) 1,000 (2,000)". The FX component inside fair-value changes also counts: investment property loss CU63,095 (20X0: gain CU22,174) |
| 30.25(b) | Exchange differences arising in the period and classified in a separate component of equity | Ex 36: SOCE column "Foreign currency translation reserve" with the OCI line, for example (75) in 20X2 and 120 in 20X1 |
| 30.26 | Presentation currency. If it differs from the functional currency, state that fact, disclose the functional currency and the **reason** | Ex 38: "In accordance with local laws, ABC entity (functional currency FCU) presents its financial statements in CU, the local currency of Z land, its country of incorporation." |
| 30.27 | A change in functional currency of the entity or of a significant foreign operation, and the reason | Ex 39 |
| 30.28 | When a spot rate is estimated because of non-exchangeability: the currency and restrictions, affected transactions, carrying amounts, rates used and whether observable or estimated, and the estimation process with its inputs | Ex 40: an LC2,000,000 loan at the estimated CU1:LC1.2 is CU1,666,667 |
| 30.29 | For a foreign operation whose currency is not exchangeable: its name, type, principal place, summarised financial information, and support arrangements | Ex 40 |
| 8.4–8.6 | Accounting policy for foreign currency when activities are significant, and judgements | Ex 37 policy wording |

## Typical journal entries
| Event | Dr | Cr | Narrative |
|---|---|---|---|
| Credit sale in foreign currency (Ex 11: FCU30,000 at CU1 = FCU2) | Trade receivables 15,000 | Revenue 15,000 | Spot rate at the revenue date |
| Foreign loan drawn (Ex 12: FCU100,000 at 2.5) | Cash 40,000 | Loan payable 40,000 | Monetary item |
| PPE bought for foreign cash (Ex 13: FCU10,000 at FCU1 = CU2) | PPE 20,000 | Cash 20,000 | Never retranslated |
| Year-end payable retranslated (Ex 18) | P&L FX loss 10,000 | Trade payable 10,000 | FCU100,000 × (2.1 − 2.0) |
| Settlement (Ex 18) | Trade payable 210,000 | Cash 205,000 / P&L FX gain 5,000 | Settled at 2.05 |
| Advance received (Ex 17: FCU200 at 1.15 → LCU300 as printed) | Cash 300 | Contract liability 300 | Non-monetary, frozen |
| Consolidation: net investment loan (Ex 21, Ex 23) | OCI 476 | P&L FX 476 | Reclassify to OCI in consolidated statements only |

Ex 19 (FCU5,000 loan at 8%; CU1 = FCU5, then 5.1, then 4.8; average rates 5.06 and 4.9): initial entry Cash 1,000. 20X1 interest CU79 at the average rate. Payment CU78 with a CU1 gain. Year-end loan CU980, an FX gain of CU20. 20X2 interest CU82. Settlement: Dr loan 1,062, Dr FX loss 63, Cr cash 1,125.

## Worked examples from the module
**Ex 17 — advance and receivable (rates 1:1.15, 1:1.7, 1:1.9):** Advance FCU200 is LCU300. On 1 March, revenue for the first product is LCU725, made up of the LCU300 released from the contract liability and FCU250 at 1.7 = LCU425 receivable. On 1 June the receivable is LCU475, a gain of LCU50. Cash of FCU800 = LCU1,520 settles the LCU475 receivable and gives LCU1,045 revenue for the second product.

**Ex 20 — fair-value investment property (FCU500,000; rates 1.1, 1.05, 1.2, 1.1):** Cost CU454,545. 20X0: FCU520,000 ÷ 1.05 = CU495,238, a gain of CU40,693. 20X1: FCU540,000 ÷ 1.2 = CU450,000, a loss of CU45,238. Sale for FCU570,000 ÷ 1.1 = CU518,182, a gain of CU68,182. The FX element sits inside the P&L fair-value change and must be disclosed under 30.25(a). Ex 35 splits 20X1 into a CU17,857 fair-value gain and a CU(63,095) translation effect.

**Case study 2 — outlet and loan (FCU500,000; CU1 = FCU2, 2.4, 2.1, 2.5):** PPE CU250,000. 20X0: loan FX gain CU41,667, interest CU9,470 at 2.2, gain on interest CU789, depreciation CU10,417. The recoverable amount of FCU550,000 ÷ 2.4 = CU229,167 is below CU239,583, so an **impairment of CU10,416 arises only in CU**. 20X1: settlement costs CU250,000 cash with a CU31,174 loss, depreciation CU11,957, and a further impairment of CU97,210 to CU120,000.

**Ex 28 — stand-alone entity, functional CU, presentation LCA required by law.** Closing rates are CU1 = LCA2 at 20X1 and LCA2.3 at 20X2. Average rates are 2.2 for 20X2 and 1.9 for 20X1. Dividend CU3,000 at 2.25. Share capital at 1.8.

| SFP | 20X2 CU | Rate | 20X2 LCA | 20X1 CU | Rate | 20X1 LCA |
|---|---|---|---|---|---|---|
| PPE | 85,000 | 2.3 | 195,500 | 90,000 | 2 | 180,000 |
| Inventory | 8,000 | 2.3 | 18,400 | 3,000 | 2 | 6,000 |
| Receivables | 12,000 | 2.3 | 27,600 | 5,000 | 2 | 10,000 |
| Cash | 5,000 | 2.3 | 11,500 | 3,000 | 2 | 6,000 |
| **Total assets** | 110,000 | | 253,000 | 101,000 | | 202,000 |
| Payables / overdraft / loan | 5,500 / 5,000 / 50,000 | 2.3 | 12,650 / 11,500 / 115,000 | 3,000 / 10,000 / 50,000 | 2 | 6,000 / 20,000 / 100,000 |
| Share capital | 20,000 | 1.8 | 36,000 | 20,000 | 1.8 | 36,000 |
| Retained earnings | 29,500 | build-up | 62,350 | 18,000 | build-up | 37,200 |
| FX reserve | – | balance | 15,500 | – | balance | 2,800 |
| **Total equity** | 49,500 | | 113,850 | 38,000 | | 76,000 |

- Retained earnings: 37,200 + profit 31,900 (CU14,500 × 2.2) − dividends 6,750 (CU3,000 × 2.25) = 62,350. For 20X1: 22,000 + 15,200 (CU8,000 × 1.9) = 37,200.
- P&L at 2.2: revenue 321,451, cost of sales (250,732), PBT 41,789, tax (9,889), profit 31,900. OCI: translation difference **12,700** (15,500 − 2,800). TCI 44,600.
- SOCE reserve column: 2,800 + 12,700 = 15,500.

**Ex 31 — proof of the 12,700 (all items retranslated to the 2.3 closing rate):**
| Component | At closing | At rate used | Difference |
|---|---|---|---|
| Opening share capital CU20,000 (2.3 vs 2.0) | 46,000 | (40,000) | 6,000 |
| Opening retained earnings CU18,000 (2.3 vs 2.0) | 41,400 | (36,000) | 5,400 |
| Profit CU14,500 (2.3 vs 2.2 average) | 33,350 | (31,900) | 1,450 |
| Dividends CU3,000 (2.3 vs 2.25) | (6,900) | 6,750 | (150) |
| **Translation difference** | | | **12,700** |

In effect the opening net assets of CU38,000 × (2.3 − 2.0) = 11,400, plus profit × (closing − average), minus dividends × (closing − dividend-date rate).

**Ex 29 / Ex 32 — share issue of CU10,000 on 1 December 20X2 at 2.6:** Cash becomes 34,500, share capital 62,000 (36,000 + 26,000), FX reserve 12,500 and total equity 136,850. The OCI difference is **9,700**. Proof: share capital 69,000 − 40,000 − 26,000 = 3,000, plus 5,400, plus 1,450, less 150 = 9,700. Issuing capital at 2.6 when the closing rate is 2.3 reduces the difference by 3,000 compared with Ex 28. The module text calls these "exchange losses", but the reserve increases.

**Ex 30 — monthly average rates:** monthly revenue × monthly rate (January 7,000 × 2.10 through December 15,614 × 2.40) totals **LCA321,363.60**, against 321,451 using the annual 2.2. Apply the same approach to all income and expenses, and use actual rates for large one-off items.

**Case study 3 — new entity, functional CU, presentation ACU. Rates: 1 January 20X0 2.0, 31 December 20X0 1.95, 31 December 20X1 2.1. Averages 2.11 (20X0) and 1.79 (20X1). Tax paid at 2.05, PPE bought at 1.9.**
- SFP 20X1: assets 1,784 × 2.1 = 3,746.4. Liabilities 543.9. Share capital 1,000 × 2 = 2,000. Retained earnings 1,033.8 (620.3 + 413.5). FX reserve 168.7. SFP 20X0: total 2,964, retained earnings 620.3, reserve (97.0).
- P&L 20X1 at 1.79: profit 413.5, OCI 265.7, TCI 679.2. P&L 20X0 at 2.11: profit 620.3, OCI (97.0).
- Proof for 20X1: opening net assets 1,294 retranslated from 1.95 to 2.1 gives 2,717.4 − 2,523.3 = 194.1. Profit 231 retranslated from 1.79 to 2.1 gives 485.1 − 413.5 = 71.6. The total gain is 265.7. The module prints these differences with negative signs, but the reserve increases.
- Proof for 20X0: profit 573.3 at closing less 620.3 at average = (47.0). The module's table shows only that line yet totals (97.0). The missing (50.0) is the CU1,000 share capital issued at 2.0 and retranslated at 1.95 in the closing net assets.
- Cash flows: operating flows at 1.79, tax paid at 2.05 = (258.3), PPE (1,710.0) at 1.9, shares 2,000 at 2. The FX effect on cash is a balancing line: 153.9 in 20X1 (labelled "loss" but positive) and (111.0) in 20X0.

**Case study 4 — consolidation (CU1 = FCU2 at acquisition, 2.1 at loan date, 2.2 closing, 2.05 average):** The subsidiary's FCU1,450 PPE is CU659 at 2.2, share capital 1,600 at 2 is CU800, and its translation loss is CU84. That loss is opening net assets of FCU1,800 at 2 versus 2.2 (900 − 818 = 82) plus profit of FCU70 at 2.05 versus 2.2 (34 − 32 = 2). Goodwill of FCU180 at 2.2 is CU82, with amortisation of CU10 at the average rate and an FX loss of CU8. The CU4 FX loss on the net-investment loan is moved from finance costs to OCI. Consolidated OCI is **(96) = 84 + 8 + 4**, profit 144 and equity 1,848.

**Ex 33 / Ex 34 — goodwill:** goodwill FCU1,000 and machine FCU4,000 at 2.0 = CU2,500, retranslated at 2.1 = CU2,381, a loss of CU119 in OCI. With 10-year amortisation and an average rate of 2.05: goodwill CU429, PPE CU1,714, amortisation CU49, depreciation CU195, OCI loss CU113.

## Review points and common errors
- The functional currency is assumed to be the local currency or the presentation currency. It must be evidenced against 30.3–30.5 (Ex 2, Ex 5).
- A contracted rate is used instead of the closing rate for monetary items (prohibited).
- Non-monetary items are retranslated: PPE, inventory, prepayments, **advances paid or received (30.8A)** and equity investments at cost.
- **Comparatives are retranslated at the current closing rate.** 30.18(a) uses each statement's own closing rate. The 20X1 column in Ex 28 stays at 2.0.
- P&L is translated at the closing rate, or at an annual average despite volatile rates (30.19). Large one-off items should use actual rates.
- The translation difference is booked to P&L, or recycled. It must be in OCI and never reclassified (30.18(c)).
- The reserve is "plugged" without a proof. Reconcile it using Ex 31 or Case study 3, including dividends and share issues.
- The FX component inside fair-value or impairment movements is omitted from the 30.25(a) disclosure (Ex 35).
- The 30.26 statement of the functional currency and the reason for a different presentation currency is missing.
- Net-investment treatment is applied to trade balances, or the P&L to OCI reclassification is missed on consolidation (Ex 21).
- An impairment that exists only in the functional currency is missed (Case study 2).

## USD functional currency / KHR presentation currency
This is the fact pattern of Ex 28 (functional CU, presentation in the local currency because of local law) and Ex 38 (wording of the reason). Quote rates as KHR per USD1 and multiply.

**Step 1 — Keep the ledger in USD.** Confirm and document that USD is the functional currency on 30.3: the currency that mainly influences prices and costs. Support this with 30.4: financing and retained cash. If the books are kept in KHR, first remeasure every balance into USD under 30.6–30.11 (module note to 30.6) before any translation.

**Step 2 — Remeasure KHR monetary items in the USD ledger.**
- Record each KHR transaction at the KHR/USD spot rate on its recognition date, or at a weekly or monthly average if rates are stable (30.7–30.8).
- At year-end, retranslate KHR cash, bank balances, receivables, payables, loans, accruals and **KHR tax liabilities**, which are monetary per the module, at the closing rate. Take the difference to **P&L** (30.9(a), 30.10).
- Do **not** retranslate KHR-denominated non-monetary items such as PPE, inventory, prepaid rent or advances (30.9(b), 30.8A).
- Accumulate the P&L FX total for the 30.25(a) note.

**Step 3 — Translate the USD statements into KHR (30.18–30.19).**
| Item | KHR/USD rate |
|---|---|
| All assets and liabilities, including deferred tax | Closing rate at that SFP date. The comparative SFP uses the prior closing rate |
| Income and expenses, including tax expense (Ex 28 used the average 2.2) | Transaction-date rate, or a period or monthly average if not significantly volatile (Ex 30). Actual rates for one-off items |
| Share capital | Historical rate at issue, which is the module's convention and should be applied consistently (Ex 28 note (a)) |
| Opening retained earnings | Prior-year closing KHR balance carried forward. Never retranslated |
| Dividends | Rate at the date paid or declared (Ex 28: 2.25) |
| Cash flows | Rates at the cash-flow dates or an approximation, with a separate line for the FX effect on cash (7.11–7.13; Case study 3) |

**Step 4 — Translation reserve.** The balancing figure is the 30.18(c) difference. Present it in **OCI** under "Exchange differences on translating from USD to KHR" and accumulate it in a **translation reserve** column in the SOCE. It is **never recycled** to P&L (30.18(c)). It is a separate item from the Step 2 P&L differences, which arise in USD and are already in profit.

**Step 5 — Prove the difference (30.20; Ex 31 method).**
```
Translation difference for the year (KHR) =
   Opening net assets USD  x (closing rate - opening rate)
 + Profit for year USD     x closing rate - profit translated in KHR (at transaction/average rates)
 - [Dividends USD          x closing rate - dividends translated at date-paid rate]
 + [Capital issued USD     x closing rate - capital translated at issue rate]
 + any other equity movement on the same basis
Check: opening reserve + this difference = closing reserve = KHR net assets - share capital (hist.) - retained earnings (KHR build-up)
```
Ex 28 ties out: 11,400 + 1,450 − 150 = 12,700. If the proof fails, look first for a comparative that was retranslated, a dividend translated at the closing rate, or share capital translated at the closing rate in one year and at the historical rate in another.

**Other points:**
- 30.26 note: "The functional currency is the US dollar. The financial statements are presented in Khmer riel [reason, for example local legal requirement — confirm]."
- Hyperinflation: 30.18 applies because the test is whether the **functional** currency (USD) is hyperinflationary. Section 31 is triggered only for a hyperinflationary functional currency (30.21, 31.1).
- Exchangeability: KHR to USD is assessed under 30A only if restrictions exist at the measurement date. Otherwise no estimate is needed (30A notes).

## Differences from full IFRS Accounting Standards
- Section 30 is drafted in simpler language with less application guidance than IAS 21 and IFRIC 22.
- **No recycling:** cumulative exchange differences on a foreign operation that were recognised in OCI are never reclassified to P&L on disposal. IAS 21 requires reclassification.
- Disclosures: IAS 21 paragraphs 52, 53, 54, A20 and A19(a)–(e) map to 30.25, 30.26, 30.27, 30.29 and 30.28. IAS 21 A19(f), 57, 57A, 57B and A18 have no equivalent in Section 30.
