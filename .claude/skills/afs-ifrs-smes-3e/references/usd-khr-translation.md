# USD functional currency, KHR presentation currency

> Sources: IFRS for SMEs Accounting Standard, third edition (2025), Section 30 paragraphs 30.2–30.27, Section 7
> paragraph 7.13, Section 3 paragraph 3.23(d). Read `sections/s30-foreign-currency.md` for the Educational Module 30
> worked examples. Exchange rates in this file are illustrative only.

## Contents
1. Two different currency steps — do not merge them
2. Step A: confirm USD is the functional currency (30.2–30.5)
3. Step B: foreign-currency transactions in the USD ledger (30.6–30.13)
4. Step C: translate the USD statements into KHR (30.17–30.20)
5. Equity components and the translation reserve
6. Cash flow statement (7.13 and 30.18)
7. Proof of the translation difference (30.20)
8. Rates: which ones, from where, documentation
9. Worked example
10. Disclosures (30.25–30.26, 3.23)
11. Review points and common errors

---

## 1. Two different currency steps — do not merge them

| Step | What | Where the difference goes | Paragraphs |
|---|---|---|---|
| **B. Remeasurement** inside the ledger | KHR (and any other non-USD) monetary items — KHR cash, KHR receivables/payables, KHR tax payable — remeasured to USD at the closing rate | **Profit or loss** (exchange gain/loss, 30.10) | 30.6–30.13 |
| **C. Translation** of the finished USD statements into KHR | Every line of the USD SoFP, SoCI, SoCE and SoCF expressed in KHR | **Other comprehensive income**, accumulated in a separate component of equity; **never reclassified** to profit or loss (30.18(c)) | 30.17–30.20 |

Step B changes the USD numbers. Step C never changes the USD numbers; it only re-expresses them. A KHR bank
balance therefore produces a P&L exchange difference in step B (because the functional currency is USD), and the
whole balance sheet then produces an OCI translation difference in step C (because the presentation currency is
KHR). Confusing the two is the commonest error in USD-functional Cambodian entities.

## 2. Step A: confirm USD is the functional currency

Functional currency is a matter of fact, not choice (30.2: "the currency of the primary economic environment in
which the entity operates"). Document the assessment in the working papers:

| Indicator | Paragraph | Evidence to record | Points to USD when… |
|---|---|---|---|
| Currency that mainly influences sales prices; country whose competitive forces and regulation set prices | 30.3(a) | price lists, contracts, invoices | sales are priced and settled in USD |
| Currency that mainly influences labour, material and other costs | 30.3(b) | payroll currency, supplier invoices, rent contracts | salaries, purchases and rent are set and paid in USD |
| Currency of financing (debt and equity) | 30.4(a) | loan agreements, share capital | loans and capital are denominated in USD |
| Currency in which operating receipts are retained | 30.4(b) | bank account mix | most cash is held in USD accounts |

The primary indicators (30.3) take precedence; 30.4 provides supporting evidence. If indicators are mixed, management
uses judgement and should disclose it under 8.6. A change of functional currency is applied prospectively from the
date of change at the rate on that date (30.14–30.16).

Hyperinflation: if the functional currency were the currency of a hyperinflationary economy, Section 31 would apply
before translation (30.21). With USD as functional currency this is not expected; record the conclusion.

## 3. Step B: foreign-currency transactions in the USD ledger

- **Initial recognition** at the spot rate on the transaction date (30.7); a weekly or monthly average may be used
  for practicality unless rates fluctuate significantly (30.8).
- **At each reporting date** (30.9):
  - monetary items (cash, receivables, payables, loans, tax payable, employee obligations payable in KHR) → closing rate;
  - non-monetary items at historical cost (prepayments, inventories, PPE, advance payments, contract liabilities for
    goods/services still to be delivered) → rate at the transaction date, **not** retranslated;
  - non-monetary items at fair value → rate at the date fair value was determined.
- **Exchange differences** on settlement or retranslation of monetary items → profit or loss (30.10). When a gain or
  loss on a non-monetary item is recognised in OCI (e.g. revaluation), its exchange component also goes to OCI (30.11).

Typical journal at year end (USD ledger, KHR bank balance of KHR 40,100,000; book value USD 10,000 at 4,010; closing
rate 4,005 → USD 10,012.48):

| Account | Dr USD | Cr USD |
|---|---|---|
| Bank – KHR account | 12.48 | |
| Foreign exchange gain (P&L) | | 12.48 |

## 4. Step C: translate the USD statements into KHR (30.18)

| Item | Rate | Paragraph |
|---|---|---|
| Assets and liabilities of each SoFP presented, including comparatives | closing rate at that SoFP date | 30.18(a) |
| Income and expenses of each period presented, including comparatives | rates at the transaction dates; an average rate is acceptable unless rates fluctuate significantly | 30.18(b), 30.19 |
| OCI items (e.g. revaluation surplus, its deferred tax) | same as income and expense: transaction-date or average rate | 30.18(b) |
| Share capital and share premium | historical rate at the date of issue (Standard silent on equity rates; this is the usual practice — document the policy) | — |
| Retained earnings | opening balance carried forward in KHR; each year's profit at that year's translated amount; dividends at the declaration-date rate | — |
| Resulting exchange difference | OCI → foreign currency translation reserve; not reclassified to P&L | 30.18(c), 30.20 |

Because the comparatives are also translated at their own closing and average rates (30.18(a)–(b) say "including
comparatives"), the prior-year KHR figures are **not** last year's figures re-expressed at this year's rate.

When an item is a one-off with a known date (share issue, land revaluation at year end, dividend declaration, large
asset purchase), using the actual transaction-date rate is more faithful than the average; decide and document the policy.

## 5. Equity components and the translation reserve

Translation reserve (KHR) = translated net assets at the closing rate − (share capital + share premium + other
reserves + retained earnings, each at its historical or accumulated KHR amount).

The reserve is presented as a separate component of equity in the SoFP and in the SoCE. The year's movement is the
"Exchange differences on translation to the presentation currency" line in OCI, under items that will not be
reclassified subsequently to profit or loss (5.5(g)(i); 30.18(c)).

## 6. Cash flow statement

- Translate cash flows at the rates at the dates of the cash flows or an approximating average (consistent with
  30.18(b); 7.11 for foreign-currency cash flows in the functional currency).
- Opening cash at the previous closing rate; closing cash at the current closing rate.
- The difference is presented as **"Effect of exchange rate changes on cash and cash equivalents"**, separately from
  operating, investing and financing activities (7.13). It has two parts: step B remeasurement of KHR cash already in
  the USD ledger (non-cash, reverse it out of profit), and step C translation effects.

## 7. Proof of the translation difference (30.20)

30.20 identifies the two sources of the exchange difference:

```
Translation difference for the year (KHR)
  = Opening net assets (USD) × (closing rate − previous closing rate)          30.20(b)
  + [Profit or loss + OCI before translation (USD)] × (closing rate − average rate)   30.20(a)
  + Owner transactions (USD) × (closing rate − rate used for them)
```

Compute the proof independently and agree it to the OCI line. A difference means a line was translated at the wrong rate.

## 8. Rates: which ones, from where, documentation

- Use a single, documented rate source for all three rates (closing, average, historical) and both years. For a
  Cambodian entity the usual sources are the official exchange rates published by the National Bank of Cambodia or
  the rates prescribed by the tax authority for tax reporting. Confirm which source local law requires — this skill's
  sources (IFRS Foundation material) are silent on local rate sources.
- Average rate: simple average of daily or monthly rates is acceptable only if rates did not fluctuate significantly
  (30.19). Record the range (min/max) to evidence that.
- Keep a rates table in the working papers: closing CY, average CY, closing PY, average PY, opening rate of the
  earliest period (needed to translate opening equity), historical rates for share capital issues.
- KHR amounts are normally presented in thousands (KHR'000). State the rounding level (3.23(e)).

## 9. Worked example (illustrative rates)

USD figures (functional currency): opening net assets 617,342; profit 90,353; revaluation OCI net of tax 40,000;
share issue 60,000; dividends 30,000; closing net assets 777,695.
Rates (KHR per USD): previous closing 4,010; closing 4,005; average 4,012.

| Line | USD | Rate | KHR'000 |
|---|---|---|---|
| Opening net assets | 617,342 | 4,010 (as translated last year) | 2,475,541 |
| Profit for the year | 90,353 | 4,012 | 362,496 |
| Revaluation OCI | 40,000 | 4,012 | 160,480 |
| Share issue | 60,000 | 4,012 | 240,720 |
| Dividends | (30,000) | 4,012 | (120,360) |
| Translation difference (balancing) | — | — | (4,209) |
| Closing net assets | 777,695 | 4,005 | 3,114,668 |

Proof: 617,342 × (4,005 − 4,010) = −3,087; (90,353 + 40,000 + 60,000 − 30,000) × (4,005 − 4,012) = −1,122; total
−4,209 thousand-KHR → agrees.

## 10. Disclosures

| Paragraph | Requirement | Wording pattern |
|---|---|---|
| 30.26 | Presentation currency; the fact that it differs from the functional currency; the functional currency; the reason | "The functional currency of the Company is the United States dollar (USD)… The financial statements are presented in Khmer Riel (KHR) because [reason]." |
| 30.25(a) | Exchange differences recognised in profit or loss (step B) | amount, in the finance costs / other gains note |
| 30.25(b) | Exchange differences classified in a separate component of equity (step C) | translation reserve movement and closing balance |
| 3.23(d)–(e) | Presentation currency and rounding level displayed prominently | heading of every statement: "KHR'000" |
| 8.5 | Translation policy as material accounting policy information | rates used for each class of item |
| 8.6 | Judgement on functional currency if indicators were mixed | — |

Present USD comparative columns for information if useful, but label them clearly as not part of the presented
financial statements unless the entity decides to present in both currencies (30.17 permits "any currency (or currencies)").

## 11. Review points and common errors

1. KHR monetary balances left at the historical rate in the USD ledger (step B skipped).
2. Translation difference posted to profit or loss instead of OCI, or later recycled.
3. Comparatives re-expressed at the current closing rate instead of their own rates.
4. Share capital translated at the closing rate, so the translation reserve is mis-stated (reserve absorbs it all).
5. SoCF "effect of exchange rate changes" line missing, so opening + movement ≠ closing cash in KHR.
6. Note totals in KHR computed by translating each line separately at mixed rates so that notes do not cast or do not
   agree to the face — translate the USD total and the components consistently, and let rounding fall in one line.
7. Average rate used for a year in which the riel moved significantly (30.19).
8. Per-share amounts presented in KHR'000 by mistake — show them in single currency units.
