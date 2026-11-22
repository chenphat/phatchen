---
name: humaniser
description: Edit academic prose so it reads as the work of a careful human scholar, by removing the stock patterns of machine-drafted text (aphoristic closers, "not X but Y" contrasts, signposting, inflated vocabulary, reflexive triplets, self-praise) without changing the argument. Use when asked to humanise, clean, polish or de-AI a research proposal, dissertation chapter, journal article, lecture note or report, especially in accounting and finance. Works on .docx, Markdown and plain text.
---

# Humaniser

Version 1.0

## Purpose

The reader of a doctoral proposal is an examiner who reads many of them. Stock phrasing costs the writer credibility because it signals that a sentence was produced without thought. This skill removes that phrasing and leaves the argument, the evidence and the author's own voice alone.

The aim is better prose. It is not to defeat detection software, and the skill must never be used to conceal that AI tools were used where a university or journal requires disclosure.

## Fixed rules

These hold for every edit. If a change would break one, do not make it.

1. Do not change meaning. Hypotheses, research questions, objectives, construct definitions, questionnaire items and interview questions keep their exact wording unless the author asks otherwise. They are instruments, and rewording them changes the research.
2. Do not touch citations, reference-list entries, numbers, statistical thresholds, dates, table and figure numbers or cross-references, except to correct an inconsistency, and then report the correction.
3. Do not add facts, sources, examples or claims. If a sentence needs evidence it lacks, flag it for the author.
4. Do not remove or weaken a declaration, an AI-use disclosure, a limitation or an ethical safeguard.
5. Do not fake a human. No invented anecdotes, deliberate errors, slang, contractions or forced informality in academic text.
6. Keep the author's conventions: spelling variety (British for NUM work: organisation, judgement, modelling), APA 7 citation form, the headings prescribed by the institution's handbook, and the defined terms and abbreviations of the document.
7. Keep hedges that carry real uncertainty ("is associated with", "may", "the design cannot establish"). Remove only hedges stacked on other hedges.
8. Edit lightly where the prose is already sound. A sentence with no pattern from the catalogue below is left as it is.

## Pattern catalogue

Each entry gives the pattern, why it reads as machine-drafted, and the repair.

### 1. Aphoristic closer

A short, quotable sentence that ends a paragraph by restating it.

- Before: "Enforcement quality, in other words, is not the same as regulatory quantity."
- After: delete it, or fold the point into the preceding sentence as a plain finding.

Test: if the paragraph loses nothing when the last sentence goes, the sentence goes.

### 2. Manufactured contrast

"Not X but Y", "X, not Y", "rather than", "and not as", used where nobody proposed X.

- Before: "Financial reporting is better understood as part of the infrastructure of a financial system than as a compliance exercise."
- After: "Financial reporting is part of the infrastructure of a financial system."

Keep the contrast when X is a real alternative that the literature or the design considered (for example, "the ten-times rule is not relied upon, because it underestimates the sample required").

### 3. Signposting and throat-clearing

"It is important to note that", "The implication is that", "In other words", "The lesson for Cambodia is that", "It is worth noting", "In doing so".

Repair: delete the announcement and state the content. Chapter and section outlines required by the handbook stay.

### 4. Reveal structures

A colon or a question that sets up a small surprise: "The result: ...", "The pattern is the one that would be expected if ...", "That condition is where trust enters the model."

Repair: say the thing directly, with its subject first.

### 5. Reflexive triplets and counted lists

Three items because three sounds complete; "for three reasons" where the reasons are of unequal weight.

Repair: keep the number of items the content has. Counted lists are legitimate in a methods chapter when each item is later used. Check that the count matches the list and any table that restates it.

### 6. Inflated and empty vocabulary

Words to cut or replace on sight: delve, crucial, pivotal, key (as adjective), robust (outside statistics), comprehensive, holistic, seamless, landscape, realm, tapestry, underscore, shed light on, foster, leverage, navigate, nuanced, multifaceted, vital, notably, importantly, indeed, moreover, furthermore (at the head of a sentence, more than once a page).

Intensifiers and sincerity markers: very, highly, precisely, exactly, clearly, truly, particularly (when nothing is being compared), "with unusual clarity", "a large body of".

Repair: delete, or replace with the specific claim the word stood in for.

### 7. Metaphor in place of a claim

"The missing link", "the puzzle", "bridges the gap", "does little work", "the scholarly problem mirrors the practical one".

Repair: write the literal statement. "Trust as the missing link" becomes "Trust as the mechanism".

### 8. Self-praise and salesmanship

"The candidate is well placed to", "this novel study", "a significant contribution", "the first study to" without a qualifier.

Repair: state the fact that supports the claim and let the examiner judge. Qualify priority claims: "the first known to the candidate".

### 9. Mechanical rhythm

Every paragraph opening with a topic sentence of the same length and closing with a summary; long runs of "First, ... Second, ... Third, ..."; interposed "however," in the middle of short sentences; semicolons joining clauses that would be clearer as two sentences.

Repair: vary only where it helps reading. Split a semicolon sentence when the halves make separate points. Move "however" to the front or use "but".

### 10. Redundant restatement

A section that ends by summarising itself, a sentence that says what the table beside it says, a definition repeated in full a third time.

Repair: keep the first full statement and the handbook-required chapter summary. Cut the rest to a cross-reference.

### 11. Stacked hedges and empty caution

"may potentially suggest", "it could be argued that it is possible".

Repair: one hedge per claim, placed on the word that is uncertain.

### 12. Typography tells

Em dashes used as a general-purpose pause, curly "scare quotes" around ordinary words, bold scattered through running text, bullet lists where an argument needs sentences.

Repair: commas, parentheses or a new sentence in place of em dashes. En dashes stay in number ranges. Bold only for run-in headings and table heads.

## Procedure

1. Read the whole document before changing anything. Note its conventions (rule 6) and the passages protected by rule 1.
2. Mark every instance of a catalogue pattern. Count them by type, so that the report to the author is specific.
3. Rewrite sentence by sentence. Prefer deletion to substitution, and substitution to restructuring.
4. Reread each edited paragraph in full. An edit that removes one pattern often creates another, such as a new contrast or an orphaned "this".
5. Run the consistency checks:
   - arithmetic in tables and sample-size statements;
   - counts stated in text against the lists and tables they introduce;
   - every in-text citation has a reference entry and every entry is cited;
   - cross-references point to the section that contains the material;
   - abbreviations are defined at first use and listed.
6. For .docx files, edit the XML in place so that styles, tables, figures and numbering survive. If the table of contents is typed text and not a field, render the file and check every page number against it.
7. Report to the author: what was changed, by pattern type; any inconsistency corrected; and everything flagged but left alone (unverified sources, placeholders, claims that need evidence).

## What is not a tell

Leave these alone. They are features of the genre, and stripping them makes the text worse.

- Passive voice in a methods chapter ("interviews will be recorded").
- Formal hedging that states the limit of a design.
- Enumerated research questions, objectives, hypotheses and limitations.
- Repeated technical terms. A construct keeps one name throughout, and elegant variation is an error in a thesis.
- Handbook-prescribed section headings and chapter summaries.
- Long sentences that carry a chain of reasoning and are punctuated correctly.

## Notes for accounting and finance texts

- Standards and bodies take their official names: "IFRS Accounting Standards", "IFRS for SMEs Accounting Standard", "International Auditing and Assurance Standards Board".
- "Associated with" and "effect on" are different claims. Do not strengthen an association into an effect while tidying a sentence, and do not weaken a cited causal finding into an association.
- Statistical terms are fixed: "significant" means statistically significant, "robust" is acceptable only for a robustness check, "impact" in a title set by the university stays.
- Currency, ratios and thresholds are copied exactly from the source text.
