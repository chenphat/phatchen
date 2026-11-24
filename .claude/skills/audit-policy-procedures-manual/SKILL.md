---
name: audit-policy-procedures-manual
description: Reads, reviews and analyses an audit methodology, practice manual, training deck or regulation (e.g. an ISA Audit Practice Manual), then drafts a firm Policies and Procedures Manual as three linked deliverables - a Word manual, an Excel companion workbook (registers, calculators, checklists) and a PowerPoint staff briefing deck. Use this skill whenever the user asks to turn audit, assurance, accounting or finance guidance into a policies and procedures (P&P) manual, an audit methodology, SOPs, a quality management manual, or a policy pack in Word/Excel/PowerPoint, or asks to "review and analyse this manual and draft policies", even if they don't say "skill" or name all three formats. Also use it for ISA, ISQM, ISA 315/240/570, materiality, sampling or working paper policy drafting for an audit firm or teaching material for accounting lecturers.
---

# Audit Policies & Procedures Manual Builder

Turns a source document (a methodology manual, workshop deck, standard or regulation) into a firm-ready
**Policies & Procedures Manual** delivered as:

| File | What it contains |
|---|---|
| `*.docx` Manual | Cover, document control and approval, contents, introduction, **review & analysis of the source** (structure, strengths, rated gaps), one section per policy (purpose, scope, numbered "shall" statements, procedures table with responsibility / WP ref / timing, documentation required, reviewer red flags), appendices |
| `*_Workbook.xlsx` | Policy register, procedures tracker (status dropdowns, % complete), RACI (formula-driven), WP index, acceptance checklist, materiality calculator, sample size calculator, risk register, ISA 450 misstatements, gap analysis, implementation tracker. All calculations are live formulas |
| `*_Briefing.pptx` | ~20-slide staff training deck: why, review findings, policy map, key requirements, completion & reporting, changes in standards, roadmap, toolkit |

All three are generated from **one content JSON**, so they never disagree with each other.

## Workflow

### 1. Get the source
- Google Drive link → use the Google Drive connector (`read_file_content` with the file ID from the URL).
  If it returns "not found", the file is probably owned by a different Google account: list recent files,
  then ask the user to share it with the connected account, upload it, or pick another file. Don't guess the content.
- Local PDF/DOCX/PPTX → extract the text (`pdftotext`, python-docx, python-pptx).
- Save the extracted text to the scratchpad and read **all** of it before drafting. Large tool outputs
  are often saved to a file. Read that file in full, because the policies rely on the detail.

### 2. Review and analyse
Build a picture of the source before writing policy. Capture:
- **Status**: is it mandatory or guidance? Who owns it? This decides how much "shall" language the firm must add.
- **Structure**: parts, phases, the file or working-paper reference system.
- **Strengths**: what is practical and testable (rules of thumb, tables, worked examples).
- **Gaps**: what a firm still needs, such as ownership and version control, scope limits with no escalation route, standards
  covered only lightly, reporting, quality review, file assembly and retention, values cited but not reproduced, and
  upcoming effective dates. Rate each gap High/Medium/Low by its effect on file defensibility, and point it to the policy that closes it.

`references/review-checklist.md` lists the questions to ask, with an ISA coverage map. Read it for audit sources.

### 3. Draft the content JSON
Copy `assets/apm_content.json` as the template and adapt it. The schema is in `references/content-schema.md`.
Principles that make the manual useful rather than decorative:
- **Policy statements are requirements**, not summaries. Use "shall", one requirement per statement, and
  keep the source's sharp rules (e.g. "a heading is not a risk", "sign Ac2 before fieldwork").
- **Every procedure has an owner, a WP reference and a timing.** This lets the Excel tracker and the RACI work,
  and gives an inspector a trail to follow.
- **Red flags** are what a reviewer would spot on a weak file. They make the manual a review tool.
- **Never invent numbers or citations.** If the source cites a table it doesn't reproduce (e.g. sample risk
  factors), record it as a gap and make it a controlled input in Excel. Use only thresholds that appear in
  the source or are standard practice ranges, and label them as such.
- Keep the firm name, owner and approver as placeholders (e.g. `[Firm Name]`) unless the user gives them.
- Respect effective dates. Say which periods each new standard applies to, with a concrete year-end example.

If the user's source isn't audit (e.g. procurement, treasury, HR), keep the same meta/analysis/policies
structure. The audit-only tables (materiality, sampling, opinion grid) can be emptied, and the Excel
calculator sheets should then be removed from `build_xlsx.py` (see `references/content-schema.md`).

### 4. Build and verify
```bash
python scripts/build_all.py <content.json> <output_dir>
```
This writes all three files, re-opens them, and recalculates the workbook with LibreOffice (if installed),
failing on any `#REF!`, `#VALUE!`, `#NAME?` or `#DIV/0!` errors. Then render a visual check:
```bash
soffice --headless --convert-to pdf <file.pptx|file.docx> --outdir <tmp>
pdftoppm -r 50 -png <file.pdf> <tmp>/p
```
Look at the slide and page images for overflowing text, wrapped labels and tables with collapsed column widths, then fix
the script or shorten the text. Spot-check key calculator outputs against the source's worked examples
(the bundled APM example must give a substantive sample of 14).

### 5. Deliver
- Put the outputs where the user can open them (the repo `outputs/` folder or a path the user names).
- Summarise in a few lines: what the source is, the top gaps found, what each file contains, and what the
  user must still decide (firm name, owners, risk-factor table, local law and retention periods, sign-off).
- The Word table of contents refreshes when the file is opened in Word (`updateFields` is set). Mention this.

## Bundled resources
- `assets/apm_content.json`: complete worked content built from the *Audit Practice Manual: The Fundamentals of an ISA Audit*
  (16 policies, 10 rated gaps, all reference tables). Use it as the default and as the schema example.
- `scripts/build_docx.py`, `build_xlsx.py`, `build_pptx.py`: the individual generators (each runs on its own).
- `scripts/build_all.py`: builds and verifies all three.
- `references/content-schema.md`: field-by-field schema and how to adapt it for non-audit manuals.
- `references/review-checklist.md`: review and gap-analysis questions plus an ISA coverage map.
- `references/source-analysis-apm.md`: the full review of the bundled source, as a model of the analysis expected.
