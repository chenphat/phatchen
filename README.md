# phatchen

## Audit Policies & Procedures Manual skill

`.claude/skills/audit-policy-procedures-manual/` is a Claude skill. It reviews an audit methodology or source document and drafts a firm Policies & Procedures Manual as Word, Excel and PowerPoint files, all built from one content JSON.

```bash
S=.claude/skills/audit-policy-procedures-manual
python3 $S/scripts/build_all.py $S/assets/apm_content.json outputs
```

### Current outputs (`outputs/`)
Built from *Audit Practice Manual: The Fundamentals of an ISA Audit* (workshop, 18-19 August 2026):

| File | Contents |
|---|---|
| `APPM-2026-01_Audit_Policies_and_Procedures_Manual.docx` | 35-page manual: document control, review and gap analysis of the source, 16 policies, appendices A-G |
| `APPM-2026-01_..._Workbook.xlsx` | 12 sheets: registers, RACI, acceptance checklist, materiality and sample size calculators, risk register, ISA 450 misstatements, gap and implementation trackers |
| `APPM-2026-01_..._Briefing.pptx` | 20-slide staff briefing / training deck |
