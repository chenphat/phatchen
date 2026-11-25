#!/usr/bin/env python3
"""Build the IFRS for SMEs compliance checklist data from the text of the Standard.

Input: the Standard converted with `pdftotext -layout 2025_IFRS_for_SMEs.pdf std.txt`.
Output: a JSON list of checklist items (presentation, disclosure, recognition and measurement, transition),
one item per lettered sub-requirement where the paragraph has them.

Usage: python extract_checklist.py std.txt --out afs_smes/data/ifrs_smes_3e_checklist.json

The requirement wording is taken from the Standard (cleaned of page furniture); keep the output for internal
working-paper use under the IFRS Foundation's licence terms that came with the user's copy.
"""
import argparse
import json
import re

PARA = re.compile(r"^(\d{1,2}[A-Z]?\.\d{1,3}[A-Z]{0,2}|A\d{1,2}[A-Z]?)\s{2,}(\S.*)$")
ITEM = re.compile(r"^\s{4,12}\(([a-z]{1,2})\)\s+(\S.*)$")
ROMAN = re.compile(r"^\s{10,24}\(((?:i|ii|iii|iv|v|vi|vii|viii|ix|x|xi|xii))\)\s+(\S.*)$")
SECTION = re.compile(r"^\s*Section (\d{1,2})\s*$")
FURNITURE = re.compile(r"IFRS Foundation|^\s*FEBRUARY 2025\s*$|^\s*\d{1,3}\s*$|IFRS FOR SMES ACCOUNTING STANDARD|^\f")

DISCLOSE = re.compile(r"\bshall (also )?disclose\b|\bshall disclose\b|disclose the following|\bdiscloses?\b.*\bin the notes\b|"
                      r"shall (also )?state\b|shall (also )?describe\b|disclose that fact|shall (also )?explain\b|"
                      r"statement of such compliance in the notes")
PRESENT = re.compile(r"\bshall (also )?present\b|\bshall include,? in the statement|\bline items that present\b|"
                     r"\bshall display\b|\bshall classify\b|\bshall clearly identify\b|shall reclassify comparative")
RM = re.compile(r"\bshall (initially |subsequently )?(recognise|measure|derecognise)\b|\bshall (apply|use|account for)\b|"
                r"\bshall not (recognise|offset|discount|capitalise)\b|shall (choose|depreciate|amortise|allocate|determine)\b")


def clean(s):
    s = s.replace("—", "-").replace("–", "-").replace("’", "'").replace("‘", "'")
    s = s.replace("“", '"').replace("”", '"')
    return re.sub(r"\s+", " ", s).strip()


def is_heading(line):
    s = line.strip()
    if not s or len(s) > 90 or line[:3].strip() == "" and len(line) - len(line.lstrip()) > 2:
        return False
    if s[0].isdigit() or s.startswith("(") or s.endswith((".", ";", ",", ":")):
        return False
    return s[0].isupper()


def parse(lines, start=900):
    paras, section, title, heading = [], None, "", ""
    cur = None
    i = start
    pending_heading = []
    while i < len(lines):
        raw = lines[i].rstrip("\n")
        i += 1
        if FURNITURE.search(raw):
            continue
        m = SECTION.match(raw)
        if m:
            section = int(m.group(1))
            j = i
            while j < len(lines) and not lines[j].strip():
                j += 1
            title = clean(lines[j])
            k = j + 1   # titles can wrap onto a second line
            if k < len(lines) and lines[k].strip() and not PARA.match(lines[k]) and is_heading(lines[k]) \
                    and lines[k].strip()[0].islower() is False and len(title) > 40:
                title += " " + clean(lines[k])
            i = k
            cur, heading, pending_heading = None, "", []
            continue
        if raw.strip().startswith("Appendix A") or "EFFECTIVE DATE AND TRANSITION" in raw.upper() and section == 35:
            section, title = "A", "Appendix A - Effective date and transition"
        m = PARA.match(raw)
        if m and section is not None:
            if pending_heading:
                heading = clean(" ".join(pending_heading))
            cur = {"ref": m.group(1), "section": section, "section_title": title, "heading": heading,
                   "intro": m.group(2), "items": []}
            paras.append(cur)
            pending_heading = []
            continue
        if not raw.strip():
            if pending_heading:
                heading = clean(" ".join(pending_heading))
                pending_heading = []
            continue
        if cur is None:
            if is_heading(raw) or (pending_heading and len(raw) - len(raw.lstrip()) <= 2 and len(raw.strip()) < 60
                                   and raw.strip()[0].islower()):
                pending_heading.append(raw)
            continue
        mi, mr = ITEM.match(raw), ROMAN.match(raw)
        if mr and cur["items"]:
            cur["items"][-1]["subs"].append([mr.group(1), mr.group(2)])
        elif mi:
            cur["items"].append({"label": mi.group(1), "text": mi.group(2), "subs": []})
        elif len(raw) - len(raw.lstrip()) <= 2 and is_heading(raw):
            pending_heading.append(raw)
            cur = None
        else:
            text = raw.strip()
            if cur["items"]:
                it = cur["items"][-1]
                if it["subs"]:
                    it["subs"][-1][1] += " " + text
                else:
                    it["text"] += " " + text
            else:
                cur["intro"] += " " + text
    return paras


def area(section, text):
    if DISCLOSE.search(text):
        return "Disclosure"
    if PRESENT.search(text) or (isinstance(section, int) and 3 <= section <= 8):
        return "Presentation"
    if RM.search(text):
        return "Recognition and measurement"
    return None


def build(paras):
    out = []
    for p in paras:
        if p["section"] in (1, 2) and p["ref"] not in ("2.30", "2.31"):
            continue   # scope and concepts: covered by the section guides, not checklist items
        intro = clean(p["intro"])
        full = intro + " " + " ".join(clean(it["text"]) for it in p["items"])
        kind = "Transition" if p["section"] in ("A", 35) and not DISCLOSE.search(full) else area(p["section"], full)
        if p["section"] in ("A", 35) and DISCLOSE.search(full):
            kind = "Disclosure"
        if not kind or "[Deleted]" in intro or intro.startswith("[Deleted]"):
            continue
        if p["heading"].startswith("Scope"):
            kind = "Scope"
        base = {"section": p["section"], "section_title": p["section_title"], "topic": p["heading"], "area": kind}
        listed = p["items"] and kind in ("Disclosure", "Presentation") and re.search(r"[:\-]\s*$|following", intro)
        if listed:
            for it in p["items"]:
                t = clean(it["text"])
                if t.lower().startswith("[deleted]"):
                    continue
                if it["subs"]:
                    t += " " + " ".join(f"({r}) {clean(x)}" for r, x in it["subs"])
                out.append(dict(base, ref=f"{p['ref']}({it['label']})", requirement=t, context=intro))
        else:
            t = intro
            if p["items"]:
                t += " " + " ".join(f"({it['label']}) {clean(it['text'])}" +
                                    "".join(f" ({r}) {clean(x)}" for r, x in it["subs"]) for it in p["items"])
            out.append(dict(base, ref=p["ref"], requirement=t, context=""))
    for n, item in enumerate(out, start=1):
        item["id"] = n
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("text")
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    lines = open(a.text, encoding="utf-8").read().split("\n")
    start = next(i for i, l in enumerate(lines) if SECTION.match(l) and l.strip() == "Section 1")
    items = build(parse(lines, start))
    json.dump(items, open(a.out, "w", encoding="utf-8"), indent=1, ensure_ascii=False)
    from collections import Counter
    print(f"{len(items)} items:", dict(Counter(i["area"] for i in items)))


if __name__ == "__main__":
    main()
