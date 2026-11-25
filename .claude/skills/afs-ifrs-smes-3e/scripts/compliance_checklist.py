#!/usr/bin/env python3
"""Create a standalone IFRS for SMEs (third edition) compliance checklist workbook.

The same checklist is built into every AFS workbook made by build_afs.py (sheets Compliance_Summary and
Compliance). Use this script when the user wants the checklist on its own, e.g. to review financial statements
prepared elsewhere, or as a teaching handout.

Usage:
  python compliance_checklist.py --entity "Angkor Bright Co., Ltd." --period "year ended 31 December 2026" \
      [--answers answers.json] --out Compliance_Checklist_2026.xlsx
answers.json (optional) maps a section to "Yes"/"No", e.g. {"16": "Yes", "34": "No", "A": "Yes"}.
Sections without an answer show "?" (always-applicable sections default to "Yes").
"""
import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from openpyxl import Workbook  # noqa: E402

from afs_smes.sheets_compliance import build_compliance  # noqa: E402
from afs_smes.xl import F_BOLD, F_GREY, F_INPUT, F_TITLE, put, widths  # noqa: E402


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--entity", default="[Entity name]")
    ap.add_argument("--period", default="[year ended 31 December 2026]")
    ap.add_argument("--answers")
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    answers = {}
    if a.answers:
        for k, v in json.load(open(a.answers)).items():
            answers[int(k) if str(k).isdigit() else k] = v
    wb = Workbook()
    info = wb.active
    info.title = "Info"
    widths(info, [4, 30, 70])
    put(info, "B1", "IFRS for SMEs Accounting Standard (third edition) - compliance checklist", F_TITLE)
    put(info, "B3", "Entity", F_BOLD)
    put(info, "C3", a.entity, F_INPUT)
    put(info, "B4", "Period", F_BOLD)
    put(info, "C4", a.period, F_INPUT)
    put(info, "B6", "How to use", F_BOLD)
    steps = ["1. On Compliance_Summary, answer the applicability question for every section (Yes / No).",
             "2. On Compliance, set Status for each applicable item: Yes, Partly, No or N/A; say where it is addressed.",
             "3. Explain every No or Partly in the comment column and raise it as a finding; the reviewer signs off.",
             "4. Items are taken from the text of the Standard (February 2025); read the full paragraph where truncated."]
    for i, t in enumerate(steps):
        put(info, f"C{7 + i}", t, F_GREY)
    build_compliance(wb, applies=answers, entity_ref="Info!$C$3", period_ref="Info!$C$4")
    wb.save(a.out)
    print("Checklist:", a.out)


if __name__ == "__main__":
    main()
