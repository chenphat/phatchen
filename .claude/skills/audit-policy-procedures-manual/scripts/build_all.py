#!/usr/bin/env python3
"""Build Word + Excel + PowerPoint from one content JSON, then verify them.

Usage: python build_all.py <content.json> <output_dir> [--basename NAME]

Verification: re-opens each file, and (if LibreOffice is installed) recalculates
the workbook headlessly and fails on any #REF!/#VALUE!/#NAME?/#DIV/0! errors.
"""
import argparse
import glob
import json
import os
import shutil
import subprocess
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build_docx  # noqa: E402
import build_pptx  # noqa: E402
import build_xlsx  # noqa: E402

ERRORS = ("#REF!", "#VALUE!", "#NAME?", "#DIV/0!", "#N/A", "#NUM!")


def check_formulas(xlsx):
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return "skipped (LibreOffice not installed)"
    from openpyxl import load_workbook
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run([soffice, "--headless", "--calc", "--convert-to", "xlsx", "--outdir", tmp, xlsx],
                       check=True, capture_output=True, timeout=180)
        wb = load_workbook(glob.glob(os.path.join(tmp, "*.xlsx"))[0], data_only=True)
        bad = [f"{ws.title}!{c.coordinate}={c.value}" for ws in wb for row in ws.iter_rows() for c in row
               if isinstance(c.value, str) and c.value in ERRORS]
    if bad:
        raise SystemExit("Formula errors found:\n  " + "\n  ".join(bad[:30]))
    return "OK - no formula errors after recalculation"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("content")
    ap.add_argument("outdir")
    ap.add_argument("--basename", default=None)
    args = ap.parse_args()
    with open(args.content, encoding="utf-8") as f:
        content = json.load(f)
    base = args.basename or content["meta"]["doc_ref"] + "_" + content["meta"]["title"].replace(" ", "_")
    os.makedirs(args.outdir, exist_ok=True)
    paths = {ext: os.path.join(args.outdir, f"{base}{suffix}.{ext}") for ext, suffix in
             (("docx", ""), ("xlsx", "_Workbook"), ("pptx", "_Briefing"))}
    build_docx.build(content, paths["docx"])
    build_xlsx.build(content, paths["xlsx"])
    build_pptx.build(content, paths["pptx"])

    from docx import Document
    from openpyxl import load_workbook
    from pptx import Presentation
    d = Document(paths["docx"])
    print(f"DOCX  {paths['docx']}  ({len(d.paragraphs)} paragraphs, {len(d.tables)} tables)")
    wb = load_workbook(paths["xlsx"])
    print(f"XLSX  {paths['xlsx']}  (sheets: {', '.join(wb.sheetnames)})")
    print(f"      formula check: {check_formulas(paths['xlsx'])}")
    print(f"PPTX  {paths['pptx']}  ({len(Presentation(paths['pptx']).slides)} slides)")


if __name__ == "__main__":
    main()
