"""End-to-end run: read -> clean -> check -> map -> classify -> build the AFS workbook."""
import json
import shutil
import subprocess
import tempfile
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook

from . import checks as chk
from .ingest import Log, harmonise_names, read_gl, read_tb
from .mapping import build_mapping, load_user_mapping
from .movements import classify
from .sheets_data import (build_aje, build_dq_log, build_entity, build_gl, build_guide, build_lead, build_lists,
                          build_mapping as sheet_mapping, build_tb)
from .sheets_notes import Notes, plan_notes
from .sheets_review import build_analysis, build_checklist, build_checks
from .sheets_statements import build_soce, build_socf, build_soci, build_soci_nature, build_sofp
from .taxonomy import LEAF
from .xl import Book

SHEETS = ["Guide", "Entity", "Checks", "DQ_Log", "SoFP", "SoCI", "SoCE", "SoCF", "Notes", "SoCI_Nature", "Analysis", "Checklist",
          "Mapping", "Lead", "TB", "AJE", "GL", "Lists"]
TAB = {"Guide": "7F7F7F", "Entity": "FFC000", "Checks": "C00000", "DQ_Log": "C00000", "SoFP": "1F3864",
       "SoCI": "1F3864", "SoCE": "1F3864", "SoCF": "1F3864", "Notes": "1F3864", "SoCI_Nature": "8EA9DB", "Analysis": "548235",
       "Checklist": "548235", "Mapping": "FFC000", "AJE": "FFC000"}


def run(tb_cy, out, tb_py=None, gl_cy=None, gl_py=None, config=None, mapping_file=None):
    cfg = json.loads(Path(config).read_text()) if config else {}
    log = Log()
    dayfirst = cfg.get("dayfirst", True)
    t_cy = read_tb(tb_cy, log)
    t_py = read_tb(tb_py, log) if tb_py else None
    g_cy = read_gl(gl_cy, log, dayfirst=dayfirst) if gl_cy else None
    g_py = read_gl(gl_py, log, dayfirst=dayfirst) if gl_py else None
    names = harmonise_names([t for t in (t_cy, t_py) if t is not None], [g_cy, g_py], log)

    if "reporting_date" not in cfg:
        last = max(d for d in g_cy["date"] if d) if g_cy is not None and g_cy["date"].notna().any() else date.today()
        cfg["reporting_date"] = date(last.year, 12, 31).isoformat()
        log.add("config", "-", "No reporting date supplied", f"Assumed {cfg['reporting_date']} from the latest GL posting")
    end = datetime.strptime(cfg["reporting_date"], "%Y-%m-%d").date()
    start = lambda e: (pd.Timestamp(e) - pd.DateOffset(years=1) + pd.Timedelta(days=1)).date()
    py_end = (pd.Timestamp(start(end)) - pd.Timedelta(days=1)).date()
    periods = {"CY": (start(end), end), "PY": (start(py_end), py_end)}

    accounts = dict(names)
    user_map = load_user_mapping(mapping_file) if mapping_file else None
    mapping = build_mapping(accounts, cfg.get("account_classes"), user_map)
    leaf_of = dict(zip(mapping["account"], mapping["leaf"]))
    for g in (g_cy, g_py):
        if g is not None:
            g["mvt"] = classify(g, leaf_of)

    results = chk.run_checks(t_cy, g_cy, mapping, periods["CY"], "CY")
    if t_py is not None:
        results += chk.run_checks(t_py, g_py, mapping, periods["PY"], "PY")
    results += chk.continuity_checks(t_cy, t_py, mapping) + chk.mapping_checks(mapping)

    # AFS components / face lines that carry balances in either year
    book = Book(Workbook())
    bal = pd.concat([t for t in (t_cy, t_py) if t is not None])
    bal = bal[(bal[["opening", "debit", "credit", "closing"]].abs().sum(axis=1) > 0)]
    book.active = {leaf_of[a] for a in bal["account"] if leaf_of.get(a) in LEAF}
    book.active_faces = {LEAF[l][2] for l in book.active}
    book.active_natures = set(mapping.loc[mapping["account"].isin(bal["account"]), "nature"]) - {""}

    wb = book.wb
    wb.remove(wb.active)
    ws = {n: wb.create_sheet(n) for n in SHEETS}
    for n, colour in TAB.items():
        ws[n].sheet_properties.tabColor = colour
    build_entity(book, ws["Entity"], cfg)
    build_guide(book, ws["Guide"])
    build_lead(book, ws["Lead"])
    build_lists(book, ws["Lists"])
    sheet_mapping(book, ws["Mapping"], mapping)
    build_tb(book, ws["TB"], mapping, t_cy, t_py)
    build_aje(book, ws["AJE"])
    build_gl(book, ws["GL"], [("PY", g_py), ("CY", g_cy)])
    plan_notes(book)
    build_soci(book, ws["SoCI"])
    build_soci_nature(book, ws["SoCI_Nature"])
    build_sofp(book, ws["SoFP"])
    build_soce(book, ws["SoCE"])
    build_socf(book, ws["SoCF"])
    Notes(book, ws["Notes"], cfg).build()
    build_analysis(book, ws["Analysis"])
    build_checklist(book, ws["Checklist"])
    build_checks(book, ws["Checks"])
    sources = [Path(p).name for p in (tb_cy, tb_py, gl_cy, gl_py) if p]
    build_dq_log(book, ws["DQ_Log"], results, log.rows, sources)
    for n in ("SoFP", "SoCI", "SoCE", "SoCF", "Notes"):
        s = ws[n]
        s.page_setup.orientation = "portrait" if n not in ("SoCE", "Notes") else "landscape"
        s.page_setup.fitToWidth, s.page_setup.fitToHeight = 1, 0
        s.sheet_properties.pageSetUpPr.fitToPage = True
    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    mapping.rename(columns={"leaf": "afs_code"})[["account", "account_name", "afs_code", "nature", "basis"]].to_csv(
        out.with_name(out.stem + "_mapping.csv"), index=False)
    return {"out": out, "checks": results, "cleaning": log.rows, "mapping": mapping, "book": book}


def live_check_results(path):
    """Return [(description, difference, status)] from the Checks sheet after a LibreOffice recalculation."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return None
    path = Path(path).resolve()
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run([soffice, "--headless", f"-env:UserInstallation=file://{tmp}/profile", "--convert-to",
                        "xlsx:Calc MS Excel 2007 XML", "--outdir", tmp, str(path)],
                       check=True, capture_output=True, timeout=300)
        ws = load_workbook(Path(tmp) / path.name, data_only=True)["Checks"]
        return [(r[1], r[2], r[3]) for r in ws.iter_rows(min_row=6, values_only=True) if r[0] is not None and r[3]]
