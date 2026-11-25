import shutil

import pytest
from openpyxl import load_workbook

from afs_smes import pipeline, sample
from afs_smes.ingest import parse_number
from afs_smes.mapping import suggest


@pytest.fixture(scope="module")
def data(tmp_path_factory):
    d = tmp_path_factory.mktemp("sample")
    sample.generate(d, 2026)
    return d


def build(d, out, **over):
    args = dict(tb_cy=d / "TB_2026.xlsx", tb_py=d / "TB_2025.xlsx", gl_cy=d / "GL_2026.xlsx",
                gl_py=d / "GL_2025.xlsx", config=d / "config.json")
    args.update(over)
    return pipeline.run(out=out, **args)


def test_parse_number():
    assert parse_number("1,234.50") == (1234.5, True)
    assert parse_number("(1,234.50)")[0] == -1234.5
    assert parse_number("1.234,50")[0] == 1234.5
    assert parse_number(None) == (0.0, False)


def test_mapping_rules():
    assert suggest("1515", "Accumulated depreciation - buildings")[0] == "PPE.LB.D"
    assert suggest("2500", "Bank loan - non-current portion")[0] == "BORR_NC.LOAN"
    assert suggest("6020", "Salaries - sales staff")[:2] == ("DIST", "EMP")
    assert "review" in suggest("2999", "Miscellaneous")[2].lower()


def test_sample_is_cleaned_and_passes(data, tmp_path):
    res = build(data, tmp_path / "afs.xlsx")
    assert not [c for c in res["checks"] if c["result"] == "FAIL"]
    issues = " ".join(c["issue"] for c in res["cleaning"])
    for expected in ("duplicate", "stored as text", "Total/summary", "Negative amount", "stray spaces"):
        assert expected in issues
    assert any(c["id"] == "CY-GL06" and c["result"] == "WARN" for c in res["checks"])  # cut-off posting flagged


def test_workbook_statements_reconcile(data, tmp_path):
    if not shutil.which("soffice"):
        pytest.skip("LibreOffice not installed")
    res = build(data, tmp_path / "afs.xlsx")
    live = pipeline.live_check_results(res["out"])
    assert len(live) > 40 and all(status == "PASS" for _, _, status in live)
    # Excel rejects a formula containing a text literal longer than 255 characters
    import re
    for ws in load_workbook(res["out"]).worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    assert all(len(lit) <= 255 for lit in re.findall(r'"((?:[^"]|"")*)"', c.value)), (ws.title, c.coordinate)


def test_broken_trial_balance_is_detected(data, tmp_path):
    bad = tmp_path / "TB_bad.xlsx"
    wb = load_workbook(data / "TB_2026.xlsx")
    ws = wb.active
    row = next(r for r in range(5, ws.max_row) if ws.cell(r, 2).value == "Office rent")
    ws.cell(row, 7).value = ws.cell(row, 7).value + 1000      # closing debit overstated
    wb.save(bad)
    res = build(data, tmp_path / "afs.xlsx", tb_cy=bad)
    failed = {c["id"] for c in res["checks"] if c["result"] == "FAIL"}
    assert {"CY-TB02", "CY-TB03"} <= failed
