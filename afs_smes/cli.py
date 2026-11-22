"""Command line: python -m afs_smes sample | build"""
import argparse
from collections import Counter

from . import pipeline, sample


def main(argv=None):
    ap = argparse.ArgumentParser(prog="afs_smes", description="IFRS for SMEs (3rd edition) AFS workbook from GL and TB")
    sub = ap.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("sample", help="write a two-year sample GL/TB data set with deliberate data defects")
    s.add_argument("--out", default="data/sample")
    s.add_argument("--year", type=int, default=2026)
    b = sub.add_parser("build", help="clean, check, map and draft the financial statements workbook")
    b.add_argument("--tb-cy", required=True, help="current-year trial balance (xlsx/csv)")
    b.add_argument("--tb-py", help="prior-year trial balance")
    b.add_argument("--gl-cy", help="current-year general ledger")
    b.add_argument("--gl-py", help="prior-year general ledger")
    b.add_argument("--config", help="entity information (json)")
    b.add_argument("--mapping", help="reviewed mapping file (csv/xlsx: account, afs_code, nature)")
    b.add_argument("--out", default="output/AFS.xlsx")
    b.add_argument("--no-verify", action="store_true", help="skip the LibreOffice recalculation of the live checks")
    a = ap.parse_args(argv)

    if a.cmd == "sample":
        out = sample.generate(a.out, a.year)
        print(f"Sample data written to {out}/")
        return 0
    res = pipeline.run(a.tb_cy, a.out, a.tb_py, a.gl_cy, a.gl_py, a.config, a.mapping)
    print(f"Workbook: {res['out']}")
    print(f"Cleaning actions logged: {len(res['cleaning'])}")
    print("Source-data checks: " + ", ".join(f"{k} {v}" for k, v in sorted(Counter(c['result'] for c in res['checks']).items())))
    for c in res["checks"]:
        if c["result"] in ("FAIL", "WARN"):
            print(f"  {c['result']:4} {c['id']} {c['check']} - {c['detail']}")
    failed = [c for c in res["checks"] if c["result"] == "FAIL"]
    if not a.no_verify:
        live = pipeline.live_check_results(res["out"])
        if live is None:
            print("Live checks: LibreOffice not found - open the workbook and read the Checks sheet")
        else:
            bad = [l for l in live if l[2] != "PASS"]
            print(f"Live workbook checks: {len(live) - len(bad)} of {len(live)} pass")
            for d, v, st in bad:
                print(f"  FAIL {d} (difference {v})")
            failed += bad
    return 1 if failed else 0
