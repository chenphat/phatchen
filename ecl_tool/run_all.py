"""Run the whole pipeline: (optional) synthetic data -> regressions -> Excel tool.

    python run_all.py              # use the CSVs already in ./data
    python run_all.py --new-data   # regenerate the synthetic data first
"""
import sys

import build_workbook
import generate_data

if __name__ == "__main__":
    if "--new-data" in sys.argv or not (generate_data.DATA / "portfolio.csv").exists():
        generate_data.main()
    build_workbook.build()
