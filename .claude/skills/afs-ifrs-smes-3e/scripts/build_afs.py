#!/usr/bin/env python3
"""Run the bundled AFS engine from anywhere.

Examples
  python build_afs.py sample --out ./sample                     # teaching data set (USD ledger, KHR presentation)
  python build_afs.py build --tb-cy TB_2026.xlsx --tb-py TB_2025.xlsx \\
      --gl-cy GL_2026.xlsx --gl-py GL_2025.xlsx --config config.json --out AFS_2026.xlsx
Requires: pandas, openpyxl. LibreOffice (soffice) is optional and used to verify the live checks.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from afs_smes.cli import main  # noqa: E402

sys.exit(main())
