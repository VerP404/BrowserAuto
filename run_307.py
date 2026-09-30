"""ASCII launcher: python run_307.py --limit 3 --building \"ГП №11\""""
from __future__ import annotations

import runpy
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
OMS = ROOT / "robots" / "oms"
SCRIPT = OMS / "cel_307" / "add_307.py"
if not SCRIPT.is_file():
    raise SystemExit(f"Script not found: {SCRIPT}")
sys.argv[0] = str(SCRIPT)
sys.path.insert(0, str(SCRIPT.parent))
sys.path.insert(0, str(OMS / "dv_opv"))
runpy.run_path(str(SCRIPT), run_name="__main__")
