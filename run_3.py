"""ASCII launcher: python run_3.py --limit 3 --building \"ГП №3\" --doctor 30097144"""
from __future__ import annotations

import runpy
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
OMS = ROOT / "robots" / "oms"
SCRIPT = OMS / "cel_3" / "add_3.py"
if not SCRIPT.is_file():
    raise SystemExit(f"Script not found: {SCRIPT}")
sys.argv[0] = str(SCRIPT)
sys.path.insert(0, str(SCRIPT.parent))
sys.path.insert(0, str(OMS / "dv_opv"))
sys.path.insert(0, str(OMS / "cel_307"))
runpy.run_path(str(SCRIPT), run_name="__main__")
