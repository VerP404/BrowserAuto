"""ASCII launcher: python run_medical_exam.py --limit 1"""
from __future__ import annotations

import runpy
import sys
from pathlib import Path

SCRIPT = (
    Path(__file__).resolve().parent / "robots" / "oms" / "dv_opv" / "add_medical_exam.py"
)
if not SCRIPT.is_file():
    raise SystemExit(f"Script not found: {SCRIPT}")
sys.argv[0] = str(SCRIPT)
# ensure package-local imports (doctors_catalog) resolve
sys.path.insert(0, str(SCRIPT.parent))
runpy.run_path(str(SCRIPT), run_name="__main__")
