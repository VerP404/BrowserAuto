"""ASCII launcher: python run_form_diagnosis.py --limit 1"""
from __future__ import annotations

import runpy
import sys
from pathlib import Path

SCRIPT = (
    Path(__file__).resolve().parent
    / "robots"
    / "oms"
    / "form_diagnosis"
    / "form_diagnosis.py"
)
if not SCRIPT.is_file():
    raise SystemExit(f"Script not found: {SCRIPT}")
sys.argv[0] = str(SCRIPT)
sys.path.insert(0, str(SCRIPT.parent))
sys.path.insert(0, str(SCRIPT.parent.parent / "dv_opv"))
runpy.run_path(str(SCRIPT), run_name="__main__")
