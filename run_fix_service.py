"""ASCII launcher: python run_fix_service.py --limit 1 --dry-run"""
from __future__ import annotations

import runpy
import sys
from pathlib import Path

SCRIPT = (
    Path(__file__).resolve().parent
    / "robots"
    / "oms"
    / "fix_service"
    / "fix_service.py"
)
if not SCRIPT.is_file():
    raise SystemExit(f"Script not found: {SCRIPT}")
sys.argv[0] = str(SCRIPT)
sys.path.insert(0, str(SCRIPT.parent))
sys.path.insert(0, str(SCRIPT.parent.parent / "dv_opv"))
sys.path.insert(0, str(SCRIPT.parent.parent / "cel_307"))
runpy.run_path(str(SCRIPT), run_name="__main__")
