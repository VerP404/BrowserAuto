"""ASCII launcher for carryover DN plan robot."""
from __future__ import annotations

import runpy
import sys
from pathlib import Path

SCRIPT = (
    Path(__file__).resolve().parent
    / "robots"
    / "iszl"
    / "carryover"
    / "add_carryover_plan.py"
)

if not SCRIPT.is_file():
    raise SystemExit(f"Script not found: {SCRIPT}")

sys.argv[0] = str(SCRIPT)
runpy.run_path(str(SCRIPT), run_name="__main__")
