"""ASCII launcher: python run_307_template.py --from-source ..."""
from __future__ import annotations

import runpy
import sys
from pathlib import Path

SCRIPT = (
    Path(__file__).resolve().parent / "robots" / "oms" / "cel_307" / "template_307.py"
)
if not SCRIPT.is_file():
    raise SystemExit(f"Script not found: {SCRIPT}")
sys.argv[0] = str(SCRIPT)
sys.path.insert(0, str(SCRIPT.parent))
runpy.run_path(str(SCRIPT), run_name="__main__")
