"""Launcher из корня BrowserAuto. Автономный запуск — robots/oms/socstatus/run.bat."""
from __future__ import annotations

import runpy
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent / "robots" / "oms" / "socstatus"
SCRIPT = ROOT / "add_socstatus.py"
if not SCRIPT.is_file():
    raise SystemExit(f"Script not found: {SCRIPT}")
sys.argv[0] = str(SCRIPT)
sys.path.insert(0, str(ROOT))
runpy.run_path(str(SCRIPT), run_name="__main__")
