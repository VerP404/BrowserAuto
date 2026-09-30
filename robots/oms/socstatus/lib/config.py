"""Локальный конфиг автономного oms_socstatus."""

from __future__ import annotations

import os
from pathlib import Path

from dotenv import load_dotenv

ROOT = Path(__file__).resolve().parent.parent
GECKODRIVER = ROOT / "geckodriver.exe"
ENV_FILE = ROOT / "credentials.env"
DATA_DIR = ROOT / "data"


def load_credentials(*extra: Path | str) -> None:
    """Загрузка credentials.env из корня пакета (+ опциональные override)."""
    if ENV_FILE.is_file():
        load_dotenv(ENV_FILE, override=True)
    for raw in extra:
        path = Path(raw)
        if path.is_file():
            load_dotenv(path, override=True)


def oms_base_url() -> str:
    return (os.getenv("OMS_BASE_URL") or "http://10.36.0.142:9000").rstrip("/")


def oms_login() -> str:
    return (os.getenv("OMS_LOGIN") or "").strip()


def oms_password() -> str:
    return (os.getenv("OMS_PASSWORD") or "").strip()


def browser_implicit_wait() -> int:
    return int(os.getenv("BROWSER_IMPLICIT_WAIT", "10"))
