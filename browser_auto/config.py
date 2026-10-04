import os
from pathlib import Path

from dotenv import load_dotenv

PROJECT_ROOT = Path(__file__).resolve().parent.parent
GECKODRIVER = PROJECT_ROOT / "geckodriver.exe"
ENV_FILE = PROJECT_ROOT / ".env"


def _apply_env() -> None:
    """Перечитать переменные окружения в модульные константы."""
    global TESSERACT_CMD, OMS_BASE_URL, OMS_LOGIN, OMS_PASSWORD
    global ISZL_BASE_URL, ISZL_LOGIN, ISZL_PASSWORD, BROWSER_IMPLICIT_WAIT

    TESSERACT_CMD = Path(
        os.getenv("TESSERACT_CMD", r"C:\Program Files\Tesseract-OCR\tesseract.exe")
    )
    OMS_BASE_URL = os.getenv("OMS_BASE_URL", "http://10.36.0.142:9000")
    OMS_LOGIN = os.getenv("OMS_LOGIN", "")
    OMS_PASSWORD = os.getenv("OMS_PASSWORD", "")
    ISZL_BASE_URL = os.getenv("ISZL_BASE_URL", "http://10.36.29.2:8088")
    ISZL_LOGIN = os.getenv("ISZL_LOGIN", "")
    ISZL_PASSWORD = os.getenv("ISZL_PASSWORD", "")
    BROWSER_IMPLICIT_WAIT = int(os.getenv("BROWSER_IMPLICIT_WAIT", "10"))


def _is_placeholder(value: str) -> bool:
    v = (value or "").strip().lower()
    if not v:
        return True
    return v.startswith("your_") or v in {
        "changeme",
        "change_me",
        "xxx",
        "todo",
        "password",
        "login",
    }


def load_credentials(*extra_env_files: Path | str) -> None:
    """
    Загрузка секретов. Приоритет у BrowserAuto/.env.

    Порядок:
      1) robot credentials.env (если есть) — только ключи, которых ещё нет
      2) BrowserAuto/.env — override=True (рабочие логины всегда главнее)
      3) выкинуть плейсхолдеры your_* и снова наложить .env

    Иначе копия credentials.env.example с your_oms_login затирала .env.
    """
    for raw in extra_env_files:
        path = Path(raw)
        if path.is_file():
            load_dotenv(path, override=False)
    if ENV_FILE.is_file():
        load_dotenv(ENV_FILE, override=True)
    for key in ("OMS_LOGIN", "OMS_PASSWORD", "ISZL_LOGIN", "ISZL_PASSWORD"):
        if _is_placeholder(os.getenv(key, "")):
            os.environ.pop(key, None)
    if ENV_FILE.is_file():
        load_dotenv(ENV_FILE, override=True)
    _apply_env()


# Первичная загрузка при импорте
load_credentials()

CAPTCHA_IMAGE_DIR = PROJECT_ROOT / "data" / "captcha"
CAPTCHA_IMAGE_DIR.mkdir(parents=True, exist_ok=True)
