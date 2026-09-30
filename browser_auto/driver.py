from __future__ import annotations

from pathlib import Path
from time import sleep
from typing import Optional

from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service

from browser_auto.config import BROWSER_IMPLICIT_WAIT, GECKODRIVER, PROJECT_ROOT

# MIME-типы отчётов ИСЗЛ (CSV / Excel / octet-stream)
_DOWNLOAD_MIME = (
    "text/csv,text/plain,application/csv,application/excel,"
    "application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,"
    "application/octet-stream"
)


class GeckoBrowser:
    """Firefox-сессия с общим geckodriver.exe из корня проекта."""

    def __init__(
        self,
        *,
        implicit_wait: int = BROWSER_IMPLICIT_WAIT,
        maximize: bool = True,
        window_position: Optional[tuple[int, int]] = None,
        window_size: Optional[tuple[int, int]] = None,
        download_dir: Optional[Path | str] = None,
    ):
        if not GECKODRIVER.is_file():
            raise FileNotFoundError(f"geckodriver не найден: {GECKODRIVER}")

        options = Options()
        if window_position:
            x, y = window_position
            options.add_argument(f"--window-position={x},{y}")
        if window_size:
            w, h = window_size
            options.set_preference("browser.window.width", w)
            options.set_preference("browser.window.height", h)

        if download_dir is not None:
            dl = Path(download_dir)
            dl.mkdir(parents=True, exist_ok=True)
            options.set_preference("browser.download.folderList", 2)
            options.set_preference("browser.download.dir", str(dl.resolve()))
            options.set_preference("browser.download.useDownloadDir", True)
            options.set_preference("browser.download.manager.showWhenStarting", False)
            options.set_preference("browser.helperApps.neverAsk.saveToDisk", _DOWNLOAD_MIME)
            options.set_preference("pdfjs.disabled", True)

        service = Service(executable_path=str(GECKODRIVER))
        self.driver = webdriver.Firefox(service=service, options=options)
        self.driver.implicitly_wait(implicit_wait)
        self.download_dir = Path(download_dir).resolve() if download_dir else None

        if maximize and not window_size:
            self.driver.maximize_window()

        sleep(1)

    def open(self, url: str) -> None:
        self.driver.get(url)
        sleep(1)

    def close(self) -> None:
        self.driver.quit()

    def __enter__(self) -> GeckoBrowser:
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        self.close()

    @staticmethod
    def project_root() -> str:
        return str(PROJECT_ROOT)
