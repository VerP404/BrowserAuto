"""Firefox + geckodriver из корня пакета."""

from __future__ import annotations

from pathlib import Path
from time import sleep
from typing import Optional

from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service

from lib.config import GECKODRIVER, browser_implicit_wait


class GeckoBrowser:
    def __init__(
        self,
        *,
        implicit_wait: int | None = None,
        maximize: bool = True,
        window_position: Optional[tuple[int, int]] = None,
        window_size: Optional[tuple[int, int]] = None,
    ):
        if not GECKODRIVER.is_file():
            raise FileNotFoundError(
                f"geckodriver.exe не найден: {GECKODRIVER}\n"
                "Положите geckodriver.exe в папку проекта или запустите install.bat"
            )
        options = Options()
        if window_position:
            x, y = window_position
            options.add_argument(f"--window-position={x},{y}")
        if window_size:
            w, h = window_size
            options.set_preference("browser.window.width", w)
            options.set_preference("browser.window.height", h)

        service = Service(executable_path=str(GECKODRIVER))
        self.driver = webdriver.Firefox(service=service, options=options)
        self.driver.implicitly_wait(implicit_wait if implicit_wait is not None else browser_implicit_wait())

        if maximize and not window_size:
            self.driver.maximize_window()
        elif window_size:
            self.driver.set_window_size(*window_size)
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
