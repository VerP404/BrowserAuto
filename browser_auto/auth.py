from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support.ui import WebDriverWait

from browser_auto.actions import input_enter_xpath
from browser_auto.config import (
    ISZL_BASE_URL,
    ISZL_LOGIN,
    ISZL_PASSWORD,
    OMS_BASE_URL,
    OMS_LOGIN,
    OMS_PASSWORD,
)


def login_oms(
    driver: WebDriver,
    *,
    login: str | None = None,
    password: str | None = None,
    base_url: str | None = None,
    timeout: float = 60,
) -> None:
    """Авторизация в веб-ОМС (как в талоны2_0.ipynb).

    Важно: после Enter идёт OAuth (:9090) — ждём возврат на приложение,
    иначе следующий driver.get() срывает сессию.
    """
    from browser_auto import config as cfg

    user = login or cfg.OMS_LOGIN or OMS_LOGIN
    pwd = password or cfg.OMS_PASSWORD or OMS_PASSWORD
    url = (base_url or cfg.OMS_BASE_URL or OMS_BASE_URL or "").rstrip("/")

    if not user or not pwd:
        raise ValueError("Задайте OMS_LOGIN и OMS_PASSWORD в .env")

    driver.get(url + "/")

    login_input = driver.find_element(By.XPATH, "/html/body/div/form/input[1]")
    login_input.clear()
    login_input.send_keys(user)
    input_enter_xpath(driver, "/html/body/div/form/input[2]", pwd)

    def _logged_in(d: WebDriver) -> bool:
        u = (d.current_url or "").lower()
        if "oauth" in u or "/auth?" in u or u.rstrip("/").endswith("/auth"):
            return False
        # форма логина ещё на экране
        try:
            forms = d.find_elements(By.XPATH, "/html/body/div/form/input[1]")
            if forms and forms[0].is_displayed() and "9000" not in (d.current_url or ""):
                return False
        except Exception:
            pass
        return "9000" in (d.current_url or "") or "claim" in u or "journal" in u or u.rstrip("/").endswith(":9000")

    WebDriverWait(driver, timeout).until(_logged_in)



def login_iszl(
    driver: WebDriver,
    *,
    login: str | None = None,
    password: str | None = None,
    base_url: str | None = None,
) -> None:
    """Авторизация в ИСЗЛ (как в mass_select.py / внести в план ИСЗЛ.ipynb)."""
    user = login or ISZL_LOGIN
    pwd = password or ISZL_PASSWORD
    url = base_url or ISZL_BASE_URL

    if not user or not pwd:
        raise ValueError("Задайте ISZL_LOGIN и ISZL_PASSWORD в .env")

    driver.get(url.rstrip("/") + "/")

    login_input = driver.find_element(By.XPATH, '//*[@id="tbLogin"]')
    login_input.clear()
    login_input.send_keys(user)
    input_enter_xpath(driver, '//*[@id="tbPwd"]', pwd)


def is_iszl_logged_in(driver: WebDriver) -> bool:
    """Проверка, что форма логина ИСЗЛ больше не видна."""
    try:
        driver.find_element(By.XPATH, '//*[@id="tbLogin"]')
        return False
    except Exception:
        return True
