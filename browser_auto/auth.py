from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webdriver import WebDriver

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
) -> None:
    """Авторизация в веб-ОМС (как в талоны2_0.ipynb)."""
    user = login or OMS_LOGIN
    pwd = password or OMS_PASSWORD
    url = base_url or OMS_BASE_URL

    if not user or not pwd:
        raise ValueError("Задайте OMS_LOGIN и OMS_PASSWORD в .env")

    driver.get(url.rstrip("/") + "/")

    login_input = driver.find_element(By.XPATH, "/html/body/div/form/input[1]")
    login_input.clear()
    login_input.send_keys(user)
    input_enter_xpath(driver, "/html/body/div/form/input[2]", pwd)


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
