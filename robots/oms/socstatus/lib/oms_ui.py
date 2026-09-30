"""Минимальные UI-хелперы Web.ОМС для /patient/edit."""

from __future__ import annotations

import re
from datetime import datetime
from time import sleep, time

from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webdriver import WebDriver


def login_oms(driver: WebDriver, *, base_url: str, login: str, password: str) -> None:
    if not login or not password:
        raise ValueError("Задайте OMS_LOGIN и OMS_PASSWORD в credentials.env")
    driver.get(base_url.rstrip("/") + "/")
    login_input = driver.find_element(By.XPATH, "/html/body/div/form/input[1]")
    login_input.clear()
    login_input.send_keys(login)
    pwd = driver.find_element(By.XPATH, "/html/body/div/form/input[2]")
    pwd.clear()
    pwd.send_keys(password, Keys.ENTER)
    sleep(2)


def wait_busy_gone(driver: WebDriver, timeout: float = 30) -> None:
    end = time() + timeout
    while time() < end:
        try:
            texts = driver.find_elements(By.CSS_SELECTOR, "h2")
            busy = any("подождите" in (el.text or "").lower() for el in texts)
            if not busy:
                return
        except Exception:
            return
        sleep(0.25)


def dismiss_overlays(driver: WebDriver) -> None:
    for xp in (
        "//div[@role='dialog']//button[contains(.,'OK') or contains(.,'Ок') or contains(.,'Закрыть') or contains(.,'Да')]",
        "//button[@aria-label='Close' or @aria-label='close']",
    ):
        for btn in driver.find_elements(By.XPATH, xp):
            try:
                btn.click()
                sleep(0.2)
            except Exception:
                continue
    wait_busy_gone(driver, timeout=10)


def js_click(driver: WebDriver, el) -> None:
    """Всегда через JS — на /patient/edit обычный click часто ловит overlay."""
    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center', inline:'nearest'});",
        el,
    )
    sleep(0.05)
    driver.execute_script("arguments[0].click();", el)


def snackbar_text(driver: WebDriver) -> str:
    chunks: list[str] = []
    try:
        el = driver.find_element(By.ID, "client-snackbar")
        try:
            p = el.find_element(By.TAG_NAME, "p")
            text = (p.text or "").strip()
        except Exception:
            text = (el.text or "").strip()
        if text:
            chunks.append(text)
    except NoSuchElementException:
        pass
    for sel in (
        ".MuiSnackbar-root .MuiAlert-message",
        ".MuiSnackbarContent-message",
        "[class*='Snackbar'] [class*='message']",
    ):
        for el in driver.find_elements(By.CSS_SELECTOR, sel):
            try:
                if not el.is_displayed():
                    continue
                t = (el.text or "").strip()
                if t and t not in chunks:
                    chunks.append(t)
            except Exception:
                continue
    return " | ".join(chunks)


def find_visible_dialogs(driver: WebDriver) -> list:
    found = []
    seen: set[str] = set()
    for sel in ("[role='dialog']", ".MuiDialog-paper"):
        for el in driver.find_elements(By.CSS_SELECTOR, sel):
            try:
                if not el.is_displayed():
                    continue
                key = (el.get_attribute("outerHTML") or "")[:120]
                if key in seen:
                    continue
                seen.add(key)
                found.append(el)
            except Exception:
                continue
    return found


def warning_dialog_text(driver: WebDriver) -> str:
    parts: list[str] = []
    for dlg in find_visible_dialogs(driver):
        try:
            t = (dlg.text or "").strip()
        except Exception:
            t = ""
        if t:
            lines = [ln for ln in t.splitlines() if "подождите" not in ln.lower()]
            parts.append("\n".join(lines).strip())
    return "\n---\n".join(parts) if parts else ""


def _click_dialog_confirm(driver: WebDriver) -> bool:
    labels = (
        "сохранить с предупреждением",
        "для талона и пациента",
        "да",
        "ок",
        "сохранить",
        "продолжить",
    )
    for dlg in find_visible_dialogs(driver):
        buttons = dlg.find_elements(By.TAG_NAME, "button")
        ranked: list[tuple[int, object]] = []
        for btn in buttons:
            try:
                if not btn.is_displayed():
                    continue
                label = (btn.text or "").strip().lower()
            except Exception:
                continue
            for i, needle in enumerate(labels):
                if needle in label:
                    ranked.append((i, btn))
                    break
        if ranked:
            ranked.sort(key=lambda x: x[0])
            js_click(driver, ranked[0][1])
            sleep(0.3)
            return True
    return False


def handle_post_save_prompts(driver: WebDriver, *, timeout: float = 20) -> tuple[bool, str]:
    """Подтвердить предупреждения после Save (упрощённо для карты пациента)."""
    deadline = time() + timeout
    notes: list[str] = []
    while time() < deadline:
        wait_busy_gone(driver, timeout=10)
        text = warning_dialog_text(driver)
        if find_visible_dialogs(driver) or text:
            if text:
                notes.append(text[:200])
            if not _click_dialog_confirm(driver):
                return False, text or "Диалог не подтверждён"
            sleep(0.4)
            continue
        msg = snackbar_text(driver)
        if msg:
            low = msg.lower()
            if any(x in low for x in ("ошибка", "обязательн", "не указан", "дубликат")):
                return False, msg
            notes.append(msg[:200])
            return True, " | ".join(notes) if notes else msg
        if notes:
            return True, " | ".join(notes)
        sleep(0.3)
    still = warning_dialog_text(driver)
    if still or find_visible_dialogs(driver):
        return False, (still or "Диалог остался открытым")[:400]
    return True, " | ".join(notes) if notes else "ok"


def birth_year(value: str) -> int | None:
    text = (value or "").strip()
    if not text:
        return None
    m = re.search(r"(\d{1,2})[-./](\d{1,2})[-./](\d{2,4})", text)
    if not m:
        m4 = re.search(r"(19|20)\d{2}", text)
        return int(m4.group(0)) if m4 else None
    y = m.group(3)
    if len(y) == 2:
        y = f"20{y}" if int(y) <= 30 else f"19{y}"
    return int(y)


def social_occupation_by_age(age: int) -> tuple[str, str]:
    """18 < age < 60 → 11/1; иначе (в т.ч. ≥60) → 22/3."""
    if 18 < age < 60:
        return "11", "1"
    return "22", "3"


def social_occupation_from_birth(birth_raw: str, *, as_of: datetime | None = None) -> tuple[str, str]:
    year = birth_year(birth_raw)
    if year is None:
        raise ValueError(f"Не разобрать дату рождения: {birth_raw!r}")
    ref = as_of or datetime.now()
    age = ref.year - year
    return social_occupation_by_age(age)


def closest_react_select(input_el):
    el = input_el
    for _ in range(10):
        try:
            el = el.find_element(By.XPATH, "..")
        except Exception:
            return None
        cls = (el.get_attribute("class") or "").split()
        if "Select" in cls:
            return el
    return None


def fill_react_select_input(driver: WebDriver, input_el, text: str) -> None:
    wrap = closest_react_select(input_el)
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", input_el)
    sleep(0.1)
    driver.execute_script("arguments[0].focus();", input_el)
    sleep(0.1)

    opened = False
    if wrap is not None:
        for sel in (".Select-arrow-zone", ".Select-control", ".Select-arrow"):
            zones = wrap.find_elements(By.CSS_SELECTOR, sel)
            if not zones:
                continue
            try:
                driver.execute_script(
                    "arguments[0].dispatchEvent(new MouseEvent('mousedown',{bubbles:true}));"
                    "arguments[0].dispatchEvent(new MouseEvent('mouseup',{bubbles:true}));"
                    "arguments[0].click();",
                    zones[0],
                )
                opened = True
                break
            except Exception:
                continue
    if not opened:
        input_el.send_keys(Keys.ARROW_DOWN)
    sleep(0.45)

    try:
        input_el.send_keys(Keys.CONTROL, "a")
        input_el.send_keys(Keys.BACKSPACE)
    except Exception:
        pass
    input_el.send_keys(str(text))
    sleep(0.7)

    options = driver.find_elements(
        By.CSS_SELECTOR,
        ".Select-option, .Select-menu .Select-option, .Select-menu-outer .Select-option, div[role='option']",
    )
    needle = str(text).strip().lower()
    for opt in options:
        label = (opt.text or "").strip().lower()
        if not label:
            continue
        if needle == label or label.startswith(needle) or needle in label:
            try:
                driver.execute_script("arguments[0].click();", opt)
                sleep(0.25)
                return
            except Exception:
                continue
    if options:
        try:
            driver.execute_script("arguments[0].click();", options[0])
            sleep(0.25)
            return
        except Exception:
            pass
    if wrap is not None:
        shown = (wrap.text or "").strip().lower()
        if needle and needle in shown:
            input_el.send_keys(Keys.TAB)
            sleep(0.2)
            return
    input_el.send_keys(Keys.ENTER)
    sleep(0.3)
