#!/usr/bin/env python
"""
Внесение плана ДН в ИСЗЛ: «Нет в текущем PlanYear, есть в прошлых».

Быстрый цикл: без переоткрытия меню между строками, implicit_wait=0,
короткий blockUI-poll. СНИЛС: SS_D (врач ДН) → SS_DOCTOR (участковый) → файл.

Пример:
  python add_carryover_plan.py --offset 0 --limit 10
  # параллель:
  python add_carryover_plan.py --worker w0 --offset 0  --limit 10 --window-x 40  --window-y 40
  python add_carryover_plan.py --worker w1 --offset 10 --limit 10 --window-x 220 --window-y 40
"""

from __future__ import annotations

import argparse
import csv
import os
import re
import sys
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from time import sleep, time

from loguru import logger
from openpyxl import load_workbook
from selenium.common.exceptions import (
    NoAlertPresentException,
    StaleElementReferenceException,
    TimeoutException,
)
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

WORK_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = WORK_DIR.parents[2]  # BrowserAuto (robots/iszl/carryover)
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from browser_auto.config import load_credentials

# Локальный credentials.env перекрывает корневой .env (удобно на другом ПК).
load_credentials(WORK_DIR / "credentials.env")

from browser_auto.actions import input_xpath
from browser_auto.auth import is_iszl_logged_in, login_iszl
from browser_auto.driver import GeckoBrowser
from browser_auto.ocr import ensure_tesseract, recognize_digits

MENU_PLANNING_DN_LI = 19
DEFAULT_FILE = WORK_DIR / "data" / "iszl_carryover_robot_2026-09-23.xlsx"
CAPTCHA_DIR = WORK_DIR / "captcha"
RESULTS_DIR = WORK_DIR / "data"

# Уникальный суффикс процесса — чтобы 5–6 роботов не перетирали капчу/CSV.
_WORKER_TAG = f"pid{os.getpid()}"

CAPTCHA_HINTS = (
    "нужно ввести цифры",
    "цифры из картинки",
    "неверно",
    "неправильн",
)


@dataclass
class CarryoverRow:
    enp: str
    dx: str
    ldw_id: str
    pid: str
    pdw_id: str
    doctor_snils: str
    plan_year: int
    source_years: str
    reason: str


@dataclass
class RowResult:
    index: int
    enp: str
    dx: str
    ldw_id: str
    pid: str
    status: str
    message: str
    plan_month: int
    plan_year: int
    snils_used: str = ""
    elapsed_sec: float = 0.0


def _wait(driver: WebDriver, timeout: float = 8) -> WebDriverWait:
    return WebDriverWait(driver, timeout, poll_frequency=0.15)


def captcha_path(name: str | None = None) -> Path:
    CAPTCHA_DIR.mkdir(parents=True, exist_ok=True)
    return CAPTCHA_DIR / (name or f"carryover_captcha_{_WORKER_TAG}.png")


def next_plan_month(now: datetime | None = None) -> int:
    now = now or datetime.now()
    return now.month + 1 if now.month < 12 else 12


def _cell_str(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _snils_digits(value: str) -> str:
    return re.sub(r"\D", "", value or "")


def _snils_variants(value: str) -> list[str]:
    d = _snils_digits(value)
    if not d:
        return []
    out = [d]
    if len(d) == 11:
        out.append(f"{d[:3]}-{d[3:6]}-{d[6:9]} {d[9:]}")
    return out


def load_carryover_rows(path: Path) -> list[CarryoverRow]:
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)
    header = [_cell_str(h) for h in next(rows_iter)]
    idx = {h.strip().lower(): i for i, h in enumerate(header) if h}

    def col(*names: str) -> int | None:
        for name in names:
            if name.lower() in idx:
                return idx[name.lower()]
        return None

    i_enp = col("ЕНП", "enp")
    i_dx = col("МКБ", "dx")
    i_ldw = col("ldwID", "ldw_id")
    i_pid = col("pID", "pid")
    i_pdw = col("pdwID", "pdw_id")
    i_snils = col("СНИЛС врача (цифры)", "СНИЛС врача", "doctor_snils")
    i_year = col("Год плана", "target_year", "plan_year")
    i_src = col("Годы источника", "source_years")
    i_reason = col("Причина", "reason")
    if i_enp is None or i_ldw is None:
        raise ValueError(f"Нет колонок ЕНП/ldwID. Заголовки: {header}")

    out: list[CarryoverRow] = []
    for raw in rows_iter:
        if not raw or all(v is None or str(v).strip() == "" for v in raw):
            continue
        enp = _cell_str(raw[i_enp])
        ldw = _cell_str(raw[i_ldw])
        if not enp or not ldw:
            continue
        snils_raw = _cell_str(raw[i_snils]) if i_snils is not None else ""
        year_raw = _cell_str(raw[i_year]) if i_year is not None else str(datetime.now().year)
        try:
            plan_year = int(float(year_raw)) if year_raw else datetime.now().year
        except ValueError:
            plan_year = datetime.now().year
        out.append(
            CarryoverRow(
                enp=enp,
                dx=_cell_str(raw[i_dx]) if i_dx is not None else "",
                ldw_id=ldw,
                pid=_cell_str(raw[i_pid]) if i_pid is not None else "",
                pdw_id=_cell_str(raw[i_pdw]) if i_pdw is not None else "",
                doctor_snils=_snils_digits(snils_raw),
                plan_year=plan_year,
                source_years=_cell_str(raw[i_src]) if i_src is not None else "",
                reason=_cell_str(raw[i_reason]) if i_reason is not None else "",
            )
        )
    return out


def ensure_main_frame(driver: WebDriver) -> None:
    if driver.find_elements(By.ID, "tbZL"):
        return
    driver.switch_to.default_content()
    frame = _wait(driver, 5).until(EC.presence_of_element_located((By.ID, "ifMain")))
    driver.switch_to.frame(frame)
    _wait(driver, 5).until(EC.presence_of_element_located((By.ID, "tbZL")))


def open_planning_dn_page(driver: WebDriver) -> None:
    driver.switch_to.default_content()
    try:
        driver.find_element(By.TAG_NAME, "body").click()
    except Exception:
        pass
    sleep(0.15)
    clickable = _wait(driver, 10).until(EC.presence_of_element_located((By.ID, "mnuMO")))
    ActionChains(driver).click_and_hold(clickable).perform()
    sleep(0.2)
    xpath = f'//*[@id="mnuMO:submenu:2"]/li[{MENU_PLANNING_DN_LI}]'
    el = _wait(driver, 5).until(EC.presence_of_element_located((By.XPATH, xpath)))
    for _ in range(2):
        try:
            el.click()
        except Exception:
            driver.execute_script("arguments[0].click();", el)
        sleep(0.1)
        el = driver.find_element(By.XPATH, xpath)
    sleep(0.4)
    ensure_main_frame(driver)
    logger.info("Открыто «Планирование ДН» (li[{}])", MENU_PLANNING_DN_LI)


def wait_blockui_gone(driver: WebDriver, timeout: float = 12) -> None:
    """Быстрый poll blockUI — без длинных пауз."""
    end = time() + timeout
    clear_hits = 0
    while time() < end:
        visible = False
        try:
            for el in driver.find_elements(By.CSS_SELECTOR, "div.blockUI"):
                try:
                    if el.is_displayed():
                        visible = True
                        break
                except StaleElementReferenceException:
                    continue
        except Exception:
            visible = False
        if not visible:
            clear_hits += 1
            if clear_hits >= 1:
                return
        else:
            clear_hits = 0
        sleep(0.12)
    # не валимся — UI часто уже готов
    logger.debug("blockUI timeout {}s — продолжаем", timeout)


def _captcha_failed(driver: WebDriver) -> bool:
    try:
        text = driver.page_source.lower()
    except Exception:
        return False
    return any(h in text for h in CAPTCHA_HINTS)


def read_captcha(driver: WebDriver) -> str:
    path = captcha_path()
    imgs = driver.find_elements(By.CSS_SELECTOR, "img[src*='wfImg']")
    if imgs:
        imgs[0].screenshot(str(path))
    else:
        frame = driver.find_element(By.ID, "ifs")
        frame.screenshot(str(path))
    # auto_trim в OCR срезает светлую рамку справа/снизу у screenshot элемента
    code = recognize_digits(path, expected_len=4, auto_trim=True)
    digits = "".join(ch for ch in code if ch.isdigit())
    logger.info("Капча OCR: {!r} → {!r}", code, digits)
    return digits


def accept_alert_if_any(driver: WebDriver, timeout: float = 0.4) -> str:
    try:
        WebDriverWait(driver, timeout, poll_frequency=0.1).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        text = alert.text
        alert.accept()
        return text
    except (TimeoutException, NoAlertPresentException):
        return ""


def click_id(driver: WebDriver, element_id: str) -> None:
    el = driver.find_element(By.ID, element_id)
    try:
        el.click()
    except Exception:
        driver.execute_script("arguments[0].click();", el)


def search_by_enp(driver: WebDriver, enp: str, *, captcha_attempts: int = 5) -> None:
    ensure_main_frame(driver)
    for attempt in range(1, captcha_attempts + 1):
        input_xpath(driver, '//*[@id="tbZL"]', enp)
        code = read_captcha(driver)
        if len(code) != 4:
            logger.warning("OCR не 4 цифры ({!r}), попытка {}", code, attempt)
            sleep(0.2)
            continue
        input_xpath(driver, '//*[@id="tbValueStr"]', code)
        clicked = False
        for btn_id in ("lbtExec", "lbtnExec"):
            if driver.find_elements(By.ID, btn_id):
                click_id(driver, btn_id)
                clicked = True
                break
        if not clicked:
            raise RuntimeError("Нет кнопки lbtExec/lbtnExec")

        wait_blockui_gone(driver, timeout=6)
        if _captcha_failed(driver):
            logger.warning("Капча не принята, попытка {}", attempt)
            continue
        if driver.find_elements(By.ID, "gvLstDW"):
            logger.success("Поиск {} ок (попытка {})", enp, attempt)
            return
        try:
            _wait(driver, 2.5).until(EC.presence_of_element_located((By.ID, "gvLstDW")))
            logger.success("Поиск {} ок (попытка {})", enp, attempt)
            return
        except TimeoutException:
            if attempt == captcha_attempts:
                raise RuntimeError(f"Нет gvLstDW после поиска ЕНП {enp}")
    raise RuntimeError(f"ЕНП {enp}: капча не прошла за {captcha_attempts} попыток")


def _table_snils_from_row(table, row_idx: str) -> str:
    """SS_D (врач на ДН) → SS_DOCTOR (участковый)."""
    if row_idx != "":
        for elem_id in (f"gvLstDW_lbSS_D_{row_idx}", f"gvLstDW_lbSS_DOCTOR_{row_idx}"):
            els = table.find_elements(By.ID, elem_id)
            if els:
                digits = _snils_digits(els[0].text)
                if digits:
                    return digits
    return ""


def select_row_by_ids(driver: WebDriver, *, ldw_id: str, pid: str) -> str:
    """Клик по строке. Возвращает СНИЛС врача ДН из таблицы (SS_D → SS_DOCTOR)."""
    table = _wait(driver, 6).until(EC.presence_of_element_located((By.ID, "gvLstDW")))
    spans = table.find_elements(By.CSS_SELECTOR, "span[id^='gvLstDW_lblstId_']")
    if not spans:
        raise RuntimeError("gvLstDW пуст")

    target_span = None
    row_idx = ""
    for span in spans:
        if _cell_str(span.text) != ldw_id:
            continue
        m = re.search(r"_(\d+)$", span.get_attribute("id") or "")
        row_idx = m.group(1) if m else ""
        if pid and row_idx:
            pid_els = table.find_elements(By.ID, f"gvLstDW_lbPID_{row_idx}")
            if pid_els and _cell_str(pid_els[0].text) not in ("", pid):
                continue
        target_span = span
        break

    if target_span is None:
        raise RuntimeError(f"Строка ldwID={ldw_id} pID={pid} не найдена")

    row = target_span.find_element(By.XPATH, "./ancestor::tr[1]")
    table_snils = _table_snils_from_row(table, row_idx)

    try:
        row.click()
    except Exception:
        driver.execute_script("arguments[0].click();", row)
    wait_blockui_gone(driver, timeout=5)
    _wait(driver, 4).until(EC.presence_of_element_located((By.ID, "tbWEditFindDoct")))
    logger.info("Строка ldwID={} pID={} | СНИЛС из таблицы={}", ldw_id, pid, table_snils or "—")
    return table_snils


def try_set_doctor(driver: WebDriver, snils: str, *, wait_add: float = 2.5) -> bool:
    """True если врач найден и выбран."""
    if not snils:
        return False
    field = driver.find_element(By.ID, "tbWEditFindDoct")
    field.clear()
    field.send_keys(snils)
    click_id(driver, "ibtnWEditFindDoct")
    sleep(0.15)
    wait_blockui_gone(driver, timeout=3)
    # AJAX без blockUI — коротко ждём кнопку «добавить»
    end = time() + wait_add
    add_btn = None
    while time() < end:
        els = driver.find_elements(By.ID, "gvWEditDoct_ibtAdd_0")
        if els:
            try:
                if els[0].is_displayed():
                    add_btn = els[0]
                    break
            except StaleElementReferenceException:
                pass
        sleep(0.12)
    if add_btn is None:
        return False
    try:
        add_btn.click()
    except Exception:
        driver.execute_script("arguments[0].click();", add_btn)
    wait_blockui_gone(driver, timeout=3)
    return True


def set_doctor(driver: WebDriver, file_snils: str, table_snils: str) -> str:
    """СНИЛС: файл (SS_D) → таблица. Возвращает использованный."""
    candidates: list[str] = []
    for s in (file_snils, table_snils):
        for v in _snils_variants(s):
            if v not in candidates:
                candidates.append(v)
    if not candidates:
        raise RuntimeError("Нет СНИЛС ни в файле (SS_D), ни в таблице")

    for snils in candidates:
        logger.info("Пробуем врача СНИЛС {!r}", snils)
        if try_set_doctor(driver, snils, wait_add=2.5):
            logger.success("Врач {!r} выбран", snils)
            return _snils_digits(snils) or snils
    raise RuntimeError(f"Врач не найден. Пробовали: {candidates}")


def fill_plan_month_year(driver: WebDriver, *, month: int, year: int) -> None:
    input_xpath(driver, '//*[@id="tbWEditMonth"]', str(month))
    y = driver.find_element(By.ID, "tbWEditYear")
    y.clear()
    y.send_keys(str(year))


def save_plan(driver: WebDriver) -> str:
    click_id(driver, "lbtnWEditSavePlanQ")
    alert_text = accept_alert_if_any(driver, timeout=0.5)
    wait_blockui_gone(driver, timeout=10)
    confirm = _wait(driver, 6).until(EC.element_to_be_clickable((By.ID, "lbtnWEditSaveList")))
    try:
        confirm.click()
    except Exception:
        driver.execute_script("arguments[0].click();", confirm)
    alert_text2 = accept_alert_if_any(driver, timeout=0.5)
    wait_blockui_gone(driver, timeout=12)
    try:
        info = _cell_str(driver.find_element(By.ID, "lbInfoPlan").text)
    except Exception:
        info = alert_text2 or alert_text or "(нет lbInfoPlan)"
    return info


def reset_to_search(driver: WebDriver) -> None:
    """Вернуться к поиску без переоткрытия меню."""
    ensure_main_frame(driver)
    if driver.find_elements(By.ID, "tbZL"):
        return
    # если ушли из формы — открываем меню заново
    open_planning_dn_page(driver)


def process_row(
    driver: WebDriver,
    row: CarryoverRow,
    *,
    index: int,
    plan_month: int,
    captcha_attempts: int,
) -> RowResult:
    t0 = time()
    logger.info("=== [{}] {} {} ldw={} ===", index, row.enp, row.dx, row.ldw_id)
    snils_used = ""
    try:
        search_by_enp(driver, row.enp, captcha_attempts=captcha_attempts)
        table_snils = select_row_by_ids(driver, ldw_id=row.ldw_id, pid=row.pid)
        snils_used = set_doctor(driver, row.doctor_snils, table_snils)
        fill_plan_month_year(driver, month=plan_month, year=row.plan_year)
        message = save_plan(driver)
        low = message.lower()
        if "уже" in low and "добавл" in low:
            status = "skip"
        elif "успеш" in low or "добавлен" in low:
            status = "ok"
        else:
            status = "fail"
        if status in ("ok", "skip"):
            logger.success("[{}] {} → {} ({:.1f}s)", index, status.upper(), message, time() - t0)
        else:
            logger.error("[{}] FAIL → {} ({:.1f}s)", index, message, time() - t0)
        return RowResult(
            index=index,
            enp=row.enp,
            dx=row.dx,
            ldw_id=row.ldw_id,
            pid=row.pid,
            status=status,
            message=message,
            plan_month=plan_month,
            plan_year=row.plan_year,
            snils_used=snils_used,
            elapsed_sec=round(time() - t0, 2),
        )
    except Exception as exc:
        msg = str(exc).split("Stacktrace:")[0].strip() or str(exc)
        logger.error("[{}] error ({:.1f}s): {}", index, time() - t0, msg)
        return RowResult(
            index=index,
            enp=row.enp,
            dx=row.dx,
            ldw_id=row.ldw_id,
            pid=row.pid,
            status="error",
            message=msg,
            plan_month=plan_month,
            plan_year=row.plan_year,
            snils_used=snils_used,
            elapsed_sec=round(time() - t0, 2),
        )


def write_results(path: Path, results: list[RowResult]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(
            f,
            fieldnames=[
                "index",
                "enp",
                "dx",
                "ldw_id",
                "pid",
                "status",
                "message",
                "plan_month",
                "plan_year",
                "snils_used",
                "elapsed_sec",
            ],
        )
        w.writeheader()
        for r in results:
            w.writerow(asdict(r))


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Внесение плана ДН (перенос), быстрый режим")
    p.add_argument("--file", type=Path, default=DEFAULT_FILE)
    p.add_argument("--limit", type=int, default=0)
    p.add_argument("--offset", type=int, default=0)
    p.add_argument("--month", type=int, default=0)
    p.add_argument("--captcha-attempts", type=int, default=5)
    p.add_argument("--keep-open", type=int, default=0)
    p.add_argument(
        "--worker",
        type=str,
        default="",
        help="Метка воркера (для параллели: уникальные капча/CSV/окно)",
    )
    p.add_argument("--window-x", type=int, default=-1)
    p.add_argument("--window-y", type=int, default=-1)
    return p.parse_args()


def main() -> int:
    global _WORKER_TAG
    args = parse_args()
    if args.worker:
        _WORKER_TAG = re.sub(r"[^\w\-]+", "_", args.worker.strip()) or _WORKER_TAG
    ensure_tesseract()
    CAPTCHA_DIR.mkdir(parents=True, exist_ok=True)
    RESULTS_DIR.mkdir(parents=True, exist_ok=True)

    xlsx = args.file
    if not xlsx.is_file():
        logger.error("Файл не найден: {}", xlsx)
        return 1

    rows = load_carryover_rows(xlsx)
    if args.offset:
        rows = rows[args.offset :]
    if args.limit and args.limit > 0:
        rows = rows[: args.limit]

    plan_month = args.month if args.month and 1 <= args.month <= 12 else next_plan_month()
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    result_path = RESULTS_DIR / f"carryover_result_{_WORKER_TAG}_o{args.offset}_{stamp}.csv"

    logger.info(
        "Воркер={} | строк={} | месяц={} | offset={}",
        _WORKER_TAG,
        len(rows),
        plan_month,
        args.offset,
    )
    results: list[RowResult] = []
    t_all = time()

    win_kwargs: dict = {"maximize": False, "window_size": (1100, 800)}
    if args.window_x >= 0 and args.window_y >= 0:
        win_kwargs["window_position"] = (args.window_x, args.window_y)
    else:
        win_kwargs["maximize"] = True
        win_kwargs.pop("window_size", None)

    with GeckoBrowser(implicit_wait=0, **win_kwargs) as session:
        driver = session.driver
        driver.implicitly_wait(0)
        login_iszl(driver)
        sleep(0.8)
        try:
            _wait(driver, 15).until(EC.presence_of_element_located((By.ID, "mnuMO")))
        except TimeoutException:
            if not is_iszl_logged_in(driver):
                logger.error("Не удалось войти в ИСЗЛ")
                return 1
            raise
        logger.success("Вход в ИСЗЛ")
        open_planning_dn_page(driver)

        for i, row in enumerate(rows, start=1 + args.offset):
            res = process_row(
                driver,
                row,
                index=i,
                plan_month=plan_month,
                captcha_attempts=args.captcha_attempts,
            )
            results.append(res)
            try:
                reset_to_search(driver)
            except Exception as reopen_exc:
                logger.error("Сброс формы: {}", reopen_exc)
                try:
                    open_planning_dn_page(driver)
                except Exception:
                    break

        if args.keep_open > 0:
            sleep(args.keep_open)

    write_results(result_path, results)
    ok_n = sum(1 for r in results if r.status in ("ok", "skip"))
    fail_n = sum(1 for r in results if r.status not in ("ok", "skip"))
    total = time() - t_all
    avg = (sum(r.elapsed_sec for r in results) / len(results)) if results else 0.0
    logger.info("——— ИТОГ ({:.0f}s, avg {:.1f}s/строка) ———", total, avg)
    for r in results:
        logger.info(
            "#{} {} | {} | {} | snils={} | {:.1f}s | {}",
            r.index,
            r.status.upper(),
            r.enp,
            r.dx,
            r.snils_used or "—",
            r.elapsed_sec,
            r.message,
        )
    logger.info("OK+SKIP={}, FAIL={}, лог: {}", ok_n, fail_n, result_path)
    return 0 if fail_n == 0 and results else 1


if __name__ == "__main__":
    sys.exit(main())
