#!/usr/bin/env python
"""
Скачивание отчётов «по месту работы» из ИСЗЛ (меню МО).

Пункты меню (без капчи):
  22 — ДН по месту работы           → Report_цель3.csv
  13 — Диспансеризация по месту работы → Report_ДВ4.csv
  14 — Проф. осмотр по месту работы → Report_ОПВ.csv
  15 — ДР по месту работы           → Report_ДР1.csv

На каждой странице: tbMonth = «1 12» → lbtExec → lbtMsExcel.

Пример:
  python download_workplace.py
  python download_workplace.py --only цель3 ОПВ
"""

from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass
from pathlib import Path
from time import sleep, time

from loguru import logger
from selenium.common.exceptions import TimeoutException
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

WORK_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = WORK_DIR.parents[2]  # BrowserAuto (robots/iszl/workplace)
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from browser_auto.actions import click_xpath, input_xpath
from browser_auto.auth import is_iszl_logged_in, login_iszl
from browser_auto.driver import GeckoBrowser

DEFAULT_MONTH = "1 12"
DEFAULT_OUT_DIR = WORK_DIR / "data"


@dataclass(frozen=True)
class ReportJob:
    menu_li: int
    suffix: str
    title: str


# Порядок как у пользователя: ДН → ДВ → ОПВ → ДР
JOBS: tuple[ReportJob, ...] = (
    ReportJob(22, "цель3", "ДН по месту работы"),
    ReportJob(13, "ДВ4", "Диспансеризация по месту работы"),
    ReportJob(14, "ОПВ", "Проф. осмотр по месту работы"),
    ReportJob(15, "ДР1", "ДР по месту работы"),
)


def _wait(driver: WebDriver, timeout: float = 30) -> WebDriverWait:
    return WebDriverWait(driver, timeout)


def wait_blockui_gone(driver: WebDriver, timeout: float = 180) -> None:
    sleep(0.8)
    end = time() + timeout
    stable = 0
    while time() < end:
        overlays = driver.find_elements(By.CSS_SELECTOR, "div.blockUI")
        visible = [el for el in overlays if el.is_displayed()]
        if not visible:
            stable += 1
            if stable >= 2:
                return
        else:
            stable = 0
        sleep(0.4)
    raise TimeoutException(f"blockUI не исчез за {timeout} сек")


def open_menu_page(driver: WebDriver, menu_li: int, title: str) -> None:
    """hold МО → пункт меню (JS-клик) → ifMain."""
    driver.switch_to.default_content()
    sleep(0.4)

    def _open_submenu() -> None:
        menu = driver.find_element(By.XPATH, '//*[@id="mnuMO"]')
        ActionChains(driver).move_to_element(menu).click_and_hold(menu).perform()
        sleep(0.55)

    def _click_item() -> None:
        xpath_a = f'//*[@id="mnuMO:submenu:2"]/li[{menu_li}]/a'
        xpath_li = f'//*[@id="mnuMO:submenu:2"]/li[{menu_li}]'
        links = driver.find_elements(By.XPATH, xpath_a)
        el = links[0] if links else driver.find_element(By.XPATH, xpath_li)
        driver.execute_script(
            "arguments[0].scrollIntoView({block:'nearest'}); arguments[0].click();",
            el,
        )

    # два независимых клика: после первого DOM часто пересобирается
    for attempt in range(2):
        _open_submenu()
        _click_item()
        sleep(0.6)

    frame = driver.find_element(By.XPATH, '//*[@id="ifMain"]')
    driver.switch_to.frame(frame)
    _wait(driver, 20).until(EC.presence_of_element_located((By.ID, "tbMonth")))
    logger.info("Открыт «{}» (li[{}])", title, menu_li)


def run_and_ready(driver: WebDriver, *, month: str) -> None:
    input_xpath(driver, '//*[@id="tbMonth"]', month)
    click_xpath(driver, '//*[@id="lbtExec"]')
    sleep(0.5)
    wait_blockui_gone(driver, timeout=300)
    _wait(driver, 30).until(EC.element_to_be_clickable((By.ID, "lbtMsExcel")))
    logger.success("Список готов")


def click_excel(driver: WebDriver) -> None:
    wait_blockui_gone(driver, timeout=60)
    el = _wait(driver, 20).until(EC.element_to_be_clickable((By.ID, "lbtMsExcel")))
    try:
        el.click()
    except Exception:
        driver.execute_script("arguments[0].click();", el)


def _report_meta(download_dir: Path) -> dict[Path, tuple[float, int]]:
    meta: dict[Path, tuple[float, int]] = {}
    for p in download_dir.glob("Report*.csv"):
        if p.is_file() and not p.name.endswith((".part", ".crdownload")):
            st = p.stat()
            meta[p.resolve()] = (st.st_mtime, st.st_size)
    return meta


def wait_for_new_report(
    download_dir: Path,
    before_meta: dict[Path, tuple[float, int]],
    *,
    timeout: float = 600,
) -> Path:
    deadline = time() + timeout
    last_candidate: Path | None = None
    stable_since: float | None = None
    last_size = -1

    while time() < deadline:
        if any(download_dir.glob("*.part")) or any(download_dir.glob("*.crdownload")):
            sleep(0.5)
            continue

        candidates: list[Path] = []
        for path, (mtime, size) in _report_meta(download_dir).items():
            prev = before_meta.get(path)
            if prev is None or mtime > prev[0] + 0.05 or size != prev[1]:
                candidates.append(path)

        if not candidates:
            sleep(0.5)
            continue

        candidate = max(candidates, key=lambda p: p.stat().st_mtime)
        size = candidate.stat().st_size
        if size <= 0:
            sleep(0.5)
            continue

        if candidate == last_candidate and size == last_size:
            if stable_since is None:
                stable_since = time()
            elif time() - stable_since >= 1.5:
                return candidate
        else:
            last_candidate = candidate
            last_size = size
            stable_since = time()
        sleep(0.4)

    raise TimeoutError(f"CSV Report не появился в {download_dir} за {timeout} сек")


def rename_with_suffix(src: Path, suffix: str) -> Path:
    dst = src.with_name(f"Report_{suffix}{src.suffix}")
    if dst.resolve() == src.resolve():
        return dst
    if dst.exists():
        dst.unlink()
    src.rename(dst)
    return dst


def download_job(
    driver: WebDriver,
    download_dir: Path,
    job: ReportJob,
    *,
    month: str,
) -> Path:
    open_menu_page(driver, job.menu_li, job.title)
    before_meta = _report_meta(download_dir)
    run_and_ready(driver, month=month)
    click_excel(driver)
    logger.info("Ожидание скачивания «{}»...", job.suffix)
    raw = wait_for_new_report(download_dir, before_meta, timeout=600)
    final = rename_with_suffix(raw, job.suffix)
    logger.success("Сохранено: {}", final)
    return final


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Скачать отчёты «по месту работы» из ИСЗЛ")
    p.add_argument(
        "--only",
        nargs="+",
        metavar="SUFFIX",
        help="Только указанные суффиксы: цель3 ДВ4 ОПВ ДР1",
    )
    p.add_argument("--month", default=DEFAULT_MONTH, help='Период, по умолчанию "1 12"')
    p.add_argument("--out-dir", type=Path, default=DEFAULT_OUT_DIR)
    p.add_argument("--force", action="store_true", help="Перескачать, даже если файл есть")
    p.add_argument("--keep-open", type=int, default=0)
    return p.parse_args()


def main() -> int:
    args = parse_args()
    out_dir: Path = args.out_dir
    out_dir.mkdir(parents=True, exist_ok=True)

    jobs = list(JOBS)
    if args.only:
        wanted = {s.lower() for s in args.only}
        jobs = [j for j in JOBS if j.suffix.lower() in wanted]
        if not jobs:
            logger.error("Неизвестные суффиксы: {}. Доступны: {}", args.only, [j.suffix for j in JOBS])
            return 1

    logger.info("Рабочая папка: {}", WORK_DIR)
    logger.info("Задачи: {}", [(j.menu_li, j.suffix) for j in jobs])
    logger.info("Папка загрузок: {}", out_dir.resolve())

    results: list[tuple[str, Path | None, str]] = []

    with GeckoBrowser(download_dir=out_dir) as session:
        driver = session.driver
        login_iszl(driver)
        sleep(2)
        try:
            _wait(driver, 25).until(EC.presence_of_element_located((By.ID, "mnuMO")))
        except TimeoutException:
            if not is_iszl_logged_in(driver):
                logger.error("Не удалось войти в ИСЗЛ")
                return 1
            raise
        logger.success("Вход в ИСЗЛ выполнен")

        for job in jobs:
            target = out_dir / f"Report_{job.suffix}.csv"
            if not args.force and target.is_file() and target.stat().st_size > 0:
                logger.info("Пропуск {}: уже есть {}", job.suffix, target.name)
                results.append((job.suffix, target, "skip"))
                continue
            try:
                path = download_job(driver, out_dir, job, month=args.month)
                results.append((job.suffix, path, "ok"))
            except Exception as exc:
                logger.exception("{}: {}", job.suffix, exc)
                results.append((job.suffix, None, str(exc)))

        if args.keep_open > 0:
            logger.info("Браузер открыт ещё {} сек...", args.keep_open)
            sleep(args.keep_open)

    failed = 0
    for suffix, path, status in results:
        if status in ("ok", "skip"):
            logger.info("{} {} → {}", status.upper(), suffix, path)
        else:
            failed += 1
            logger.error("FAIL {} → {}", suffix, status)

    return 0 if failed == 0 and results else 1


if __name__ == "__main__":
    sys.exit(main())
