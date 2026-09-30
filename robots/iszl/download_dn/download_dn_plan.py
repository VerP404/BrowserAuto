#!/usr/bin/env python
"""
Скачивание CSV «План ДН общий» из ИСЗЛ (меню МО).

Расположение: robots/iszl/download_dn/

Пример:
  python download_dn_plan.py
  python download_dn_plan.py --year 2025
  python download_dn_plan.py --from-year 2026 --to-year 2018
"""

from __future__ import annotations

import argparse
import re
import sys
from datetime import datetime
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
# BrowserAuto/robots/iszl/download_dn → parents[2] = BrowserAuto
PROJECT_ROOT = WORK_DIR.parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from browser_auto.actions import click_xpath, input_xpath
from browser_auto.auth import is_iszl_logged_in, login_iszl
from browser_auto.driver import GeckoBrowser
from browser_auto.ocr import ensure_tesseract, recognize_digits

# МО → «План ДН общий» (li[20] в submenu:2)
MENU_DN_PLAN_LI = 20
DEFAULT_MONTH = "1 12"
DEFAULT_OUT_DIR = WORK_DIR / "data"
CAPTCHA_DIR = WORK_DIR / "captcha"
CAPTCHA_HINTS = (
    "нужно ввести цифры",
    "цифры из картинки",
    "неверно",
    "неправильн",
)


def _wait(driver: WebDriver, timeout: float = 30) -> WebDriverWait:
    return WebDriverWait(driver, timeout)


def captcha_path(name: str = "dn_plan_captcha.png") -> Path:
    CAPTCHA_DIR.mkdir(parents=True, exist_ok=True)
    return CAPTCHA_DIR / name


def existing_year_file(download_dir: Path, year: int) -> Path | None:
    """Уже скачанный файл года: Report_2026.csv или * _2026.csv."""
    direct = download_dir / f"Report_{year}.csv"
    if direct.is_file() and direct.stat().st_size > 0:
        return direct
    for p in download_dir.glob(f"*_{year}.csv"):
        if p.is_file() and p.stat().st_size > 0:
            return p
    return None


def open_dn_plan_page(driver: WebDriver) -> None:
    """Как в «внести в план ИСЗЛ.ipynb»: hold МО → клик пункта ×2 → ifMain."""
    driver.switch_to.default_content()
    clickable = driver.find_element(By.XPATH, '//*[@id="mnuMO"]')
    ActionChains(driver).click_and_hold(clickable).perform()
    xpath = f'//*[@id="mnuMO:submenu:2"]/li[{MENU_DN_PLAN_LI}]'
    perehod = driver.find_element(By.XPATH, xpath)
    perehod.click()
    perehod = driver.find_element(By.XPATH, xpath)
    perehod.click()
    sleep(1.0)

    perehod_v_frame = driver.find_element(By.XPATH, '//*[@id="ifMain"]')
    driver.switch_to.frame(perehod_v_frame)
    _wait(driver, 20).until(EC.presence_of_element_located((By.ID, "tbYear")))
    logger.info("Открыт «План ДН общий» (li[{}], double-click)", MENU_DN_PLAN_LI)


def _page_hint(driver: WebDriver) -> str:
    try:
        return driver.page_source.lower()
    except Exception:
        return ""


def _captcha_failed(driver: WebDriver) -> bool:
    text = _page_hint(driver)
    return any(h in text for h in CAPTCHA_HINTS)


def read_captcha(driver: WebDriver) -> str:
    """Скрин iframe#ifs (wfImg.aspx) — без координат экрана."""
    frame = _wait(driver, 15).until(EC.presence_of_element_located((By.ID, "ifs")))
    path = captcha_path()
    frame.screenshot(str(path))

    code = recognize_digits(path, expected_len=4)
    digits = "".join(ch for ch in code if ch.isdigit())
    logger.info("Капча OCR: {!r} → {!r}", code, digits)
    return digits


def fill_filters(driver: WebDriver, *, year: int, month: str) -> None:
    input_xpath(driver, '//*[@id="tbMonth"]', month)
    input_xpath(driver, '//*[@id="tbYear"]', str(year))


def wait_blockui_gone(driver: WebDriver, timeout: float = 120) -> None:
    """Ждём исчезновения jQuery blockUI поверх формы."""
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


def run_report(driver: WebDriver, *, year: int, month: str, captcha_attempts: int) -> None:
    for attempt in range(1, captcha_attempts + 1):
        fill_filters(driver, year=year, month=month)
        code = read_captcha(driver)
        if len(code) != 4:
            logger.warning("Попытка {}: OCR не дал 4 цифры ({!r})", attempt, code)
            sleep(0.8)
            continue

        input_xpath(driver, '//*[@id="tbValueStr"]', code)
        click_xpath(driver, '//*[@id="lbtnExec"]')
        sleep(0.5)
        try:
            wait_blockui_gone(driver, timeout=300)
            _wait(driver, 30).until(EC.element_to_be_clickable((By.ID, "lbtnMsExcel")))
        except TimeoutException:
            if _captcha_failed(driver):
                logger.warning("Попытка {}: капча не принята", attempt)
                continue
            raise TimeoutException(
                f"Год {year}: отчёт не готов за отведённое время (blockUI / Excel)"
            )

        if _captcha_failed(driver):
            logger.warning("Попытка {}: Excel есть, но страница намекает на ошибку капчи", attempt)
            continue

        logger.success("Отчёт за {} готов (капча с попытки {})", year, attempt)
        return

    raise RuntimeError(f"Год {year}: капча не распознана за {captcha_attempts} попыток")


def click_excel(driver: WebDriver) -> None:
    wait_blockui_gone(driver, timeout=60)
    el = _wait(driver, 20).until(EC.element_to_be_clickable((By.ID, "lbtnMsExcel")))
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
    """Ждём новый или перезаписанный Report*.csv со стабильным размером."""
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


def rename_with_year(src: Path, year: int) -> Path:
    # Report.csv → Report_2026.csv; Report - ....csv → Report - ...._2026.csv
    if src.stem.lower() == "report":
        dst = src.with_name(f"Report_{year}{src.suffix}")
    else:
        stem = re.sub(r"_\d{4}$", "", src.stem)
        dst = src.with_name(f"{stem}_{year}{src.suffix}")
    if dst.resolve() == src.resolve():
        return dst
    if dst.exists():
        dst.unlink()
    src.rename(dst)
    return dst


def download_year(
    driver: WebDriver,
    download_dir: Path,
    *,
    year: int,
    month: str,
    captcha_attempts: int,
) -> Path:
    before_meta = _report_meta(download_dir)
    run_report(driver, year=year, month=month, captcha_attempts=captcha_attempts)

    click_excel(driver)
    logger.info("Ожидание скачивания CSV за {}...", year)
    raw = wait_for_new_report(download_dir, before_meta, timeout=600)
    final = rename_with_year(raw, year)
    logger.success("Сохранено: {}", final)
    return final


def year_range(from_year: int, to_year: int) -> list[int]:
    if from_year >= to_year:
        return list(range(from_year, to_year - 1, -1))
    return list(range(from_year, to_year + 1))


def parse_args() -> argparse.Namespace:
    current = datetime.now().year
    p = argparse.ArgumentParser(description="Скачать «План ДН общий» из ИСЗЛ")
    p.add_argument("--year", type=int, help="Один год (например 2026)")
    p.add_argument("--from-year", type=int, default=None, help="Начальный год цикла")
    p.add_argument("--to-year", type=int, default=2018, help="Конечный год цикла (по умолчанию 2018)")
    p.add_argument("--month", default=DEFAULT_MONTH, help='Период месяцев, по умолчанию "1 12"')
    p.add_argument(
        "--out-dir",
        type=Path,
        default=DEFAULT_OUT_DIR,
        help="Папка для CSV (по умолчанию ./data рядом со скриптом)",
    )
    p.add_argument("--captcha-attempts", type=int, default=8)
    p.add_argument("--force", action="store_true", help="Перескачать даже если файл года уже есть")
    p.add_argument("--keep-open", type=int, default=0, help="Секунд держать браузер после работы")
    args = p.parse_args()

    if args.year is not None:
        args.years = [args.year]
    else:
        start = args.from_year if args.from_year is not None else current
        args.years = year_range(start, args.to_year)
    return args


def main() -> int:
    args = parse_args()
    ensure_tesseract()
    out_dir: Path = args.out_dir
    out_dir.mkdir(parents=True, exist_ok=True)
    CAPTCHA_DIR.mkdir(parents=True, exist_ok=True)

    logger.info("Рабочая папка: {}", WORK_DIR)
    logger.info("Годы: {}", args.years)
    logger.info("Папка загрузок: {}", out_dir.resolve())

    results: list[tuple[int, Path | None, str]] = []

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

        open_dn_plan_page(driver)

        for year in args.years:
            if not args.force:
                existing = existing_year_file(out_dir, year)
                if existing:
                    logger.info("Пропуск {}: уже есть {}", year, existing.name)
                    results.append((year, existing, "skip"))
                    continue
            try:
                path = download_year(
                    driver,
                    out_dir,
                    year=year,
                    month=args.month,
                    captcha_attempts=args.captcha_attempts,
                )
                results.append((year, path, "ok"))
            except Exception as exc:
                logger.exception("Год {}: {}", year, exc)
                results.append((year, None, str(exc)))

        if args.keep_open > 0:
            logger.info("Браузер открыт ещё {} сек...", args.keep_open)
            sleep(args.keep_open)

    failed = 0
    for year, path, status in results:
        if status in ("ok", "skip"):
            logger.info("{} {} → {}", status.upper(), year, path)
        else:
            failed += 1
            logger.error("FAIL {} → {}", year, status)

    return 0 if failed == 0 and results else 1


if __name__ == "__main__":
    sys.exit(main())
