#!/usr/bin/env python
"""Проверка входа в ИСЗЛ через общий geckodriver."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from time import sleep

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from loguru import logger

from browser_auto.auth import is_iszl_logged_in, login_iszl
from browser_auto.driver import GeckoBrowser
from browser_auto.ocr import ensure_tesseract


def main() -> int:
    parser = argparse.ArgumentParser(description="Авторизация в ИСЗЛ (Firefox + geckodriver)")
    parser.add_argument(
        "--keep-open",
        type=int,
        default=30,
        metavar="SEC",
        help="Сколько секунд держать браузер открытым после входа (0 = сразу закрыть)",
    )
    parser.add_argument(
        "--check-tesseract",
        action="store_true",
        help="Только проверить установку Tesseract",
    )
    args = parser.parse_args()

    if args.check_tesseract:
        path = ensure_tesseract()
        logger.info("Tesseract OK: {}", path)
        return 0

    try:
        ensure_tesseract()
        logger.info("Tesseract: OK")
    except FileNotFoundError as exc:
        logger.warning("{}", exc)

    logger.info("Запуск Firefox (geckodriver)...")
    with GeckoBrowser() as session:
        driver = session.driver
        login_iszl(driver)
        sleep(2)

        if is_iszl_logged_in(driver):
            logger.success("Авторизация в ИСЗЛ прошла успешно")
            title = driver.title
            logger.info("Заголовок страницы: {}", title)
        else:
            logger.error("Не удалось войти — форма логина всё ещё на экране")
            return 1

        if args.keep_open > 0:
            logger.info("Браузер открыт ещё {} сек...", args.keep_open)
            sleep(args.keep_open)

    return 0


if __name__ == "__main__":
    sys.exit(main())
