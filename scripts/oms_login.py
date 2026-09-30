#!/usr/bin/env python
"""Проверка входа в веб-ОМС через общий geckodriver."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from time import sleep

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from loguru import logger

from browser_auto.auth import login_oms
from browser_auto.driver import GeckoBrowser


def main() -> int:
    parser = argparse.ArgumentParser(description="Авторизация в веб-ОМС")
    parser.add_argument("--keep-open", type=int, default=30, metavar="SEC")
    args = parser.parse_args()

    logger.info("Запуск Firefox (geckodriver)...")
    with GeckoBrowser() as session:
        login_oms(session.driver)
        sleep(3)
        logger.success("Вход выполнен, URL: {}", session.driver.current_url)

        if args.keep_open > 0:
            sleep(args.keep_open)

    return 0


if __name__ == "__main__":
    sys.exit(main())
