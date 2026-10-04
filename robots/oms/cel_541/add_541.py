# -*- coding: utf-8 -*-
"""
Ввод талонов цели 541 (КТ) в Web.ОМС — амбулаторный талон.

Несколько услуг на талон (как cel_3), заполнение услуг устойчивое (как cel_307).
Явка = 1; даты дд-мм-гг.

  python run_541.py --limit 3
  python run_541.py --talons data/talons_541.xlsx --services data/services_541.xlsx --dry-run
"""
from __future__ import annotations

import argparse
import csv
import sys
from dataclasses import asdict
from datetime import datetime
from pathlib import Path
from time import sleep, time

from loguru import logger
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webdriver import WebDriver

WORK_DIR = Path(__file__).resolve().parent
OMS_DIR = WORK_DIR.parent
PROJECT_ROOT = WORK_DIR.parents[2]
for p in (PROJECT_ROOT, WORK_DIR, OMS_DIR / "dv_opv", OMS_DIR / "cel_307", OMS_DIR / "cel_3"):
    if str(p) not in sys.path:
        sys.path.insert(0, str(p))

from browser_auto.config import load_credentials
from browser_auto.auth import login_oms
from browser_auto.driver import GeckoBrowser

import add_medical_exam as dv  # noqa: E402
import add_307 as base  # noqa: E402
import add_3 as cel3  # noqa: E402

load_credentials(WORK_DIR / "credentials.env")

RESULTS_DIR = WORK_DIR / "data"
DEFAULT_TALONS = RESULTS_DIR / "talons_541.xlsx"
DEFAULT_SERVICES = RESULTS_DIR / "services_541.xlsx"

DEFAULTS_541 = {
    "цель": "541",
    "случай": "Первичный",
    "случай_2": "Законченный",
    "ДИСП. НАБЛ": "1",
    "тип": "3",
    "результат": "301",
    "исход": "304",
    "характер": "3",
    "вид_мп": "13",  # #type-medical-care
}


def write_results(path: Path, results: list[cel3.RowResult]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fields = list(asdict(results[0]).keys()) if results else [
        "index", "talon_id", "enp", "services", "status", "message", "elapsed_sec"
    ]
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for r in results:
            w.writerow(asdict(r))


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Ввод талонов цели 541 (КТ, Web.ОМС ambulatory)")
    p.add_argument("--talons", type=Path, default=DEFAULT_TALONS)
    p.add_argument("--services", type=Path, default=DEFAULT_SERVICES)
    p.add_argument("--doctor", default="", help="Врач по умолчанию, если в Excel пусто")
    p.add_argument("--building", default="", help="Корпус по умолчанию")
    p.add_argument("--limit", type=int, default=0)
    p.add_argument("--offset", type=int, default=0)
    p.add_argument("--dry-run", action="store_true")
    p.add_argument("--keep-open", type=int, default=0)
    p.add_argument("--worker", default="")
    p.add_argument("--window-x", type=int, default=None)
    p.add_argument("--window-y", type=int, default=None)
    return p.parse_args()


def _apply_541_defaults(rows: list[cel3.Talon3]) -> None:
    for row in rows:
        row.goal = (row.goal or "").strip() or DEFAULTS_541["цель"]
        row.occurrence = (row.occurrence or "").strip() or DEFAULTS_541["случай"]
        row.occurrence2 = (row.occurrence2 or "").strip() or DEFAULTS_541["случай_2"]
        row.mo_visits = (row.mo_visits or "").strip() or "1"
        row.dn = (row.dn or "").strip() or DEFAULTS_541["ДИСП. НАБЛ"]
        row.claim_type = (row.claim_type or "").strip() or DEFAULTS_541["тип"]
        row.result = (row.result or "").strip() or DEFAULTS_541["результат"]
        row.outcome = (row.outcome or "").strip() or DEFAULTS_541["исход"]
        row.character = (row.character or "").strip() or DEFAULTS_541["характер"]


def _wait_service_combobox(driver: WebDriver, timeout: float = 15) -> None:
    """Для 541 строка услуги часто не на main: вкладка + add-service."""
    if driver.find_elements(By.ID, "services-tab"):
        try:
            dv.click_id(driver, "services-tab")
            dv.wait_busy_gone(driver, timeout=8)
            sleep(0.3)
        except Exception:
            pass

    deadline = time() + timeout
    while time() < deadline:
        if driver.find_elements(By.ID, "combobox-service-0"):
            return
        add = driver.find_elements(By.ID, "add-service")
        if add:
            try:
                add[0].click()
            except Exception:
                driver.execute_script("arguments[0].click();", add[0])
            sleep(0.6)
            dv.wait_busy_gone(driver, timeout=8)
            if driver.find_elements(By.ID, "combobox-service-0"):
                return
        sleep(0.3)
    raise RuntimeError("Нет combobox-service-0 после цели/вкладки услуг")


def fill_services_541(driver: WebDriver, row: cel3.Talon3) -> int:
    """Как cel_3.service_notebook, но даты услуги заполняем и для index=0 (541 требует)."""
    if not row.services:
        raise RuntimeError(f"Нет услуг для талона {row.talon_id}")

    filled = 0
    for idx, svc in enumerate(row.services):
        code = (svc.code or "").strip()
        if not code:
            continue
        begin = dv.normalize_service_date(svc.begin or row.begin)
        end = dv.normalize_service_date(svc.end or row.end)
        amount = (svc.amount or "1").strip() or "1"
        doctor = (svc.doctor or row.doctor or "").strip()

        if idx > 0:
            cel3._click_xpath(driver, '//*[@id="add-service"]')
            sleep(0.5)

        cel3._input_enter_xpath(driver, '//*[@id="combobox-service-0"]', code)
        sleep(0.35)

        # отличие от цели 3: для 541 даты на услуге обязательны всегда
        for sid, val in (("begin-date-service-0", begin), ("end-date-service-0", end)):
            if not driver.find_elements(By.ID, sid):
                logger.warning("нет поля {}", sid)
                continue
            try:
                dv.fill_date_field(driver, sid, val)
            except Exception as exc:
                logger.warning("{} fill_date_field: {}", sid, exc)
                base._force_input(driver, sid, val)
            actual = ""
            try:
                actual = (driver.find_element(By.ID, sid).get_attribute("value") or "").strip()
            except Exception:
                pass
            logger.info("{} after fill → {!r}", sid, actual)
            if not actual:
                base._force_input(driver, sid, val)

        cel3._input_enter_xpath(driver, '//*[@id="amount-service-0"]', amount)
        perehod = driver.find_element(By.XPATH, '//*[@id="amount-service-0"]')
        perehod.clear()
        perehod.send_keys(str(amount), Keys.ENTER)

        if doctor and (idx > 0 or driver.find_elements(By.ID, "doctor-service-0")):
            try:
                cel3._input_enter_xpath(driver, '//*[@id="doctor-service-0"]', doctor)
            except Exception as exc:
                logger.warning("doctor-service-0: {}", exc)

        sleep(0.5)
        filled += 1
        logger.info(
            "service[{}] {} amt={} dates={}/{} doctor={}",
            idx,
            code,
            amount,
            begin,
            end,
            doctor or "(главный)",
        )

    logger.info("Услуг введено: {}/{}", filled, len(row.services))
    return filled


def process_row_541(driver: WebDriver, row: cel3.Talon3, *, index: int, dry_run: bool) -> cel3.RowResult:
    """Как cel_3.process_row, но цель 541 и один проход соцстатуса."""
    t0 = time()
    logger.info(
        "=== [{}] цель541 talon={} enp={} services={} ===",
        index,
        row.talon_id,
        row.enp,
        len(row.services),
    )
    r307 = row.as_307()
    r307.goal = row.goal or "541"
    try:
        if not (row.building or "").strip():
            raise RuntimeError("Пустой Корпус — укажите в Excel или --building")
        if not (row.doctor or "").strip():
            raise RuntimeError("Пустой врач — укажите --doctor")

        base.open_ambulatory(driver)
        base.fill_patient_search(driver, r307)
        base.ensure_social_status(driver, r307)
        row.social_status, row.occupation = r307.social_status, r307.occupation

        base.fill_referral(driver, r307)
        soc_keep, occ_keep = r307.social_status, r307.occupation
        r307.social_status, r307.occupation = "", ""
        base.fill_main(driver, r307)
        r307.social_status, r307.occupation = soc_keep, occ_keep

        # как cel_3: повтор цели send_keys+ENTER → sleep(2) → шаблон услуги
        cel3._input_enter_id(driver, "requestPurpose", row.goal or "541")
        sleep(2.0)
        dv.wait_busy_gone(driver, timeout=20)

        # вид МП для цели 541
        if driver.find_elements(By.ID, "type-medical-care"):
            base._enter_select(driver, "type-medical-care", DEFAULTS_541["вид_мп"])
            logger.info("type-medical-care → {}", DEFAULTS_541["вид_мп"])
            sleep(0.3)
            dv.wait_busy_gone(driver, timeout=10)

        _wait_service_combobox(driver)
        n_svc = fill_services_541(driver, row)

        if dry_run:
            msg = f"dry-run: форма+{n_svc} услуг, без Save"
            logger.success("[{}] {}", index, msg)
            return cel3.RowResult(index, row.talon_id, row.enp, n_svc, "dry_run", msg, round(time() - t0, 2))

        ok, msg = base.save_ambulatory(driver)
        logger.info("[{}] Save: {}", index, msg)

        # повтор дат только при ошибке (соцстатус / обязательные поля с датами)
        need_date_retry = (not ok) and (
            dv._is_social_status_error(msg)
            or "обязательн" in (msg or "").lower()
            or "begin-date" in (msg or "").lower()
            or "end-date" in (msg or "").lower()
        )
        if need_date_retry:
            logger.warning("[{}] ошибка Save — дозаполняем даты и повторяем", index)
            if dv._is_social_status_error(msg):
                base.ensure_social_status(driver, r307)
                base.fill_main(driver, r307)
                cel3._input_enter_id(driver, "requestPurpose", row.goal or "541")
                sleep(1.0)
                if driver.find_elements(By.ID, "type-medical-care"):
                    base._enter_select(driver, "type-medical-care", DEFAULTS_541["вид_мп"])
                _wait_service_combobox(driver)
                n_svc = fill_services_541(driver, row)
            begin = dv.normalize_service_date(row.begin)
            end = dv.normalize_service_date(row.end)
            if driver.find_elements(By.ID, "main-tab"):
                try:
                    dv.click_id(driver, "main-tab")
                    dv.wait_busy_gone(driver, timeout=8)
                except Exception:
                    pass
            for sid, val in (("begin-date", begin), ("end-date", end), ("reference-date", begin)):
                if driver.find_elements(By.ID, sid) and val:
                    try:
                        dv.fill_date_field(driver, sid, val)
                    except Exception:
                        base._force_input(driver, sid, val)
            for sid, val in (("begin-date-service-0", begin), ("end-date-service-0", end)):
                if driver.find_elements(By.ID, sid) and val:
                    try:
                        dv.fill_date_field(driver, sid, val)
                    except Exception:
                        base._force_input(driver, sid, val)
            ok, msg = base.save_ambulatory(driver)
            logger.info("[{}] Save (повтор): {}", index, msg)

        if not ok:
            soft0 = dv.collect_form_errors(driver) or ""
            soft1 = base.dump_claim_errors(driver, click_button=False) or ""
            soft = " | ".join(p for p in (soft0, soft1, msg) if p)
            if soft:
                msg = soft
                logger.warning("[{}] детали ошибок: {}", index, soft[:500])
            try:
                (RESULTS_DIR / "last_save_fail.html").write_text(driver.page_source or "", encoding="utf-8")
            except Exception:
                pass
            sleep(3.0)
        status = "ok" if ok else "fail"
        log = logger.success if ok else logger.error
        log("[{}] {} → {} ({:.1f}s)", index, status.upper(), msg, time() - t0)
        return cel3.RowResult(index, row.talon_id, row.enp, n_svc, status, msg, round(time() - t0, 2))
    except Exception as exc:
        msg = str(exc).split("Stacktrace:")[0].strip() or str(exc)
        logger.error("[{}] error ({:.1f}s): {}", index, time() - t0, msg)
        try:
            (RESULTS_DIR / "last_541_fail.html").write_text(driver.page_source or "", encoding="utf-8")
        except Exception:
            pass
        try:
            driver.refresh()
            dv.wait_busy_gone(driver)
        except Exception:
            pass
        return cel3.RowResult(index, row.talon_id, row.enp, 0, "error", msg, round(time() - t0, 2))


def main() -> int:
    args = parse_args()
    RESULTS_DIR.mkdir(parents=True, exist_ok=True)

    talons_path = args.talons
    services_path = args.services
    if not talons_path.is_file():
        talons_path = WORK_DIR / "data" / talons_path.name
    if not services_path.is_file():
        services_path = WORK_DIR / "data" / services_path.name
    if not talons_path.is_file() or not services_path.is_file():
        logger.error(
            "Нет файлов. Сначала: python robots/oms/cel_541/template_541.py\n  talons={} services={}",
            talons_path,
            services_path,
        )
        return 1

    rows = cel3.load_talons_bundle(
        talons_path=talons_path,
        services_path=services_path,
        default_building=args.building,
        default_doctor=args.doctor,
        category="",
    )
    _apply_541_defaults(rows)

    if args.offset:
        rows = rows[args.offset :]
    if args.limit and args.limit > 0:
        rows = rows[: args.limit]

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    worker_tag = (args.worker or "").strip() or f"o{args.offset}"
    result_path = RESULTS_DIR / f"oms_541_result_{stamp}_{worker_tag}.csv"
    logger.info(
        "[{}] цель541 талонов={} doctor={!r} building={!r} dry_run={}",
        worker_tag,
        len(rows),
        args.doctor,
        args.building,
        args.dry_run,
    )
    if not rows:
        logger.error("Нет строк")
        return 1

    results: list[cel3.RowResult] = []
    t_all = time()
    win_kwargs: dict = {"implicit_wait": 0, "maximize": True}
    if args.window_x is not None or args.window_y is not None:
        win_kwargs["maximize"] = False
        win_kwargs["window_position"] = (args.window_x or 20, args.window_y or 20)
        win_kwargs["window_size"] = (1100, 850)

    with GeckoBrowser(**win_kwargs) as session:
        driver = session.driver
        driver.implicitly_wait(0)
        login_oms(driver)
        sleep(0.5)
        dv.wait_busy_gone(driver, timeout=30)

        for i, row in enumerate(rows, start=1 + args.offset):
            res = process_row_541(driver, row, index=i, dry_run=args.dry_run)
            results.append(res)
            write_results(result_path, results)

        if args.keep_open > 0:
            sleep(args.keep_open)

    write_results(result_path, results)
    ok_n = sum(1 for r in results if r.status in ("ok", "dry_run"))
    fail_n = len(results) - ok_n
    logger.info(
        "——— [{}] ИТОГ ({:.0f}s) OK={}, FAIL={} → {} ———",
        worker_tag,
        time() - t_all,
        ok_n,
        fail_n,
        result_path,
    )
    for r in results:
        logger.info("  #{} talon={} {} | {}", r.index, r.talon_id, r.status, (r.message or "")[:160])
    return 0 if fail_n == 0 and results else 1


if __name__ == "__main__":
    sys.exit(main())
