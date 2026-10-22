#!/usr/bin/env python
"""
Замена кода услуги в уже созданных амбулаторных талонах Web.ОМС (цель 3 и др.).

Excel (колонки):
  Талон | было | стало
  39927692 | B04.015.003 | B04.047.003

Открываем {OMS}/claim/ambulatory/{Талон}, ищем строку услуги с кодом «было»
в #combobox-service-{i}, ставим «стало», затем Save.

Примеры:
  python run_fix_service.py --limit 1 --dry-run --keep-open 20
  python run_fix_service.py --file services.xlsx --limit 10
  python run_fix_service.py --file talons.xlsx --from B04.015.003 --to B04.047.003
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from time import sleep, time

from loguru import logger
from openpyxl import Workbook, load_workbook
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

WORK_DIR = Path(__file__).resolve().parent
OMS_DIR = WORK_DIR.parent
PROJECT_ROOT = WORK_DIR.parents[2]
for p in (PROJECT_ROOT, WORK_DIR, OMS_DIR / "dv_opv", OMS_DIR / "cel_307"):
    if str(p) not in sys.path:
        sys.path.insert(0, str(p))

from browser_auto.config import OMS_BASE_URL, load_credentials
from browser_auto.auth import login_oms
from browser_auto.driver import GeckoBrowser

import add_medical_exam as dv  # noqa: E402
import add_307 as base  # noqa: E402

load_credentials(WORK_DIR / "credentials.env", PROJECT_ROOT / ".env")

RESULTS_DIR = WORK_DIR / "data"
DEFAULT_FILE = RESULTS_DIR / "fix_service.xlsx"

_SERVICE_CODE_RE = re.compile(
    r"[A-Za-z]\d{2}\.\d{2,3}\.\d{3}(?:\.\d+)*",
    re.IGNORECASE,
)


@dataclass
class FixRow:
    talon: str
    code_from: str
    code_to: str


@dataclass
class RowResult:
    index: int
    talon: str
    code_from: str
    code_to: str
    status: str
    message: str
    replaced: int
    elapsed_sec: float


def _cell(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v).strip()


def normalize_service_code(value: object) -> str:
    text = _cell(value).upper().replace(" ", "")
    if not text:
        return ""
    m = _SERVICE_CODE_RE.search(text)
    return m.group(0).upper() if m else text


def load_rows(
    path: Path,
    *,
    default_from: str = "",
    default_to: str = "",
    skip_done: bool = True,
) -> list[FixRow]:
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    headers = [
        (_cell(c.value).lower() if c.value is not None else "")
        for c in next(ws.iter_rows(min_row=1, max_row=1))
    ]
    col = {h: i for i, h in enumerate(headers) if h}

    def idx(*names: str) -> int | None:
        for n in names:
            if n in col:
                return col[n]
        return None

    i_talon = idx("талон", "номер", "id", "claim", "n", "web-oms", "web_oms")
    i_from = idx(
        "было",
        "from",
        "old",
        "услуга_было",
        "код_было",
        "service_from",
        "from_code",
    )
    i_to = idx(
        "стало",
        "to",
        "new",
        "услуга_стало",
        "код_стало",
        "service_to",
        "to_code",
    )
    i_status = idx("status", "статус")

    if i_talon is None:
        raise SystemExit(f"Нужна колонка «Талон», сейчас: {headers}")

    def_from = normalize_service_code(default_from)
    def_to = normalize_service_code(default_to)
    if i_from is None and not def_from:
        raise SystemExit(
            f"Нужна колонка «было» или флаг --from. Сейчас: {headers}"
        )
    if i_to is None and not def_to:
        raise SystemExit(
            f"Нужна колонка «стало» или флаг --to. Сейчас: {headers}"
        )

    out: list[FixRow] = []
    for raw in ws.iter_rows(min_row=2, values_only=True):
        if not raw or not any(v is not None and str(v).strip() != "" for v in raw):
            continue
        talon = _cell(raw[i_talon])
        if not talon:
            continue
        if talon.endswith(".0"):
            talon = talon[:-2]
        # CLI --from/--to перекрывают ячейки Excel (удобно для одного правила на весь файл)
        code_from = def_from or (
            normalize_service_code(raw[i_from]) if i_from is not None else ""
        )
        code_to = def_to or (
            normalize_service_code(raw[i_to]) if i_to is not None else ""
        )
        if not code_from or not code_to:
            continue
        if skip_done and i_status is not None:
            st = _cell(raw[i_status]).lower()
            if st in ("ok", "done", "готово", "обработан", "обработано"):
                continue
        out.append(FixRow(talon=talon, code_from=code_from, code_to=code_to))
    return out


def make_template(path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "замена"
    ws.append(["Талон", "было", "стало"])
    ws.append(["39927692", "B04.015.003", "B04.047.003"])
    wb.save(path)
    return path


def open_claim(driver: WebDriver, talon: str) -> str:
    url = f"{OMS_BASE_URL.rstrip('/')}/claim/ambulatory/{talon}"
    logger.info("Открываем {}", url)
    driver.get(url)
    dv.wait_busy_gone(driver, timeout=45)
    sleep(0.5)
    WebDriverWait(driver, 25).until(
        EC.any_of(
            EC.presence_of_element_located((By.ID, "add-service")),
            EC.presence_of_element_located((By.ID, "begin-date")),
            EC.presence_of_element_located((By.ID, "mainDiagnosis")),
        )
    )
    return url


def open_services_section(driver: WebDriver) -> None:
    """На ambulatory услуги на основной форме; на DV/ОПВ — вкладка services-tab."""
    if driver.find_elements(By.ID, "services-tab"):
        try:
            if dv._tab_is_enabled(driver, "services-tab"):
                dv.click_id(driver, "services-tab")
                dv.wait_busy_gone(driver, timeout=20)
                sleep(0.4)
                return
        except Exception as exc:
            logger.debug("services-tab: {}", exc)
    # дождаться появления строк услуг на ambulatory
    try:
        WebDriverWait(driver, 20).until(
            lambda d: bool(d.find_elements(By.CSS_SELECTOR, '[id^="combobox-service-"]'))
            or bool(d.find_elements(By.ID, "add-service"))
        )
    except Exception:
        pass
    sleep(0.35)


def list_service_indices(driver: WebDriver) -> list[int]:
    ids: set[int] = set()
    for el in driver.find_elements(By.CSS_SELECTOR, '[id^="combobox-service-"]'):
        m = re.search(r"combobox-service-(\d+)$", el.get_attribute("id") or "")
        if m:
            ids.add(int(m.group(1)))
    for el in driver.find_elements(By.CSS_SELECTOR, '[id^="amount-service-"]'):
        m = re.search(r"amount-service-(\d+)$", el.get_attribute("id") or "")
        if m:
            ids.add(int(m.group(1)))
    return sorted(ids)


def wait_service_rows(driver: WebDriver, timeout: float = 25.0) -> list[int]:
    deadline = time() + timeout
    last: list[int] = []
    while time() < deadline:
        last = list_service_indices(driver)
        if last:
            # ещё чуть подождать догрузку остальных строк
            sleep(0.6)
            last2 = list_service_indices(driver)
            if last2 == last:
                return last2
            last = last2
        sleep(0.35)
    return last


def service_code_at(driver: WebDriver, index: int) -> str:
    els = driver.find_elements(By.ID, f"combobox-service-{index}")
    if not els:
        return ""
    el = els[0]
    wrap = dv._closest_react_select(el)
    shown = ""
    if wrap is not None:
        chunks: list[str] = []
        for sel in wrap.find_elements(
            By.CSS_SELECTOR, ".Selected-item, .Select-value, .Select-value-label"
        ):
            text = (sel.text or sel.get_attribute("title") or "").strip()
            if text:
                chunks.append(text)
        shown = " ".join(chunks) or (wrap.text or "")
    if not shown:
        shown = (el.get_attribute("value") or "").strip()
    return normalize_service_code(shown)


def set_service_code(driver: WebDriver, index: int, code: str) -> str:
    code = normalize_service_code(code)
    els = driver.find_elements(By.ID, f"combobox-service-{index}")
    if not els:
        raise RuntimeError(f"Нет combobox-service-{index}")
    el = els[0]
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    sleep(0.1)
    base._fill_select_fast(driver, el, code)
    sleep(0.25)
    dv.wait_busy_gone(driver, timeout=12)
    shown = service_code_at(driver, index)
    if shown != code:
        # повтор через прямой ввод
        try:
            el.click()
        except Exception:
            pass
        el.send_keys(Keys.CONTROL, "a")
        el.send_keys(Keys.BACKSPACE)
        el.send_keys(code, Keys.ENTER)
        sleep(0.45)
        dv.wait_busy_gone(driver, timeout=12)
        shown = service_code_at(driver, index)
    if shown != code:
        raise RuntimeError(
            f"combobox-service-{index} не принял «{code}», сейчас «{shown or '?'}»"
        )
    return shown


def _service_row_delete_button(driver: WebDriver, index: int):
    """Кнопка «Удалить сведение» (X) в строке услуги."""
    els = driver.find_elements(By.ID, f"combobox-service-{index}")
    if not els:
        return None
    inp = els[0]

    xpaths = [
        # типичный контейнер строки (jss424 в claim_bot; класс может меняться)
        "./ancestor::div[contains(@class,'jss424')][1]"
        "//button[@title='Удалить сведение']",
        "./ancestor::div[contains(@class,'jss424')][1]"
        "//button[contains(@title,'Удал')]",
        "./ancestor::div[contains(@class,'jss')][1]"
        "//button[@title='Удалить сведение' or contains(@title,'Удал')]",
        "./ancestor::div[contains(@class,'jss')][2]"
        "//button[@title='Удалить сведение' or contains(@title,'Удал')]",
        "./ancestor::tr[1]//button[@title='Удалить сведение' or contains(@title,'Удал')]",
        "./ancestor::li[1]//button[@title='Удалить сведение' or contains(@title,'Удал')]",
    ]
    for xp in xpaths:
        for b in inp.find_elements(By.XPATH, xp):
            try:
                if b.is_displayed() and b.is_enabled():
                    return b
            except Exception:
                continue

    # JS: ближайшая кнопка удаления в общем предке с amount/doctor той же строки
    try:
        btn = driver.execute_script(
            """
            const inp = arguments[0];
            const idx = arguments[1];
            const amount = document.getElementById('amount-service-' + idx);
            let node = inp;
            for (let depth = 0; depth < 8 && node; depth++) {
              node = node.parentElement;
              if (!node) break;
              if (amount && !node.contains(amount)) continue;
              const buttons = node.querySelectorAll('button');
              for (const b of buttons) {
                const t = (b.getAttribute('title') || b.getAttribute('aria-label') || '').toLowerCase();
                if (t.includes('удал') || t.includes('delete') || t.includes('remove')) {
                  return b;
                }
              }
            }
            return null;
            """,
            inp,
            index,
        )
        if btn is not None:
            return btn
    except Exception:
        pass
    return None


def delete_service_row(driver: WebDriver, index: int) -> bool:
    """Удалить строку услуги по индексу. True если число строк уменьшилось."""
    before_n = len(list_service_indices(driver))
    before_code = service_code_at(driver, index)
    btn = _service_row_delete_button(driver, index)
    if btn is None:
        logger.warning("нет кнопки удаления у service-{}", index)
        return False
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
    except Exception:
        pass
    try:
        btn.click()
    except Exception:
        try:
            driver.execute_script("arguments[0].click();", btn)
        except Exception:
            return False
    sleep(0.35)
    dv.wait_busy_gone(driver, timeout=15)
    deadline = time() + 6.0
    while time() < deadline:
        now_n = len(list_service_indices(driver))
        if now_n < before_n:
            return True
        # иногда индекс остаётся, но код ушёл / строка опустела
        if before_code and service_code_at(driver, index) != before_code:
            return True
        if not driver.find_elements(By.ID, f"combobox-service-{index}"):
            return True
        sleep(0.15)
    return len(list_service_indices(driver)) < before_n


def click_add_service(driver: WebDriver) -> None:
    """#add-service; если disabled — '+' в последней заполненной строке (как claim_bot)."""
    btns = driver.find_elements(By.ID, "add-service")
    if btns:
        btn = btns[0]
        try:
            if btn.is_enabled():
                try:
                    btn.click()
                except Exception:
                    driver.execute_script("arguments[0].click();", btn)
                sleep(0.45)
                dv.wait_busy_gone(driver, timeout=15)
                return
        except Exception:
            pass

    # fallback: строчный '+'
    indices = list_service_indices(driver)
    for i in reversed(indices):
        if not service_code_at(driver, i):
            continue
        els = driver.find_elements(By.ID, f"combobox-service-{i}")
        if not els:
            continue
        inp = els[0]
        plus = inp.find_elements(
            By.XPATH,
            "./ancestor::div[contains(@class,'jss')][1]"
            "//button[@aria-label='add' and not(@id='add-service')]",
        )
        if not plus:
            continue
        try:
            plus[0].click()
        except Exception:
            driver.execute_script("arguments[0].click();", plus[0])
        sleep(0.45)
        dv.wait_busy_gone(driver, timeout=15)
        return
    raise RuntimeError("Не удалось добавить строку услуги (#add-service / +)")


def replace_service_code(
    driver: WebDriver,
    code_from: str,
    code_to: str,
    *,
    replace_all: bool = False,
) -> tuple[int, str]:
    code_from = normalize_service_code(code_from)
    code_to = normalize_service_code(code_to)
    if not code_from or not code_to:
        raise RuntimeError("Пустой код услуги (было/стало)")
    if code_from == code_to:
        raise RuntimeError(f"было и стало одинаковы: {code_from}")

    open_services_section(driver)
    indices = wait_service_rows(driver, timeout=25)
    if not indices:
        raise RuntimeError("На талоне не найдено строк услуг (combobox-service-*)")

    catalog: list[tuple[int, str]] = []
    for i in indices:
        code = service_code_at(driver, i)
        catalog.append((i, code))
        logger.info("услуга[{}] = {}", i, code or "(пусто)")

    matches = [i for i, c in catalog if c == code_from]
    if not matches:
        present = ", ".join(c for _, c in catalog if c) or "—"
        raise RuntimeError(f"Код {code_from} не найден среди услуг: {present}")

    targets = matches if replace_all else matches[:1]
    replaced = 0
    for i in targets:
        shown = set_service_code(driver, i, code_to)
        logger.info("замена service-{}: {} → {}", i, code_from, shown)
        replaced += 1

    msg = f"{code_from} → {code_to} (строк: {replaced}/{len(matches)})"
    return replaced, msg


def _left_claim_form(driver: WebDriver, talon: str) -> bool:
    url = (driver.current_url or "").lower()
    if f"/ambulatory/{talon}" not in url:
        return True
    if not driver.find_elements(By.ID, "add-service") and not driver.find_elements(
        By.ID, "mainDiagnosis"
    ):
        return True
    return False


def save_existing_claim(driver: WebDriver, talon: str) -> tuple[bool, str]:
    """Save уже открытого талона (как в fix_diagnosis)."""
    url_before = driver.current_url or ""
    snack_before = base.read_snackbar(driver)

    try:
        base._click_save_button(driver)
    except RuntimeError as exc:
        return False, str(exc)

    dv.wait_busy_gone(driver, timeout=40)
    sleep(0.35)

    snack = base.wait_snackbar(driver, timeout=4, ignore=snack_before)
    verdict, detail = base.classify_snackbar(snack)

    dlg = dv.warning_dialog_text(driver)
    if (not dlg) and detail.startswith("[dialog]"):
        dlg = detail[len("[dialog]") :].strip()

    if dlg or dv.find_visible_dialogs(driver):
        low = (dlg or "").lower()
        # кнопки диалога: если OMS дал «Сохранить с предупреждениями» — подтверждаем
        # (в т.ч. пересечение цели 3 на уже существующем талоне)
        try:
            btn_blob = " | ".join(
                (dv._btn_label(driver, b) or "").lower().replace("\n", " ")
                for b in driver.find_elements(By.XPATH, "//button")
                if b.is_displayed()
            )
        except Exception:
            btn_blob = ""
        has_save_warn = dv._has_save_with_warning_label(btn_blob)
        if "дубликат" in low and not has_save_warn:
            dv.dismiss_warning_dialog(driver)
            return False, dlg or "дубликат"
        if dv._is_intersection_warning(dlg or "") and not has_save_warn:
            dv.dismiss_warning_dialog(driver)
            return False, dlg or "блокирующее предупреждение"
        if "результат поиска в црп" in low or "отсутствует пациент" in low:
            dv.dismiss_warning_dialog(driver)
            return False, dlg
        ok, wmsg = dv.handle_post_save_prompts(driver, timeout=20)
        if not ok:
            return False, wmsg or dlg or "диалог не подтверждён"
        dv.wait_busy_gone(driver, timeout=20)
        sleep(0.3)
        if _left_claim_form(driver, talon):
            return True, wmsg or "сохранено (уход в журнал)"
        snack2 = base.wait_snackbar(driver, timeout=3, ignore=snack)
        v2, d2 = base.classify_snackbar(snack2)
        if v2 is False:
            return False, d2 or wmsg
        return True, d2 or wmsg or "сохранено"

    if verdict is True:
        return True, detail

    if verdict is False:
        if _left_claim_form(driver, talon) and (detail or "").strip().lower() in (
            "error",
            "",
        ):
            logger.info("Snackbar Error, но ушли с формы талона {} — считаем OK", talon)
            return True, "сохранено (журнал после Error)"
        return False, detail or snack or "ошибка по уведомлению"

    if _left_claim_form(driver, talon) or (driver.current_url or "") != url_before:
        logger.info("Нет snackbar, уход с формы {} → OK (журнал)", talon)
        return True, "сохранено (журнал, без snackbar)"

    for _ in range(8):
        sleep(0.35)
        if _left_claim_form(driver, talon):
            return True, "сохранено (журнал)"
        sn = base.read_snackbar(driver)
        v, d = base.classify_snackbar(sn)
        if v is True:
            return True, d
        if v is False and (d or "").strip().lower() != "error":
            return False, d

    return False, "нет подтверждения Save — талон остался на форме"


def process_row(
    driver: WebDriver,
    row: FixRow,
    *,
    index: int,
    dry_run: bool,
    replace_all: bool,
) -> RowResult:
    t0 = time()
    logger.info(
        "=== [{}] talon={} {} → {} ===",
        index,
        row.talon,
        row.code_from,
        row.code_to,
    )
    try:
        open_claim(driver, row.talon)
        n, detail = replace_service_code(
            driver,
            row.code_from,
            row.code_to,
            replace_all=replace_all,
        )
        if dry_run:
            msg = f"dry-run: {detail}, без Save"
            logger.success("[{}] {}", index, msg)
            return RowResult(
                index,
                row.talon,
                row.code_from,
                row.code_to,
                "dry_run",
                msg,
                n,
                round(time() - t0, 2),
            )

        ok, msg = save_existing_claim(driver, row.talon)
        status = "ok" if ok else "fail"
        full = f"{detail} | {msg}"
        log = logger.success if ok else logger.error
        log("[{}] {} → {}", index, status.upper(), full[:300])
        return RowResult(
            index,
            row.talon,
            row.code_from,
            row.code_to,
            status,
            full,
            n,
            round(time() - t0, 2),
        )
    except Exception as exc:
        msg = str(exc).split("Stacktrace:")[0].strip() or str(exc)
        logger.error("[{}] error: {}", index, msg[:400])
        return RowResult(
            index,
            row.talon,
            row.code_from,
            row.code_to,
            "error",
            msg,
            0,
            round(time() - t0, 2),
        )


def write_results(path: Path, results: list[RowResult]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fields = list(asdict(results[0]).keys()) if results else [
        "index",
        "talon",
        "code_from",
        "code_to",
        "status",
        "message",
        "replaced",
        "elapsed_sec",
    ]
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for r in results:
            w.writerow(asdict(r))


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Замена кода услуги в амбулаторных талонах (find & replace)"
    )
    p.add_argument("--file", type=Path, default=DEFAULT_FILE, help="Excel: Талон + было + стало")
    p.add_argument("--from", dest="code_from", default="", help="Код услуги «было» (если нет в Excel)")
    p.add_argument("--to", dest="code_to", default="", help="Код услуги «стало» (если нет в Excel)")
    p.add_argument(
        "--replace-all",
        action="store_true",
        help="Менять все вхождения «было», не только первое",
    )
    p.add_argument("--make-template", action="store_true", help="Создать шаблон Excel и выйти")
    p.add_argument("--limit", type=int, default=0)
    p.add_argument("--offset", type=int, default=0)
    p.add_argument("--dry-run", action="store_true")
    p.add_argument("--keep-open", type=int, default=0)
    p.add_argument("--worker", default="")
    p.add_argument("--window-x", type=int, default=None)
    p.add_argument("--window-y", type=int, default=None)
    return p.parse_args()


def main() -> int:
    args = parse_args()
    RESULTS_DIR.mkdir(parents=True, exist_ok=True)

    if args.make_template:
        path = make_template(args.file if args.file != DEFAULT_FILE else DEFAULT_FILE)
        logger.info("Шаблон: {}", path)
        return 0

    path = args.file
    if not path.is_file():
        alt = PROJECT_ROOT / path.name
        if alt.is_file():
            path = alt
        else:
            logger.error("Нет файла {}", args.file)
            return 1

    rows = load_rows(path, default_from=args.code_from, default_to=args.code_to)
    if args.offset:
        rows = rows[args.offset :]
    if args.limit and args.limit > 0:
        rows = rows[: args.limit]

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    worker_tag = (args.worker or "").strip() or f"o{args.offset}"
    result_path = RESULTS_DIR / f"fix_service_result_{stamp}_{worker_tag}.csv"
    logger.info(
        "[{}] строк={} | dry_run={} | replace_all={} | offset={} | file={}",
        worker_tag,
        len(rows),
        args.dry_run,
        args.replace_all,
        args.offset,
        path,
    )
    if not rows:
        logger.warning("Нечего обрабатывать")
        return 0

    win_kwargs: dict = {"implicit_wait": 0, "maximize": False}
    if args.window_x is not None and args.window_y is not None:
        win_kwargs["window_position"] = (args.window_x, args.window_y)
        win_kwargs["window_size"] = (1100, 850)
    else:
        win_kwargs["maximize"] = True

    results: list[RowResult] = []
    t_all = time()
    with GeckoBrowser(**win_kwargs) as session:
        driver = session.driver
        driver.implicitly_wait(0)
        login_oms(driver)
        sleep(0.8)
        dv.wait_busy_gone(driver, timeout=40)

        for i, row in enumerate(rows, start=1 + args.offset):
            results.append(
                process_row(
                    driver,
                    row,
                    index=i,
                    dry_run=args.dry_run,
                    replace_all=args.replace_all,
                )
            )
            write_results(result_path, results)

        if args.keep_open > 0:
            logger.info("keep-open {}s", args.keep_open)
            sleep(args.keep_open)

    ok_n = sum(1 for r in results if r.status in ("ok", "dry_run"))
    fail_n = sum(1 for r in results if r.status not in ("ok", "dry_run"))
    logger.info(
        "——— [{}] ИТОГ ({:.0f}s) OK/DRY={}, FAIL={} → {} ———",
        worker_tag,
        time() - t_all,
        ok_n,
        fail_n,
        result_path,
    )
    return 0 if fail_n == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
