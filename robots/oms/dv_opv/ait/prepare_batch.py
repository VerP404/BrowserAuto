# -*- coding: utf-8 -*-
"""
Подготовка пачки ДВ4/ОПВ из папки ait/.

Inbox:     robots/oms/dv_opv/ait/*.xlsx
Done:      robots/oms/dv_opv/ait/done/   — исходники после обработки
Ready:     robots/oms/dv_opv/ait/ready/talons_ready.xlsx — общий файл для робота
State:     robots/oms/dv_opv/ait/state/done_enps.txt — уже введённые ЕНП

Что делает:
  • нормализует «начало»/«окончание» в текст DD-MM-YY
    (datetime, serial Excel 46258, «46258», dd.mm.yyyy → 15-10-26)
  • пишет даты как текст (@), чтобы Excel снова не превратил в дату
  • подставляет врач=30151001, корпус=ГП №3 (если пусто или --force)
  • склеивает все inbox-файлы в один ready
  • пропускает ЕНП, уже отмеченные как ok (state + result CSV)
  • переносит обработанные xlsx в done/

Запуск из BrowserAuto:
  .\\.venv\\Scripts\\python robots\\oms\\dv_opv\\ait\\prepare_batch.py
  .\\.venv\\Scripts\\python robots\\oms\\dv_opv\\ait\\prepare_batch.py --force-doctor
  .\\.venv\\Scripts\\python run_medical_exam.py --file robots\\oms\\dv_opv\\ait\\ready\\talons_ready.xlsx
"""
from __future__ import annotations

import argparse
import re
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, numbers

AIT_DIR = Path(__file__).resolve().parent
DONE_DIR = AIT_DIR / "done"
READY_DIR = AIT_DIR / "ready"
STATE_DIR = AIT_DIR / "state"
READY_FILE = READY_DIR / "talons_ready.xlsx"
DONE_ENPS_FILE = STATE_DIR / "done_enps.txt"
RESULTS_DIR = AIT_DIR.parent / "data"

DEFAULT_DOCTOR = "30151001"
DEFAULT_BUILDING = "ГП №3"

# Excel: день 0 ≈ 1899-12-30 (с учётом бага Lotus 1900)
_EXCEL_EPOCH = datetime(1899, 12, 30)

DATE_COLS = ("начало", "окончание")
ROBOT_HEADER = [
    "ФИО",
    "Участок",
    "Возраст",
    "Телефон",
    "Группа здоровья",
    "ИСЗЛ ДН",
    "Дата явки",
    "Сотрудник явки",
    "Подразделение явки",
    "ЕНП",
    "начало",
    "окончание",
    "тип",
    "врач",
    "результат",
    "период",
    "диагноз",
    "ДН",
    "диагноз2",
    "ДН2",
    "диагноз3",
    "ДН3",
    "диагноз4",
    "ДН4",
    "диагноз5",
    "ДН5",
    "диагноз6",
    "ДН6",
    "корпус",
    "место",
    "выполнено",
    "наз_вид",
    "наз_от_кого",
    "наз_к_кому",
    "_source_file",
]


def _log(msg: str) -> None:
    # Windows cp1251-консоль не любит стрелки/unicode
    safe = (
        str(msg)
        .replace("→", "->")
        .replace("←", "<-")
        .replace("—", "-")
        .replace("…", "...")
    )
    try:
        print(safe, flush=True)
    except UnicodeEncodeError:
        print(safe.encode("cp1251", errors="replace").decode("cp1251"), flush=True)


def excel_serial_to_date(n: float | int) -> datetime | None:
    try:
        val = float(n)
    except (TypeError, ValueError):
        return None
    if not (20000 <= val <= 80000):
        return None
    return _EXCEL_EPOCH + timedelta(days=int(val))


def to_oms_date(value: object) -> str:
    """Любой Excel-мусор → DD-MM-YY текстом."""
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d-%m-%y")
    if isinstance(value, date):
        return value.strftime("%d-%m-%y")
    if isinstance(value, bool):
        return ""
    if isinstance(value, (int, float)):
        dt = excel_serial_to_date(value)
        if dt:
            return dt.strftime("%d-%m-%y")
        if float(value).is_integer():
            # не дата — на всякий случай строка
            return str(int(value))
        return str(value).strip()

    text = str(value).strip()
    if not text or text.lower() in {"none", "nan", "nat"}:
        return ""

    # serial как текст: 46258 / 46258.0
    if re.fullmatch(r"\d{5}(?:\.0+)?", text):
        dt = excel_serial_to_date(float(text))
        if dt:
            return dt.strftime("%d-%m-%y")

    m = re.match(r"^(\d{1,2})[-./](\d{1,2})[-./](\d{2,4})(?:\s+.*)?$", text)
    if m:
        d, mo, y = m.groups()
        if len(y) == 4:
            y = y[-2:]
        return f"{int(d):02d}-{int(mo):02d}-{y}"

    digits = re.sub(r"\D", "", text)
    if len(digits) == 6:
        return f"{digits[0:2]}-{digits[2:4]}-{digits[4:6]}"
    if len(digits) == 8:
        return f"{digits[0:2]}-{digits[2:4]}-{digits[6:8]}"
    return text


def _cell_str(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d-%m-%y")
    if isinstance(value, date):
        return value.strftime("%d-%m-%y")
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, int) and not isinstance(value, bool):
        return str(value)
    return str(value).strip()


def load_done_enps() -> set[str]:
    import csv

    done: set[str] = set()
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    if DONE_ENPS_FILE.is_file():
        for line in DONE_ENPS_FILE.read_text(encoding="utf-8").splitlines():
            enp = re.sub(r"\D", "", line.strip())
            if len(enp) >= 11:
                done.add(enp)

    # подтянуть ok из result CSV робота
    if RESULTS_DIR.is_dir():
        for path in RESULTS_DIR.glob("dv_opv_result_*.csv"):
            try:
                text = path.read_text(encoding="utf-8-sig", errors="replace")
            except OSError:
                continue
            try:
                for row in csv.DictReader(text.splitlines()):
                    if (row.get("status") or "").strip().lower() != "ok":
                        continue
                    enp = re.sub(r"\D", "", str(row.get("enp") or ""))
                    if len(enp) >= 11:
                        done.add(enp)
            except csv.Error:
                # fallback на regex, если CSV битый
                for m in re.finditer(r",(\d{14,}),(?:dv4|opv|ud1),ok,", text, re.I):
                    done.add(m.group(1))
    return done


def save_done_enps(enps: set[str]) -> None:
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    lines = sorted(enps)
    DONE_ENPS_FILE.write_text("\n".join(lines) + ("\n" if lines else ""), encoding="utf-8")


def inbox_files() -> list[Path]:
    files = []
    for p in sorted(AIT_DIR.glob("*.xlsx")):
        if p.name.startswith("~$"):
            continue
        files.append(p)
    return files


def _is_blank_header(header: list[str]) -> bool:
    return not any(h.strip() for h in header)


def read_rows(path: Path) -> tuple[list[str], list[dict[str, object]]]:
    wb = load_workbook(path, data_only=False)
    ws = wb["талоны"] if "талоны" in wb.sheetnames else wb.active
    rows_iter = ws.iter_rows(values_only=False)
    # некоторые выгрузки начинаются с пустой строки — ищем строку с ЕНП/начало
    header: list[str] = []
    header_cells = None
    for _ in range(8):
        try:
            header_cells = next(rows_iter)
        except StopIteration:
            break
        header = [_cell_str(c.value).strip() for c in header_cells]
        if _is_blank_header(header):
            continue
        lower_probe = [h.lower() for h in header]
        if any(n in lower_probe for n in ("енп", "enp")) and any(
            n in lower_probe for n in ("начало", "begin", "дата начала")
        ):
            break
    else:
        wb.close()
        raise ValueError(f"{path.name}: не найдена строка заголовка с ЕНП/начало")

    # normalize header keys lower for lookup, keep original labels
    lower = [h.lower() for h in header]

    def col(*names: str) -> int | None:
        for n in names:
            if n.lower() in lower:
                return lower.index(n.lower())
        return None

    i_enp = col("енп", "enp")
    i_beg = col("начало", "begin", "дата начала")
    i_end = col("окончание", "end", "дата окончания")
    if i_enp is None or i_beg is None or i_end is None:
        wb.close()
        raise ValueError(f"{path.name}: нужны колонки ЕНП, начало, окончание. Сейчас: {header}")

    out: list[dict[str, object]] = []
    for row in rows_iter:
        values = [c.value for c in row]
        if not values or all(v is None or str(v).strip() == "" for v in values):
            continue
        enp_raw = _cell_str(values[i_enp] if i_enp < len(values) else None)
        enp = re.sub(r"\D", "", enp_raw.replace("`", ""))
        if len(enp) < 11:
            continue
        rec: dict[str, object] = {h: (values[i] if i < len(values) else None) for i, h in enumerate(header) if h}
        rec["ЕНП"] = enp
        rec["начало"] = to_oms_date(values[i_beg] if i_beg < len(values) else None)
        rec["окончание"] = to_oms_date(values[i_end] if i_end < len(values) else None)
        rec["_source_file"] = path.name
        out.append(rec)
    wb.close()
    return header, out


def apply_defaults(rec: dict[str, object], *, force: bool) -> None:
    doc = _cell_str(rec.get("врач"))
    bld = _cell_str(rec.get("корпус"))
    if force or not doc:
        rec["врач"] = DEFAULT_DOCTOR
    else:
        rec["врач"] = doc
    if force or not bld:
        rec["корпус"] = DEFAULT_BUILDING
    else:
        rec["корпус"] = bld
    # тип по умолчанию ДВ4
    if not _cell_str(rec.get("тип")):
        rec["тип"] = "ДВ4"


def write_ready(rows: list[dict[str, object]], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "талоны"
    # заголовок: стандартный набор + всё лишнее из строк
    extra_keys: list[str] = []
    seen = {h.lower() for h in ROBOT_HEADER}
    for rec in rows:
        for k in rec:
            if k.lower() not in seen and not k.startswith("_"):
                extra_keys.append(k)
                seen.add(k.lower())
    header = list(ROBOT_HEADER) + extra_keys
    ws.append(header)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    date_idx = {header.index(c) for c in DATE_COLS if c in header}
    for rec in rows:
        line = []
        for h in header:
            if h == "_source_file":
                line.append(_cell_str(rec.get("_source_file")))
            elif h in DATE_COLS:
                line.append(to_oms_date(rec.get(h)))
            elif h == "ЕНП":
                line.append(_cell_str(rec.get("ЕНП")))
            else:
                line.append(_cell_str(rec.get(h)))
        ws.append(line)
        r = ws.max_row
        for ci in date_idx:
            cell = ws.cell(r, ci + 1)
            cell.number_format = "@"
            cell.alignment = Alignment(horizontal="center")
        # ЕНП тоже текстом
        if "ЕНП" in header:
            c = ws.cell(r, header.index("ЕНП") + 1)
            c.number_format = "@"

    wb.save(path)


def move_to_done(path: Path) -> Path:
    DONE_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = DONE_DIR / f"{stamp}__{path.name}"
    if dest.exists():
        dest = DONE_DIR / f"{stamp}__{path.stem}_{path.suffix}"
    path.rename(dest)
    return dest


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Подготовка пачки ДВ4 из ait/")
    p.add_argument(
        "--force-doctor",
        action="store_true",
        help=f"Всегда ставить врача {DEFAULT_DOCTOR} и корпус «{DEFAULT_BUILDING}»",
    )
    p.add_argument(
        "--keep-inbox",
        action="store_true",
        help="Не переносить исходники в done/",
    )
    p.add_argument(
        "--include-done",
        action="store_true",
        help="Не пропускать уже отработанные ЕНП (для повторной выгрузки)",
    )
    p.add_argument(
        "--dry-run",
        action="store_true",
        help="Только показать, что будет сделано",
    )
    return p.parse_args()


def main() -> int:
    args = parse_args()
    files = inbox_files()
    if not files:
        _log(f"Inbox пуст: {AIT_DIR}")
        _log("Положите xlsx карт врача и запустите снова.")
        return 0

    done_enps = set() if args.include_done else load_done_enps()
    if not args.dry_run and not args.include_done:
        save_done_enps(done_enps)
    _log(f"Inbox: {len(files)} файл(ов) | уже отработано ЕНП (ok): {len(done_enps)}")

    merged: dict[str, dict[str, object]] = {}
    date_fixes = 0
    skipped_done = 0
    kept_prev = 0
    per_file: list[tuple[str, int, int]] = []

    # незавершённые из текущего ready (ещё не ok) — чтобы не потерять при новой пачке
    if READY_FILE.is_file() and not args.include_done:
        try:
            _h, prev_rows = read_rows(READY_FILE)
            for rec in prev_rows:
                enp = str(rec["ЕНП"])
                if enp in done_enps:
                    skipped_done += 1
                    continue
                apply_defaults(rec, force=args.force_doctor)
                merged[enp] = rec
                kept_prev += 1
            _log(f"  ready (остаток): оставлено {kept_prev}, уже ok пропуск {skipped_done}")
        except Exception as exc:  # noqa: BLE001
            _log(f"  ! {READY_FILE.name}: {exc}")

    for path in files:
        try:
            _header, rows = read_rows(path)
        except Exception as exc:  # noqa: BLE001
            _log(f"  ! {path.name}: {exc}")
            continue
        added = 0
        skipped = 0
        for rec in rows:
            enp = str(rec["ЕНП"])
            if enp in done_enps:
                skipped += 1
                skipped_done += 1
                continue
            apply_defaults(rec, force=args.force_doctor)
            if re.fullmatch(r"\d{2}-\d{2}-\d{2}", str(rec.get("начало") or "")) and re.fullmatch(
                r"\d{2}-\d{2}-\d{2}", str(rec.get("окончание") or "")
            ):
                date_fixes += 1
            # inbox перекрывает остаток ready / дубли inbox
            if enp in merged:
                skipped += 1
            else:
                added += 1
            merged[enp] = rec
        per_file.append((path.name, added, skipped))
        _log(f"  {path.name}: новых {added}, пропуск/дубли {skipped}")

    ready_rows = list(merged.values())
    # стабильный порядок: по дате начала, потом ЕНП
    ready_rows.sort(key=lambda r: (str(r.get("начало") or ""), str(r.get("ЕНП") or "")))

    _log(
        f"Итого в ready: {len(ready_rows)} | из прошлого ready: {kept_prev} | "
        f"пропуск уже ok: {skipped_done} | дат нормализовано (оценка): {date_fixes}"
    )

    if args.dry_run:
        _log("dry-run: файл не пишем, inbox не трогаем")
        for r in ready_rows[:5]:
            _log(
                f"  sample {r.get('ЕНП')} {r.get('начало')}...{r.get('окончание')} "
                f"врач={r.get('врач')} корпус={r.get('корпус')} <- {r.get('_source_file')}"
            )
        return 0

    if not ready_rows:
        _log("Нечего выгружать (все ЕНП уже отработаны или файлы пустые).")
        if not args.keep_inbox:
            for path in files:
                dest = move_to_done(path)
                _log(f"  -> done/{dest.name}")
        return 0

    write_ready(ready_rows, READY_FILE)
    _log(f"Ready: {READY_FILE}")

    if not args.keep_inbox:
        for path in files:
            if not path.exists():
                continue
            dest = move_to_done(path)
            _log(f"  inbox {path.name} -> done/{dest.name}")

    _log("")
    _log("Дальше:")
    _log(
        r'  .\.venv\Scripts\python run_medical_exam.py --file robots\oms\dv_opv\ait\ready\talons_ready.xlsx'
        r' --catalog robots\oms\dv_opv\catalog_dv_opv.xlsx'
    )
    _log("После успешного прогона ЕНП подтянутся из result CSV автоматически.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
