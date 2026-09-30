"""
Справочник врачей для услуг ДВ4/ОПВ.

Источник (по приоритету):
  1) лист «справочник_врачей» в файле талонов
  2) отдельный Excel (--catalog / doctors_by_building.xlsx)

Колонки: корпус | код_услуги | роль | врач

Логика для каждой строки услуги в ОМС:
  - если код услуги есть в справочнике (для корпуса талона) → врач из справочника;
  - если нет (или врач в правиле пустой) → основной врач из карты талона.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import Workbook, load_workbook


@dataclass
class DoctorRule:
    building: str
    service_pattern: str  # точный код, префикс%, или *
    role: str  # therapist | specialist | lab | …
    doctor: str  # код врача; пусто → main_doctor из карты


@dataclass
class DoctorCatalog:
    rules: list[DoctorRule] = field(default_factory=list)

    @classmethod
    def load(cls, path: Path, sheet: str | None = None) -> DoctorCatalog:
        if not path.is_file():
            return cls()
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            if sheet and sheet in wb.sheetnames:
                ws = wb[sheet]
            elif "справочник_врачей" in wb.sheetnames:
                ws = wb["справочник_врачей"]
            else:
                ws = wb.active
            rows = ws.iter_rows(values_only=True)
            header = [str(h or "").strip().lower() for h in next(rows)]
            idx = {h: i for i, h in enumerate(header) if h}

            def col(*names: str) -> int | None:
                for n in names:
                    if n in idx:
                        return idx[n]
                return None

            i_b = col("корпус", "building")
            i_s = col("код_услуги", "услуга", "service", "service_code")
            i_r = col("роль", "role")
            i_d = col("врач", "doctor", "код врача")
            if i_s is None:
                return cls()

            rules: list[DoctorRule] = []
            for raw in rows:
                if not raw:
                    continue
                svc = str(raw[i_s] or "").strip()
                if not svc:
                    continue
                rules.append(
                    DoctorRule(
                        building=str(raw[i_b] or "").strip() if i_b is not None else "",
                        service_pattern=svc,
                        role=str(raw[i_r] or "").strip().lower() if i_r is not None else "",
                        doctor=str(raw[i_d] or "").strip() if i_d is not None else "",
                    )
                )
            return cls(rules=rules)
        finally:
            wb.close()

    @classmethod
    def load_for_talons_file(cls, talons_path: Path, fallback: Path | None = None) -> DoctorCatalog:
        """Сначала лист «справочник_врачей» в файле талонов, иначе --catalog."""
        if talons_path.is_file():
            cat = cls.load(talons_path, sheet="справочник_врачей")
            if cat.rules:
                return cat
            # файл без нужного листа — попробовать active
            cat = cls.load(talons_path)
            if cat.rules:
                return cat
        if fallback:
            return cls.load(fallback)
        return cls()


    @staticmethod
    def ensure_example(path: Path) -> None:
        if path.is_file():
            return
        path.parent.mkdir(parents=True, exist_ok=True)
        wb = Workbook()
        ws = wb.active
        ws.title = "doctors"
        ws.append(["корпус", "код_услуги", "роль", "врач"])
        ws.append(["1", "*", "therapist", ""])
        ws.append(["1", "A09.05.%", "lab", "22222"])
        ws.append(["1", "B04.047.004", "therapist", ""])
        ws.append(["2", "*", "therapist", ""])
        wb.save(path)


def _match_pattern(pattern: str, service_code: str) -> bool:
    p = (pattern or "").strip()
    code = (service_code or "").strip()
    if not p or p == "*":
        return True
    if p.endswith("%"):
        return code.startswith(p[:-1])
    if "*" in p or "?" in p:
        rx = re.escape(p).replace(r"\*", ".*").replace(r"\?", ".")
        return re.fullmatch(rx, code, flags=re.IGNORECASE) is not None
    return p.replace(" ", "").lower() == code.replace(" ", "").lower()


def _norm_building(value: str) -> str:
    text = (value or "").strip().lower().replace("№", "no").replace("#", "no")
    return re.sub(r"\s+", "", text)


def resolve_service_doctor(
    catalog: DoctorCatalog,
    *,
    building: str,
    service_code: str,
    main_doctor: str,
) -> str:
    """Врач услуги: справочник (корпус+код) → иначе врач карты талона."""
    b = _norm_building(building)
    code = (service_code or "").strip()

    # 1) точное совпадение корпус + код_услуги
    # 2) код без корпуса / wildcard
    scored: list[tuple[int, DoctorRule]] = []
    for rule in catalog.rules:
        if not _match_pattern(rule.service_pattern, code):
            continue
        rb = _norm_building(rule.building)
        if rb and b and rb != b:
            continue
        exact = 0 if rule.service_pattern not in ("*", "") and "%" not in rule.service_pattern and "*" not in rule.service_pattern else 1
        building_rank = 0 if rb and b and rb == b else (1 if not rb else 2)
        scored.append((building_rank * 10 + exact, rule))

    scored.sort(key=lambda x: x[0])
    for _, rule in scored:
        if rule.doctor:
            return rule.doctor
        return main_doctor
    return main_doctor
