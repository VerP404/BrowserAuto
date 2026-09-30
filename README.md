# BrowserAuto

Автоматизация **Web.ОМС** и **ИСЗЛ** через Firefox + geckodriver.

## Структура

```text
BrowserAuto/
├── browser_auto/          # общее ядро (драйвер, логин, OCR)
├── robots/
│   ├── oms/               # роботы Web.ОМС
│   │   ├── cel_3/         # цель 3 (БСК) — талоны + услуги
│   │   ├── cel_307/       # цель 307
│   │   ├── dv_opv/        # ДВ4 / ОПВ
│   │   └── socstatus/     # соцстатус (можно копировать отдельно)
│   └── iszl/              # роботы ИСЗЛ
│       ├── carryover/     # перенос плана ДН
│       ├── download_dn/   # скачать план ДН
│       ├── workplace/     # отчёты «по месту работы»
│       └── oms_to_iszl/   # связка ОМС→ИСЗЛ
├── scripts/               # login-проверки, утилиты
├── archive/               # старые ноутбуки и черновики (не для прода)
├── _local/                # локальные данные/логи (не в git)
├── run_*.py               # точки входа из корня
└── requirements.txt
```

## Установка

```powershell
cd D:\Projects\BrowserAuto
python -m venv .venv
.\.venv\Scripts\pip install -r requirements.txt
copy env_example.txt .env   # логины/пароли
```

Нужны: **Python 3.10+**, **Firefox**, `geckodriver.exe` в корне, сеть до ОМС/ИСЗЛ.
Для капчи ИСЗЛ — Tesseract OCR.

Креды роботов можно класть в `robots/.../credentials.env` (см. `credentials.env.example` рядом).

## Запуск роботов

| Робот | Команда |
|-------|---------|
| Цель 3 (БСК) | `.\.venv\Scripts\python run_3.py --limit 3 --building "ГП №3" --doctor 30097144` |
| Цель 307 | `.\.venv\Scripts\python run_307.py --limit 3 --building "ГП №11"` |
| ДВ4 / ОПВ | `.\.venv\Scripts\python run_medical_exam.py --limit 1 --dry-run` |
| Соцстатус | `.\.venv\Scripts\python run_socstatus.py --limit 3` |
| Перенос ДН | `.\.venv\Scripts\python run_carryover_plan.py --offset 0 --limit 10` |
| Вход ИСЗЛ | `.\.venv\Scripts\python scripts\iszl_login.py` |
| Вход ОМС | `.\.venv\Scripts\python scripts\oms_login.py` |

Подробности — в `README.md` каждой папки робота.

## Git

В репозиторий **не** попадают: `.env`, `credentials.env`, `.venv`, Excel/CSV с данными,
логи, zip-бандлы, `_local/`. Шаблоны и код — да. См. `.gitignore`.

Старые ноутбуки и прототипы лежат в `archive/` — можно не коммитить или держать как историю.
