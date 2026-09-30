# Перенос плана ДН в ИСЗЛ («нет в текущем — есть в прошлых»)

Робот открывает Firefox, входит в ИСЗЛ, ищет пациента по ЕНП, выбирает строку по `ldwID`/`pID`, подставляет врача (`SS_D`, иначе `SS_DOCTOR`) и добавляет месяц плана.

Папка: `robots/iszl/carryover/`

---

## Что нужно на компьютере

| Компонент | Зачем |
|-----------|--------|
| **Windows 10/11** | Проверено на Win |
| **Python 3.10+** | Скрипт |
| **Mozilla Firefox** | Браузер для Selenium |
| **geckodriver.exe** | В корне репозитория `BrowserAuto/` (рядом с `run_carryover_plan.py`) |
| **Tesseract OCR** | Распознавание капчи ИСЗЛ |
| **Сеть до ИСЗЛ** | Обычно `http://10.36.29.2:8088` (VPN/локальная сеть МО) |
| **Файл Excel** | Выгрузка робота из портала → `data/iszl_carryover_robot_*.xlsx` |

Скачать geckodriver: https://github.com/mozilla/geckodriver/releases (версия под вашу Firefox).  
Tesseract: `winget install UB-Mannheim.TesseractOCR` (или установщик с github.com/UB-Mannheim/tesseract).

---

## Быстрая установка на новый ПК

1. Скопируйте папку **BrowserAuto** целиком (или клонируйте репозиторий).
2. Убедитесь, что в корне есть `geckodriver.exe`.
3. Откройте PowerShell в корне `BrowserAuto`:

```powershell
cd D:\Projects\BrowserAuto
python -m venv .venv
.\.venv\Scripts\pip install -r "ИСЗЛ\Нет в текущем есть в прошлых\requirements.txt"
```

4. Учётные данные — скопируйте пример и заполните:

```powershell
copy "ИСЗЛ\Нет в текущем есть в прошлых\credentials.env.example" "ИСЗЛ\Нет в текущем есть в прошлых\credentials.env"
notepad "ИСЗЛ\Нет в текущем есть в прошлых\credentials.env"
```

Файл `credentials.env` **не коммитьте** (в `.gitignore`). На другом ПК просто положите свой логин/пароль туда.

Можно также использовать корневой `.env` (как `env_example.txt`) — локальный `credentials.env` **перекрывает** его.

5. Положите Excel в:

`ИСЗЛ\Нет в текущем есть в прошлых\data\iszl_carryover_robot_YYYY-MM-DD.xlsx`

Имя по умолчанию в скрипте: `data/iszl_carryover_robot_2026-09-23.xlsx`  
(при другом имени укажите `--file "полный\путь\к\файлу.xlsx"`).

6. Проверка входа:

```powershell
.\.venv\Scripts\python scripts\iszl_login.py
```

---

## Запуск (один робот)

Из корня `BrowserAuto` (рекомендуется — путь без проблем с кириллицей):

```powershell
cd D:\Projects\BrowserAuto
.\.venv\Scripts\python run_carryover_plan.py --offset 0 --limit 10
```

Или из папки скрипта:

```powershell
cd "D:\Projects\BrowserAuto\ИСЗЛ\Нет в текущем есть в прошлых"
..\..\..\.venv\Scripts\python add_carryover_plan.py --offset 0 --limit 10
```

### Полезные ключи

| Ключ | Описание |
|------|----------|
| `--file PATH` | Excel со списком |
| `--offset N` | Пропустить первые N строк |
| `--limit N` | Обработать только N строк (`0` = все) |
| `--month M` | Месяц плана (1–12). По умолчанию — следующий календарный |
| `--worker NAME` | Метка воркера (уникальные капча/CSV при параллели) |
| `--window-x` / `--window-y` | Позиция окна Firefox |
| `--captcha-attempts N` | Попыток капчи на строку (по умолчанию 5) |
| `--keep-open N` | Секунд держать браузер после окончания |

Результаты: `data/carryover_result_<worker>_o<offset>_<дата>.csv`  
Логи параллели (если запускали через скрипт): `data/logs/`

---

## Параллельный запуск (5–6 роботов)

Каждый процесс — свой Firefox, свой offset, своя метка `--worker`.  
Капчи и CSV не пересекаются.

Пример 6 × 10 строк (offsets 0, 10, 20, 30, 40, 50):

```powershell
cd D:\Projects\BrowserAuto
.\ИСЗЛ\Нет в текущем есть в прошлых\run_parallel.ps1 -Workers 6 -Limit 10
```

Или вручную (6 окон):

```powershell
$py = ".\.venv\Scripts\python.exe"
$launch = ".\run_carryover_plan.py"
for ($w = 0; $w -lt 6; $w++) {
  $offset = $w * 10
  Start-Process $py -ArgumentList @(
    $launch, "--worker", "w$w", "--offset", "$offset", "--limit", "10",
    "--window-x", "$(20+$w*160)", "--window-y", "40"
  ) -WorkingDirectory (Get-Location)
}
```

Для полного файла (~40k строк) разбейте диапазоны, например по 7000:

```text
w0: --offset 0     --limit 7000
w1: --offset 7000  --limit 7000
...
```

Ориентир по времени (замер на 60 строках, 6 воркеров): **~13 с/строка**, wall ~3 мин на 60 строк → на ~40k при 6 роботах порядка **сутки–полтора** (если ИСЗЛ не тормозит сильнее).

---

## Ожидания и подвисания

Робот **не ждёт бесконечно**:

| Шаг | Потолок | Поведение |
|-----|---------|-----------|
| blockUI (поиск/сохранение) | ~3–12 с | После timeout идёт дальше |
| Поиск врача | ~2.5 с на вариант СНИЛС | Если нет — error, следующая строка |
| Капча | до `--captcha-attempts` | Потом error по строке |
| Элементы формы | WebDriverWait 2.5–15 с | Exception → error |

При параллели ИСЗЛ может чаще «сыпать» капчу — это нормально; строки с error можно перегнать отдельно по CSV.

---

## Excel: какие колонки нужны

| Колонка | Назначение |
|---------|------------|
| ЕНП | Поиск пациента |
| МКБ | Диагноз (в лог) |
| ldwID | Выбор строки в `gvLstDW` |
| pID | Уточнение пациента |
| pdwID | Справочно |
| СНИЛС врача / СНИЛС врача (цифры) | Из **SS_D** (врач на ДН); если пусто — SS_DOCTOR |
| Год плана | Обычно текущий год плана |

Выгрузка из MosaicPortal (вкладка чистки ИСЗЛ → файл для робота) уже после правки бэка отдаёт `SS_D`.

---

## Типичные статусы в CSV

| status | Значение |
|--------|----------|
| `ok` | План добавлен |
| `skip` | Уже был в плане |
| `fail` | Ответ ИСЗЛ без «успешно» |
| `error` | Исключение (нет врача, капча, нет строки и т.п.) |

---

## Структура папки

```text
BrowserAuto/
  geckodriver.exe          ← обязательно
  .venv/                   ← venv
  run_carryover_plan.py    ← лаунчер без кириллицы в argv
  browser_auto/            ← общий код (логин, Firefox, OCR)
  Папка: `robots/iszl/carryover/`
    README.md              ← эта инструкция
    requirements.txt
    credentials.env.example
    credentials.env        ← ваш логин/пароль (создать сами)
    add_carryover_plan.py
    run_parallel.ps1
    data/
      iszl_carryover_robot_*.xlsx
      carryover_result_*.csv
    captcha/               ← скрины капчи по воркерам
```

---

## Частые проблемы

1. **`Задайте ISZL_LOGIN и ISZL_PASSWORD`** — нет `credentials.env` или пустые поля.  
2. **`geckodriver не найден`** — положите `geckodriver.exe` в корень `BrowserAuto`.  
3. **Tesseract / OCR** — установите Tesseract; при другом пути укажите `TESSERACT_CMD` в `credentials.env`.  
4. **Кириллица в пути при Start-Process** — всегда запускайте через `run_carryover_plan.py` из корня.  
5. **Врач не найден** — СНИЛС из `SS_D` отсутствует в справочнике FindDoct МО (уволен и т.п.).  
6. **Нет доступа к 10.36.29.2** — нужен VPN/сеть поликлиники.
