# Цель 541 (КТ) — Web.ОМС ambulatory, несколько услуг на талон

Путь: `robots/oms/cel_541/`

| | |
|--|--|
| Цель | `541` |
| Явка | `1` |
| Даты | `дд-мм-гг` |
| Excel | `data/talons_541.xlsx` + `data/services_541.xlsx` |
| Бандл | `data/talons_541_bundle.xlsx` (оба листа) |

## Сборка из КТ-файла

```powershell
cd D:\Projects\BrowserAuto
.\.venv\Scripts\python robots\oms\cel_541\template_541.py
# опционально корпус по умолчанию:
.\.venv\Scripts\python robots\oms\cel_541\template_541.py --building "ГП №3"
```

Правила объединения (ЕНП + дата):
- несколько услуг → несколько строк в `services_541.xlsx`
- несколько МКБ → `Диагноз` + `Диагноз 2` (+ `Диагнозы_все`)
- одинаковый код услуги несколько раз → `Кол-во` = число повторов

## Запуск

```powershell
copy robots\oms\cel_541\credentials.env.example robots\oms\cel_541\credentials.env

.\.venv\Scripts\python run_541.py --limit 3 --building "ГП №3"
.\.venv\Scripts\python run_541.py --limit 1 --dry-run --keep-open 30 --building "ГП №3"

# параллельно:
.\robots\oms\cel_541\run_parallel.ps1 -Workers 5 -Building "ГП №3"
```

Результат: `robots/oms/cel_541/data/oms_541_result_*.csv`.
