# Ввод талонов цели 307 (Web.ОМС ambulatory)

Путь: `robots/oms/cel_307/`

| | |
|--|--|
| URL | `http://10.36.0.142:9000/claim/ambulatory` |
| Excel | один файл `data/talons_307.xlsx` |

## Сборка из короткого файла портала

```powershell
cd D:\Projects\BrowserAuto
.\.venv\Scripts\python run_307_template.py --from-source "robots\oms\cel_307\data\талоны для 307.xlsx" --building "ГП №11"
```

## Запуск

```powershell
copy robots\oms\cel_307\credentials.env.example robots\oms\cel_307\credentials.env

.\.venv\Scripts\python run_307.py --limit 3 --building "ГП №11"
.\.venv\Scripts\python run_307.py --limit 1 --dry-run --keep-open 30

# параллельно:
.\robots\oms\cel_307\run_parallel.ps1 -Workers 5 -StartOffset 0
```

Результат: `robots/oms/cel_307/data/oms_307_result_*.csv`.
