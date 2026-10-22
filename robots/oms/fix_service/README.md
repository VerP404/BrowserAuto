# Замена услуги в амбулаторных талонах

Папка: `robots/oms/fix_service/`

Ищет **одну** услугу по коду среди многих строк и меняет её на другую  
(цель 3 / ambulatory — поля `#combobox-service-{i}`).

URL: `{OMS}/claim/ambulatory/{Талон}`

Пример: на талоне `39927692` заменить `B04.015.003` → `B04.047.003`.

---

## Excel

| Колонка | Обязательно | Пример |
|---------|-------------|--------|
| Талон | да | 39927692 |
| было | да* | B04.015.003 |
| стало | да* | B04.047.003 |

\* или общие флаги `--from` / `--to` для всего файла.

### Из файла услуг (цель 3 / DN)

Архив или Excel услуг → уникальные талоны с нужным кодом:

```powershell
# посмотреть топ кодов
.\.venv\Scripts\python robots\oms\fix_service\prepare_from_services.py `
  --zip robots\oms\dn_fix_3_2026_2026-10-20_012704.zip --list-codes

# собрать Excel для робота (пример: терапевт → нужный B04)
.\.venv\Scripts\python robots\oms\fix_service\prepare_from_services.py `
  --zip robots\oms\dn_fix_3_2026_2026-10-20_012704.zip `
  --from B04.015.003 --to B04.047.003 `
  --out robots\oms\fix_service\data\fix_service_dn3.xlsx
```

Робот сам ищет **одну** строку «было» среди многих услуг на талоне и меняет на «стало».

Шаблон:

```powershell
.\.venv\Scripts\python run_fix_service.py --make-template
```

---

## Запуск

```powershell
cd D:\Projects\BrowserAuto
copy robots\oms\fix_service\credentials.env.example robots\oms\fix_service\credentials.env

# тест без Save — проверить, что код находится и меняется
.\.venv\Scripts\python run_fix_service.py --file robots\oms\fix_service\data\fix_service.xlsx --limit 1 --dry-run --keep-open 20

# боевой
.\.venv\Scripts\python run_fix_service.py --file robots\oms\fix_service\data\fix_service.xlsx

# если в Excel только колонка «Талон»:
.\.venv\Scripts\python run_fix_service.py --file talons.xlsx --from B04.015.003 --to B04.047.003
```

`--replace-all` — менять все вхождения «было», не только первое.

Параллельно:

```powershell
.\robots\oms\fix_service\run_parallel.ps1 -Workers 4 -File "D:\Projects\BrowserAuto\robots\oms\fix_service\data\fix_service.xlsx"
```

Результат: `robots/oms/fix_service/data/fix_service_result_*.csv`.
