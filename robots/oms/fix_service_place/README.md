# Замена «Место обслуживания» (#service-place)

Папка: `robots/oms/fix_service_place/`

Открывает амбулаторный талон и ставит значение в `#service-place`
(поле «Место обслуживания» / «Место обращения»).

URL: `{OMS}/claim/ambulatory/{Талон}`

Не путать с `#medicalExaminationPlace` («Место прохождения диспансеризации») —
его правит `form_diagnosis/form_clear_place.py`.

---

## Excel

| Колонка | Обязательно | Пример |
|---------|-------------|--------|
| Талон | да | 39927692 |
| Место обслуживания | да* | 9 |
| было | нет | 1 |

\* или общий флаг `--place` для всего файла.

Исходный файл: `robots/oms/Место обсл.xlsx` (копия в `data/fix_service_place.xlsx`).

Шаблон:

```powershell
.\.venv\Scripts\python run_fix_service_place.py --make-template
```

---

## Запуск

```powershell
cd D:\Projects\BrowserAuto
copy robots\oms\fix_service_place\credentials.env.example robots\oms\fix_service_place\credentials.env

# тест без Save
.\.venv\Scripts\python run_fix_service_place.py --file "robots\oms\Место обсл.xlsx" --limit 1 --dry-run --keep-open 20

# боевой
.\.venv\Scripts\python run_fix_service_place.py --file "robots\oms\Место обсл.xlsx"

# если в Excel только колонка «Талон»:
.\.venv\Scripts\python run_fix_service_place.py --file talons.xlsx --place 9
```

Параллельно:

```powershell
.\robots\oms\fix_service_place\run_parallel.ps1 -Workers 4 -File "D:\Projects\BrowserAuto\robots\oms\Место обсл.xlsx"
```

Результат: `robots/oms/fix_service_place/data/fix_service_place_result_*.csv`.
