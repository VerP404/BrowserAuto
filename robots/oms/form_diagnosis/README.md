# Исправление диагноза (ds1) в амбулаторных талонах

Папка: `robots/oms/form_diagnosis/`

## Что делает

1. Читает Excel: **Талон** + **ds1** (или `--diagnosis`, если колонки нет)
2. Открывает `{OMS}/claim/ambulatory/{Талон}`
3. Ставит `#mainDiagnosis` = ds1
4. Жмёт Save (как в cel_307)

> Ранее сюда временно положили очистку `medicalExaminationPlace` — она сохранена в `form_clear_place.py`.

## Запуск

```powershell
cd D:\Projects\BrowserAuto
copy robots\oms\form_diagnosis\credentials.env.example robots\oms\form_diagnosis\credentials.env

# тест без Save
.\.venv\Scripts\python run_form_diagnosis.py --file robots\oms\form_diagnosis\data\Книга15.xlsx --limit 1 --dry-run --keep-open 20

# боевой
.\.venv\Scripts\python run_form_diagnosis.py --file robots\oms\form_diagnosis\data\Книга15.xlsx

# если в Excel только колонка «Талон»:
.\.venv\Scripts\python run_form_diagnosis.py --file ...\Книга15.xlsx --diagnosis Z72.4
```

Параллельно:

```powershell
.\robots\oms\form_diagnosis\run_parallel.ps1 -Workers 5 -File "D:\Projects\BrowserAuto\robots\oms\form_diagnosis\data\Книга15.xlsx"
```

Результат: `robots/oms/form_diagnosis/data/fix_diagnosis_result_*.csv`.
