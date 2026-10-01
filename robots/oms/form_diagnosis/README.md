# Убрать «Место прохождения диспансеризации» с готовых талонов.

Папка: `robots/oms/form_diagnosis/`

## Что делает

1. Читает Excel: **Талон** (id) + **Цель** (ДВ4/ОПВ)
2. Открывает `http://10.36.0.142:9000/claim/medicalExamination/{id}`
3. Очищает `#medicalExaminationPlace` (ничего туда не пишет)
4. Жмёт Save (`#id-save`) и подтверждает диалоги как в dv_opv

## Запуск

```powershell
cd D:\Projects\BrowserAuto
copy robots\oms\form_diagnosis\credentials.env.example robots\oms\form_diagnosis\credentials.env

# проверить 1 талон без сохранения
.\.venv\Scripts\python run_form_diagnosis.py --limit 1 --dry-run --keep-open 20

# боевой прогон
.\.venv\Scripts\python run_form_diagnosis.py --file robots\oms\form_diagnosis\data\ubrat_mesto.xlsx
```

Файл по умолчанию: `data/ubrat_mesto.xlsx` (копия `robots/oms/убрать место.xlsx`).

Результат: `robots/oms/form_diagnosis/data/clear_place_result_*.csv`.
