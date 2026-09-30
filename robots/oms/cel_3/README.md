# Цель 3 (БСК) — Web.ОМС ambulatory, несколько услуг на талон.

Путь: `robots/oms/cel_3/` (как `cel_307`, но два Excel: талоны + услуги).

## Запуск

```powershell
cd D:\Projects\BrowserAuto
copy robots\oms\cel_3\credentials.env.example robots\oms\cel_3\credentials.env

.\.venv\Scripts\python run_3.py `
  --talons robots\oms\cel_3\data\talons_3_bsk_test.xlsx `
  --services robots\oms\cel_3\data\services_3_bsk_test.xlsx `
  --building "ГП №3" --doctor 30097144 --limit 3

# или из zip-бандла:
.\.venv\Scripts\python run_3.py --zip robots\oms\cel_3\data\talon_bundle_….zip `
  --category БСК --building "ГП №3" --doctor 30097144 --limit 3
```

Результат: `robots/oms/cel_3/data/oms_3_result_*.csv`.
