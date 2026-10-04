# Цель 3 (БСК) — Web.ОМС ambulatory, несколько услуг на талон.

Путь: `robots/oms/cel_3/` (как `cel_307`, но два Excel: талоны + услуги).

## Запуск

```powershell
cd D:\Projects\BrowserAuto
copy robots\oms\cel_3\credentials.env.example robots\oms\cel_3\credentials.env

# один процесс (smoke)
.\.venv\Scripts\python run_3.py `
  --talons robots\oms\cel_3\data\талоны.xlsx `
  --services robots\oms\cel_3\data\услуги.xlsx `
  --category БСК --limit 1

# 5 воркеров — все категории (по умолчанию)
.\robots\oms\cel_3\run_parallel.ps1 -Workers 5 `
  -Talons "D:\Projects\BrowserAuto\robots\oms\cel_3\data\талоны.xlsx" `
  -Services "D:\Projects\BrowserAuto\robots\oms\cel_3\data\услуги.xlsx"

# только БСК:
.\robots\oms\cel_3\run_parallel.ps1 -Workers 5 -Category BSK
```

По умолчанию **все категории**.  
Результат: `robots/oms/cel_3/data/oms_3_result_*.csv`.  
Логи: `robots/oms/cel_3/data/logs/w*.err.log`.
