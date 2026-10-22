# Синхронизация услуг по zip портала

Читает на талоне текущие коды → сравнивает с файлом услуг → приводит к целевому набору:

1. **лишние** строки → **заменить** на недостающие коды  
2. если недостающих нет → **удалить** лишние («Удалить сведение»)  
3. если строк не хватает → **#add-service** и заполнить новую (как `cel_3`)  
4. если уже совпало — **SKIP** (без Save)

```powershell
cd D:\Projects\BrowserAuto

# 1–3 талона
.\.venv\Scripts\python run_sync_services.py --zip robots\oms\fix_service\data\dn_sync_….zip --limit 3

# или напрямую xlsx
.\.venv\Scripts\python run_sync_services.py `
  --talons robots\oms\fix_service\data\талоны_….xlsx `
  --services robots\oms\fix_service\data\услуги_….xlsx `
  --limit 3

# 4 потока
.\robots\oms\fix_service\run_sync_parallel.ps1 -Workers 4 -Zip "D:\Projects\BrowserAuto\robots\oms\fix_service\data\dn_sync_….zip"
```

Статусы в CSV: `skip` | `ok` | `fail` | `error` | `dry_run`.  
Лог каждые 10 талонов: `прогресс N/M | skip=… patch_ok=… fail=…`.
