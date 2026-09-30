# socstatus — автономный робот соцстатуса Web.ОМС

Только соцстатус. Копируете **эту папку** (`robots/oms/socstatus`) на другой компьютер — без всего BrowserAuto.

| | |
|--|--|
| URL | `http://10.36.0.142:9000/patient/edit` |
| Логика | **18 < возраст < 60 → 11 / 1**, иначе **22 / 3** |
| Браузер | Firefox + `geckodriver.exe` |

---

## Установка (один раз)

1. Скопируйте папку `socstatus` целиком.
2. Дважды щёлкните **`install.bat`**
3. Заполните **`credentials.env`**
4. Положите Excel в **`data/list.xlsx`**

## Запуск

Из корня BrowserAuto:

```powershell
.\.venv\Scripts\python run_socstatus.py --limit 3
```

Или автономно:

```bat
run.bat
run.bat --limit 3 --dry-run --keep-open 30
```

Результат: `data/oms_socstatus_result_*.csv`.
