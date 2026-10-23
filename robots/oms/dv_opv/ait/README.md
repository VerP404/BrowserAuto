# ait — пачка карт врача (ДВ4 / ОПВ)

Сюда складываете выгрузки по врачу **30151001**, корпус **ГП №3**.

## Папки

| Путь | Назначение |
|------|------------|
| `ait/*.xlsx` | Inbox — новые файлы |
| `ait/done/` | Исходники после подготовки |
| `ait/ready/talons_ready.xlsx` | Один общий файл для робота |
| `ait/state/done_enps.txt` | Кэш уже введённых ЕНП (+ подтягивается из `data/dv_opv_result_*.csv`) |

## Подготовка

```powershell
cd D:\Projects\BrowserAuto

# собрать inbox → ready, починить даты, учесть уже ok
.\.venv\Scripts\python robots\oms\dv_opv\ait\prepare_batch.py --force-doctor

# только посмотреть, без записи
.\.venv\Scripts\python robots\oms\dv_opv\ait\prepare_batch.py --dry-run
```

Что делает `prepare_batch.py`:
- **даты** `начало` / `окончание` → текст `DD-MM-YY`  
  (чинит `datetime`, serial Excel `46258`, текст `"46258"`);
- пишет даты форматом `@` (текст), чтобы Excel снова не сломал;
- врач / корпус по умолчанию `30151001` / `ГП №3` (`--force-doctor` — всегда);
- склеивает все xlsx inbox в один `ready/talons_ready.xlsx`;
- пропускает ЕНП, уже успешно введённые;
- переносит обработанные файлы в `done/`.

## Запуск робота

```powershell
.\.venv\Scripts\python run_medical_exam.py `
  --file robots\oms\dv_opv\ait\ready\talons_ready.xlsx `
  --catalog robots\oms\dv_opv\catalog_dv_opv.xlsx
```

После прогона ЕНП со статусом `ok` в result CSV больше не попадут в следующую пачку.
