# robots/iszl — роботы ИСЗЛ

| Папка | Назначение | Launcher / вход |
|-------|------------|-----------------|
| `carryover/` | перенос плана ДН («нет в текущем — есть в прошлых») | `run_carryover_plan.py` |
| `download_dn/` | скачать «План ДН общий» | `python robots/iszl/download_dn/download_dn_plan.py` |
| `workplace/` | отчёты «по месту работы» | `python robots/iszl/workplace/download_workplace.py` |
| `oms_to_iszl/` | связка ОМС→ИСЗЛ (локально, не в публичном git) | см. файлы внутри |

Общее ядро: `browser_auto/` + `scripts/iszl_login.py`.
