# robots/oms — роботы Web.ОМС

| Папка | Цель | Launcher |
|-------|------|----------|
| `cel_3/` | цель 3 (БСК), талоны+услуги | `run_3.py` |
| `cel_307/` | цель 307 | `run_307.py` |
| `dv_opv/` | ДВ4 / ОПВ | `run_medical_exam.py` |
| `socstatus/` | соцстатус в карте пациента | `run_socstatus.py` |

Общие хелперы ОМС живут в `dv_opv/add_medical_exam.py` — их импортируют `cel_307` и `cel_3`.
