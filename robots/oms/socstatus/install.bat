@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo === oms_socstatus: установка ===

where py >nul 2>&1
if %ERRORLEVEL%==0 (
  set PY=py -3
) else (
  where python >nul 2>&1
  if %ERRORLEVEL%==0 (
    set PY=python
  ) else (
    echo ERROR: Python 3 не найден. Установите с https://www.python.org/downloads/
    echo При установке отметьте "Add python.exe to PATH".
    pause
    exit /b 1
  )
)

if not exist ".venv\Scripts\python.exe" (
  echo Создаю виртуальное окружение .venv ...
  %PY% -m venv .venv
  if errorlevel 1 (
    echo ERROR: не удалось создать venv
    pause
    exit /b 1
  )
)

echo Устанавливаю зависимости...
".venv\Scripts\python.exe" -m pip install --upgrade pip
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
  echo ERROR: pip install failed
  pause
  exit /b 1
)

if not exist "credentials.env" (
  copy /Y "credentials.env.example" "credentials.env" >nul
  echo Создан credentials.env — заполните OMS_LOGIN и OMS_PASSWORD.
)

if not exist "data" mkdir data

if not exist "geckodriver.exe" (
  echo.
  echo geckodriver.exe не найден — скачиваю...
  ".venv\Scripts\python.exe" tools\download_geckodriver.py
  if errorlevel 1 (
    echo.
    echo Не удалось скачать автоматически.
    echo Скачайте geckodriver для Windows с:
    echo   https://github.com/mozilla/geckodriver/releases
    echo и положите geckodriver.exe в эту папку.
  )
)

echo.
echo Готово.
echo 1^) Заполните credentials.env
echo 2^) Положите Excel в data\list.xlsx ^(колонки: ЕНП, Соцстатус, Вид занятости^)
echo 3^) Запустите run.bat
echo.
pause
