@echo off
chcp 65001 >nul
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo Сначала запустите install.bat
  pause
  exit /b 1
)

if not exist "credentials.env" (
  echo Нет credentials.env — скопируйте credentials.env.example и заполните логин/пароль.
  pause
  exit /b 1
)

if not exist "geckodriver.exe" (
  echo Нет geckodriver.exe — запустите install.bat или положите драйвер вручную.
  pause
  exit /b 1
)

REM Можно передать аргументы: run.bat --limit 3
".venv\Scripts\python.exe" add_socstatus.py %*
set EXITCODE=%ERRORLEVEL%
echo.
pause
exit /b %EXITCODE%
