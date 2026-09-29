@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
set "PORT=8789"

if defined FENG_SILENT goto :skip_banner
echo ===============================================
echo  F-Engineering Launcher v3.2 (НОВОЕ ЯДРО)
echo  app/rendering: DOCX-HTML, Excel-вкладки, 150 DPI
echo ===============================================
:skip_banner

set "PY=python"
if exist "%~dp0.venv\Scripts\python.exe" set "PY=%~dp0.venv\Scripts\python.exe"
if exist "%~dp0venv\Scripts\python.exe" set "PY=%~dp0venv\Scripts\python.exe"
set "PYW="
if exist "%~dp0.venv\Scripts\pythonw.exe" set "PYW=%~dp0.venv\Scripts\pythonw.exe"
if exist "%~dp0venv\Scripts\pythonw.exe" set "PYW=%~dp0venv\Scripts\pythonw.exe"
if exist "%~dp0runtime\python\pythonw.exe" set "PYW=%~dp0runtime\python\pythonw.exe"
if exist "C:\Python314\pythonw.exe" set "PYW=C:\Python314\pythonw.exe"
where python >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
  if not defined FENG_SILENT echo [ERROR] Python не найден.
  if not defined FENG_SILENT pause
  exit /b 1
)
if not exist "%~dp0app\backend\server.py" (
  if not defined FENG_SILENT echo [ERROR] app\backend\server.py не найден.
  if not defined FENG_SILENT pause
  exit /b 1
)
if defined FENG_SILENT goto :start_silent
start "FEng-Server-8789" /min cmd /c ""%PY%" "%~dp0app\backend\server.py" --port %PORT%"
goto :wait_ready
:start_silent
rem SILENT: pythonw = вообще без окна; fallback — минимизированный cmd
if defined PYW (
  start "" "%PYW%" "%~dp0app\backend\server.py" --port %PORT%
) else (
  start "FEng-Server-8789" /min cmd /c ""%PY%" "%~dp0app\backend\server.py" --port %PORT%"
)
:wait_ready

set "READY=0"
for /L %%i in (1,1,20) do (
  curl -s -o nul http://127.0.0.1:%PORT%/api/health >nul 2>nul
  if %ERRORLEVEL% EQU 0 (
    set "READY=1"
    goto :server_up
  )
  timeout /t 1 /nobreak >nul
)
:server_up
if defined FENG_SILENT goto :launch_browser
if "%READY%"=="0" (
  echo [WARN] Сервер не ответил за 20 сек.
) else (
  echo [OK] Backend: http://127.0.0.1:%PORT%/api/health
)
:launch_browser
if "%FENG_NO_BROWSER%"=="1" goto :status
where msedge >nul 2>nul
if %ERRORLEVEL% EQU 0 (
  start "" msedge.exe --app=http://127.0.0.1:%PORT%/
  goto :status
)
where chrome >nul 2>nul
if %ERRORLEVEL% EQU 0 (
  start "" chrome.exe --app=http://127.0.0.1:%PORT%/
  goto :status
)
start "" http://127.0.0.1:%PORT%/
:status
if defined FENG_SILENT (
  endlocal
  exit /b 0
)
echo.
echo Запущен F-Engineering Launcher v3.2 (Новое ядро app/rendering, 150 DPI)
echo Сервер: http://127.0.0.1:%PORT%/
echo.
if not "%FENG_NO_BROWSER%"=="1" pause
endlocal