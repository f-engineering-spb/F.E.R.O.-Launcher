@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"

:: Очистка зависших невидимых фоновых процессов AutoCAD без окон
taskkill /F /FI "WINDOWTITLE eq " /IM acad.exe 2>nul

:: 1. Автономный портативный Python (внутри перенесенной папки или на флешке)
if exist "%~dp0runtime\python\pythonw.exe" (
    start "" "%~dp0runtime\python\pythonw.exe" "%~dp0app\flauncher.pyw"
    exit /b 0
)
if exist "%~dp0runtime\python\python.exe" (
    start "" "%~dp0runtime\python\python.exe" "%~dp0app\flauncher.pyw"
    exit /b 0
)

:: 2. Системный Python (если запускают из репозитория разработчика)
if exist "C:\Python314\pythonw.exe" (
    start "" "C:\Python314\pythonw.exe" "%~dp0app\flauncher.pyw"
    exit /b 0
)
where pythonw >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    start "" pythonw "%~dp0app\flauncher.pyw"
    exit /b 0
)
where python >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    start "" python "%~dp0app\flauncher.pyw"
    exit /b 0
)

:: 3. Если ничего не найдено - мастер проверки
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\setup_pc.ps1"
endlocal