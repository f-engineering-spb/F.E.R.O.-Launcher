@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
cd /d "%~dp0.."

timeout /t 2 /nobreak >nul

rmdir /s /q app_backup 2>nul
xcopy /E /I /Q /Y app app_backup

set "ARCHIVE=%~1"
if "%ARCHIVE%"=="" set "ARCHIVE=runtime\cache\update_payload.zip"

powershell -NoProfile -Command "Expand-Archive -LiteralPath '%ARCHIVE%' -DestinationPath '.' -Force"

if %ERRORLEVEL%==0 (
    del "%ARCHIVE%" >nul 2>&1
    if exist "FEngineeringLauncher.exe" (start "" "FEngineeringLauncher.exe") else (start "" "Запуск_Лаунчера.cmd")
) else (
    echo Ошибка распаковки. Восстановление из бэкапа...
    xcopy /E /I /Q /Y app_backup app
    if exist "FEngineeringLauncher.exe" (start "" "FEngineeringLauncher.exe") else (start "" "Запуск_Лаунчера.cmd")
)

exit
