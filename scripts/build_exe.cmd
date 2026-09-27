@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
cd /d "%~dp0.."

where pyinstaller >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo [INFO] Установка PyInstaller...
    python -m pip install pyinstaller
)

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

python -m PyInstaller --noconfirm --onedir --windowed --name "FEngineeringLauncher" --icon "app\frontend\assets\flauncher.ico" --add-data "app;app" --distpath "dist" --workpath "build" "app\flauncher.pyw"

if %ERRORLEVEL% EQU 0 (
    echo [OK] FEngineeringLauncher.exe собран в dist\FEngineeringLauncher\
    dir "dist\FEngineeringLauncher\FEngineeringLauncher.exe"
) else (
    echo [ERROR] Сборка завершилась с ошибкой.
)
