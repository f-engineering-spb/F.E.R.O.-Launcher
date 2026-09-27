@echo off
chcp 65001 >nul
cd /d "%~dp0.."

if not exist "dist\FEngineeringLauncher\FEngineeringLauncher.exe" (
    echo [INFO] dist\FEngineeringLauncher not found. Building EXE first...
    call "%~dp0build_exe.cmd"
)

set "ISCC_PATH="
if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" set "ISCC_PATH=C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if not defined ISCC_PATH if exist "C:\Program Files\Inno Setup 6\ISCC.exe" set "ISCC_PATH=C:\Program Files\Inno Setup 6\ISCC.exe"
if not defined ISCC_PATH (
    for /f "delims=" %%i in ('where iscc.exe 2^>nul') do set "ISCC_PATH=%%i"
)

if not defined ISCC_PATH (
    echo [ERROR] Inno Setup 6 ^(ISCC.exe^) not found!
    echo Download free Inno Setup 6: https://jrsoftware.org/isdl.php
    exit /b 1
)

"%ISCC_PATH%" "installer\launcher_setup.iss"

if %ERRORLEVEL% EQU 0 (
    echo [OK] Installer built in dist_setup\
    dir "dist_setup\FEngineering_Launcher_v3_Setup.exe"
) else (
    echo [ERROR] Installer build failed.
)
