param(
    [switch]$NoLaunch
)

# F-Engineering Launcher v3 — Мастер установки и проверки системы
# Скрипт проверяет готовность ПК, устанавливает зависимости и создаёт ярлык на Рабочем столе.

[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$env:PYTHONUTF8 = "1"
$env:PYTHONIOENCODING = "utf-8"

$repoRoot = Split-Path -Parent $PSScriptRoot
$requirementsFile = Join-Path $repoRoot "requirements.txt"
$icoPath = Join-Path $repoRoot "app\frontend\assets\flauncher.ico"
$startCmd = Join-Path $repoRoot "scripts\start_windows.cmd"
$startAppPs1 = Join-Path $repoRoot "scripts\start_app.ps1"

Write-Host ""
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host "   F-Engineering Launcher v3 — Мастер установки и проверки системы      " -ForegroundColor Cyan
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host ""

$hasErrors = $false
$warningsCount = 0

# ----------------------------------------------------------------------
# 1. Поиск и проверка Python
# ----------------------------------------------------------------------
Write-Host "[1/6] Проверка интерпретатора Python..." -NoNewline

$pythonExe = $null
$candidates = @(
    (Join-Path $repoRoot "runtime\python\python.exe"),
    "python.exe",
    "C:\Python314\python.exe",
    "C:\Python313\python.exe",
    "C:\Python312\python.exe",
    "C:\Python311\python.exe",
    "C:\Python310\python.exe"
)

$localPythons = Get-ChildItem -Path "$env:LOCALAPPDATA\Programs\Python" -Filter "python.exe" -Recurse -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
if ($localPythons) {
    $candidates += $localPythons
}

foreach ($cand in $candidates) {
    try {
        $res = & $cand --version 2>&1
        if ($res -match "Python (\d+)\.(\d+)") {
            $major = [int]$matches[1]
            $minor = [int]$matches[2]
            if ($major -ge 3 -and $minor -ge 10) {
                $pythonExe = (Get-Command $cand -ErrorAction SilentlyContinue).Source
                if (-not $pythonExe) { $pythonExe = $cand }
                break
            }
        }
    } catch {}
}

if (-not $pythonExe) {
    Write-Host " [НЕ НАЙДЕН]" -ForegroundColor Red
    Write-Host ""
    Write-Host "  ОШИБКА: Python 3.10+ не найден на этом компьютере!" -ForegroundColor Red
    Write-Host "  Для работы лаунчера требуется Python." -ForegroundColor Yellow
    Write-Host "  Как установить:" -ForegroundColor Yellow
    Write-Host "   1. Скачайте официальный установщик: https://www.python.org/downloads/" -ForegroundColor Gray
    Write-Host "   2. ВАЖНО: в первом окне установщика включите галочку [x] 'Add python.exe to PATH'" -ForegroundColor White
    Write-Host "   3. Либо выполните в консоли: winget install Python.Python.3.12" -ForegroundColor Gray
    Write-Host ""
    $hasErrors = $true
} else {
    $verString = (& $pythonExe --version 2>&1).ToString().Trim()
    Write-Host " [OK]" -ForegroundColor Green
    Write-Host "      Обнаружен: $pythonExe ($verString)" -ForegroundColor Gray
}

# ----------------------------------------------------------------------
# 2. Проверка и установка библиотек Python (requirements.txt)
# ----------------------------------------------------------------------
if ($pythonExe) {
    Write-Host "[2/6] Проверка библиотек Python (PyMuPDF, openpyxl, pypdf)..." -NoNewline
    $checkLibsCmd = "import sys; import fitz, openpyxl, pypdf; sys.exit(0)"
    $libCheck = & $pythonExe -c $checkLibsCmd 2>&1

    if ($LASTEXITCODE -ne 0) {
        Write-Host " [ТРЕБУЕТСЯ УСТАНОВКА]" -ForegroundColor Yellow
        Write-Host "      Устанавливаю необходимые библиотеки через pip..." -ForegroundColor Cyan
        & $pythonExe -m pip install --quiet -r $requirementsFile
        
        # Повторная проверка
        $libCheck2 = & $pythonExe -c $checkLibsCmd 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Host "      Библиотеки успешно установлены." -ForegroundColor Green
        } else {
            Write-Host "  ОШИБКА установки библиотек: $libCheck2" -ForegroundColor Red
            $hasErrors = $true
        }
    } else {
        Write-Host " [OK]" -ForegroundColor Green
        Write-Host "      Все модули (PyMuPDF, openpyxl, pypdf) на месте." -ForegroundColor Gray
    }
} else {
    Write-Host "[2/6] Пропуск проверки библиотек (нет Python)." -ForegroundColor DarkGray
}

# ----------------------------------------------------------------------
# 3. Проверка САПР (AutoCAD)
# ----------------------------------------------------------------------
Write-Host "[3/6] Проверка САПР-движка для рендеринга DWG..." -NoNewline
$hasAutocad = Test-Path "Registry::HKEY_CLASSES_ROOT\AutoCAD.Application"

if ($hasAutocad) {
    Write-Host " [OK: AutoCAD]" -ForegroundColor Green
    Write-Host "      Обнаружен AutoCAD. Автоматический рендеринг DWG -> PDF активен." -ForegroundColor Gray
} else {
    Write-Host " [ПРЕДУПРЕЖДЕНИЕ]" -ForegroundColor Yellow
    Write-Host "      САПР (AutoCAD/AutoCAD) не зарегистрирован в реестре Windows." -ForegroundColor Yellow
    Write-Host "      Лаунчер будет работать со всеми PDF, Excel и Word, но автоматическая" -ForegroundColor Gray
    Write-Host "      генерация PDF из DWG на этом ПК будет отключена до установки САПР." -ForegroundColor Gray
    $warningsCount++
}

# ----------------------------------------------------------------------
# 4. Проверка Microsoft Office (Excel и Word)
# ----------------------------------------------------------------------
Write-Host "[4/6] Проверка MS Office для превью таблиц и документов..." -NoNewline
$hasExcel = Test-Path "Registry::HKEY_CLASSES_ROOT\Excel.Application"
$hasWord = Test-Path "Registry::HKEY_CLASSES_ROOT\Word.Application"

if ($hasExcel -and $hasWord) {
    Write-Host " [OK]" -ForegroundColor Green
    Write-Host "      MS Excel и Word обнаружены. Быстрый предпросмотр документов готов." -ForegroundColor Gray
} elseif ($hasExcel) {
    Write-Host " [OK: Только Excel]" -ForegroundColor Green
    Write-Host "      MS Excel доступен. MS Word не найден." -ForegroundColor Gray
} else {
    Write-Host " [ПРЕДУПРЕЖДЕНИЕ]" -ForegroundColor Yellow
    Write-Host "      MS Office не найден. Просмотр Excel/Word будет ограничен." -ForegroundColor Gray
    $warningsCount++
}

# ----------------------------------------------------------------------
# 5. Проверка браузера / оконного режима (Edge / Chrome)
# ----------------------------------------------------------------------
Write-Host "[5/6] Проверка браузерного движка для окна приложения..." -NoNewline
$edgePaths = @(
    "$env:ProgramFiles(x86)\Microsoft\Edge\Application\msedge.exe",
    "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe"
)
$hasEdge = $false
foreach ($ep in $edgePaths) {
    if (Test-Path -LiteralPath $ep) { $hasEdge = $true; break }
}
$chromePath = "$env:ProgramFiles\Google\Chrome\Application\chrome.exe"
$hasChrome = Test-Path -LiteralPath $chromePath

if ($hasEdge -or $hasChrome) {
    Write-Host " [OK]" -ForegroundColor Green
    $browserName = if ($hasEdge) { "Microsoft Edge" } else { "Google Chrome" }
    Write-Host "      $browserName готов к запуску в режиме отдельного окна." -ForegroundColor Gray
} else {
    Write-Host " [ПРЕДУПРЕЖДЕНИЕ]" -ForegroundColor Yellow
    Write-Host "      Edge/Chrome не найдены. Лаунчер откроется в браузере по умолчанию." -ForegroundColor Gray
    $warningsCount++
}

# ----------------------------------------------------------------------
# 6. Инициализация служебных папок и создание ярлыка
# ----------------------------------------------------------------------
Write-Host "[6/6] Настройка рабочего места..." -NoNewline

New-Item -ItemType Directory -Force -Path (Join-Path $repoRoot "runtime\cache") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $repoRoot "runtime\logs") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $repoRoot "runtime\manifests") | Out-Null

Write-Host "[6/6] Создание ярлыков (Рабочий стол и меню Пуск)..." -NoNewline
try {
    $wsh = New-Object -ComObject WScript.Shell

    # 1. Рабочий стол
    $desktopPath = [Environment]::GetFolderPath("Desktop")
    $desktopShortcut = $wsh.CreateShortcut((Join-Path $desktopPath "F-Engineering Launcher.lnk"))
    $desktopShortcut.TargetPath = $startCmd
    $desktopShortcut.WorkingDirectory = $repoRoot
    $desktopShortcut.Description = "F-Engineering Launcher v3"
    if (Test-Path -LiteralPath $icoPath) {
        $desktopShortcut.IconLocation = "$icoPath, 0"
    }
    $desktopShortcut.Save()

    # 2. Главное меню Пуск (в корне списка программ Programs)
    $startMenuPrograms = [Environment]::GetFolderPath("Programs")
    if ($startMenuPrograms -and (Test-Path -LiteralPath $startMenuPrograms)) {
        $startMenuShortcut = $wsh.CreateShortcut((Join-Path $startMenuPrograms "F-Engineering Launcher.lnk"))
        $startMenuShortcut.TargetPath = $startCmd
        $startMenuShortcut.WorkingDirectory = $repoRoot
        $startMenuShortcut.Description = "F-Engineering Launcher v3"
        if (Test-Path -LiteralPath $icoPath) {
            $startMenuShortcut.IconLocation = "$icoPath, 0"
        }
        $startMenuShortcut.Save()
    }

    # 3. Список недавних / последних установленных (Recent / Start Menu Root)
    $startMenuRoot = [Environment]::GetFolderPath("StartMenu")
    if ($startMenuRoot -and (Test-Path -LiteralPath $startMenuRoot)) {
        $rootShortcut = $wsh.CreateShortcut((Join-Path $startMenuRoot "F-Engineering Launcher.lnk"))
        $rootShortcut.TargetPath = $startCmd
        $rootShortcut.WorkingDirectory = $repoRoot
        $rootShortcut.Description = "F-Engineering Launcher v3"
        if (Test-Path -LiteralPath $icoPath) {
            $rootShortcut.IconLocation = "$icoPath, 0"
        }
        $rootShortcut.Save()
    }

    # Регистрация в списке последних установленных приложений Windows (Recent)
    $recentPath = [Environment]::GetFolderPath("Recent")
    if ($recentPath -and (Test-Path -LiteralPath $recentPath)) {
        $recentShortcut = $wsh.CreateShortcut((Join-Path $recentPath "F-Engineering Launcher.lnk"))
        $recentShortcut.TargetPath = $startCmd
        $recentShortcut.WorkingDirectory = $repoRoot
        $recentShortcut.Save()
    }

    Write-Host " [OK]" -ForegroundColor Green
    Write-Host "      ✓ Ярлык создан на Рабочем столе: 'F-Engineering Launcher'" -ForegroundColor Cyan
    Write-Host "      ✓ Ярлык зарегистрирован в меню Пуск (вверху и в списке последних установленных)" -ForegroundColor Cyan
} catch {
    Write-Host " [ВНИМАНИЕ]" -ForegroundColor Yellow
    Write-Host "      Не удалось создать некоторые ярлыки: $($_.Exception.Message)" -ForegroundColor Gray
}
# ----------------------------------------------------------------------
# Итоги и запуск
# ----------------------------------------------------------------------
Write-Host ""
Write-Host "========================================================================" -ForegroundColor Cyan

if ($hasErrors) {
    Write-Host "  ВНИМАНИЕ: Найдено несколько критических замечаний (см. выше)." -ForegroundColor Red
    Write-Host "  Пожалуйста, установите недостающие компоненты и запустите SETUP.cmd повторно." -ForegroundColor Red
    Write-Host "========================================================================" -ForegroundColor Cyan
    Write-Host ""
    return
}

Write-Host "  ГОТОВО: Лаунчер полностью настроен и готов к работе!" -ForegroundColor Green
if ($warningsCount -gt 0) {
    Write-Host "  (Есть $warningsCount некритических предупреждений, лаунчер будет работать)" -ForegroundColor Yellow
}
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host ""

if (-not $NoLaunch) {
    $choice = Read-Host "Запустить F-Engineering Launcher прямо сейчас? [Y/n]"
    if ($choice -ne "n" -and $choice -ne "N" -and $choice -ne "т" -and $choice -ne "Т") {
        Write-Host "Запускаю лаунчер..." -ForegroundColor Cyan
        & $startAppPs1
    }
}
