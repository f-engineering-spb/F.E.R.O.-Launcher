@echo off
setlocal
chcp 65001 >nul

powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ^
    "$wsh = New-Object -ComObject WScript.Shell; " ^
    "$desktop = [Environment]::GetFolderPath('Desktop'); " ^
    "$programs = [Environment]::GetFolderPath('Programs'); " ^
    "$startMenu = [Environment]::GetFolderPath('StartMenu'); " ^
    "$recent = [Environment]::GetFolderPath('Recent'); " ^
    "$pyw1 = Join-Path '%~dp0' 'runtime\python\pythonw.exe'; " ^
    "$pyw2 = 'C:\Python314\pythonw.exe'; " ^
    "$pyw = if (Test-Path $pyw1) { $pyw1 } elseif (Test-Path $pyw2) { $pyw2 } else { '' }; " ^
    "$app = Join-Path '%~dp0' 'app\flauncher.pyw'; " ^
    "$vbs = Join-Path '%~dp0' 'Запуск_Лаунчера.vbs'; " ^
    "$runCmd = Join-Path '%~dp0' 'Запуск_Лаунчера.cmd'; " ^
    "$ico = Join-Path '%~dp0' 'app\frontend\assets\flauncher.ico'; " ^
    "$makeShortcut = { param($folder, $name); " ^
    "   if (-not $folder -or -not (Test-Path $folder)) { return }; " ^
    "   $s = $wsh.CreateShortcut((Join-Path $folder $name)); " ^
    "   if (Test-Path $vbs) { $s.TargetPath = 'wscript.exe'; $s.Arguments = '\"' + $vbs + '\"'; } elseif ($pyw) { $s.TargetPath = $pyw; $s.Arguments = '\"' + $app + '\"'; } else { $s.TargetPath = $runCmd; }; " ^
    "   $s.WorkingDirectory = '%~dp0'; " ^
    "   $s.Description = 'F-Engineering Launcher v3'; " ^
    "   if (Test-Path $ico) { $s.IconLocation = $ico + ',0'; }; " ^
    "   $s.Save(); " ^
    "}; " ^
    "& $makeShortcut $desktop 'F-Engineering Launcher.lnk'; " ^
    "& $makeShortcut $programs 'F-Engineering Launcher.lnk'; " ^
    "& $makeShortcut $startMenu 'F-Engineering Launcher.lnk'; " ^
    "& $makeShortcut $recent 'F-Engineering Launcher.lnk'; " ^
    "Write-Host ''; " ^
    "Write-Host '  [OK] Ярлык успешно создан:' -ForegroundColor Green; " ^
    "Write-Host '       1. На Рабочем столе' -ForegroundColor Cyan; " ^
    "Write-Host '       2. В меню Пуск (в списке программ и последних установленных)' -ForegroundColor Cyan; " ^
    "Write-Host ''"

pause
endlocal
