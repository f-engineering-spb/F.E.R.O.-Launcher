# Invoke-NativeDwgPdfExport.ps1 — shared native AutoCAD PDF export helper via accoreconsole.exe
# ASCII-only encoding for PowerShell 5.1 compatibility.

$script:NativeExportAccorePath = ""

function Find-NativeAccoreConsole {
  if ($script:NativeExportAccorePath -and (Test-Path -LiteralPath $script:NativeExportAccorePath -PathType Leaf)) {
    return $script:NativeExportAccorePath
  }
  $candidates = @()
  
  # 1. Launcher config app setting
  try {
    $cfgPath = Join-Path $env:APPDATA 'F-Engineering Launcher\config.json'
    if (Test-Path -LiteralPath $cfgPath) {
      $cfg = Get-Content -LiteralPath $cfgPath -Raw -Encoding UTF8 | ConvertFrom-Json
      $cadExe = $cfg.nativeApps.dwg
      if ($cadExe -and (Test-Path -LiteralPath $cadExe)) {
        $c = Join-Path (Split-Path -Parent $cadExe) 'accoreconsole.exe'
        if (Test-Path -LiteralPath $c) { $candidates += $c }
      }
    }
  } catch {}

  # 2. Standard paths for AutoCAD 2021-2026
  foreach ($ver in @("2026", "2025", "2024", "2023", "2022", "2021")) {
    $candidates += (Join-Path $env:ProgramFiles "Autodesk\AutoCAD $ver\accoreconsole.exe")
  }
  try {
    $x86 = [System.Environment]::GetEnvironmentVariable('ProgramFiles(x86)')
    if ($x86) {
      foreach ($ver in @("2026", "2025", "2024", "2023", "2022", "2021")) {
        $candidates += (Join-Path $x86 "Autodesk\AutoCAD $ver\accoreconsole.exe")
      }
    }
  } catch {}

  foreach ($c in $candidates) {
    if ($c -and (Test-Path -LiteralPath $c -PathType Leaf)) {
      $script:NativeExportAccorePath = $c
      return $c
    }
  }
  return $null
}

function Invoke-NativeDwgPdfExport {
  param(
    [Parameter(Mandatory = $true)][string]$InputPath,
    [Parameter(Mandatory = $true)][string]$OutputPdf,
    [Parameter(Mandatory = $true)][string]$WorkDir,
    [int]$TimeoutSec = 300,
    [string]$LayoutName = "Layout1"
  )
  if (-not (Test-Path -LiteralPath $InputPath -PathType Leaf)) {
    throw "DWG-файл не найден: $InputPath"
  }
  $accore = Find-NativeAccoreConsole
  if (-not $accore) {
    throw "accoreconsole.exe не найден"
  }
  try { New-Item -ItemType Directory -Force -Path $WorkDir | Out-Null } catch {}
  try { Remove-Item -LiteralPath $OutputPdf -Force -ErrorAction SilentlyContinue } catch {}

  $scrFile = Join-Path $WorkDir "native_export.scr"
  $consoleLog = Join-Path $WorkDir "native_export.log"

  # AutoCAD Script: check layout, export to vector PDF, exit without saving
  $scriptLines = @(
    '(setvar "EXPERT" 5)',
    '(setvar "BACKGROUNDPLOT" 0)',
    ('(princ (strcat "LAYOUTCHECK:" (if (tblsearch "LAYOUT" "{0}") "FOUND" "MISSING")))' -f $LayoutName),
    ('if (= (getvar "CTAB") "Model") (progn (if (tblsearch "LAYOUT" "{0}") (setvar "CTAB" "{0}"))))' -f $LayoutName),
    '_.-EXPORT',
    '_PDF',
    '_C',
    '_N',
    ('"{0}"' -f $OutputPdf),
    '_QUIT',
    '_N'
  )
  [System.IO.File]::WriteAllLines($scrFile, $scriptLines, [System.Text.UTF8Encoding]::new($false))

  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  $runStart = Get-Date

  $psi = New-Object System.Diagnostics.ProcessStartInfo
  $psi.FileName = $accore
  $psi.Arguments = ('/i "{0}" /s "{1}"' -f $InputPath, $scrFile)
  $psi.UseShellExecute = $false
  $psi.CreateNoWindow = $true
  $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
  $psi.RedirectStandardOutput = $true
  $psi.RedirectStandardError = $true

  $proc = [System.Diagnostics.Process]::Start($psi)
  try {
    $outTask = $proc.StandardOutput.ReadToEndAsync()
    $errTask = $proc.StandardError.ReadToEndAsync()
    if (-not $proc.WaitForExit($TimeoutSec * 1000)) {
      try { $proc.Kill() } catch {}
      throw ("Таймаут консоли AutoCAD Core Console ({0} сек)" -f $TimeoutSec)
    }
    try { ($outTask.Result + "`n" + $errTask.Result) | Out-File -LiteralPath $consoleLog -Encoding utf8 -Force } catch {}
  } finally {
    try { if (-not $proc.HasExited) { $proc.Kill() } } catch {}
    try { $proc.Dispose() } catch {}
  }

  if (-not (Test-Path -LiteralPath $OutputPdf)) {
    # If Layout1 export didn't create file (e.g. model space drawing), try exporting Model space directly
    $scrModel = Join-Path $WorkDir "native_model.scr"
    $scriptModelLines = @(
      '(setvar "EXPERT" 5)',
      '(setvar "BACKGROUNDPLOT" 0)',
      '(setvar "CTAB" "Model")',
      '_.-EXPORT',
      '_PDF',
      '_D',
      '_N',
      ('"{0}"' -f $OutputPdf),
      '_QUIT',
      '_N'
    )
    [System.IO.File]::WriteAllLines($scrModel, $scriptModelLines, [System.Text.UTF8Encoding]::new($false))
    $psiModel = New-Object System.Diagnostics.ProcessStartInfo
    $psiModel.FileName = $accore
    $psiModel.Arguments = ('/i "{0}" /s "{1}"' -f $InputPath, $scrModel)
    $psiModel.UseShellExecute = $false
    $psiModel.CreateNoWindow = $true
    $psiModel.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
    $procM = [System.Diagnostics.Process]::Start($psiModel)
    try {
      if (-not $procM.WaitForExit($TimeoutSec * 1000)) {
        try { $procM.Kill() } catch {}
      }
    } finally {
      try { if (-not $procM.HasExited) { $procM.Kill() } } catch {}
      try { $procM.Dispose() } catch {}
    }
  }

  if (-not (Test-Path -LiteralPath $OutputPdf)) {
    throw "AutoCAD Core Console завершила работу, но PDF-файл не был создан."
  }
  $info = Get-Item -LiteralPath $OutputPdf
  if ($info.Length -le 1024) {
    throw "Созданный PDF-файл пуст или поврежден (размер менее 1 КБ)."
  }
  return [ordered]@{
    ok = $true
    pdfPath = $OutputPdf
    exportMs = [int]$sw.ElapsedMilliseconds
    pageCount = 1
  }
}
