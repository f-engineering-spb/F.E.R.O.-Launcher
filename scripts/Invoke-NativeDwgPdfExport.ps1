# Invoke-NativeDwgPdfExport.ps1 - shared native AutoCAD PDF export helper via accoreconsole.exe
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
    throw "DWG file not found: $InputPath"
  }
  $accore = Find-NativeAccoreConsole
  if (-not $accore) {
    throw "accoreconsole.exe not found"
  }
  try { New-Item -ItemType Directory -Force -Path $WorkDir | Out-Null } catch {}
  try { Remove-Item -LiteralPath $OutputPdf -Force -ErrorAction SilentlyContinue } catch {}

  $scrFile = Join-Path $WorkDir "native_export.scr"
  $consoleLog = Join-Path $WorkDir "native_export.log"

  # Stage network/cloud drives (H:, G:, \\) to local SSD WorkDir for fast I/O
  $effectiveInput = $InputPath
  $stagedLocal = $false
  $rootPath = [System.IO.Path]::GetPathRoot($InputPath)
  if ($rootPath -and ($rootPath.StartsWith('\\') -or $rootPath.StartsWith('H:') -or $rootPath.StartsWith('G:') -or $rootPath.StartsWith('Z:'))) {
    try {
      $localCopy = Join-Path $WorkDir ([System.IO.Path]::GetFileName($InputPath))
      Copy-Item -LiteralPath $InputPath -Destination $localCopy -Force -ErrorAction Stop
      $effectiveInput = $localCopy
      $stagedLocal = $true
    } catch {}
  }

  $normalizedOut = $OutputPdf.Replace('\', '/')
  $scriptLines = @(
    '(setvar "EXPERT" 5)',
    '(setvar "BACKGROUNDPLOT" 0)',
    '(setvar "FILEDIA" 0)',
    '(setvar "CMDDIA" 0)',
    '(if (= (getvar "CTAB") "Model")',
    ('  (command "_.-EXPORT" "_PDF" "_E" "_N" "{0}")' -f $normalizedOut),
    ('  (command "_.-EXPORT" "_PDF" "_C" "_N" "{0}")' -f $normalizedOut),
    ')',
    '_.QUIT',
    '_N'
  )
  [System.IO.File]::WriteAllLines($scrFile, $scriptLines, [System.Text.Encoding]::ASCII)

  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  $runStart = Get-Date

  $psi = New-Object System.Diagnostics.ProcessStartInfo
  $psi.FileName = $accore
  $psi.Arguments = ('/i "{0}" /s "{1}" /readonly' -f $effectiveInput, $scrFile)
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
      throw ("AutoCAD Core Console timeout ({0} sec)" -f $TimeoutSec)
    }
    try { ($outTask.Result + "`n" + $errTask.Result) | Out-File -LiteralPath $consoleLog -Encoding utf8 -Force } catch {}
  } finally {
    try { if (-not $proc.HasExited) { $proc.Kill() } } catch {}
    try { $proc.Dispose() } catch {}
    if ($stagedLocal -and (Test-Path -LiteralPath $effectiveInput)) {
      try { Remove-Item -LiteralPath $effectiveInput -Force -ErrorAction SilentlyContinue } catch {}
    }
  }

  if (-not (Test-Path -LiteralPath $OutputPdf)) {
    throw "AutoCAD Core Console finished but PDF was not created"
  }
  $info = Get-Item -LiteralPath $OutputPdf
  if ($info.Length -le 1024) {
    throw "Created PDF is empty or corrupt (<1KB)"
  }
  return [ordered]@{
    ok = $true
    pdfPath = $OutputPdf
    exportMs = [int]$sw.ElapsedMilliseconds
    pageCount = 1
  }
}
