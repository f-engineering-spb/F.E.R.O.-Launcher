param(
  [Parameter(Mandatory = $true)][string]$InputPath,
  [Parameter(Mandatory = $true)][string]$OutputPath
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)

if (-not (Test-Path -LiteralPath $InputPath -PathType Leaf)) {
  throw "DWG file was not found: $InputPath"
}

New-Item -ItemType Directory -Force -Path (Split-Path -Parent $OutputPath) | Out-Null
Remove-Item -LiteralPath $OutputPath -Force -ErrorAction SilentlyContinue

if (-not ([System.Management.Automation.PSTypeName]'LauncherWin32').Type) {
  Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class LauncherWin32 {
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
}
"@
}

$app = $null
$document = $null
$cadPid = 0
try {
  # Open read-only.  The original DWG and its source directory are never a
  # write target for the preview pipeline.
  $comProgIds = @(
    "AutoCAD.Application.24",
    "AutoCAD.Application"
  )
  $app = $null
  foreach ($progId in $comProgIds) {
    try {
      $app = New-Object -ComObject $progId -ErrorAction Stop
      if ($app) { break }
    } catch {}
  }
  if (-not $app) {
    throw "CAD COM error: AutoCAD could not be initialized."
  }
  try {
    $hwnd = [IntPtr]::new([long]$app.HWND)
    [uint32]$pidOut = 0
    [void][LauncherWin32]::GetWindowThreadProcessId($hwnd, [ref]$pidOut)
    if ($pidOut -gt 0) { $cadPid = [int]$pidOut }
  } catch {}

  $app.Visible = $false
  $document = $app.Documents.Open($InputPath, $true)
  $document.SetVariable("BACKGROUNDPLOT", 0)
  try { $document.SetVariable("EXPERT", 5) } catch {}

  # Model Space overview: full A0 page, extents, scale-to-fit and no plotted
  # lineweights.  It is a fast visual map, not a replacement for CAD layouts.
  $layout = $document.ModelSpace.Layout
  $devices = @($layout.GetPlotDeviceNames())
  $preferredDevices = @(
    "DWG To PDF.pc3",
    "AutoCAD PDF (General Documentation).pc3",
    "AutoCAD PDF (High Quality Print).pc3",
    "Microsoft Print to PDF"
  )
  foreach ($dev in $preferredDevices) {
    if ($devices -contains $dev) {
      $layout.ConfigName = $dev
      $layout.RefreshPlotDeviceInfo()
      break
    }
  }
  $layout.RefreshPlotDeviceInfo()
  $a0Media = @($layout.GetCanonicalMediaNames() | Where-Object { $_ -match "A0" } | Select-Object -First 1)
  if ($a0Media.Count -ne 1) {
    throw "CAD PDF-плоттер не предоставил формат A0."
  }
  $layout.CanonicalMediaName = $a0Media[0]
  $layout.PlotType = 1 # acExtents
  $layout.CenterPlot = $true
  $layout.UseStandardScale = $true
  $layout.StandardScale = 0 # acScaleToFit
  $layout.PlotWithLineweights = $false
  $layout.PlotWithPlotStyles = $true

  if (-not $document.Plot.PlotToFile($OutputPath)) {
    throw "CAD-система не подтвердила создание PDF для пространства модели."
  }
  if (-not (Test-Path -LiteralPath $OutputPath) -or (Get-Item -LiteralPath $OutputPath).Length -le 1024) {
    throw "CAD-система не создала PDF-файл превью."
  }
} finally {
  try { [LauncherMessageFilter]::Revoke() } catch {}
  if ($document) {
    try { $document.Close($false) } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($document) | Out-Null } catch {}
  }
  if ($app) {
    try { $app.Quit() } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}
  }
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()

  if ($cadPid -gt 0) {
    $deadline = (Get-Date).AddSeconds(3)
    while ((Get-Date) -lt $deadline) {
      $p = Get-Process -Id $cadPid -ErrorAction SilentlyContinue
      if (-not $p -or $p.HasExited) { break }
      Start-Sleep -Milliseconds 200
    }
    try {
      $p = Get-Process -Id $cadPid -ErrorAction SilentlyContinue
      if ($p -and -not $p.HasExited) {
        Stop-Process -Id $cadPid -Force -ErrorAction SilentlyContinue
      }
    } catch {}
  }
}
