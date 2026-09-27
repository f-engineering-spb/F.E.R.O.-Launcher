param(
  [Parameter(Mandatory = $true)][string]$JobDir,
  [Parameter(Mandatory = $false)][string]$ReadyFile = "",
  [Parameter(Mandatory = $false)][string]$LogFile = "",
  [Parameter(Mandatory = $false)][int]$IdleMinutes = 15
)

# Долгоживущий конвертер DWG->PDF: ОДНА CAD-сессия на все файлы.
# Раньше каждый файл поднимал свой AutoCAD (старт 30-90 с, параллельные
# сессии дерутся за COM) — теперь сессия одна, задания идут очередью.
#
# Протокол (файлы в $JobDir, имена-гуиды от сервера):
#   <id>.job.json  -> { inputPath, outputPath, fallbackCachePath, pythonExe }
#   <id>.done.json -> { ok: true, result: {...} } | { ok: false, error: "..." }
#   daemon.fatal   -> { error } если CAD вообще не поднялся (сервер увидит и откатится).
# Без заданий дольше IdleMinutes демон гасит CAD и выходит сам (сервер поднимет заново).

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)

if (-not (Test-Path -LiteralPath $JobDir -PathType Container)) {
  New-Item -ItemType Directory -Force -Path $JobDir | Out-Null
}

function Write-DaemonLog([string]$message) {
  if ([string]::IsNullOrWhiteSpace($LogFile)) { return }
  try {
    $line = "{0:yyyy-MM-dd HH:mm:ss} [dwg-daemon] {1}" -f (Get-Date), $message
    Add-Content -LiteralPath $LogFile -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
  } catch {}
}

if (-not ([System.Management.Automation.PSTypeName]'LauncherMessageFilter').Type) {
  Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

[ComImport(), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("00000016-0000-0000-C000-000000000046")]
public interface IOleMessageFilter
{
    [PreserveSig] int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);
    [PreserveSig] int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);
    [PreserveSig] int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
}

public class LauncherMessageFilter : IOleMessageFilter
{
    [DllImport("ole32.dll")] private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);
    public static void Register() {
        IOleMessageFilter newFilter = new LauncherMessageFilter();
        IOleMessageFilter oldFilter = null;
        CoRegisterMessageFilter(newFilter, out oldFilter);
    }
    public static void Revoke() {
        IOleMessageFilter oldFilter = null;
        CoRegisterMessageFilter(null, out oldFilter);
    }
    public int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo) { return 0; }
    public int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType) {
        if (dwRejectType == 2) return 100;
        return -1;
    }
    public int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType) { return 2; }
}
"@
}
[LauncherMessageFilter]::Register()

$comProgIds = @(
  "AutoCAD.Application.25",
  "AutoCAD.Application.24.3",
  "AutoCAD.Application.24.2",
  "AutoCAD.Application.24.1",
  "AutoCAD.Application.24.0",
  "AutoCAD.Application.24",
  "AutoCAD.Application.23.1",
  "AutoCAD.Application.23",
  "AutoCAD.Application.22",
  "AutoCAD.Application"
)

function Get-AcadProcessIds {
  $ids = @()
  foreach ($procName in @("acad")) {
    try {
      foreach ($p in (Get-Process -Name $procName -ErrorAction SilentlyContinue)) {
        $ids += $p.Id
      }
    } catch {}
  }
  return $ids
}

function Connect-CadApp {
  foreach ($progId in $comProgIds) {
    try {
      $candidate = New-Object -ComObject $progId -ErrorAction Stop
      if ($candidate) {
        return @{ App = $candidate; ProgId = $progId }
      }
    } catch {}
  }
  return $null
}

function Get-AppPid($cadApp) {
  try {
    if (-not ([System.Management.Automation.PSTypeName]'LauncherWin32').Type) {
      Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class LauncherWin32 {
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
}
"@
    }
    $hwnd = [IntPtr]::new([long]$cadApp.HWND)
    [uint32]$pidOut = 0
    [void][LauncherWin32]::GetWindowThreadProcessId($hwnd, [ref]$pidOut)
    if ($pidOut -gt 0) { return [int]$pidOut }
  } catch {}
  return 0
}

function Test-CadAlive($cadApp) {
  if (-not $cadApp) { return $false }
  try {
    [void]$cadApp.Version
    return $true
  } catch { return $false }
}

# --- Тело конвертации: та же логика, что в render_dwg_smart.ps1, ---
# --- но сессия переиспользуется: Quit здесь НЕТ, документ закрывается. ---
function Convert-DwgJob($job, $cadApp, $usedProgId) {
  $InputPath = [string]$job.inputPath
  $OutputPath = [string]$job.outputPath
  $FallbackCachePath = [string]$job.fallbackCachePath
  $PythonExe = [string]$job.pythonExe

  if (-not (Test-Path -LiteralPath $InputPath -PathType Leaf)) {
    throw "DWG-файл не найден: $InputPath"
  }
  if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = [System.IO.Path]::ChangeExtension($InputPath, ".pdf")
  }
  $outputDir = Split-Path -Parent $OutputPath
  if (-not [string]::IsNullOrWhiteSpace($outputDir) -and -not (Test-Path -LiteralPath $outputDir)) {
    try {
      New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
    } catch {}
  }

  $document = $null
  $sessionGuid = [System.Guid]::NewGuid().ToString("N").Substring(0, 10)
  $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) "FEng_dwg_render_$sessionGuid"
  New-Item -ItemType Directory -Force -Path $tempDir | Out-Null

  try {
    $document = $cadApp.Documents.Open($InputPath, $true)
    $document.SetVariable("BACKGROUNDPLOT", 0)
    try { $document.SetVariable("EXPERT", 5) } catch {}

    $candidateLayouts = @($document.Layouts | Where-Object { -not $_.ModelType } | Sort-Object TabOrder)
    $nonEmptyLayouts = @($candidateLayouts | Where-Object { $_.Block.Count -gt 1 })

    $pagePdfPaths = [System.Collections.Generic.List[string]]::new()

    if ($nonEmptyLayouts.Count -gt 0) {
      foreach ($layout in $nonEmptyLayouts) {
        $document.ActiveLayout = $layout
        $layout.RefreshPlotDeviceInfo()

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

        $availableMedia = @($layout.GetCanonicalMediaNames())
        if ($availableMedia.Count -gt 0) {
          $curMedia = $layout.CanonicalMediaName
          if (-not ($curMedia -and ($availableMedia -contains $curMedia))) {
            $matched = $null
            foreach ($code in @("A0", "A1", "A2", "A3", "A4")) {
              if ($curMedia -match $code) {
                $matched = @($availableMedia | Where-Object { $_ -match $code } | Select-Object -First 1)
                if ($matched.Count -gt 0) { break }
              }
            }
            if ($matched -and $matched.Count -gt 0) {
              $layout.CanonicalMediaName = $matched[0]
            } elseif ($availableMedia -contains "ISO_full_bleed_A3_(420.00_x_297.00_MM)") {
              $layout.CanonicalMediaName = "ISO_full_bleed_A3_(420.00_x_297.00_MM)"
            } else {
              $layout.CanonicalMediaName = $availableMedia[0]
            }
          }
        }

        $layout.PlotType = 4 # acLayout
        $layout.PlotWithLineweights = $true
        $layout.PlotWithPlotStyles = $true

        $pageFile = Join-Path $tempDir ("page_{0:D4}.pdf" -f $layout.TabOrder)
        if ($document.Plot.PlotToFile($pageFile)) {
          if ((Test-Path -LiteralPath $pageFile) -and (Get-Item -LiteralPath $pageFile).Length -gt 1024) {
            $pagePdfPaths.Add($pageFile)
          }
        }
      }
    }

    if ($pagePdfPaths.Count -eq 0) {
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
      $allMedia = @($layout.GetCanonicalMediaNames())
      $a0 = @($allMedia | Where-Object { $_ -match "A0" } | Select-Object -First 1)
      if ($a0.Count -gt 0) {
        $layout.CanonicalMediaName = $a0[0]
      } elseif ($allMedia.Count -gt 0) {
        $layout.CanonicalMediaName = $allMedia[0]
      }

      $layout.PlotType = 1 # acExtents
      $layout.CenterPlot = $true
      $layout.UseStandardScale = $true
      $layout.StandardScale = 0 # acScaleToFit
      $layout.PlotWithLineweights = $false
      $layout.PlotWithPlotStyles = $true

      $modelPageFile = Join-Path $tempDir "page_model.pdf"
      if ($document.Plot.PlotToFile($modelPageFile)) {
        if ((Test-Path -LiteralPath $modelPageFile) -and (Get-Item -LiteralPath $modelPageFile).Length -gt 1024) {
          $pagePdfPaths.Add($modelPageFile)
        }
      }
    }

    if ($pagePdfPaths.Count -eq 0) {
      throw "CAD не смог сгенерировать ни одного листа PDF для чертежа."
    }

    $tempCombinedPdf = Join-Path $tempDir "combined.pdf"
    if ($pagePdfPaths.Count -eq 1) {
      Copy-Item -LiteralPath $pagePdfPaths[0] -Destination $tempCombinedPdf -Force
    } else {
      $mergeScript = @'
import sys
try:
    import fitz
    doc = fitz.open()
    for p in sys.argv[2:]:
        with fitz.open(p) as page_doc:
            doc.insert_pdf(page_doc)
    doc.save(sys.argv[1])
    doc.close()
except ImportError:
    import pypdf
    writer = pypdf.PdfWriter()
    for p in sys.argv[2:]:
        reader = pypdf.PdfReader(p)
        for page in reader.pages:
            writer.add_page(page)
    with open(sys.argv[1], "wb") as f:
        writer.write(f)
'@
      $mergePyFile = Join-Path $tempDir "merge.py"
      [System.IO.File]::WriteAllText($mergePyFile, $mergeScript, [System.Text.Encoding]::UTF8)

      $psi = [System.Diagnostics.ProcessStartInfo]::new()
      $psi.FileName = $PythonExe
      $psi.Arguments = "`"$mergePyFile`" `"$tempCombinedPdf`" " + (($pagePdfPaths | ForEach-Object { "`"$_`"" }) -join " ")
      $psi.UseShellExecute = $false
      $psi.CreateNoWindow = $true
      $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
      $psi.RedirectStandardError = $true
      $pyProc = [System.Diagnostics.Process]::Start($psi)
      $pyProc.WaitForExit(120000)
      if ($pyProc.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $tempCombinedPdf)) {
        $err = $pyProc.StandardError.ReadToEnd()
        throw "Ошибка объединения страниц PDF: $err"
      }
    }

    $finalDestination = $OutputPath
    $writeSuccess = $false
    try {
      Copy-Item -LiteralPath $tempCombinedPdf -Destination $OutputPath -Force -ErrorAction Stop
      $writeSuccess = $true
    } catch {
      if (-not [string]::IsNullOrWhiteSpace($FallbackCachePath)) {
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $FallbackCachePath) | Out-Null
        Copy-Item -LiteralPath $tempCombinedPdf -Destination $FallbackCachePath -Force
        $finalDestination = $FallbackCachePath
        $writeSuccess = $true
      } else {
        throw "Не удалось сохранить PDF в целевую папку '$OutputPath': $_"
      }
    }

    return @{
      ok = $true
      finalPath = $finalDestination
      isLocalFolder = ($finalDestination -eq $OutputPath)
      pageCount = $pagePdfPaths.Count
      progId = $usedProgId
    }
  } finally {
    # Сессию НЕ гасим (переиспользуем), документ закрываем, мусор чистим.
    if ($document) {
      try { $document.Close($false) } catch {}
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    if (Test-Path -LiteralPath $tempDir) {
      Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
  }
}

# --- Старт: поднимаем ОДНУ сессию ---

function Remove-StrayCadSessions($ProtectedPids, $Since) {
  # AutoCAD при пакетной печати из автоматизации порождает отдельные
  # сессии acad.exe /Automation, которые иногда не выходят сами (~1 ГБ каждая).
  # Добиваем только их: свои (ProtectedPids) и чужие интерактивные не трогаем.
  try {
    $procs = Get-CimInstance Win32_Process -Filter "Name='acad.exe'" -ErrorAction SilentlyContinue
    foreach ($p in $procs) {
      if ($ProtectedPids -contains $p.ProcessId) { continue }
      if ([string]$p.CommandLine -notlike "*/Automation*") { continue }
      try { $started = [System.Management.ManagementDateTimeConverter]::ToDateTime($p.CreationDate) }
      catch { continue }
      if ($started -lt $Since) { continue }
      try { Stop-Process -Id $p.ProcessId -Force -ErrorAction SilentlyContinue } catch {}
      Write-DaemonLog ("reaped stray acad pid={0}" -f $p.ProcessId)
    }
  } catch {}
}

$acadBefore = Get-AcadProcessIds
$connected = Connect-CadApp
if (-not $connected) {
  $fatal = @{ error = "Не удалось подключиться к AutoCAD через COM." }
  try {
    $fatal | ConvertTo-Json -Compress | Out-File -LiteralPath (Join-Path $JobDir "daemon.fatal") -Encoding utf8 -Force
  } catch {}
  Write-DaemonLog "FATAL: CAD connect failed"
  exit 1
}
$script:cadApp = $connected.App
$script:usedProgId = $connected.ProgId
try { $script:cadApp.Visible = $false } catch {}

$ourAcadPid = Get-AppPid $script:cadApp
if (-not $ourAcadPid) {
  $after = Get-AcadProcessIds
  $ourAcadPid = @($after | Where-Object { $acadBefore -notcontains $_ } | Select-Object -First 1)[0]
}
if (-not $ourAcadPid) { $ourAcadPid = 0 }

$ready = @{
  pid = $PID
  acadPid = [int]$ourAcadPid
  progId = $script:usedProgId
  startedAt = (Get-Date).ToString("o")
}
try {
  if ($ReadyFile) {
    $ready | ConvertTo-Json -Compress | Out-File -LiteralPath $ReadyFile -Encoding utf8 -Force
  }
} catch {}
Write-DaemonLog ("ready progId={0} acadPid={1}" -f $script:usedProgId, $ourAcadPid)
$script:lastReapTime = Get-Date

# --- Главный цикл ---
$idleSeconds = 0
$idleLimit = [Math]::Max(60, $IdleMinutes * 60)
try {
  while ($true) {
    $jobFile = Get-ChildItem -LiteralPath $JobDir -Filter "*.job.json" -File -ErrorAction SilentlyContinue |
      Sort-Object CreationTime |
      Select-Object -First 1
    if (-not $jobFile) {
      Start-Sleep -Seconds 1
      $idleSeconds += 1
      if (($idleSeconds % 60) -eq 0) {
        Remove-StrayCadSessions -ProtectedPids (@($acadBefore) + @($script:ourAcadPid)) -Since $script:lastReapTime
        $script:lastReapTime = Get-Date
      }
      if ($idleSeconds -ge $idleLimit) {
        Write-DaemonLog "idle timeout, exiting"
        break
      }
      continue
    }
    $idleSeconds = 0
    $jobStartTime = Get-Date
    $jobId = [System.IO.Path]::GetFileNameWithoutExtension([System.IO.Path]::GetFileNameWithoutExtension($jobFile.Name))
    $doneFile = Join-Path $JobDir ($jobId + ".done.json")
    try {
      $job = Get-Content -LiteralPath $jobFile.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
      if (-not (Test-CadAlive $script:cadApp)) {
        Write-DaemonLog "CAD session lost, reconnecting"
        $reconnected = Connect-CadApp
        if (-not $reconnected) { throw "CAD-сессия потеряна, переподключение не удалось." }
        $script:cadApp = $reconnected.App
        $script:usedProgId = $reconnected.ProgId
        try { $script:cadApp.Visible = $false } catch {}
        try {
          $rePid = Get-AppPid $script:cadApp
          if ($rePid) { $script:ourAcadPid = $rePid }
        } catch {}
      }
      Remove-StrayCadSessions -ProtectedPids (@($acadBefore) + @($script:ourAcadPid)) -Since $script:lastReapTime
      $script:lastReapTime = Get-Date
      $result = Convert-DwgJob $job $script:cadApp $script:usedProgId
      @{ ok = $true; result = $result } | ConvertTo-Json -Compress -Depth 6 |
        Out-File -LiteralPath $doneFile -Encoding utf8 -Force
      Write-DaemonLog ("job {0} ok pages={1}" -f $jobId, $result.pageCount)
    } catch {
      $errText = $_.Exception.Message
      if (-not $errText) { $errText = "$_" }
      @{ ok = $false; error = [string]$errText } | ConvertTo-Json -Compress |
        Out-File -LiteralPath $doneFile -Encoding utf8 -Force
      Write-DaemonLog ("job {0} FAIL: {1}" -f $jobId, $errText)
    } finally {
      try { Remove-Item -LiteralPath $jobFile.FullName -Force -ErrorAction SilentlyContinue } catch {}
    }
    Remove-StrayCadSessions -ProtectedPids (@($acadBefore) + @($script:ourAcadPid)) -Since $script:lastReapTime
    Start-Sleep -Seconds 5
    Remove-StrayCadSessions -ProtectedPids (@($acadBefore) + @($script:ourAcadPid)) -Since $script:lastReapTime
    $script:lastReapTime = Get-Date
  }
} finally {
  try { [LauncherMessageFilter]::Revoke() } catch {}
  if ($script:cadApp) {
    try { $script:cadApp.Quit() } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:cadApp) | Out-Null } catch {}
  }
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()

  if ($script:ourAcadPid -and $script:ourAcadPid -gt 0) {
    $deadline = (Get-Date).AddSeconds(3)
    while ((Get-Date) -lt $deadline) {
      $p = Get-Process -Id $script:ourAcadPid -ErrorAction SilentlyContinue
      if (-not $p -or $p.HasExited) { break }
      Start-Sleep -Milliseconds 200
    }
    try {
      $p = Get-Process -Id $script:ourAcadPid -ErrorAction SilentlyContinue
      if ($p -and -not $p.HasExited) {
        Stop-Process -Id $script:ourAcadPid -Force -ErrorAction SilentlyContinue
      }
    } catch {}
  }
}
Write-DaemonLog "stopped"
