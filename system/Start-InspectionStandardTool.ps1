#requires -version 3
<#
  Start-InspectionStandardTool.ps1 - the one way the tool is started.

  Runs inside a HEADLESS console host (conhost.exe --headless), started by
  "Launch Inspection Standard Tool.cmd" or by the desktop shortcut, so no
  PowerShell window appears (the same mechanism as the WVN Stage Gate
  Command Center). On Windows 11 the default terminal is Windows Terminal,
  which takes the console over before -WindowStyle Hidden can hide it, so
  the headless host is the only reliable way to get no window at all.

  First run on a PC (from the .cmd in the synced SharePoint/OneDrive folder):
    1. clears the "downloaded from the internet" mark on the tool folder
    2. creates the desktop shortcut "Inspection Standard Tool" for THIS PC,
       pointing at wherever OneDrive synced the folder here, with the tool's
       own icon
    3. runs inspection-standard-launcher.ps1 in this same hidden process
       (it prepares the BOM data and opens the browser)
  From the desktop shortcut (-FromShortcut) only step 3 runs.

  Nothing is ever shown in a console. While the data is being rebuilt a
  small progress window (progress-window.ps1) shows the launcher log, and
  problems come up as an ordinary message box; details are in system\logs.

  Extra arguments are passed straight through to the launcher:
    --update        rebuild the BOM data without asking
    --skip-update   open straight away with the data from the last run
    --force-rebuild / --clear-cache / --prepare-only  (as before)
#>
[CmdletBinding()]
param(
  # set by the desktop shortcut: skip the first-run work
  [switch]$FromShortcut,
  # tests only: put the shortcut here instead of the Desktop, bake extra
  # arguments into it, and never open a message box
  [string]$ShortcutDir = '',
  [string]$ShortcutExtraArgs = '',
  [switch]$Quiet,
  # everything else goes to inspection-standard-launcher.ps1
  [Parameter(ValueFromRemainingArguments = $true)]
  [string[]]$LauncherArgs = @()
)

$ErrorActionPreference = 'Stop'
$SystemDir = $PSScriptRoot
if (-not $SystemDir) { $SystemDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
$RootDir = Split-Path -Parent $SystemDir
$AppTitle = 'Inspection Standard Tool'
$ShortcutFile = $AppTitle + '.lnk'
$IconFile = Join-Path $SystemDir 'assets\inspection_standard_tool.ico'
$LogDir = Join-Path $SystemDir 'logs'

function Write-StartLog { param([string]$Message)
  try {
    if (-not (Test-Path -LiteralPath $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }
    $who = ([string]$env:COMPUTERNAME + '_' + [string]$env:USERNAME) -replace '[^A-Za-z0-9_-]', '_'
    Add-Content -LiteralPath (Join-Path $LogDir ('start_' + $who + '.log')) -Value ('[' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '] ' + $Message) -Encoding UTF8
  } catch { }
}
function Show-Message { param([string]$Text, [string]$Icon = 'Error')
  Write-StartLog ('message: ' + $Text)
  if ($Quiet) { return }
  try {
    Add-Type -AssemblyName System.Windows.Forms
    [void][System.Windows.Forms.MessageBox]::Show($Text, $AppTitle, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::$Icon)
  } catch { }
}

# A path inside the user's profile is written relative to %USERPROFILE%, so a
# shortcut that OneDrive syncs to the person's other PCs (Desktop folder
# backup) still works there even when the Windows user name differs.
function Get-ProfileRelative { param([string]$Path)
  $prof = ([Environment]::GetFolderPath('UserProfile')).TrimEnd('\')
  if ($prof -and $Path.StartsWith($prof + '\', [StringComparison]::OrdinalIgnoreCase)) { return $Path.Substring($prof.Length + 1) }
  return $null
}

# What the desktop shortcut runs. conhost.exe --headless gives PowerShell a
# console that has no window (Windows 10 1809 and later); without it the
# shortcut falls back to a hidden, minimised PowerShell.
function Get-ShortcutSpec {
  $ps = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
  $conhost = Join-Path $env:SystemRoot 'System32\conhost.exe'
  $me = Join-Path $SystemDir 'Start-InspectionStandardTool.ps1'
  $extra = if ($ShortcutExtraArgs) { ' ' + $ShortcutExtraArgs } else { '' }
  $rel = Get-ProfileRelative $me
  if ($rel) {
    $cmd = "& (Join-Path `$env:USERPROFILE '" + $rel.Replace("'", "''") + "') -FromShortcut" + $extra
    $psArgs = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command "' + $cmd + '"'
  } else {
    $psArgs = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + $me + '" -FromShortcut' + $extra
  }
  $iconRel = Get-ProfileRelative $IconFile
  $spec = @{
    target   = $ps
    args     = $psArgs
    icon     = if ($iconRel) { '%USERPROFILE%\' + $iconRel + ',0' } else { $IconFile + ',0' }
    workdir  = '%USERPROFILE%'
    style    = 7   # minimised: only matters for the fallback without conhost
  }
  if (Test-Path -LiteralPath $conhost) { $spec.target = $conhost; $spec.args = '--headless "' + $ps + '" ' + $psArgs }
  return $spec
}

# Creates the desktop shortcut, or repairs it when it no longer matches this
# PC (the folder moved, a newer version changed how it starts, or it came
# from another PC through OneDrive with a path that does not exist here).
# An identical shortcut is left untouched, so two PCs syncing the same
# Desktop never rewrite each other's file.
function Install-DesktopShortcut {
  $desk = if ($ShortcutDir) { $ShortcutDir } else { [Environment]::GetFolderPath('Desktop') }
  if (-not $desk) { Write-StartLog 'no Desktop folder - shortcut skipped'; return 'skipped' }
  if (-not (Test-Path -LiteralPath $desk)) { New-Item -ItemType Directory -Path $desk -Force | Out-Null }
  $path = Join-Path $desk $ShortcutFile
  $want = Get-ShortcutSpec
  $sh = New-Object -ComObject WScript.Shell
  $state = 'created'
  if (Test-Path -LiteralPath $path) {
    $have = $sh.CreateShortcut($path)
    $iconOk = ($have.IconLocation -eq $want.icon) -or ($have.IconLocation -eq [Environment]::ExpandEnvironmentVariables($want.icon))
    if ($have.TargetPath -eq $want.target -and $have.Arguments -eq $want.args -and $iconOk) { return 'unchanged' }
    $state = 'updated'
  }
  $lnk = $sh.CreateShortcut($path)
  $lnk.TargetPath = $want.target
  $lnk.Arguments = $want.args
  $lnk.WorkingDirectory = $want.workdir
  $lnk.IconLocation = $want.icon
  $lnk.WindowStyle = $want.style
  $lnk.Description = 'WAHL Vietnam - Inspection Standard Tool (' + $RootDir + ')'
  $lnk.Save()
  Write-StartLog ('desktop shortcut ' + $state + ': ' + $path + ' -> ' + $RootDir)
  return $state
}

Write-StartLog ('start (' + $(if ($FromShortcut) { 'desktop shortcut' } else { 'Launch Inspection Standard Tool.cmd' }) + ') in ' + $RootDir)

if (-not $FromShortcut) {
  try {
    Get-ChildItem -LiteralPath $RootDir -Recurse -File -ErrorAction SilentlyContinue | Unblock-File -ErrorAction SilentlyContinue
  } catch { Write-StartLog ('unblock: ' + $_.Exception.Message) }
  try { [void](Install-DesktopShortcut) }
  catch { Write-StartLog ('desktop shortcut failed: ' + $_.Exception.Message) }   # never a reason not to start
}

$launcher = Join-Path $SystemDir 'inspection-standard-launcher.ps1'
$la = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + $launcher + '"'
foreach ($a in @($LauncherArgs)) { if ($a) { $la += ' "' + ([string]$a).Replace('"', '""') + '"' } }
$psExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
# Tells the launcher there is no console to print progress on, so it opens
# the small progress window while the data is rebuilt.
$env:IST_HEADLESS = '1'
try {
  # The launcher runs as its OWN PowerShell process on this same hidden
  # console (UseShellExecute=false and no new window = it inherits this
  # headless console, so nothing is shown). WaitForExit() waits for the
  # launcher only - not for the browser it opens - so this process ends as
  # soon as the tool is open instead of lingering until Chrome is closed.
  $psi = New-Object System.Diagnostics.ProcessStartInfo
  $psi.FileName = $psExe
  $psi.Arguments = $la
  $psi.UseShellExecute = $false
  $psi.WorkingDirectory = $RootDir
  $child = [System.Diagnostics.Process]::Start($psi)
  $child.WaitForExit()
  Write-StartLog ('launcher exit code ' + $child.ExitCode)
  if ($child.ExitCode -ne 0) {
    Show-Message ('Không mở được Inspection Standard Tool (mã lỗi ' + $child.ExitCode + ').' + "`n" +
                  'The Inspection Standard Tool could not start (exit code ' + $child.ExitCode + ').' + "`n`n" +
                  'Chi tiết / details: ' + (Join-Path $LogDir 'Inspection_Standard_Tool_launcher.log'))
  }
} catch {
  Write-StartLog ('launcher failed: ' + $_.Exception.Message)
  Show-Message ('Không mở được ứng dụng. / The app could not start.' + "`n`n" + $_.Exception.Message + "`n`n" +
                'Chi tiết / details: ' + $LogDir)
}
