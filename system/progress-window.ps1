#requires -version 3
<#
  progress-window.ps1 - a small always-on-top window shown while the launcher
  rebuilds the BOM data in a headless (window-less) process. It tails the
  launcher log so the technician can see something is happening, and closes
  by itself when the launcher process (-ParentPid) exits or when the
  "done" marker file appears.
#>
param(
  [string]$LogPath,
  [int]$ParentPid = 0,
  [string]$DoneFile = '',
  [string]$Title = 'Inspection Standard Tool'
)
$ErrorActionPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object System.Windows.Forms.Form
$form.Text = $Title
$form.FormBorderStyle = 'FixedToolWindow'
$form.StartPosition = 'Manual'
$form.TopMost = $true
$form.ShowInTaskbar = $true
$form.ClientSize = New-Object System.Drawing.Size(520, 118)
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$form.Location = New-Object System.Drawing.Point(($wa.Right - $form.Width - 16), ($wa.Bottom - $form.Height - 16))

$heading = New-Object System.Windows.Forms.Label
$heading.Text = 'Đang cập nhật dữ liệu BOM... / Updating BOM data...'
$heading.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$heading.Location = New-Object System.Drawing.Point(14, 10)
$heading.Size = New-Object System.Drawing.Size(492, 22)
$form.Controls.Add($heading)

$bar = New-Object System.Windows.Forms.ProgressBar
$bar.Style = 'Marquee'
$bar.MarqueeAnimationSpeed = 30
$bar.Location = New-Object System.Drawing.Point(14, 38)
$bar.Size = New-Object System.Drawing.Size(492, 16)
$form.Controls.Add($bar)

$status = New-Object System.Windows.Forms.Label
$status.Text = 'Vui lòng chờ, trình duyệt sẽ tự mở khi xong. / Please wait, the browser opens when done.'
$status.ForeColor = [System.Drawing.Color]::DimGray
$status.Location = New-Object System.Drawing.Point(14, 62)
$status.Size = New-Object System.Drawing.Size(492, 48)
$status.AutoEllipsis = $true
$form.Controls.Add($status)

$started = Get-Date
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 700
$timer.Add_Tick({
  $elapsed = [int]((Get-Date) - $started).TotalSeconds
  $last = ''
  if ($LogPath -and (Test-Path -LiteralPath $LogPath)) {
    try {
      $tail = Get-Content -LiteralPath $LogPath -Tail 1 -Encoding UTF8
      if ($tail) { $last = ([string]$tail) -replace '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} ', '' }
    } catch { }
  }
  $status.Text = ('[' + [string]$elapsed + ' s] ' + $last)
  $exit = $false
  if ($DoneFile -and (Test-Path -LiteralPath $DoneFile)) { $exit = $true }
  if ($ParentPid -gt 0) { $p = Get-Process -Id $ParentPid -ErrorAction SilentlyContinue; if (-not $p) { $exit = $true } }
  if ($exit) { $timer.Stop(); $form.Close() }
})
$timer.Start()
[void]$form.ShowDialog()
