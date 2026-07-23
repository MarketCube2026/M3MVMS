$ErrorActionPreference = "Stop"

$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Python = "C:\Users\admin\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\pythonw.exe"
$Server = Join-Path $ProjectDir "server.py"

$Existing = Get-NetTCPConnection -LocalPort 8765 -State Listen -ErrorAction SilentlyContinue
if ($Existing) {
  exit 0
}

$ProcessInfo = New-Object System.Diagnostics.ProcessStartInfo
$ProcessInfo.FileName = $Python
$ProcessInfo.Arguments = "`"$Server`""
$ProcessInfo.WorkingDirectory = $ProjectDir
$ProcessInfo.UseShellExecute = $false
$ProcessInfo.CreateNoWindow = $true
[System.Diagnostics.Process]::Start($ProcessInfo) | Out-Null
