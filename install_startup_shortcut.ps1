$ErrorActionPreference = "Stop"

$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Launcher = Join-Path $ProjectDir "watchdog_dashboard.vbs"
$StartupDir = [Environment]::GetFolderPath("Startup")
$ShortcutPath = Join-Path $StartupDir "MeetingDashboard.lnk"

$Shell = New-Object -ComObject WScript.Shell
$Shortcut = $Shell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = $Launcher
$Shortcut.WorkingDirectory = $ProjectDir
$Shortcut.WindowStyle = 7
$Shortcut.Description = "Keep Meeting Dashboard local service running"
$Shortcut.Save()

Write-Output $ShortcutPath
