# FIX 1: Register Windows Task Scheduler job that wakes the PC at 5:00 AM daily
# and runs the automation. Run this script ONCE as Administrator.
#
# Usage (run as Administrator):
#   powershell -ExecutionPolicy Bypass -File "C:\laragon\www\automation-project - Main\register_task.ps1"

$TaskName    = "AutomationProjectDailyRun"
$ScriptDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
$BatFile     = Join-Path $ScriptDir "run_scheduled.bat"
$RunAt       = "05:00"

# Remove existing task with the same name if it exists
Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue

$Action  = New-ScheduledTaskAction -Execute "cmd.exe" -Argument "/c `"$BatFile`"" -WorkingDirectory $ScriptDir
$Trigger = New-ScheduledTaskTrigger -Daily -At $RunAt

$Settings = New-ScheduledTaskSettingsSet `
    -WakeToRun `
    -ExecutionTimeLimit (New-TimeSpan -Hours 4) `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable:$false `
    -MultipleInstances IgnoreNew

# Run as the current logged-in user (keeps desktop session, needed for Edge/COM)
$Principal = New-ScheduledTaskPrincipal `
    -UserId ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) `
    -LogonType Interactive `
    -RunLevel Highest

Register-ScheduledTask `
    -TaskName  $TaskName `
    -Action    $Action `
    -Trigger   $Trigger `
    -Settings  $Settings `
    -Principal $Principal `
    -Force

Write-Host "[OK] Task '$TaskName' registered — runs daily at $RunAt with WakeToRun enabled."
Write-Host "[NOTE] Ensure the PC is plugged in and sleep/hibernate is allowed in Power Options."
Write-Host "[NOTE] To allow wake timers: Control Panel -> Power Options -> Change plan settings -> Change advanced -> Sleep -> Allow wake timers -> Enable."
