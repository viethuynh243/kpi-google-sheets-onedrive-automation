$ErrorActionPreference = "Stop"

$TaskName = "KPI GoogleSheet To OneDrive Automation"
$WorkDir = "C:\Users\admin\OneDrive\Documents\Excel"
$Runner = Join-Path $WorkDir "run_kpi_automation.ps1"

if (-not (Test-Path $Runner)) {
    throw "Runner script not found: $Runner"
}

$Action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$Runner`"" `
    -WorkingDirectory $WorkDir

# Default: run on day 3 of every month at 09:00 local time.
$Trigger = New-ScheduledTaskTrigger `
    -Monthly `
    -DaysOfMonth 3 `
    -At 9:00AM

$Settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Hours 2)

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $Action `
    -Trigger $Trigger `
    -Settings $Settings `
    -Description "Download Google Sheet KPI files, normalize data, and write staging outputs into the OneDrive sync folder." `
    -Force | Out-Null

Write-Host "Scheduled task created/updated: $TaskName"
Write-Host "Schedule: day 3 of every month at 09:00"
Write-Host "Runner: $Runner"
