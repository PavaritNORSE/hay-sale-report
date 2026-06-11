#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Register Windows Scheduled Task for HAY Sale Report -- Daily 6 AM (Bangkok)

.DESCRIPTION
    Creates a scheduled task that runs run_scheduled.bat every day at 06:00
    local time. The bat will run daily_report.py --fetch-odoo (no --date
    arg -> defaults to YESTERDAY per v3.9). Stdout/stderr captured to
    _Archive\scheduled_logs\<timestamp>.log per run.

    Settings:
      - Trigger: Daily at 06:00
      - Wake the computer to run this task (laptop closed-lid scenario)
      - Start when available (catches missed runs if computer was unreachable)
      - Allow on battery, don't stop on AC/battery transition
      - Run only when user is logged on (Excel COM requirement)
      - Max execution: 1 hour
      - Restart on failure: no (failure email is sent via Graph API instead)

.USAGE
    Open PowerShell AS ADMINISTRATOR, then run:
        cd "C:\Users\USER\Desktop\Claude Cowork-Workspace\02-Projects\01-Sale Report"
        .\register_scheduled_task.ps1

    To remove the task later:
        .\unregister_scheduled_task.ps1
#>

$ErrorActionPreference = 'Stop'

$TaskName  = "Sale Report - Daily 6 AM"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$BatFile   = Join-Path $ScriptDir "run_scheduled.bat"

if (-not (Test-Path $BatFile)) {
    Write-Error "[FATAL] run_scheduled.bat not found at: $BatFile"
    exit 1
}

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Registering Scheduled Task" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Task name:     $TaskName"
Write-Host "  Script:        $BatFile"
Write-Host "  User:          $env:USERDOMAIN\$env:USERNAME"
Write-Host "  Schedule:      Daily at 06:00 (local time)"
Write-Host ""

# Action -- run the bat with the project directory as working dir
$Action = New-ScheduledTaskAction `
    -Execute $BatFile `
    -WorkingDirectory $ScriptDir

# Trigger -- daily at 06:00 local time
$Trigger = New-ScheduledTaskTrigger -Daily -At 06:00

# Settings -- important: WakeToRun + StartWhenAvailable for closed-lid laptops
$Settings = New-ScheduledTaskSettingsSet `
    -WakeToRun `
    -StartWhenAvailable `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Hours 1) `
    -MultipleInstances IgnoreNew

# Principal -- Interactive (user must be logged on; required for Excel COM)
# RunLevel: Limited (not Highest) -- Excel COM works under normal user privileges
$Principal = New-ScheduledTaskPrincipal `
    -UserId "$env:USERDOMAIN\$env:USERNAME" `
    -LogonType Interactive `
    -RunLevel Limited

# Remove any existing task with the same name (idempotent re-registration)
$existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existing) {
    Write-Host "[INFO] Existing task found -- unregistering first..." -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}

# Register
$task = Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $Action `
    -Trigger $Trigger `
    -Settings $Settings `
    -Principal $Principal `
    -Description "Auto-runs HAY Sale Report daily at 6 AM for yesterday's sales data. Manual runs: double-click run_fetch_odoo.bat instead. Logs: _Archive\scheduled_logs\."

Write-Host "[OK] Task registered" -ForegroundColor Green

# Show next run time
$info = Get-ScheduledTask -TaskName $TaskName | Get-ScheduledTaskInfo
Write-Host ""
Write-Host "  Next scheduled run: $($info.NextRunTime)"
Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Next steps" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  1. Test the task manually (recommended):"
Write-Host "       Open Task Scheduler -> Task Scheduler Library"
Write-Host "       Right-click 'Sale Report - Daily 6 AM' -> Run"
Write-Host ""
Write-Host "  2. Check the log file under:"
Write-Host "       $ScriptDir\_Archive\scheduled_logs\"
Write-Host ""
Write-Host "  3. To remove the task later:"
Write-Host "       .\unregister_scheduled_task.ps1  (run as admin)"
Write-Host ""
Write-Host "  4. Manual runs (with date picker):"
Write-Host "       Double-click run_fetch_odoo.bat as usual."
Write-Host ""
