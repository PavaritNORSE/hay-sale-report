#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Remove the HAY Sale Report scheduled task.

.USAGE
    Open PowerShell AS ADMINISTRATOR, then run:
        .\unregister_scheduled_task.ps1
#>

$TaskName = "Sale Report - Daily 6 AM"

$existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if (-not $existing) {
    Write-Host "[INFO] Task '$TaskName' not found -- nothing to remove." -ForegroundColor Yellow
    exit 0
}

Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
Write-Host "[OK] Task removed: $TaskName" -ForegroundColor Green
Write-Host ""
Write-Host "Manual runs (run_fetch_odoo.bat) still work as before."
