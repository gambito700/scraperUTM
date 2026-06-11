$taskName = "ScraperIndicadoresPrevired"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptPath = Join-Path $scriptDir "run_scraper.ps1"

$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
$trigger = New-ScheduledTaskTrigger -Daily -At 08:00
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -RunLevel Highest -Force

Write-Host "Tarea programada '$taskName' creada exitosamente." -ForegroundColor Green
Write-Host "Se ejecutar� diariamente a las 08:00." -ForegroundColor Yellow
Write-Host "Para ejecutar manualmente: Start-ScheduledTask -TaskName '$taskName'" -ForegroundColor Cyan
Write-Host "Para ver logs: Get-Content '$(Join-Path $scriptDir scraper_log.txt)'" -ForegroundColor Cyan
