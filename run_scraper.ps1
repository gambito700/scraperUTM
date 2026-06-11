$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$logFile = Join-Path $scriptDir "scraper_log.txt"

$date = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
"[$date] Iniciando scraping..." | Out-File $logFile -Append

try {
    python (Join-Path $scriptDir "scraper.py") 2>&1 | Out-File $logFile -Append
    "[$date] Scraping completado exitosamente" | Out-File $logFile -Append
} catch {
    "[$date] ERROR: $_" | Out-File $logFile -Append
}
