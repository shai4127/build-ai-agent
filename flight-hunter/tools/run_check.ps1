# Scheduled entry point: runs one price check headlessly with Claude Code.
# Registered in Windows Task Scheduler by tools\register_tasks.ps1.
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root
$LogDir = Join-Path $Root 'logs'
New-Item -ItemType Directory -Force $LogDir | Out-Null
$Log = Join-Path $LogDir ("run_" + (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd_HH-mm') + "UTC.log")
$Lock = Join-Path $Root '.run.lock'

if ((Get-Date).ToUniversalTime() -gt [datetime]'2027-08-15T23:59:59') {
    "Tracking period ended (15/08/2027). Nothing to do." | Out-File $Log -Encoding utf8
    exit 0
}
if ((Test-Path $Lock) -and ((Get-Item $Lock).LastWriteTime -gt (Get-Date).AddMinutes(-45))) {
    "Skipped: a previous run is still active." | Out-File $Log -Encoding utf8
    exit 0
}
New-Item -ItemType File -Force $Lock | Out-Null
try {
    $Prompt = 'Scheduled run. Perform one price check exactly as defined in automations/check-prices.md, then stop.'
    & claude -p $Prompt `
        --allowedTools "Read" "Edit(RUN_LOG.md)" "WebSearch" "WebFetch" `
            "Bash(python tools/fetch_page.py:*)" "Bash(python tools/flights_log.py:*)" `
            "Bash(python tools/notify_whatsapp.py:*)" *>&1 | Out-File $Log -Encoding utf8
}
finally {
    Remove-Item $Lock -Force -ErrorAction SilentlyContinue
}
