# Updates an existing Flight Hunter folder from this (freshly downloaded) copy without touching your data:
# flights.xlsx, backups, .env, the WhatsApp session, logs, RUN_LOG.md, MEMORY.md and INTEGRATIONS.md stay as they are.
# Usage: powershell -ExecutionPolicy Bypass -File <download>\flight-hunter\tools\update.ps1 -Target "<Desktop>\flight-hunter"
param([Parameter(Mandatory = $true)][string]$Target)
$Source = Split-Path -Parent $PSScriptRoot
if (-not (Test-Path (Join-Path $Target 'AGENTS.md'))) { "Not a Flight Hunter folder: $Target - nothing was changed."; exit 1 }

robocopy $Source $Target /E /NFL /NDL /NJH /NJS `
    /XF flights.xlsx 'flights_backup_*.xlsx' .env RUN_LOG.md MEMORY.md INTEGRATIONS.md pending_rows.jsonl .last_daily_summary .run.lock `
    /XD .whatsapp-session logs .git | Out-Null

Set-Location $Target
python -m pip install --quiet -r requirements.txt
python tools/flights_log.py migrate
"Files updated. Next: open this folder in Claude Code and type the start phrase (see START_HERE.md)."
