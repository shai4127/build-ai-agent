# One-time local setup: Python packages, headless Chromium, .env and the Excel file.
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root
python -m pip install --upgrade -r requirements.txt
python -m playwright install chromium
if (-not (Test-Path '.env')) { Copy-Item '.env.example' '.env'; "Created .env - fill in WHATSAPP_PHONE, then run: python tools/notify_whatsapp.py --login" }
python tools/flights_log.py init
