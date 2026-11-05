# Registers the hourly price check (every hour on the hour, UTC) in Windows Task Scheduler.
# Run only after the owner approved activation. Safe to re-run: it replaces any earlier
# Flight Hunter tasks (including the old 3-times-a-day tasks).
$Root = Split-Path -Parent $PSScriptRoot
$Script = Join-Path $Root 'tools\run_check.ps1'
$EndUtc = [datetime]'2027-08-16T00:00:00'

Get-ScheduledTask -TaskPath '\FlightHunter\' -ErrorAction SilentlyContinue |
    Unregister-ScheduledTask -Confirm:$false

$Action = New-ScheduledTaskAction -Execute 'powershell.exe' `
    -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$Script`"" `
    -WorkingDirectory $Root
$Settings = New-ScheduledTaskSettingsSet -WakeToRun -StartWhenAvailable -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Minutes 50)

$NowUtc = (Get-Date).ToUniversalTime()
$First = $NowUtc.Date.AddHours($NowUtc.Hour + 1)   # next full hour, UTC
$Trigger = New-ScheduledTaskTrigger -Once -At (Get-Date) `
    -RepetitionInterval (New-TimeSpan -Hours 1) -RepetitionDuration ($EndUtc - $First)
# A trailing "Z" makes Task Scheduler use UTC, so the hours never shift with Israel summer/winter time.
$Trigger.StartBoundary = $First.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'")
$Trigger.EndBoundary = $EndUtc.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'")

Register-ScheduledTask -TaskPath '\FlightHunter\' -TaskName 'FlightHunter-Hourly' -Action $Action `
    -Trigger $Trigger -Settings $Settings `
    -Description 'Flight Hunter: hourly TLV-HER price check (15-21/08/2027) + daily 08:00 summary' -Force | Out-Null

Get-ScheduledTask -TaskPath '\FlightHunter\' | ForEach-Object {
    $Info = $_ | Get-ScheduledTaskInfo
    [pscustomobject]@{ Task = $_.TaskName; State = $_.State; NextRun = $Info.NextRunTime }
} | Format-Table -AutoSize
