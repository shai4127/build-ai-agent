# Registers the three daily price checks (00:00, 12:00, 20:00 UTC) in Windows Task Scheduler.
# Run only after the owner approved activation. Safe to re-run (overwrites the same tasks).
$Root = Split-Path -Parent $PSScriptRoot
$Script = Join-Path $Root 'tools\run_check.ps1'
$Action = New-ScheduledTaskAction -Execute 'powershell.exe' `
    -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$Script`"" `
    -WorkingDirectory $Root
$Settings = New-ScheduledTaskSettingsSet -WakeToRun -StartWhenAvailable -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Minutes 40)
$NowUtc = (Get-Date).ToUniversalTime()

foreach ($Hour in 0, 12, 20) {
    $First = $NowUtc.Date.AddHours($Hour)
    if ($First -le $NowUtc) { $First = $First.AddDays(1) }
    $Trigger = New-ScheduledTaskTrigger -Daily -At (Get-Date)
    # A trailing "Z" makes Task Scheduler fire at this UTC time regardless of Israel summer/winter time.
    $Trigger.StartBoundary = $First.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'")
    $Trigger.EndBoundary = '2027-08-16T00:00:00Z'
    $Name = 'FlightHunter-{0:D2}00UTC' -f $Hour
    Register-ScheduledTask -TaskPath '\FlightHunter\' -TaskName $Name -Action $Action -Trigger $Trigger `
        -Settings $Settings -Description 'Flight Hunter: TLV-HER price check (15-21/08/2027)' -Force | Out-Null
}

Get-ScheduledTask -TaskPath '\FlightHunter\' | ForEach-Object {
    $Info = $_ | Get-ScheduledTaskInfo
    [pscustomobject]@{ Task = $_.TaskName; State = $_.State; NextRun = $Info.NextRunTime }
} | Format-Table -AutoSize
