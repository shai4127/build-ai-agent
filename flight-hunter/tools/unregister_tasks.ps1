# Removes all Flight Hunter scheduled tasks.
Get-ScheduledTask -TaskPath '\FlightHunter\' -ErrorAction SilentlyContinue |
    Unregister-ScheduledTask -Confirm:$false
"Flight Hunter scheduled tasks removed."
