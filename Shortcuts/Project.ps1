Clear-Host
Echo "Toggling ScrollLock...end by exiting window or press Ctrl + C"
$WShell = New-Object -com "Wscript.Shell"
while ($true) {
[datetime]$exitTime = $(get-date -Format "dd-MMM-yyyy 17:00:00")
if ( (get-date) -ge $exitTime ) # Shutdown at 5:00 EST
{
  break # Close Connections and shutdown
}
$WShell.sendkeys("{SCROLLLOCK}")
Start-Sleep -Milliseconds 200
$WShell.sendkeys("{SCROLLLOCK}")
Start-Sleep -Seconds 240
}
