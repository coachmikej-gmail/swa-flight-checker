<#
.SYNOPSIS
  Tests whether this PC can wake itself from sleep for a scheduled task, in a few
  minutes instead of waiting for the real watcher's next run.

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File .\test_wake.ps1             # arm a test 5 minutes from now
  powershell -ExecutionPolicy Bypass -File .\test_wake.ps1 -Minutes 3  # ...or a different delay
  powershell -ExecutionPolicy Bypass -File .\test_wake.ps1 -Cleanup    # remove the test task and its log

.NOTES
  The test task only writes one line to wake_test.log. It deliberately does NOT run
  "when available" after a missed start, so if the PC didn't wake by itself the log
  simply won't exist - there's no way to confuse a real wake with a late catch-up.
#>
param([int]$Minutes = 5, [switch]$Cleanup)

$ErrorActionPreference = "Stop"
$name = "SWA Wake Test"
$log = Join-Path $PSScriptRoot "wake_test.log"

if ($Cleanup) {
    Unregister-ScheduledTask -TaskName $name -Confirm:$false -ErrorAction SilentlyContinue
    Remove-Item $log -ErrorAction SilentlyContinue
    Write-Host "Removed the test task and wake_test.log."
    return
}

Remove-Item $log -ErrorAction SilentlyContinue
$fireAt = (Get-Date).AddMinutes($Minutes)

$action = New-ScheduledTaskAction -Execute "cmd.exe" -Argument "/c echo fired at %date% %time% >> `"$log`"" -WorkingDirectory $PSScriptRoot
$trigger = New-ScheduledTaskTrigger -Once -At $fireAt
$settings = New-ScheduledTaskSettingsSet -WakeToRun -StartWhenAvailable:$false `
    -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 1)
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Limited
Register-ScheduledTask -TaskName $name -Action $action -Trigger $trigger -Settings $settings `
    -Principal $principal -Force -Description "One-time test: can this PC wake itself from sleep?" | Out-Null

$when = $fireAt.ToString("h:mm:ss tt")
Write-Host ""
Write-Host "Wake test armed for $when."
Write-Host ""
Write-Host "  1. Put the PC to SLEEP now (Start > Power > Sleep)."
Write-Host "  2. Don't touch the mouse or keyboard until at least $($fireAt.AddMinutes(3).ToString('h:mm tt'))."
Write-Host "     (If you wake it yourself before $when, the test is void - just run this again.)"
Write-Host "  3. Then run:   Get-Content '$log'"
Write-Host ""
Write-Host "  A line like 'fired at ... $($fireAt.ToString('HH:mm'))...' = the PC woke itself. WORKS."
Write-Host "  No log file at all                                       = it did NOT wake. Check the BIOS."
Write-Host "  (The screen stays off during a timer wake and the PC falls back asleep after a couple of"
Write-Host "   minutes, so the log file is the proof, not what you see.)"
Write-Host ""
Write-Host "Clean up afterwards with:  powershell -ExecutionPolicy Bypass -File .\test_wake.ps1 -Cleanup"
