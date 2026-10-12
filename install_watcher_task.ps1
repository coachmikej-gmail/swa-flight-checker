<#
.SYNOPSIS
  Installs (or removes) the Windows scheduled task that runs swa_watcher.py every
  hour from 6:00 to 23:00 and wakes the PC from sleep to do it.

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File .\install_watcher_task.ps1
  powershell -ExecutionPolicy Bypass -File .\install_watcher_task.ps1 -Uninstall

.NOTES
  - Runs as you, only while you're logged on (sleeping/locked is fine; logged OFF is not).
  - Needs "Allow wake timers" enabled in your power plan for the PC to wake from sleep
    (Control Panel > Power Options > Change plan settings > Advanced > Sleep > Allow wake timers).
  - Doesn't change any power or system settings itself.
#>
[CmdletBinding()]
param(
    [switch]$Uninstall,
    [string]$TaskName = "SWA Flight Watcher",
    [int]$StartHour = 6,
    [int]$EndHour = 23,
    # Extra arguments for swa_watcher.py, e.g. "--dry-run" when testing the task itself.
    [string]$ExtraArgs = ""
)

$ErrorActionPreference = "Stop"

if ($Uninstall) {
    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        Write-Host "Removed scheduled task '$TaskName'."
    } else {
        Write-Host "No scheduled task named '$TaskName' found."
    }
    return
}

if ($StartHour -lt 0 -or $EndHour -gt 23 -or $StartHour -gt $EndHour) { throw "Hours must satisfy 0 <= StartHour <= EndHour <= 23." }

$projectDir = $PSScriptRoot
$script = Join-Path $projectDir "swa_watcher.py"
if (-not (Test-Path $script)) { throw "Can't find $script" }

# Use the real interpreter (not the Microsoft Store alias, which doesn't work from Task Scheduler),
# and pythonw.exe so no console window flashes up.
$python = (& python -c "import sys; print(sys.executable)" 2>$null | Select-Object -First 1)
if (-not $python -or -not (Test-Path $python)) { throw "Couldn't find Python. Is it installed and on PATH?" }
$pythonw = Join-Path (Split-Path $python) "pythonw.exe"
if (-not (Test-Path $pythonw)) { $pythonw = $python }

$arguments = "`"$script`" --only-between $StartHour-$EndHour"
if ($ExtraArgs) { $arguments += " $ExtraArgs" }
$action = New-ScheduledTaskAction -Execute $pythonw -Argument $arguments -WorkingDirectory $projectDir

# One plain daily trigger per hour (06:00, 07:00, ... 23:00) rather than a single trigger
# that "repeats every hour". With the repeating form, `powercfg /waketimers` showed Windows
# arming a wake timer only for the next day's start time, not for each hourly repeat, so the
# PC never woke for the hourly runs. Separate triggers each get their own wake timer.
$triggers = $StartHour..$EndHour | ForEach-Object {
    New-ScheduledTaskTrigger -Daily -At ([datetime]::Today.AddHours($_))
}

$settings = New-ScheduledTaskSettingsSet `
    -WakeToRun `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable `
    -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 15) `
    -MultipleInstances IgnoreNew

$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Limited

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $triggers -Settings $settings `
    -Principal $principal -Force `
    -Description "Checks SWA Flight Checker price watches hourly ($StartHour:00-$EndHour:00) and sends Telegram alerts. Wakes the PC from sleep." | Out-Null

$runs = $EndHour - $StartHour + 1
Write-Host "Installed scheduled task '$TaskName'."
Write-Host "  Runs hourly at $($StartHour):00 ... $($EndHour):00 ($runs times a day), waking the PC from sleep."
Write-Host "  Command: $pythonw $arguments"
Write-Host "  Log:     $(Join-Path $projectDir 'watcher.log')"
Write-Host "  Next run: $((Get-ScheduledTaskInfo -TaskName $TaskName).NextRunTime)"
Write-Host "Remove it any time with: powershell -ExecutionPolicy Bypass -File .\install_watcher_task.ps1 -Uninstall"
