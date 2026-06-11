param(
    [string]$TaskName = "PDF_DDD hourly batches",
    [ValidateSet("Daily", "Weekly", "Once")]
    [string]$Schedule = "Daily",
    [string]$At = "02:30",
    [string[]]$DaysOfWeek = @("Monday"),
    [string]$ProjectDir = "",
    [string]$WhCodes = "US02",
    [string]$Statuses = "10,15,20,30",
    [int]$Workers = 5,
    [int]$StartHour = 0,
    [int]$EndHour = 24,
    [string]$Channel = "",
    [switch]$BrowserMode,
    [switch]$Ocr
)

$ErrorActionPreference = "Stop"

if (-not $ProjectDir) {
    $ProjectDir = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
}

$runner = Join-Path $ProjectDir "scripts\run_hourly_batches.ps1"
if (-not (Test-Path -LiteralPath $runner)) {
    throw "Runner script not found: $runner"
}

$taskArgs = @(
    "-NoProfile",
    "-ExecutionPolicy", "Bypass",
    "-File", "`"$runner`"",
    "-ProjectDir", "`"$ProjectDir`"",
    "-WhCodes", "`"$WhCodes`"",
    "-Statuses", "`"$Statuses`"",
    "-Workers", $Workers,
    "-StartHour", $StartHour,
    "-EndHour", $EndHour
)

if ($Channel) {
    $taskArgs += @("-Channel", "`"$Channel`"")
}
if ($BrowserMode) {
    $taskArgs += "-BrowserMode"
}
if ($Ocr) {
    $taskArgs += "-Ocr"
}

$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument ($taskArgs -join " ") -WorkingDirectory $ProjectDir
$runAt = [datetime]::Parse($At)

switch ($Schedule) {
    "Daily" {
        $trigger = New-ScheduledTaskTrigger -Daily -At $runAt
    }
    "Weekly" {
        $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $DaysOfWeek -At $runAt
    }
    "Once" {
        $trigger = New-ScheduledTaskTrigger -Once -At $runAt
    }
}

$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
$task = New-ScheduledTask -Action $action -Trigger $trigger -Principal $principal -Settings $settings

Register-ScheduledTask -TaskName $TaskName -InputObject $task -Force | Out-Null

Write-Host "Scheduled task registered: $TaskName" -ForegroundColor Green
Write-Host "Schedule: $Schedule at $At"
Write-Host "Command : powershell.exe $($taskArgs -join ' ')"
Write-Host ""
Write-Host "Run now:"
Write-Host "  Start-ScheduledTask -TaskName `"$TaskName`""
Write-Host ""
Write-Host "Delete task:"
Write-Host "  Unregister-ScheduledTask -TaskName `"$TaskName`" -Confirm:`$false"
