param(
    [string]$TaskName = "PDF_DDD previous hour batch",
    [string]$StartAt = "01:00",
    [string]$EndAt = "22:30",
    [string]$ProjectDir = "",
    [string]$WhCodes = "US02",
    [string]$Statuses = "10,15,20,30",
    [int]$Workers = 5,
    [string]$Channel = "",
    [switch]$BrowserMode,
    [switch]$Ocr,
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"

if (-not $ProjectDir) {
    $ProjectDir = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
}

$runner = Join-Path $ProjectDir "scripts\run_previous_hour_batch.ps1"
if (-not (Test-Path -LiteralPath $runner)) {
    throw "Runner script not found: $runner"
}

$startTime = [datetime]::Parse($StartAt)
$endTime = [datetime]::Parse($EndAt)
if ($endTime -le $startTime) {
    $endTime = $endTime.AddDays(1)
}
$windowMinute = $startTime.Minute

$taskArgs = @(
    "-NoProfile",
    "-ExecutionPolicy", "Bypass",
    "-File", "`"$runner`"",
    "-ProjectDir", "`"$ProjectDir`"",
    "-WhCodes", "`"$WhCodes`"",
    "-Statuses", "`"$Statuses`"",
    "-Workers", $Workers,
    "-WindowMinute", $windowMinute
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
$triggers = @()
$triggerTimes = @()
$cursor = $startTime
while ($cursor -le $endTime) {
    $triggers += New-ScheduledTaskTrigger -Daily -At $cursor
    $triggerTimes += $cursor.ToString("HH:mm")
    $cursor = $cursor.AddHours(1)
}

if ($DryRun) {
    Write-Host "Dry run only. Scheduled task was not registered."
    Write-Host "TaskName: $TaskName"
    Write-Host "Daily triggers: $($triggerTimes -join ', ')"
    Write-Host "Command: powershell.exe $($taskArgs -join ' ')"
    exit 0
}

$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
$task = New-ScheduledTask -Action $action -Trigger $triggers -Principal $principal -Settings $settings

Register-ScheduledTask -TaskName $TaskName -InputObject $task -Force | Out-Null

Write-Host "Scheduled task registered: $TaskName" -ForegroundColor Green
Write-Host "Daily window: $StartAt ~ $EndAt"
Write-Host "Daily triggers: $($triggerTimes -join ', ')"
Write-Host "Run behavior: at HH:00, process previous full hour."
Write-Host "Examples:"
Write-Host "  01:00 -> 00:00:00 ~ 00:59:59"
Write-Host "  02:00 -> 01:00:00 ~ 01:59:59"
Write-Host "  22:00 -> 21:00:00 ~ 21:59:59"
Write-Host ""
Write-Host "Command : powershell.exe $($taskArgs -join ' ')"
Write-Host ""
Write-Host "Run now:"
Write-Host "  Start-ScheduledTask -TaskName `"$TaskName`""
Write-Host ""
Write-Host "Delete task:"
Write-Host "  Unregister-ScheduledTask -TaskName `"$TaskName`" -Confirm:`$false"
