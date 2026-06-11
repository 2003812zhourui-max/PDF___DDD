param(
    [string]$BatchDate = (Get-Date).AddDays(-1).ToString("yyyy-MM-dd"),
    [int]$StartHour = 0,
    [int]$EndHour = 24,
    [string]$WhCodes = "US02",
    [string]$Statuses = "10,15,20,30",
    [int]$Workers = 5,
    [int]$Limit = 0,
    [string]$Channel = "",
    [string]$ProjectDir = "",
    [string]$PythonExe = "",
    [switch]$BrowserMode,
    [switch]$Ocr,
    [switch]$DryRun,
    [switch]$StopOnError
)

$ErrorActionPreference = "Stop"

if (-not $ProjectDir) {
    $ProjectDir = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
}

if (-not $PythonExe) {
    $venvPython = Join-Path $ProjectDir ".venv\Scripts\python.exe"
    if (Test-Path -LiteralPath $venvPython) {
        $PythonExe = $venvPython
    } else {
        $PythonExe = "python"
    }
}

if ($StartHour -lt 0 -or $StartHour -gt 23) {
    throw "StartHour must be between 0 and 23."
}
if ($EndHour -lt 1 -or $EndHour -gt 24 -or $EndHour -le $StartHour) {
    throw "EndHour must be between 1 and 24, and greater than StartHour."
}

$envPath = Join-Path $ProjectDir ".env"
if (Test-Path -LiteralPath $envPath) {
    Get-Content -LiteralPath $envPath | ForEach-Object {
        if ($_ -match "^\s*([^#][^=]*?)=(.*)$") {
            [Environment]::SetEnvironmentVariable($matches[1].Trim(), $matches[2].Trim(), "Process")
        }
    }
}

$batchDay = [datetime]::ParseExact($BatchDate, "yyyy-MM-dd", $null)
$batchStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$mainPy = Join-Path $ProjectDir "main.py"

if (-not (Test-Path -LiteralPath $mainPy)) {
    throw "main.py not found under ProjectDir: $ProjectDir"
}

Set-Location -LiteralPath $ProjectDir

Write-Host "ProjectDir: $ProjectDir"
Write-Host "PythonExe : $PythonExe"
Write-Host "BatchDate : $BatchDate"
Write-Host "Hours     : $StartHour-$EndHour"
Write-Host "Statuses  : $Statuses"
Write-Host "WhCodes   : $WhCodes"

$failed = @()

for ($hour = $StartHour; $hour -lt $EndHour; $hour++) {
    $start = $batchDay.Date.AddHours($hour)
    $end = $start.AddHours(1).AddSeconds(-1)
    $runName = "run_{0}_{1:00}00_{1:00}59_{2}" -f $batchDay.ToString("yyyyMMdd"), $hour, $batchStamp
    $outputDir = Join-Path $ProjectDir ("output\pdf\" + $runName)

    $argsList = @(
        $mainPy,
        "--start-time", $start.ToString("yyyy-MM-dd HH:mm:ss"),
        "--end-time", $end.ToString("yyyy-MM-dd HH:mm:ss"),
        "--wh-codes", $WhCodes,
        "--statuses", $Statuses,
        "--workers", [string]$Workers,
        "--force",
        "--output-name", $runName,
        "--output-dir", $outputDir
    )

    if ($Limit -gt 0) {
        $argsList += @("--limit", [string]$Limit)
    }
    if ($Channel) {
        $argsList += @("--channel", $Channel)
    }
    if ($BrowserMode) {
        $argsList += "--browser-mode"
    }
    if ($Ocr) {
        $argsList += "--ocr"
    }

    Write-Host ""
    Write-Host "===== $($start.ToString('HH:mm:ss')) ~ $($end.ToString('HH:mm:ss')) | $runName ====="
    Write-Host "& $PythonExe $($argsList -join ' ')"
    if ($DryRun) {
        continue
    }
    & $PythonExe @argsList
    $exitCode = $LASTEXITCODE

    if ($exitCode -ne 0) {
        $failed += $runName
        Write-Host "[FAILED] $runName exit=$exitCode" -ForegroundColor Red
        if ($StopOnError) {
            throw "Stopped because batch failed: $runName"
        }
    } else {
        Write-Host "[OK] $runName" -ForegroundColor Green
    }
}

Write-Host ""
if ($failed.Count -gt 0) {
    Write-Host "Finished with failed batches:" -ForegroundColor Yellow
    $failed | ForEach-Object { Write-Host "  $_" -ForegroundColor Yellow }
    exit 1
}

Write-Host "All hourly batches finished." -ForegroundColor Green
