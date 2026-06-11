param(
    [Parameter(Mandatory = $true)]
    [string]$StartTime,
    [Parameter(Mandatory = $true)]
    [string]$EndTime,
    [string]$MasterOutputName = "",
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

$envPath = Join-Path $ProjectDir ".env"
if (Test-Path -LiteralPath $envPath) {
    Get-Content -LiteralPath $envPath | ForEach-Object {
        if ($_ -match "^\s*([^#][^=]*?)=(.*)$") {
            [Environment]::SetEnvironmentVariable($matches[1].Trim(), $matches[2].Trim(), "Process")
        }
    }
}

$startAt = [datetime]::Parse($StartTime)
$endAt = [datetime]::Parse($EndTime)
if ($endAt -le $startAt) {
    throw "EndTime must be later than StartTime."
}

if (-not $MasterOutputName) {
    $MasterOutputName = "master_{0}_{1}_{2}" -f $startAt.ToString("yyyyMMdd_HHmm"), $endAt.ToString("yyyyMMdd_HHmm"), (Get-Date -Format "yyyyMMdd_HHmmss")
}

$masterOutputDir = Join-Path $ProjectDir ("output\pdf\" + $MasterOutputName)
$mainPy = Join-Path $ProjectDir "main.py"
if (-not (Test-Path -LiteralPath $mainPy)) {
    throw "main.py not found under ProjectDir: $ProjectDir"
}

Set-Location -LiteralPath $ProjectDir

Write-Host "ProjectDir       : $ProjectDir"
Write-Host "PythonExe        : $PythonExe"
Write-Host "Window           : $($startAt.ToString('yyyy-MM-dd HH:mm:ss')) ~ $($endAt.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Host "MasterOutputName : $MasterOutputName"
Write-Host "MasterOutputDir  : $masterOutputDir"
Write-Host "Statuses         : $Statuses"
Write-Host "WhCodes          : $WhCodes"

$failed = @()
$cursor = $startAt
$batchStamp = Get-Date -Format "yyyyMMdd_HHmmss"

while ($cursor -lt $endAt) {
    $batchStart = $cursor
    $batchEndExclusive = $cursor.AddHours(1)
    if ($batchEndExclusive -gt $endAt) {
        $batchEndExclusive = $endAt
    }
    $batchEnd = $batchEndExclusive.AddSeconds(-1)
    $downloadName = "run_{0}_{1}_{2}_{3}" -f $batchStart.ToString("yyyyMMdd"), $batchStart.ToString("HHmm"), $batchEnd.ToString("HHmm"), $batchStamp

    $argsList = @(
        $mainPy,
        "--start-time", $batchStart.ToString("yyyy-MM-dd HH:mm:ss"),
        "--end-time", $batchEnd.ToString("yyyy-MM-dd HH:mm:ss"),
        "--wh-codes", $WhCodes,
        "--statuses", $Statuses,
        "--workers", [string]$Workers,
        "--force",
        "--download-name", $downloadName,
        "--output-name", $MasterOutputName,
        "--output-dir", $masterOutputDir
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
    Write-Host "===== $($batchStart.ToString('yyyy-MM-dd HH:mm:ss')) ~ $($batchEnd.ToString('yyyy-MM-dd HH:mm:ss')) | $downloadName ====="
    Write-Host "& $PythonExe $($argsList -join ' ')"

    if (-not $DryRun) {
        & $PythonExe @argsList
        $exitCode = $LASTEXITCODE
        if ($exitCode -ne 0) {
            $failed += $downloadName
            Write-Host "[FAILED] $downloadName exit=$exitCode" -ForegroundColor Red
            if ($StopOnError) {
                throw "Stopped because batch failed: $downloadName"
            }
        } else {
            Write-Host "[OK] $downloadName" -ForegroundColor Green
        }
    }

    $cursor = $batchEndExclusive
}

Write-Host ""
if ($failed.Count -gt 0) {
    Write-Host "Finished with failed batches:" -ForegroundColor Yellow
    $failed | ForEach-Object { Write-Host "  $_" -ForegroundColor Yellow }
    exit 1
}

Write-Host "Master run finished." -ForegroundColor Green
Write-Host "Excel: $(Join-Path $masterOutputDir ($MasterOutputName + '.xlsx'))"
