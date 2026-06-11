param(
    [string]$RunAt = "",
    [string]$WhCodes = "US02",
    [string]$Statuses = "10,15,20,30",
    [int]$Workers = 5,
    [int]$Limit = 0,
    [string]$Channel = "",
    [string]$ProjectDir = "",
    [string]$PythonExe = "",
    [switch]$BrowserMode,
    [switch]$Ocr,
    [switch]$DryRun
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

if ($RunAt) {
    $referenceTime = [datetime]::Parse($RunAt)
} else {
    $referenceTime = Get-Date
}

$currentHourStart = $referenceTime.Date.AddHours($referenceTime.Hour)
$start = $currentHourStart.AddHours(-1)
$end = $currentHourStart.AddSeconds(-1)
$batchStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$runName = "run_{0}_{1:00}00_{1:00}59_{2}" -f $start.ToString("yyyyMMdd"), $start.Hour, $batchStamp
$outputDir = Join-Path $ProjectDir ("output\pdf\" + $runName)
$mainPy = Join-Path $ProjectDir "main.py"

if (-not (Test-Path -LiteralPath $mainPy)) {
    throw "main.py not found under ProjectDir: $ProjectDir"
}

Set-Location -LiteralPath $ProjectDir

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

Write-Host "ProjectDir: $ProjectDir"
Write-Host "PythonExe : $PythonExe"
Write-Host "RunAt     : $($referenceTime.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Host "Batch     : $($start.ToString('yyyy-MM-dd HH:mm:ss')) ~ $($end.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Host "RunName   : $runName"
Write-Host "Command   : & $PythonExe $($argsList -join ' ')"

if ($DryRun) {
    exit 0
}

& $PythonExe @argsList
exit $LASTEXITCODE
