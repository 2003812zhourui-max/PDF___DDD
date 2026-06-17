param(
    [string]$RunAt = "",
    [string]$WhCodes = "US02",
    [string]$Statuses = "10,15,20,30",
    [int]$Workers = 8,
    [int]$DownloadRetries = 5,
    [double]$RetryBaseDelay = 0.8,
    [int]$WindowMinute = 0,
    [int]$Limit = 0,
    [string]$Channel = "",
    [string]$MasterOutputName = "",
    [string]$MasterOutputDir = "",
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

if ($WindowMinute -lt 0 -or $WindowMinute -gt 59) {
    throw "WindowMinute must be between 0 and 59."
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

$currentWindowStart = $referenceTime.Date.AddHours($referenceTime.Hour).AddMinutes($WindowMinute)
if ($referenceTime -lt $currentWindowStart) {
    $currentWindowStart = $currentWindowStart.AddHours(-1)
}
$start = $currentWindowStart.AddHours(-1)
$end = $currentWindowStart.AddSeconds(-1)
$batchStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$runName = "run_{0}_{1}_{2}_{3}" -f $start.ToString("yyyyMMdd"), $start.ToString("HHmm"), $end.ToString("HHmm"), $batchStamp
$outputName = if ($MasterOutputName) { $MasterOutputName } else { $runName }
$outputDir = if ($MasterOutputDir) { $MasterOutputDir } else { Join-Path $ProjectDir ("output\pdf\" + $outputName) }
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
    "--download-retries", [string]$DownloadRetries,
    "--retry-base-delay", [string]$RetryBaseDelay,
    "--force",
    "--download-name", $runName,
    "--output-name", $outputName,
    "--output-dir", $outputDir
)

# 优化备注：这些参数用于服务器/飞书部署时按网络质量调优下载容错。

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
Write-Host "WindowMin : $WindowMinute"
Write-Host "Batch     : $($start.ToString('yyyy-MM-dd HH:mm:ss')) ~ $($end.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Host "RunName   : $runName"
Write-Host "Output    : $outputName"
Write-Host "Command   : & $PythonExe $($argsList -join ' ')"

if ($DryRun) {
    exit 0
}

& $PythonExe @argsList
exit $LASTEXITCODE
