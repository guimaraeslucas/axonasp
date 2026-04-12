param(
    [string]$TestsPath = ".\\www\\tests",
    [string]$CliPath = ".\\axonasp-cli.exe",
    [string]$LogPath = ".\\temp\\www-tests-vbscript-errors.log",
    [int]$LogWriteRetries = 10,
    [int]$LogWriteRetryDelayMs = 150
)

$ErrorActionPreference = "Stop"

if (Get-Variable -Name PSNativeCommandUseErrorActionPreference -ErrorAction SilentlyContinue) {
    $PSNativeCommandUseErrorActionPreference = $false
}

if (-not (Test-Path $CliPath)) {
    Write-Error "CLI executable not found: $CliPath"
    exit 1
}

if (-not (Test-Path $TestsPath)) {
    Write-Error "Tests path not found: $TestsPath"
    exit 1
}

$logDir = Split-Path -Parent $LogPath
if ($logDir -and -not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir | Out-Null
}

function Write-LogLine {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Value,
        [Parameter(Mandatory = $true)]
        [bool]$Overwrite
    )

    for ($attempt = 1; $attempt -le $LogWriteRetries; $attempt++) {
        try {
            if ($Overwrite) {
                Set-Content -Path $Path -Encoding UTF8 -Value $Value
            }
            else {
                Add-Content -Path $Path -Encoding UTF8 -Value $Value
            }
            return
        }
        catch {
            if ($attempt -ge $LogWriteRetries) {
                throw
            }
            [System.Threading.Thread]::Sleep($LogWriteRetryDelayMs)
        }
    }
}

$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$effectiveLogPath = $LogPath
try {
    Write-LogLine -Path $effectiveLogPath -Value "[$timestamp] Starting CLI scan for ASP tests in $TestsPath" -Overwrite $true
}
catch {
    $fallbackName = "www-tests-vbscript-errors-{0}.log" -f (Get-Date -Format "yyyyMMdd-HHmmss")
    $effectiveLogPath = Join-Path -Path $logDir -ChildPath $fallbackName
    Write-LogLine -Path $effectiveLogPath -Value "[$timestamp] Starting CLI scan for ASP tests in $TestsPath" -Overwrite $true
}

$files = Get-ChildItem -Path $TestsPath -Filter "*.asp" -File | Sort-Object FullName
$total = $files.Count
$withVBScriptError = 0
$failedRuns = 0

if ($total -eq 0) {
    Write-Output "No ASP files found in $TestsPath"
    Write-LogLine -Path $effectiveLogPath -Value "No ASP files found." -Overwrite $false
    exit 0
}

foreach ($file in $files) {
    Write-Output "Running: $($file.FullName)"

    $previousErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        $outputLines = & $CliPath -r $file.FullName 2>&1
    }
    finally {
        $ErrorActionPreference = $previousErrorActionPreference
    }
    $exitCode = $LASTEXITCODE
    $output = (($outputLines | ForEach-Object { $_.ToString() }) -join [Environment]::NewLine)

    $hasVBScriptRuntimeError = $false
    if ($output -match "VBScript runtime error") {
        $hasVBScriptRuntimeError = $true
    }

    if ($exitCode -ne 0) {
        $failedRuns++
    }

    if ($hasVBScriptRuntimeError) {
        $withVBScriptError++

        Write-LogLine -Path $effectiveLogPath -Value "" -Overwrite $false
        Write-LogLine -Path $effectiveLogPath -Value ("=" * 80) -Overwrite $false
        Write-LogLine -Path $effectiveLogPath -Value ("File: {0}" -f $file.FullName) -Overwrite $false
        Write-LogLine -Path $effectiveLogPath -Value ("ExitCode: {0}" -f $exitCode) -Overwrite $false
        Write-LogLine -Path $effectiveLogPath -Value "Detected: VBScript runtime error" -Overwrite $false
        Write-LogLine -Path $effectiveLogPath -Value "Output:" -Overwrite $false
        Write-LogLine -Path $effectiveLogPath -Value $output.TrimEnd() -Overwrite $false
    }
}

$summary = "Total files: $total | CLI non-zero exits: $failedRuns | VBScript runtime errors: $withVBScriptError"
Write-Output $summary
Write-LogLine -Path $effectiveLogPath -Value "" -Overwrite $false
Write-LogLine -Path $effectiveLogPath -Value $summary -Overwrite $false

if ($withVBScriptError -gt 0) {
    Write-Output "VBScript runtime errors were logged to: $effectiveLogPath"
    exit 2
}

exit 0
