# AxonASP Build Script
# Copyright (c) 2025 Lucas RH Guimaraes (G3Pix)
# Licensed under the MIT License

param(
    [Parameter(Mandatory=$false)]
    [ValidateSet("windows", "linux", "darwin", "all")]
    [string]$Platform = "windows",
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("amd64", "arm64", "386")]
    [string]$Architecture = "amd64",
    
    [Parameter(Mandatory=$false)]
    [switch]$Clean,
    
    [Parameter(Mandatory=$false)]
    [switch]$Test
)

# Color output functions
function Write-Success {
    param([string]$Message)
    Write-Host $Message -ForegroundColor Green
}

function Write-Info {
    param([string]$Message)
    Write-Host $Message -ForegroundColor Cyan
}

function Write-Error-Custom {
    param([string]$Message)
    Write-Host $Message -ForegroundColor Red
}

function Write-Warning-Custom {
    param([string]$Message)
    Write-Host $Message -ForegroundColor Yellow
}

# Script header
Write-Host ""
Write-Host "=======================================================" -ForegroundColor Magenta
Write-Host "  G3Pix AxonASP Build Script" -ForegroundColor White
Write-Host "=======================================================" -ForegroundColor Magenta
Write-Host ""

# Get script directory
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ScriptDir

# Clean previous builds
if ($Clean) {
    Write-Info "Cleaning previous builds..."
    Remove-Item -Path "axonasp.exe" -ErrorAction SilentlyContinue
    Remove-Item -Path "axonaspcgi.exe" -ErrorAction SilentlyContinue
    Write-Success "Cleaned build files in current directory"
    Write-Host ""
}

# Build function
function Build-Binary {
    param(
        [string]$TargetOS,
        [string]$TargetArch,
        [string]$OutputName,
        [string]$SourcePath
    )
    
    $env:GOOS = $TargetOS
    $env:GOARCH = $TargetArch
    
    $Extension = ""
    if ($TargetOS -eq "windows") {
        $Extension = ".exe"
    }
    
    $OutputFile = "${OutputName}${Extension}"
    $DisplayName = "$OutputName ($TargetOS/$TargetArch)"
    
    Write-Info "Formating..."
    $BuildCommand = "gofmt -w ./.."

    Write-Info "Generating $DisplayName..."
    $BuildCommand = "go generate ./.."

    Write-Info "Building $DisplayName..."
    
    $BuildCommand = "go build -o `"$OutputFile`" $SourcePath"
    
    $Output = Invoke-Expression $BuildCommand 2>&1
    
    if ($LASTEXITCODE -eq 0) {
        if (Test-Path $OutputFile) {
            $FileSize = (Get-Item $OutputFile).Length
            $FileSizeMB = [math]::Round($FileSize / 1MB, 2)
            Write-Success "[OK] Built $DisplayName successfully ($FileSizeMB MB)"
            return $true
        }
        else {
            Write-Error-Custom "[FAIL] Build succeeded but output file not found: $OutputFile"
            return $false
        }
    }
    else {
        Write-Error-Custom "[FAIL] Build failed for $DisplayName"
        Write-Host $Output
        return $false
    }
}

# Track build success
$BuildSuccess = $true

# Build based on platform parameter
if ($Platform -eq "windows" -or $Platform -eq "all") {
    Write-Host "-------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host " Building for Windows ($Architecture)" -ForegroundColor Yellow
    Write-Host "-------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host ""
    
    $Result1 = Build-Binary -TargetOS "windows" -TargetArch $Architecture -OutputName "axonasp" -SourcePath "."
    $Result2 = Build-Binary -TargetOS "windows" -TargetArch $Architecture -OutputName "axonaspcgi" -SourcePath "./axonaspcgi"
    
    $BuildSuccess = $BuildSuccess -and $Result1 -and $Result2
    Write-Host ""
}

if ($Platform -eq "linux" -or $Platform -eq "all") {
    Write-Host "-------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host " Building for Linux ($Architecture)" -ForegroundColor Yellow
    Write-Host "-------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host ""
    
    # Create build directory if it doesn't exist
    New-Item -ItemType Directory -Force -Path "build/linux-$Architecture" | Out-Null
    
    $Result1 = Build-Binary -TargetOS "linux" -TargetArch $Architecture -OutputName "build/linux-$Architecture/axonasp" -SourcePath "."
    $Result2 = Build-Binary -TargetOS "linux" -TargetArch $Architecture -OutputName "build/linux-$Architecture/axonaspcgi" -SourcePath "./axonaspcgi"
    
    $BuildSuccess = $BuildSuccess -and $Result1 -and $Result2
    Write-Host ""
}

if ($Platform -eq "darwin" -or $Platform -eq "all") {
    Write-Host "-------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host " Building for macOS ($Architecture)" -ForegroundColor Yellow
    Write-Host "-------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host ""
    
    # Create build directory if it doesn't exist
    New-Item -ItemType Directory -Force -Path "build/darwin-$Architecture" | Out-Null
    
    $Result1 = Build-Binary -TargetOS "darwin" -TargetArch $Architecture -OutputName "build/darwin-$Architecture/axonasp" -SourcePath "."
    $Result2 = Build-Binary -TargetOS "darwin" -TargetArch $Architecture -OutputName "build/darwin-$Architecture/axonaspcgi" -SourcePath "./axonaspcgi"
    
    $BuildSuccess = $BuildSuccess -and $Result1 -and $Result2
    Write-Host ""
}

# Run tests if requested
if ($Test) {
    Write-Host "-------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host " Running Tests" -ForegroundColor Yellow
    Write-Host "-------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host ""
    
    Write-Info "Running Go tests..."
    $TestOutput = go test ./... 2>&1
    
    if ($LASTEXITCODE -eq 0) {
        Write-Success "[OK] All tests passed"
    }
    else {
        Write-Error-Custom "[FAIL] Some tests failed"
        Write-Host $TestOutput
        $BuildSuccess = $false
    }
    Write-Host ""
}

# Build summary
Write-Host "=======================================================" -ForegroundColor Magenta

if ($BuildSuccess) {
    Write-Success "  BUILD SUCCESSFUL!"
    Write-Host ""
    Write-Host "  Executables ready:" -ForegroundColor White
    
    if ($Platform -eq "windows" -or $Platform -eq "all") {
        if (Test-Path "axonasp.exe") {
            Write-Host "    - axonasp.exe (Standalone Server)" -ForegroundColor Cyan
        }
        if (Test-Path "axonaspcgi.exe") {
            Write-Host "   - axonaspcgi.exe (FastCGI Server)" -ForegroundColor Cyan
        }
    }
    
    if (Test-Path "build") {
        Get-ChildItem -Path "build" -Recurse -File | ForEach-Object {
            $RelativePath = $_.FullName.Replace($ScriptDir + "\", "")
            Write-Host "    - $RelativePath" -ForegroundColor Cyan
        }
    }
    
    Write-Host ""
    Write-Host "  Quick Start:" -ForegroundColor White
    Write-Host "    Standalone:  .\axonasp.exe" -ForegroundColor Gray
    Write-Host "    FastCGI:     .\axonaspcgi.exe -listen :9000 -root ./www" -ForegroundColor Gray
}
else {
    Write-Error-Custom "  BUILD FAILED!"
    Write-Host ""
    Write-Host "  Check the error messages above for details." -ForegroundColor Yellow
}

Write-Host "=======================================================" -ForegroundColor Magenta
Write-Host ""

# Exit with appropriate code
if ($BuildSuccess) {
    exit 0
}
else {
    exit 1
}
