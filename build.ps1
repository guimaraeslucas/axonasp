#                  AxonASP Build Script
#
# AxonASP Server
# Copyright (C) 2026 G3pix Ltda. All rights reserved.
#
# Developed by Lucas Guimarães - G3pix Ltda
# Contact: https://g3pix.com.br
# Project URL: https://g3pix.com.br/axonasp
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.
#
# Attribution Notice:
# If this software is used in other projects, the name "AxonASP Server"
# must be cited in the documentation or "About" section.
#
# Contribution Policy:
# Modifications to the core source code of AxonASP Server must be
# made available under this same license terms.
#

param(
    [Parameter(Mandatory = $false)]
    [ValidateSet("windows", "linux", "darwin", "all")]
    [string]$Platform = "windows",
    
    [Parameter(Mandatory = $false)]
    [ValidateSet("amd64", "arm64", "386")]
    [string]$Architecture = "amd64",
    
    [Parameter(Mandatory = $false)]
    [switch]$Clean,
    
    [Parameter(Mandatory = $false)]
    [switch]$Test
)

# --- AUTOMATIC VERSION CONFIGURATION ---
$Major = "1"
$Minor = "0"
$Patch = "0"
$Revision = "0"

try {
    # Try to get the commit count (Patch)
    $GitCount = git rev-list --count HEAD
    if ($LASTEXITCODE -eq 0) { $Patch = $GitCount.Trim() }

    # Try to get the short hash (Revision)
    $GitHash = git rev-parse --short HEAD
    if ($LASTEXITCODE -eq 0) { $Revision = $GitHash.Trim() }
}
catch {
    Write-Warning "Git not found or not a valid repository. Using default versioning."
}

# Final format: 0.0.150.a1b2c
$FullVersion = "$Major.$Minor.$Patch.$Revision"
Write-Host "Build Version: $FullVersion" -ForegroundColor Cyan
# ------------------------------------------------

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
    
    # --- ALTERAÇÃO AQUI: Definindo as Flags do Linker ---
    # Injeta a versão na variável 'main.Version' (Ajuste o caminho do pacote se necessário)
    # O ` escapa as aspas para o PowerShell passar corretamente para o Go
    $LdFlags = "-X main.Version=$FullVersion"
    # ----------------------------------------------------

    Write-Info "Formating..."
    $BuildCommand = "gofmt -w ./.." 
    Invoke-Expression $BuildCommand | Out-Null

    Write-Info "Generating $DisplayName..."
    $BuildCommand = "go generate ./..."
    Invoke-Expression $BuildCommand | Out-Null

    Write-Info "Building $DisplayName with version $FullVersion..."
    
    $BuildCommand = "go build -trimpath -ldflags `"$LdFlags`" -o `"$OutputFile`" $SourcePath"
    
    # Executa o comando e captura saída de erro
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
