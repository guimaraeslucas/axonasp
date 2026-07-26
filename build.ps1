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
    [ValidateSet("windows", "linux", "darwin", "wasm", "all")]
    [string]$Platform = "windows",

    [Parameter(Mandatory = $false)]
    [ValidateSet("amd64", "arm64", "386")]
    [string]$Architecture = "amd64",

    [Parameter(Mandatory = $false)]
    [switch]$Clean,

    [Parameter(Mandatory = $false)]
    [switch]$Test,

    [Parameter(Mandatory = $false)]
    [string]$Tags = ""
)

# --- AUTOMATIC VERSION CONFIGURATION ---
$Major = "2"
$Minor = "2"
$Patch = "0"
$Revision = "0"

# --- AUTOMATIC VERSION CONFIGURATION ---
$Major = "2"
$Minor = "2"
$Patch = "0"
$Revision = "0"

try {
    # Procura uma tag APENAS se ela estiver exatamente no commit atual (HEAD)
    
    $GitTag = git describe --tags --abbrev=0 2>$null
    
    # Se encontrou a tag no commit atual e ela segue o padrão
    if ($LASTEXITCODE -eq 0 -and $GitTag -match '^v?(\d+)\.(\d+)\.(\d+)$') {
        $Major = $matches[1]
        $Minor = $matches[2]
        $Patch = $matches[3]
    }
    else {
        $GitCount = git rev-list --count HEAD 2>$null
        if ($LASTEXITCODE -eq 0) { $Patch = $GitCount.Trim() }
    }

    $GitHash = git rev-parse --short HEAD 2>$null
    if ($LASTEXITCODE -eq 0) { $Revision = $GitHash.Trim() }
}
catch {
    Write-Warning "Git not found or not a valid repository. Using default versioning."
}

$FullVersion = "$Major.$Minor.$Patch.$Revision"

# Normalize Go build tags so users can pass comma/semicolon/space-separated values.
$NormalizedTags = ($Tags -replace '[,;]+', ' ').Trim()

# Color output functions
function Write-Success { param([string]$Message); Write-Host $Message -ForegroundColor Green }
function Write-Info { param([string]$Message); Write-Host $Message -ForegroundColor Cyan }
function Write-Err { param([string]$Message); Write-Host $Message -ForegroundColor Red }
function Write-Warn { param([string]$Message); Write-Host $Message -ForegroundColor Yellow }

# Script header
Write-Host ""
Write-Host "=======================================================" -ForegroundColor Magenta
Write-Host "  G3Pix AxonASP Build Script" -ForegroundColor White
Write-Host "  Version: $FullVersion" -ForegroundColor Cyan
if ($NormalizedTags) {
    Write-Host "  Build Tags: $NormalizedTags" -ForegroundColor Yellow
}
Write-Host "=======================================================" -ForegroundColor Magenta
Write-Host ""

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ScriptDir

# Clean previous builds
if ($Clean) {
    Write-Info "Cleaning previous builds..."
    Remove-Item -Path "axonasp-http.exe"    -ErrorAction SilentlyContinue
    Remove-Item -Path "axonasp-fastcgi.exe" -ErrorAction SilentlyContinue
    Remove-Item -Path "axonasp-cli.exe"     -ErrorAction SilentlyContinue
    Remove-Item -Path "axonasp-testsuite.exe" -ErrorAction SilentlyContinue
    Remove-Item -Path "axonasp-mcp.exe"     -ErrorAction SilentlyContinue
    Remove-Item -Path "axonasp-service.exe" -ErrorAction SilentlyContinue
    Remove-Item -Path "axonasp-admin.exe"   -ErrorAction SilentlyContinue
    Remove-Item -Path "axonhta.exe"        -ErrorAction SilentlyContinue
    Remove-Item -Path "build" -Recurse -Force -ErrorAction SilentlyContinue
    Write-Success "Cleaned."
    Write-Host ""
}

# Targets: label, output name, source path
$Targets = @(
    @{ Label = "HTTP Server"; Output = "axonasp-http"; Source = "./server" },
    @{ Label = "FastCGI Server"; Output = "axonasp-fastcgi"; Source = "./fastcgi" },
    @{ Label = "CLI"; Output = "axonasp-cli"; Source = "./cli" },
    @{ Label = "Test Suite"; Output = "axonasp-testsuite"; Source = "./testsuite" },
    @{ Label = "MCP"; Output = "axonasp-mcp"; Source = "./mcp" },
    @{ Label = "Service Wrapper"; Output = "axonasp-service"; Source = "./service" },
    @{ Label = "Admin Tool"; Output = "axonasp-admin"; Source = "./admin" },
    @{ Label = "HTA Desktop"; Output = "axonhta"; Source = "./axonhta" }
)

function Build-Binary {
    param(
        [string]$TargetOS,
        [string]$TargetArch,
        [string]$OutputName,
        [string]$SourcePath,
        [string]$Label
    )

    $env:GOOS = $TargetOS
    $env:GOARCH = $TargetArch

    $Extension = if ($TargetOS -eq "windows") { ".exe" } else { "" }
    $OutputFile = "${OutputName}${Extension}"
    # Linker flags: -s (strip symbol table), -w (omit DWARF), -X main.Version=... (embed version)
    $LdFlags = "-s -w -X main.Version=$FullVersion"
    # AxonHTA is a desktop app: hide the console window on Windows.
    if ($OutputName -eq "axonhta" -and $TargetOS -eq "windows") {
        $LdFlags += " -H windowsgui"
    }
    $BuildArgs = @("build", "-trimpath", "-ldflags", $LdFlags, "-o", $OutputFile)
    if ($NormalizedTags) {
        $BuildArgs += @("-tags", $NormalizedTags)
    }
    $BuildArgs += $SourcePath

    Write-Info "Building $Label ($TargetOS/$TargetArch) -> $OutputFile ..."

    $Output = & go @BuildArgs 2>&1

    if ($LASTEXITCODE -eq 0 -and (Test-Path $OutputFile)) {
        $Size = [math]::Round((Get-Item $OutputFile).Length / 1MB, 2)
        Write-Success "  [OK] $OutputFile ($Size MB)"
        return $true
    }
    else {
        Write-Err "  [FAIL] $Label"
        if ($Output) { Write-Host $Output }
        return $false
    }
}

$BuildSuccess = $true

# --- Fix, format and generate before any build pass ---
Write-Info "Fixing source..."
go fix ./... | Out-Null

Write-Info "Formatting source..."
gofmt -w . | Out-Null

Write-Info "Running go generate..."
go generate ./... | Out-Null
Write-Host ""

function Run-Platform {
    param([string]$OS, [string]$Arch)

    $OutDir = ""
    if ($OS -ne "windows") {
        $OutDir = "build/$OS-$Arch/"
        New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
    }

    Write-Host "-------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host " Building for $OS/$Arch" -ForegroundColor Yellow
    Write-Host "-------------------------------------------------------" -ForegroundColor DarkGray

    foreach ($t in $Targets) {
        $out = "$OutDir$($t.Output)"
        $ok = Build-Binary -TargetOS $OS -TargetArch $Arch -OutputName $out -SourcePath $t.Source -Label $t.Label
        $script:BuildSuccess = $script:BuildSuccess -and $ok
    }

    if ($OS -ne "windows") {
        $fpmOut = "${OutDir}axonasp-fpm"
        $fpmOk = Build-Binary -TargetOS $OS -TargetArch $Arch -OutputName $fpmOut -SourcePath "./fpm" -Label "FPM Manager"
        $script:BuildSuccess = $script:BuildSuccess -and $fpmOk
    }
    Write-Host ""
}

if ($Platform -eq "windows" -or $Platform -eq "all") { Run-Platform "windows" $Architecture }
if ($Platform -eq "linux" -or $Platform -eq "all") { Run-Platform "linux"   $Architecture }
if ($Platform -eq "darwin" -or $Platform -eq "all") { Run-Platform "darwin"  $Architecture }
if ($Platform -eq "wasm" -or $Platform -eq "all") {
    Write-Host "-------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host " Building for js/wasm" -ForegroundColor Yellow
    Write-Host "-------------------------------------------------------" -ForegroundColor DarkGray

    $OutDir = "./wasm/"
    $ok = Build-Binary -TargetOS "js" -TargetArch "wasm" -OutputName "${OutDir}axonasp.wasm" -SourcePath "./wasm/main.go" -Label "WASM Web"
    $script:BuildSuccess = $script:BuildSuccess -and $ok
    $goroot = go env GOROOT
    Copy-Item "$goroot\lib\wasm\wasm_exec.js" -Destination $OutDir -ErrorAction SilentlyContinue
    Write-Success "  [OK] Copied wasm_exec.js"
    Write-Host ""
}

# Restore to host OS after cross-compiling
$env:GOOS = "windows"
$env:GOARCH = "amd64"

# --- Tests ---
if ($Test) {
    Write-Host "-------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host " Running Tests" -ForegroundColor Yellow
    Write-Host "-------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host ""

    Write-Info "Running go test ./..."
    $TestOutput = go test ./... 2>&1

    if ($LASTEXITCODE -eq 0) {
        Write-Success "[OK] All tests passed"
    }
    else {
        Write-Err "[FAIL] Some tests failed"
        Write-Host $TestOutput
        $BuildSuccess = $false
    }
    Write-Host ""
}

# --- Summary ---
Write-Host "=======================================================" -ForegroundColor Magenta

if ($BuildSuccess) {
    Write-Success "  BUILD SUCCESSFUL  (v$FullVersion)"
    Write-Host ""
    Write-Host "  Executables:" -ForegroundColor White

    @("axonasp-http.exe", "axonasp-fastcgi.exe", "axonasp-cli.exe", "axonasp-testsuite.exe", "axonasp-mcp.exe", "axonasp-service.exe", "axonasp-admin.exe", "axonhta.exe") | ForEach-Object {
        if (Test-Path $_) { Write-Host "    - $_" -ForegroundColor Cyan }
    }

    if (Test-Path "build") {
        Get-ChildItem -Path "build" -Recurse -File | ForEach-Object {
            Write-Host "    - $($_.FullName.Replace($ScriptDir+'\',''))" -ForegroundColor Cyan
        }
    }

    Write-Host ""
    Write-Host "  Quick Start:" -ForegroundColor White
    Write-Host "    HTTP Server : .\axonasp-http.exe" -ForegroundColor Gray
    Write-Host "    FastCGI     : .\axonasp-fastcgi.exe" -ForegroundColor Gray
    Write-Host "    CLI         : .\axonasp-cli.exe" -ForegroundColor Gray
    Write-Host "    Test Suite  : .\axonasp-testsuite.exe .\www\tests" -ForegroundColor Gray
    Write-Host "    MCP         : .\axonasp-mcp.exe" -ForegroundColor Gray
    Write-Host "    Service     : .\axonasp-service.exe install|start|stop|uninstall" -ForegroundColor Gray
    Write-Host "    Admin Tool  : .\axonasp-admin.exe" -ForegroundColor Gray
    Write-Host "    HTA Desktop : .\axonhta.exe --app .\axonhta\HTAtest\axonhta-todo-app" -ForegroundColor Gray
}
else {
    Write-Err "  BUILD FAILED!"
    Write-Host ""
    Write-Host "  Check the error messages above for details." -ForegroundColor Yellow
}

Write-Host "=======================================================" -ForegroundColor Magenta
Write-Host ""

exit $(if ($BuildSuccess) { 0 } else { 1 })
