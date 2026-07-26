#                  AxonHTA Build Script
#
# AxonASP Server
# Copyright (C) 2026 G3pix Ltda. All rights reserved.
#
# Developed by Jeffrey He (@jeffreyheping)
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

<#
.SYNOPSIS
    Build script for AxonHTA desktop runtime.

.DESCRIPTION
    Compiles the AxonHTA executable with platform-appropriate linker flags.
    On Windows, the -H windowsgui flag is automatically added to hide the
    console window, producing a clean GUI application.

.PARAMETER Output
    Output executable name (default: axonhta).

.EXAMPLE
    .\build.ps1
    Builds axonhta.exe (Windows, no console window).

.EXAMPLE
    .\build.ps1 -Output myapp
    Builds myapp.exe.
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$Output = "axonhta"
)

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ScriptDir

# --- Detect version from git ---
$Version = "dev"
try {
    $GitTag = git describe --tags --abbrev=0 2>$null
    if ($LASTEXITCODE -eq 0 -and $GitTag -match '^v?(\d+\.\d+\.\d+)') {
        $Version = $matches[1]
    } else {
        $GitCount = git rev-list --count HEAD 2>$null
        if ($LASTEXITCODE -eq 0) { $Version = "0.0.$($GitCount.Trim())" }
    }
} catch {}

Write-Host ""
Write-Host "  AxonHTA Build Script" -ForegroundColor White
Write-Host "  Version: $Version" -ForegroundColor Cyan
Write-Host ""

# --- Build flags ---
$LdFlags = "-s -w -X main.Version=$Version"
$Extension = ""

if ($env:GOOS -eq "windows" -or (-not $env:GOOS -and [System.Environment]::OSVersion.Platform -eq "Win32NT")) {
    # Windows: hide console window via GUI subsystem flag
    $LdFlags = "-s -w -H windowsgui -X main.Version=$Version"
    $Extension = ".exe"
    Write-Host "  Mode: Windows GUI (no console)" -ForegroundColor Yellow
} else {
    Write-Host "  Mode: Console" -ForegroundColor Yellow
}

$OutputFile = "${Output}${Extension}"

Write-Host "  Output: $OutputFile" -ForegroundColor Gray
Write-Host ""

# --- Build ---
Write-Host "Building..." -ForegroundColor Cyan
& go build -trimpath -ldflags "$LdFlags" -o $OutputFile . 2>&1

if ($LASTEXITCODE -eq 0 -and (Test-Path $OutputFile)) {
    $Size = [math]::Round((Get-Item $OutputFile).Length / 1MB, 2)
    Write-Host ""
    Write-Host "  [OK] $OutputFile ($Size MB)" -ForegroundColor Green
    Write-Host ""
} else {
    Write-Host ""
    Write-Host "  [FAIL] Build failed" -ForegroundColor Red
    Write-Host ""
    exit 1
}
