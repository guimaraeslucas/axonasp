#            AxonASP Service Uninstallation Script
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

# --- Script Configuration ---
$ServiceExecutable = ".\axonasp-service.exe"
$ServiceName = "AxonASPServer"

# --- Function to check if running as Administrator ---
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# --- Function to display error message ---
function Write-ErrorMessage {
    param([string]$Message)
    Write-Host "ERROR: $Message" -ForegroundColor Red
}

# --- Function to display success message ---
function Write-SuccessMessage {
    param([string]$Message)
    Write-Host "SUCCESS: $Message" -ForegroundColor Green
}

# --- Function to display info message ---
function Write-InfoMessage {
    param([string]$Message)
    Write-Host "INFO: $Message" -ForegroundColor Cyan
}

# --- Main Script Execution ---
Write-InfoMessage "Starting AxonASP Service Uninstallation..."

# Check if running as Administrator
if (-not (Test-Administrator)) {
    Write-ErrorMessage "This script must be run as Administrator."
    Write-InfoMessage "Please run PowerShell as Administrator and try again."
    exit 1
}

# Check if the service executable exists
if (-not (Test-Path $ServiceExecutable)) {
    Write-ErrorMessage "Service executable not found at: $ServiceExecutable"
    Write-InfoMessage "Please ensure axonasp-service.exe is in the current directory."
    exit 1
}

# Check if service exists before attempting to stop/uninstall
$service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if (-not $service) {
    Write-ErrorMessage "Service '$ServiceName' not found in the system."
    Write-InfoMessage "There is nothing to uninstall."
    exit 1
}

# Check current service status
Write-InfoMessage "Checking current service status..."
$currentStatus = $service.Status
Write-InfoMessage "Service Status: $currentStatus"

# Stop the service if it is running
if ($currentStatus -eq "Running") {
    Write-InfoMessage "Stopping AxonASP Service..."
    & $ServiceExecutable stop
    if ($LASTEXITCODE -eq 0) {
        Write-SuccessMessage "Service stopped successfully."
    } else {
        Write-ErrorMessage "Failed to stop service. Exit code: $LASTEXITCODE"
        exit 1
    }
} else {
    Write-InfoMessage "Service is not running. Skipping stop command."
}

# Uninstall the service
Write-InfoMessage "Uninstalling AxonASP Service..."
& $ServiceExecutable uninstall
if ($LASTEXITCODE -eq 0) {
    Write-SuccessMessage "Service uninstalled successfully."
} else {
    Write-ErrorMessage "Failed to uninstall service. Exit code: $LASTEXITCODE"
    exit 1
}

# Verify service is removed
Write-InfoMessage "Verifying service removal..."
Start-Sleep -Seconds 1
$serviceAfterRemoval = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($serviceAfterRemoval) {
    Write-ErrorMessage "Service still exists after uninstall attempt."
    exit 1
} else {
    Write-SuccessMessage "Service has been successfully removed from the system."
}

Write-SuccessMessage "AxonASP Service uninstallation completed successfully!"
exit 0
