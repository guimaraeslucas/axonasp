#!/usr/bin/env bash

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
SERVICE_EXECUTABLE="./axonasp-service"
SERVICE_NAME="axonasp"

# --- Color Codes for Output ---
RED='\033[0;31m'
GREEN='\033[0;32m'
CYAN='\033[0;36m'
NC='\033[0m' # No Color

# --- Function to display error message ---
write_error() {
    echo -e "${RED}ERROR: $1${NC}"
}

# --- Function to display success message ---
write_success() {
    echo -e "${GREEN}SUCCESS: $1${NC}"
}

# --- Function to display info message ---
write_info() {
    echo -e "${CYAN}INFO: $1${NC}"
}

# --- Function to check if running as root ---
check_root() {
    if [[ $EUID -ne 0 ]]; then
        write_error "This script must be run as root (use sudo)."
        write_info "Please try again with: sudo ./uninstall-service.sh"
        exit 1
    fi
}

# --- Function to check if systemd is available ---
check_systemd() {
    if ! command -v systemctl &> /dev/null; then
        write_error "systemd is not available on this system."
        write_info "This script requires systemd for service management."
        exit 1
    fi
}

# --- Main Script Execution ---
write_info "Starting AxonASP Service Uninstallation..."

# Check if running as root
check_root

# Check if systemd is available
check_systemd

# Check if the service executable exists
if [[ ! -f $SERVICE_EXECUTABLE ]]; then
    write_error "Service executable not found at: $SERVICE_EXECUTABLE"
    write_info "Please ensure axonasp-service is in the current directory."
    exit 1
fi

# Check if service exists
write_info "Checking current service status..."
if ! systemctl is-enabled --quiet "$SERVICE_NAME" 2>/dev/null; then
    write_error "Service '$SERVICE_NAME' is not registered in the system."
    write_info "There is nothing to uninstall."
    exit 1
fi

# Display current service status
write_info "Service Status: $(systemctl is-active $SERVICE_NAME)"

# Stop the service if it is running
if systemctl is-active --quiet "$SERVICE_NAME"; then
    write_info "Stopping AxonASP Service..."
    "$SERVICE_EXECUTABLE" stop
    if [[ $? -eq 0 ]]; then
        write_success "Service stopped successfully."
    else
        write_error "Failed to stop service."
        exit 1
    fi
else
    write_info "Service is not running. Skipping stop command."
fi

# Uninstall the service
write_info "Uninstalling AxonASP Service..."
"$SERVICE_EXECUTABLE" uninstall
if [[ $? -eq 0 ]]; then
    write_success "Service uninstalled successfully."
else
    write_error "Failed to uninstall service."
    exit 1
fi

# Verify service is removed
write_info "Verifying service removal..."
sleep 1
if ! systemctl is-enabled --quiet "$SERVICE_NAME" 2>/dev/null; then
    write_success "Service has been successfully removed from the system."
else
    write_error "Service still exists after uninstall attempt."
    exit 1
fi

write_success "AxonASP Service uninstallation completed successfully!"
exit 0
