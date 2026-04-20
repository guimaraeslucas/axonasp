#!/usr/bin/env bash

#              AxonASP Service Installation Script
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
# Post-install script for AxonASP Server

echo "Configuring G3pix AxonASP server environment..."

# 1. Cross-platform user and group creation
if ! id "axonasp" >/dev/null 2>&1; then
    echo "Creating 'axonasp' system user and group..."
    
    # Check for standard Linux tools (Debian/RPM)
    if command -v useradd >/dev/null 2>&1; then
        groupadd -r axonasp 2>/dev/null || true
        useradd -r -g axonasp -s /sbin/nologin -d /opt/axonasp -c "AxonASP Web Server" axonasp
        
    # Check for BusyBox/Alpine tools (APK)
    elif command -v adduser >/dev/null 2>&1; then
        addgroup -S axonasp 2>/dev/null || true
        adduser -S -D -H -G axonasp -h /opt/axonasp -s /sbin/nologin -g "AxonASP Web Server" axonasp
        
    else
        echo "Error: Could not find user management tools (useradd/adduser)."
        exit 1
    fi
else
    echo "User 'axonasp' already exists. Skipping creation."
fi

# 2. Debian specific: Add to www-data group if it exists
# This allows AxonASP to write to standard /var/www/ directories in Debian/Ubuntu
if grep -q "^www-data:" /etc/group; then
    echo "Detected 'www-data' group. Adding 'axonasp' to it..."
    if command -v usermod >/dev/null 2>&1; then
        usermod -aG www-data axonasp
    elif command -v adduser >/dev/null 2>&1; then # Alpine fallback just in case
        adduser axonasp www-data
    fi
fi

# 3. Apply Ownership
echo "Setting ownership for /opt/axonasp..."
chown -R axonasp:axonasp /opt/axonasp

# 4. Apply Group-Write Permissions
echo "Applying group-write permissions..."
# ug+rwX gives User and Group Read/Write access. 
# The uppercase 'X' ensures directories are accessible without making normal files executable.
chmod -R ug+rwX /opt/axonasp

# Explicitly ensure binaries are executable (in case they lost the +x flag)
chmod +x /opt/axonasp/axonasp-* 2>/dev/null || true
chmod +x /opt/axonasp/*.sh 2>/dev/null || true

echo "G3pix AxonASP installation setup completed successfully!\n"
echo "If you want to install the systemd service, please run: sudo ./install-service.sh"
echo "The server is located at /opt/axonasp and runs under the 'axonasp' user for security."