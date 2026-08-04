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

echo "Configuring G3pix ❖ AxonASP server environment..."

# 1. Cross-platform user and group creation
if ! id "axonasp" >/dev/null 2>&1; then
    echo "Creating 'axonasp' system user and group..."
    
    # Check for standard Linux tools (Debian/RPM)
    if command -v useradd >/dev/null 2>&1; then
        groupadd -r axonasp 2>/dev/null || true
        useradd -r -g axonasp -s /sbin/nologin -d /opt/axonasp -c "AxonASP Server" axonasp
        
    # Check for BusyBox/Alpine tools (APK)
    elif command -v adduser >/dev/null 2>&1; then
        addgroup -S axonasp 2>/dev/null || true
        adduser -S -D -H -G axonasp -h /opt/axonasp -s /sbin/nologin -g "AxonASP Server" axonasp
        
    else
        echo "Error: Could not find user management tools (useradd/adduser)."
        exit 1
    fi
else
    echo "User 'axonasp' already exists. Skipping creation."
fi

# 2. Add 'axonasp' to standard web server groups if present (Debian: www-data, Fedora/RPM: apache/nginx, Arch: http)
for webgroup in www-data apache nginx http; do
    if grep -q "^${webgroup}:" /etc/group 2>/dev/null; then
        echo "Detected '${webgroup}' group. Adding 'axonasp' to it..."
        if command -v usermod >/dev/null 2>&1; then
            usermod -aG "$webgroup" axonasp 2>/dev/null || true
        elif command -v adduser >/dev/null 2>&1; then # Alpine fallback
            adduser axonasp "$webgroup" 2>/dev/null || true
        fi
    fi
done

# 3. Apply Ownership
echo "Setting ownership for /opt/axonasp..."
if [ -d /opt/axonasp ]; then
    chown -R axonasp:axonasp /opt/axonasp
fi

# 4. Apply Write Permissions for /opt/axonasp and /opt/axonasp/www (and all subfolders)
echo "Applying permissions for /opt/axonasp and /opt/axonasp/www..."
chmod -R ug+rwX /opt/axonasp

if [ -d /opt/axonasp/www ]; then
    echo "Ensuring recursive write permissions on /opt/axonasp/www and all subdirectories..."
    chown -R axonasp:axonasp /opt/axonasp/www
    find /opt/axonasp/www -type d -exec chmod 775 {} + 2>/dev/null || chmod -R ug+rwX /opt/axonasp/www
    find /opt/axonasp/www -type f -exec chmod 664 {} + 2>/dev/null || true
fi

if [ -d /opt/axonasp/temp ]; then
    echo "Ensuring write permissions on /opt/axonasp/temp..."
    chown -R axonasp:axonasp /opt/axonasp/temp
    find /opt/axonasp/temp -type d -exec chmod 775 {} + 2>/dev/null || true
    find /opt/axonasp/temp -type f -exec chmod 664 {} + 2>/dev/null || true
fi

# Explicitly ensure binaries are executable (in case they lost the +x flag)
chmod +x /opt/axonasp/axonasp-* 2>/dev/null || true
chmod +x /opt/axonasp/*.sh 2>/dev/null || true

# 5. SELinux configuration (Fedora / RHEL / CentOS / Rocky / AlmaLinux)
if command -v selinuxenabled >/dev/null 2>&1 && selinuxenabled; then
    echo "SELinux detected. Configuring security contexts for /opt/axonasp..."
    
    # Register persistent file context rules if semanage is installed
    if command -v semanage >/dev/null 2>&1; then
        semanage fcontext -a -t httpd_sys_rw_content_t "/opt/axonasp/www(/.*)?" 2>/dev/null || true
        semanage fcontext -a -t httpd_sys_rw_content_t "/opt/axonasp/temp(/.*)?" 2>/dev/null || true
    fi
    
    # Restore contexts using restorecon if available
    if command -v restorecon >/dev/null 2>&1; then
        restorecon -R /opt/axonasp/www 2>/dev/null || true
        restorecon -R /opt/axonasp/temp 2>/dev/null || true
    fi
    
    # Apply immediate context via chcon as fallback/guarantee
    if command -v chcon >/dev/null 2>&1; then
        chcon -R -t httpd_sys_rw_content_t /opt/axonasp/www 2>/dev/null || true
        chcon -R -t httpd_sys_rw_content_t /opt/axonasp/temp 2>/dev/null || true
    fi
elif command -v chcon >/dev/null 2>&1; then
    # Direct chcon attempt if SELinux tools are partially present
    chcon -R -t httpd_sys_rw_content_t /opt/axonasp/www 2>/dev/null || true
    chcon -R -t httpd_sys_rw_content_t /opt/axonasp/temp 2>/dev/null || true
fi

echo "G3pix ❖ AxonASP installation setup completed successfully!"
echo ""
echo "If you want to install the systemd service, please run: sudo ./install-service.sh"
echo "The server is located at /opt/axonasp and runs under the 'axonasp' user for security."
echo "If you're upgrading from a previous version, please ensure that your axonasp.toml is updated with the latest configuration keys."
echo "Check the manual for further configuration and usage instructions: https://g3pix.com.br/axonasp/manual/"
echo "You can also interactively test ASP code by running 'axonasp-cli' from the command line."
