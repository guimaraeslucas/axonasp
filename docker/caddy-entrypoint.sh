#!/bin/sh
set -e

# Ensure runtime directories exist and have correct ownership
mkdir -p /opt/axonasp/temp
chown axonasp:axonasp /opt/axonasp/temp

# Drop privileges from root to axonasp user and execute the main binary
exec su-exec axonasp "$@"
