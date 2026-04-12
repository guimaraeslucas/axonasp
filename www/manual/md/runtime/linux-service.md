# Running AxonASP as a Linux Service

## Overview

This page describes how to deploy the AxonASP HTTP or FastCGI server as a persistent background service on Linux using `systemd`. Running AxonASP as a systemd service ensures it starts automatically on boot, restarts after crashes, and integrates with standard Linux service management tools.

## Prerequisites

- Compiled `axonasp-http` or `axonasp-fastcgi` binary (built with `build.sh`)
- A Linux system with `systemd` (Ubuntu 18+, Debian 10+, CentOS 7+, and similar)
- A dedicated service account (recommended)

## Building for Linux

Use `build.sh` in the project root:

```bash
./build.sh
```

The script builds the following binaries for the current platform:
- `axonasp-http` — HTTP web server
- `axonasp-fastcgi` — FastCGI application server
- `axonasp-cli` — Command-line interpreter
- `axonasp-mcp` — MCP server
- `axonasp-service` — service wrapper helper

For cross-compilation to a specific platform/architecture:

```bash
./build.sh --platform linux --arch amd64
./build.sh --platform linux --arch arm64
```

## Creating the Service User

```bash
sudo useradd --system --no-create-home --shell /usr/sbin/nologin axonasp
```

## Deploying Files

Copy the binary, configuration, and web root to the target directory:

```bash
sudo mkdir -p /opt/axonasp
sudo cp axonasp-http /opt/axonasp/
sudo cp -r config /opt/axonasp/
sudo cp -r www /opt/axonasp/
sudo chown -R axonasp:axonasp /opt/axonasp
```

## Creating the systemd Unit File

Create a unit file at `/etc/systemd/system/axonasp.service`:

```ini
[Unit]
Description=AxonASP HTTP Server
After=network.target
Wants=network-online.target

[Service]
Type=simple
User=axonasp
Group=axonasp
WorkingDirectory=/opt/axonasp
ExecStart=/opt/axonasp/axonasp-http
Restart=on-failure
RestartSec=5
StandardOutput=journal
StandardError=journal
SyslogIdentifier=axonasp

# Resource limits
LimitNOFILE=65536
LimitNPROC=4096

# Security hardening
NoNewPrivileges=true
PrivateTmp=true

[Install]
WantedBy=multi-user.target
```

For the **FastCGI server**, replace `ExecStart` with:

```ini
ExecStart=/opt/axonasp/axonasp-fastcgi
```

## Enabling and Starting the Service

```bash
# Reload systemd to pick up the new unit file
sudo systemctl daemon-reload

# Enable the service to start on boot
sudo systemctl enable axonasp

# Start the service now
sudo systemctl start axonasp
```

## Managing the Service

```bash
# Check status
sudo systemctl status axonasp

# Stop the service
sudo systemctl stop axonasp

# Restart the service
sudo systemctl restart axonasp

# View logs
sudo journalctl -u axonasp -f

# View last 100 log lines
sudo journalctl -u axonasp -n 100 --no-pager
```

## Configuring the Web Root and Port

Edit `/opt/axonasp/config/axonasp.toml` to set the web root and port:

```toml
[server]
port = 8801
web_root = "/opt/axonasp/www"
```

These values can also be passed via environment variables in the unit file:

```ini
[Service]
Environment="WEB_ROOT=/opt/axonasp/www"
Environment="SERVER_PORT=8801"
```

## Nginx Reverse Proxy Integration

Run AxonASP behind Nginx for TLS and public-facing traffic:

```nginx
server {
    listen 443 ssl http2;
    server_name myapp.example.com;

    ssl_certificate     /etc/ssl/certs/myapp.crt;
    ssl_certificate_key /etc/ssl/private/myapp.key;

    location / {
        proxy_pass         http://127.0.0.1:8801;
        proxy_set_header   Host $host;
        proxy_set_header   X-Real-IP $remote_addr;
        proxy_set_header   X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header   X-Forwarded-Proto $scheme;
    }
}
```

## Remarks

- Use `sudo journalctl -u axonasp` for real-time log monitoring instead of separate log files.
- The `PrivateTmp=true` directive gives the service its own isolated `/tmp` directory, preventing conflicts with other processes.
- If AxonASP writes session files or cache to `./temp/`, ensure that directory is owned by the `axonasp` user and writable.
- To update the binary, stop the service, replace the file, and restart: `sudo systemctl restart axonasp`.
- For high availability, multiple AxonASP instances can be run on different ports and load-balanced by the front-end proxy.
