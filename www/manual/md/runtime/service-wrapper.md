# AxonASP Service Wrapper

## Overview

The AxonASP service wrapper starts and supervises either axonasp-http or axonasp-fastcgi as an operating system service.

This wrapper is intended to simplify service deployment, especially on Windows and small installations.

For advanced production environments, use the platform-native service strategy documented in the runtime manual pages, especially on Unix systems.

## Binary Name

The build process generates:

- Windows: axonasp-service.exe
- Linux and macOS: axonasp-service

## How It Works

- Reads service settings from the [service] section in config/axonasp.toml using the shared axonconfig loader.
- Resolves relative executable paths from the wrapper executable directory, not from the current working directory.
- This behavior is important on Windows, where services often start under system directories.
- Starts the configured child executable and monitors process lifetime.

## Install and Manage the Service

## Windows

Run commands from the deployment folder where axonasp-service.exe, config, and target binaries exist.

```powershell
.\axonasp-service.exe install
.\axonasp-service.exe start
.\axonasp-service.exe stop
.\axonasp-service.exe uninstall
```

## Linux or macOS

Use equivalent commands:

```bash
./axonasp-service install
./axonasp-service start
./axonasp-service stop
./axonasp-service uninstall
```

Notes:

- Service registration behavior depends on the host service manager.
- For complex Unix deployments, prefer dedicated systemd or launchd units and reverse-proxy orchestration described in other runtime pages.

## Automated Installation and Uninstallation Scripts

Helper scripts are provided to simplify service installation and uninstallation with automatic status verification:

### Windows

Use the PowerShell scripts from the project root:

**Installation:**
```powershell
.\install-service.ps1
```

This script:
- Checks for Administrator privileges.
- Installs the service using axonasp-service.exe.
- Starts the service automatically.
- Verifies service status and reports success or failure.

**Uninstallation:**
```powershell
.\uninstall-service.ps1
```

This script:
- Checks for Administrator privileges.
- Stops the running service (if active).
- Uninstalls the service.
- Verifies service removal from the system.

### Linux and Unix Systems

Use the Bash scripts from the project root. These scripts require root privileges and assume systemd is available:

**Installation:**
```bash
sudo ./install-service.sh
```

This script:
- Verifies root privileges.
- Confirms systemd availability.
- Installs the service using axonasp-service.
- Starts the service automatically.
- Verifies service status using systemctl.

**Uninstallation:**
```bash
sudo ./uninstall-service.sh
```

This script:
- Verifies root privileges.
- Confirms systemd availability.
- Stops the running service (if active).
- Uninstalls the service.
- Verifies service removal from the system.

## Service Configuration Flags in axonasp.toml

Configure the [service] section:

```toml
[service]
service_name = "AxonASPServer"
service_display_name = "G3pix AxonASP Server"
service_description = "AxonASP Service running AxonASP Server. This is a wrapper used by axonasp-http or axonasp-fastcgi."
service_executable_path = "./axonasp-http"
service_environment_variables = []
```

### service_name

Internal service identifier used by the operating system service manager.

### service_display_name

Human-readable name shown in service management tools.

### service_description

Service description displayed by the service manager.

### service_executable_path

Path to child executable started by the wrapper.

- You can point to axonasp-http or axonasp-fastcgi.
- Relative paths are resolved from the wrapper executable directory.
- On Windows, .exe is automatically appended when no extension is provided.

### service_environment_variables

List of environment variable entries in KEY=VALUE format applied to the child process.

Example:

```toml
service_environment_variables = ["SERVER_SERVER_PORT=9901"]
```

## Deployment Recommendations

- Keep axonasp-service binary, config folder, and selected child executable in the same deployment root.
- Validate service_executable_path after each deployment.
- Keep this wrapper for straightforward local and Windows service setups.
- For advanced Unix operations, use native service manager definitions and process policies from the runtime Linux service guidance.
