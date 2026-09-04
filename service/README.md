# AxonASP Service Runner (`service`)

The `service` package provides automated background service daemon management across Windows, Linux (systemd), and macOS (launchd) platforms.

## Key Features

- **Cross-Platform Daemon Management**: Installs, manages, and runs AxonASP as a background OS service (Windows Service, Linux systemd daemon, or macOS launchd service).
- **Target Executable Customization**: Allows configuring which AxonASP binary (HTTP server or FastCGI server) is supervised by the service daemon.
- **Lifecycle Management**: Handles start, stop, pause, restart signals, and auto-restart policies.

## Target Binary

Compiles into `./axonasp-service.exe`.
