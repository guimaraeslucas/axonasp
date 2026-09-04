# AxonASP Admin (`admin`)

The `admin` package implements the visual administrative interface and configuration manager for AxonASP.

## Features

- **Configuration Management**: Programmatically reads, validates, and writes TOML configuration files (`config/axonasp.toml`) with descriptions and default settings.
- **Instance Management**: Starts, stops, and manages AxonASP server instances.
- **Diagnostics & Status**: Inspects active server runtime status, metrics, and logs.

## Target Binary

Compiles into `./axonasp-admin.exe` (or `axonasp-admin` on POSIX systems).
