# Manage AxonASP Global and FPM Configuration with AxonAdmin

## Overview

The axonadmin executable is the dedicated configuration manager for G3Pix AxonASP. It supports command-line creation flows and a browser-based dashboard for:

- Editing axonasp.toml sections.
- Editing AxonASP FPM pool files in the fpm/fpm.d directory.
- Creating new global and FPM configuration files from a documented default schema.
- Viewing real-time host and AxonASP process telemetry on the Home dashboard.

## Syntax

Run axonadmin from the command line using the following syntax:

```cmd
axonadmin.exe [flags]
```

## Parameters and Arguments

The tool accepts the following command-line flags:

- **-edit <path>**
  Selects the axonasp.toml target file for UI editing mode. If omitted, axonadmin uses the standard path resolution sequence.
- **-create <path>**
  Creates a new default axonasp.toml file with documented fields at the provided path.
- **-create-fpm <path>**
  Creates a new default FPM pool configuration file (.conf) with documented fields at the provided path.
- **-noui**
  Runs in headless mode. Use it together with -create or -create-fpm.
- **-h, --help**
  Displays the command help output.

## Return Values

Axonadmin returns the following command execution behavior:

- **Exit code 0:** Command completed successfully, including file creation and UI startup.
- **Exit code 1:** Command failed due to invalid combinations, file system permission issues, or configuration write errors.
- **Headless output:** Prints the final created path when -create or -create-fpm succeeds.

## Remarks

### UI Workspace Layout

When started without headless mode, axonadmin starts on localhost:8088 and opens the default browser. The sidebar is partitioned into three groups:

- **Home:** Main dashboard with host telemetry and monitored process summary.
- **axonasp.toml:** Lists the available axonasp.toml sections and opens the selected section editor.
- **AxonASP FPM:** Lists all pool files from fpm/fpm.d/*.conf and opens the selected pool editor.

At the bottom of the sidebar, the dashboard provides two creation actions:

- **+ Global Configuration:** Generates a new axonasp.toml file from the default schema.
- **+ FPM Configuration:** Prompts for a pool filename and creates fpm/fpm.d/<name>.conf from the default schema.

### FPM Pool Editing and Generation

FPM pool files are handled as TOML documents. Axonadmin loads and saves properties using a documented schema so generated files keep descriptive comments for each setting. Existing pool files can be edited and saved directly from the AxonASP FPM section.

Before writing updates, axonadmin creates a .bak backup for the target file when the file already exists.

### Home Telemetry

The Home dashboard includes machine and runtime telemetry, including:

- Machine name.
- Operating system information.
- Processor model/type.
- Total, available, and used memory.
- Memory usage percentage.
- Active process count and aggregated RSS memory for these executable families:
  - axonasp-http
  - axonasp-fastcgi
  - axonasp-fpm
  - axonasp-service

### Path Resolution Sequence for axonasp.toml

If -edit is omitted, axonadmin resolves axonasp.toml in this order:

1. config/axonasp.toml (current working directory)
2. ../config/axonasp.toml
3. ../../config/axonasp.toml
4. config/axonasp.toml relative to the executable directory

### Creation Behavior Notes

- Do not combine -create and -create-fpm in the same command.
- If -noui is used without a creation flag, the command shows help output.

## Code Example

```cmd
:: Create a new default global configuration in headless mode
axonadmin.exe -create C:\axonasp\config\axonasp.toml -noui

:: Create a new default FPM pool configuration in headless mode
axonadmin.exe -create-fpm C:\axonasp\fpm\fpm.d\example.conf -noui

:: Open the visual manager for a specific global configuration file
axonadmin.exe -edit C:\axonasp\config\axonasp.toml
```
