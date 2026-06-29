# Use Custom Configuration Files

## Overview
This document explains how to launch and configure AxonASP executables using a custom configuration file path via the command-line flag. This enables running multiple distinct ASP applications simultaneously on the same machine using a single set of executables.

## Syntax
To run any of the AxonASP executables with a custom configuration file, use the `-c` or `--config.config_file` flag followed by the path to the target configuration file:

```shell
axonasp-http.exe -c <path_to_config_file>
axonasp-http.exe --config.config_file <path_to_config_file>
```

This syntax applies identically to all AxonASP executables:
- **axonasp-http.exe** (HTTP Web Server)
- **axonasp-fastcgi.exe** (FastCGI Server)
- **axonasp-cli.exe** (Interactive Command Line Interface)
- **axonasp-mcp.exe** (Model Context Protocol Server)
- **axonasp-service.exe** (Windows Service wrapper)
- **axonasp-testsuite.exe** (Integration Test Suite runner)

## Parameters and Arguments
- **-c, --config.config_file** (String): Required flag argument. Specifies the absolute or relative path to the custom TOML configuration file (e.g., `config/axonasp.toml` or `C:\my_app\custom-config.toml`).

## Return Values
Not applicable. The command returns standard system shell output or executes the server process in the foreground/background as configured.

## Remarks
By default, the AxonASP configuration loader searches for the configuration file named `axonasp.toml` in several relative folders (such as the current working directory, parent directories, or the executable directory). Specifying the custom configuration flag overrides this behavior entirely, instructing the shared Viper configuration manager to load the configuration from the designated path.

### Multi-Application Hosting Strategy
Because the configuration file controls critical service options—including the web root directory (`server.web_root`), the HTTP listener port (`server.server_port`), and temporary directory settings—you can host multiple independent ASP applications on the same server simultaneously:
1. Create separate TOML configuration files for each application (e.g. `app1.toml` and `app2.toml`).
2. Set different `server.web_root` and `server.server_port` values in each file (e.g., Port `8801` for App 1 and Port `8802` for App 2).
3. Start the executables pointing to their respective configuration files.

This avoids file conflicts and separates application pools, all while utilizing the same AxonASP binary files.

## Code Example
The following examples demonstrate how to start two independent AxonASP HTTP instances serving different applications on different ports:

### Configuration for Application A (app_a.toml)
```toml
[server]
server_port = 8801
web_root = "C:\\inetpub\\wwwroot\\app_a"
enable_directory_listing = false
```

### Configuration for Application B (app_b.toml)
```toml
[server]
server_port = 8802
web_root = "C:\\inetpub\\wwwroot\\app_b"
enable_directory_listing = true
```

### Execution Commands
Start both instances from separate shell terminals:

```powershell
# Start Application A
.\axonasp-http.exe --config.config_file .\config\app_a.toml

# Start Application B
.\axonasp-http.exe -c .\config\app_b.toml
```
