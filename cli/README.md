# AxonASP CLI (`cli`)

The `cli` package implements the Command Line Interface (CLI) and Terminal User Interface (TUI) for executing, debugging, and inspecting Classic ASP scripts directly from the console.

## Features

- **Direct Script Execution**: Execute standalone `.asp` files directly using the `-r` flag (e.g. `./axonasp-cli.exe -r path/to/script.asp`).
- **Global.asa Support**: Honors application and session lifecycle events defined in local `global.asa` files.
- **Interactive TUI**: Provides console-based script runner utilities and diagnostics without requiring a running web server.

## Target Binary

Compiles into `./axonasp-cli.exe`.
