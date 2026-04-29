# Understand the AxonASP Project Structure

## Overview

AxonASP is a Go-based Classic ASP and VBScript execution runtime that provides multiple deployment modes. The project contains a shared VM/compiler core and multiple executable entry points for different use cases. All executables are built using the **build.ps1** (Windows) or **build.sh** (Linux/macOS) script and placed in the project root directory.

## Building

**Windows (PowerShell):**

```powershell
.\build.ps1
```

**Linux / macOS (Bash):**

```bash
./build.sh
```

For Linux cross-compilation to a specific architecture:

```bash
./build.sh --platform linux --arch amd64
./build.sh --platform linux --arch arm64
```

Both scripts build the following executables:

| Executable | Description |
|-----------|-------------|
| `axonasp-http` / `axonasp-http.exe` | HTTP/HTTPS web server (port 8801) |
| `axonasp-fastcgi` / `axonasp-fastcgi.exe` | FastCGI application server (port 9000) |
| `axonasp-cli` / `axonasp-cli.exe` | Command-line interpreter and TUI |
| `axonasp-mcp` / `axonasp-mcp.exe` | Model Context Protocol server for AI integration |
| `axonasp-testsuite` / `axonasp-testsuite.exe` | Automated ASP test suite runner |

## Project Directory Structure

```
axonasp2/
├── axonvm/                    # Core VM, compiler, and intrinsic objects
│   ├── asp/                   # ASP intrinsic objects (Request, Response, Session, etc.)
│   ├── lib_*.go               # Native object libraries (ADODB, MSXML, FSO, G3*, etc.)
│   ├── compiler*.go           # Single-pass compiler emitting bytecode
│   ├── vm.go                  # Stack-based virtual machine execution
│   ├── opcode.go              # VM opcodes and bytecode definitions
│   └── value.go               # VM value type system
│
├── vbscript/                  # Lexical analyzer and VBScript parser
│   ├── lexer.go               # Token generation from ASP source
│   ├── parser.go              # ASP parsing and validation
│   ├── token.go               # Token definitions
│   └── vberrorcodes.go        # VBScript error numbers/messages
│
├── jscript/                   # Lexical analyzer and JScript parser
│   ├── ast/ast.go             # Abstract syntax tree definitions
│   ├── parser/parser.go       # ASP parsing and validation
│   ├── token/token.go         # Token definitions
│   └── jserrorcodes.go        # JScript error numbers/messages
│
├── server/                    # HTTP web server runtime
│   ├── main.go                # HTTP listener and request handler
│   ├── web_host.go            # ASP execution and routing
│   ├── webconfig.go           # web.config parsing and rules
│   └── directorylisting.go    # Directory listing UI generation
│
├── fastcgi/                   # FastCGI application server runtime
│   └── main.go                # FastCGI listener and protocol handler
│
├── cli/                       # Command-line interface runtime
│   └── main.go                # TUI and script execution
│
├── mcp/                       # Model Context Protocol server
│   └── main.go                # MCP stdio/SSE server
│
├── testsuite/                 # Automated ASP test suite runner
│   └── main.go                # Test runner entry point
│
├── service/                   # Background service runtime
│   └── main.go                # Service entry point
│
├── config/
│   └── axonasp.toml           # Configuration file (shared by all runtimes)
│
├── axonconfig/
│   └── loader.go              # Configuration loader (shared by all runtimes)
│
├── www/                       # Web root directory
│   ├── manual/                # Built-in documentation
│   ├── tests/                 # ASP test suite pages
│   ├── error-pages/           # Custom HTTP error page templates
│   ├── axonasp-pages/         # System pages, CSS, and assets
│   ├── database-convert/      # Access database conversion tool
│   ├── mvc/                   # MVC example application
│   ├── mvvm/                  # MVVM example application
│   ├── rest/                  # REST example application
│   └── restful/               # RESTful example application
│
├── resources/                 # Static resources and data
├── docker/                    # Docker deployment files
├── temp/                      # Runtime cache and session storage
├── global.asa                 # Application-level event handlers
├── build.ps1                  # PowerShell build script (Windows)
├── build.sh                   # Bash build script (Linux/macOS)
└── go.mod, go.sum             # Go module dependencies
```

## Required Runtime Files

When deploying AxonASP executables, ensure the following are in the **same directory** as the executable:

1. **config/axonasp.toml** - Configuration file (absolute or relative path, configurable)
2. **www/** - Web root directory containing your ASP applications

Both paths are configured in `axonasp.toml` and can be set via environment variables:
```powershell
$env:WEB_ROOT = "C:\myapp\www"
$env:CONFIG_PATH = "C:\myapp\config\axonasp.toml"
```

## Environment Variable Override

All configuration values in `axonasp.toml` can be overridden via **environment variables** (requires `viper_automatic_env = true` in config):

```powershell
# Format: SECTION_SETTING or SECTION_SUBSETTING (uppercase, underscores replace dots)
$env:DEFAULT_CHARSET = "UTF-8"
$env:DEFAULT_SCRIPT_TIMEOUT = "120"
$env:SERVER_PORT = "8802"
```

## Service Ports and Endpoints

| Service | Executable | Default Port | Purpose |
|---------|-----------|--------------|---------|
| **HTTP** | axonasp-http.exe / axonasp-http | 8801 | Direct web server (development/proxy backend) |
| **FastCGI** | axonasp-fastcgi.exe / axonasp-fastcgi | 9000 | FastCGI application server |
| **CLI** | axonasp-cli.exe / axonasp-cli | N/A | Command-line script execution |
| **MCP** | axonasp-mcp.exe / axonasp-mcp | stdio / SSE | AI model integration server |
| **Test Suite** | axonasp-testsuite.exe / axonasp-testsuite | N/A | Automated ASP test runner |

## Docker Support

AxonASP includes Docker deployment support via **docker-compose.yml** and **Dockerfile**. Build and run containerized instances:

```bash
docker-compose up -d
```

This creates isolated, easily scalable AxonASP instances suitable for cloud deployment or load-balancing scenarios.

## Deployment Architecture

### NOT Recommended: Direct Web Exposure
❌ **DO NOT** expose axonasp-http.exe directly to public internet traffic
- Increases security surface area
- No centralized TLS/SSL termination
- No rate-limiting or DDoS protection
- No request logging/monitoring integration

### Recommended: Reverse Proxy Mode (axonasp-http)

Run AxonASP behind a reverse proxy (Nginx, Apache, IIS, Caddy) for security, TLS offloading, and load-balancing.

**Benefits:**
- Centralized authentication and authorization
- TLS/SSL certificate management
- Request rate-limiting and throttling
- Web Application Firewall (WAF) integration
- Multiple backend instances for redundancy
- Static asset caching and compression

### Recommended: FastCGI Mode (axonasp-fastcgi)

Use FastCGI protocol for direct integration with other web servers without reverse proxy overhead.

**Benefits:**
- Native protocol support in Nginx, Apache, IIS
- Lower latency than proxy forwarding
- Direct request handling with minimal translation
- Better integration with native web server features

---

## Deployment Architecture

AxonASP should not be exposed directly to public internet traffic. The recommended deployment patterns are:

- **Reverse Proxy Mode** — Run `axonasp-http` behind Nginx, Apache, or Caddy. The proxy handles TLS, rate-limiting, and static content. AxonASP handles only ASP execution.
- **FastCGI Mode** — Use `axonasp-fastcgi` as a backend process integrated directly with Nginx, Apache, or IIS using the FastCGI protocol.
- **Linux Service** — Run any executable as a persistent `systemd` service for production Linux deployments.

For detailed configuration examples see:

- Reverse Proxy Setup 
- FastCGI Setup 
- Running as a Linux Service
- MCP Server and VS Code Integration

## Configuration Reference

For detailed information about each configuration option, see the axonasp.toml documentation.

## Remarks

- **HTTP Server** (`axonasp-http`): High-performance web server suitable as a proxy backend or development server.
- **FastCGI Server** (`axonasp-fastcgi`): Lightweight FastCGI handler for native web server integration.
- **CLI** (`axonasp-cli`): Development and maintenance tool with interactive TUI and batch script support.
- **MCP Server** (`axonasp-mcp`): AI model integration for code generation, analysis, and documentation.
- **Test Suite** (`axonasp-testsuite`): Automated runner that discovers and executes ASP test files using the `G3TestSuite` object.
- Runtime feature parity is maintained across all deployment modes.
- All runtimes share a single `config/axonasp.toml` configuration file.
- On Linux and macOS, executables are named without the `.exe` extension.