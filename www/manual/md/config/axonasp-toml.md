# AxonASP Configuration Reference

## Overview

The `axonasp.toml` file contains all configuration settings for the AxonASP server. This file is shared across all runtime modes: HTTP server, FastCGI server, CLI, and MCP server. All settings can be overridden using environment variables, making it easy to manage different configurations for development and production environments.

## Configuration Location

The configuration file should be located at:
```
./config/axonasp.toml
```

Both the configuration folder and web root (`./www/`) must be in the same directory as the executable (or use absolute paths).

## Environment Variable Override

All configuration values can be overridden via environment variables when `viper_automatic_env = true` in `axonasp.toml`:

**Format:** `SECTION_KEY` (uppercase, with an underscore joining the `[section]` name and the setting key; dots are replaced by underscores)

**Examples:**
```powershell
# Override global settings ([global])
$env:GLOBAL_DEFAULT_CHARSET = "UTF-8"
$env:GLOBAL_DEFAULT_SCRIPT_TIMEOUT = "120"

# Override server settings ([server])
$env:SERVER_SERVER_PORT = "8802"
$env:SERVER_WEB_ROOT = "C:\myapp\www"

# Override FastCGI settings ([fastcgi])
$env:FASTCGI_SERVER_PORT = "9001"

# Override database settings ([g3db])
$env:G3DB_MYSQL_HOST = "db.example.com"
$env:G3DB_MYSQL_USER = "appuser"
$env:G3DB_MYSQL_PASS = "securepassword"
```

**Using .env File:**

Create a `.env` file in the same directory as the executable:
```
GLOBAL_DEFAULT_CHARSET=UTF-8
SERVER_SERVER_PORT=8802
G3DB_MYSQL_HOST=db.example.com
G3DB_MYSQL_USER=appuser
G3DB_MYSQL_PASS=securepassword
```

---

## Global Settings `[global]`

Global settings apply to all AxonASP runtime modes.

### default_charset

**Type:** String  
**Default:** `"UTF-8"`  
**Environment Variable:** `GLOBAL_DEFAULT_CHARSET`

Sets the character encoding used when serving content. Controls the `Content-Type: text/html; charset=` HTTP header.

**Notes:**
- Only UTF-8 is supported internally; this setting is for compatibility with legacy ASP applications
- HTTP server and FastCGI server include this charset in response headers
- CLI always uses UTF-8 regardless of this setting

**Example:**
```toml
default_charset = "UTF-8"
```

### default_mslcid

**Type:** Integer (LCID - Locale ID)  
**Default:** `1033` (English - United States)  
**Environment Variable:** `GLOBAL_DEFAULT_MSLCID`

Sets the default locale for ASP applications. Affects date/time formatting, number formatting, and language-specific functions.

**Common LCIDs:**
- `1033` - English (United States)
- `2057` - English (United Kingdom)
- `1046` - Portuguese (Brazil)
- `1051` - German (Germany)
- `1035` - Finnish
- `1040` - Italian
- `3082` - Spanish (Spain)
- `1036` - French (France)

**Example:**
```toml
default_mslcid = 1046  # Portuguese (Brazil)
```

### default_script_timeout

**Type:** Integer (seconds)  
**Default:** `60`  
**Environment Variable:** `GLOBAL_DEFAULT_SCRIPT_TIMEOUT`

Maximum execution time before the server terminates an ASP script. Prevents runaway scripts from consuming resources indefinitely.

**Guidelines:**
- **Development:** 120-300 seconds (allow for interactive debugging)
- **Production:** 30-60 seconds (strict resource control)
- **Long-Running Tasks:** 300+ seconds (but consider background jobs instead)

**Example:**
```toml
default_script_timeout = 60
```

### response_buffer_limit_mb

**Type:** Integer (megabytes)  
**Default:** `4`  
**Environment Variable:** `GLOBAL_RESPONSE_BUFFER_LIMIT_MB`

Maximum size of buffered output when `Response.Buffer` is enabled before aborting with a runtime error. Protects against unbounded memory consumption.

**Guidelines:**
- **Small Sites:** 2-4 MB (default)
- **Medium Sites:** 8-16 MB
- **Large Sites:** 32-64 MB (requires proportional memory allocation)

**Example:**
```toml
response_buffer_limit_mb = 4
```

### default_timezone

**Type:** String (IANA timezone)  
**Default:** `"UTC"`  
**Environment Variable:** `GLOBAL_DEFAULT_TIMEZONE`

Set the default timezone for the server. Make sure to use a valid timezone identifier from the IANA Time Zone Database. You can also use UTC+offsets like "UTC+2" or "UTC-5" to specify a timezone relative to UTC, but it's generally recommended to use named timezones for better clarity and to account for daylight saving time changes. The server will use this timezone setting for date and time functions in your ASP scripts, as well as for logging and other time-related operations. Note that this setting does not affect the system timezone of the server itself, which may be different from the timezone used by the AxonASP server.

**Common Timezones:**
- `"UTC"`
- `"America/New_York"`
- `"America/Chicago"`
- `"America/Los_Angeles"`
- `"Europe/London"`
- `"Europe/Paris"`
- `"Asia/Tokyo"`
- `"Asia/Shanghai"`
- `"Australia/Sydney"`
- `"America/Sao_Paulo"`

**Example:**
```toml
default_timezone = "America/New_York"
```

### enable_asp_debugging

**Type:** Boolean  
**Default:** `true`  
**Environment Variable:** `GLOBAL_ENABLE_ASP_DEBUGGING`

When enabled, the HTTP/FastCGI server provides detailed error messages, stack traces, and debugging information. Helpful during development but exposes sensitive information.

**⚠️ Security Warning:** *ALWAYS* disable in production environments! This will also enable go pprof endpoints on the proxy server version, which can be accessed at /debug/pprof and can provide detailed information about the server's performance and resource usage, but it can also pose a security risk if exposed to unauthorized users. If for any reason you need to enable ASP debugging in production, make sure to secure the pprof endpoints properly.

**Example:**
```toml
enable_asp_debugging = true  # Development only!
```

### enable_log_files

**Type:** Boolean  
**Default:** `true`  
**Environment Variable:** `GLOBAL_ENABLE_LOG_FILES` (the internal error logger also accepts `ENABLE_LOG_FILES` as a fallback)

When enabled, AxonASP writes runtime diagnostic files in `./temp/` or the directory specified by `GLOBAL_TEMP_DIR`. These files include:
- `error.log` for processed ASP/VBScript/runtime errors, plus `console.error` and `console.warn` output
- `console.log` for `console.log` and `console.info` output

Use this setting during development and incident diagnostics. In production, monitor file size and use rotation policies to prevent unbounded disk growth.

**Example:**
```toml
enable_log_files = true
```

### enable_error_log_file *(Deprecated)*

**Type:** Boolean  
**Environment Variable:** `GLOBAL_ENABLE_ERROR_LOG_FILE`  
**Status:** **Deprecated since version 2.1** — replaced by `enable_log_files`

This setting is retained for backward compatibility only. If `enable_log_files` is present in the configuration, this setting is ignored. If `enable_log_files` is absent, the server falls back to this key.

**Action required:** Replace `enable_error_log_file` with `enable_log_files` in your configuration file.

```toml
# Old (deprecated)
enable_error_log_file = true

# New (use this instead)
enable_log_files = true
```

---

### dump_preprocessed_source

**Type:** Boolean  
**Default:** `false`  
**Environment Variable:** `GLOBAL_DUMP_PREPROCESSED_SOURCE` (the compiler also checks `DUMP_PREPROCESSED_SOURCE` directly)

When enabled, saves the preprocessed ASP source code to `./temp/` before compilation. Useful for debugging parser issues but consumes disk space on high-traffic sites.

**⚠️ Performance Warning:** Disable in production! Can create thousands of temporary files.

**Example:**
```toml
dump_preprocessed_source = false
```

### clean_sessions_on_startup

**Type:** Boolean  
**Default:** `true`  
**Environment Variable:** `GLOBAL_CLEAN_SESSIONS_ON_STARTUP`

When enabled, clears all existing sessions when the server starts. Prevents stale session data but logs out active users.

**Use Cases:**
- `true` - Development, testing, or security-critical restarts
- `false` - Production with persistent user sessions

**Example:**
```toml
clean_sessions_on_startup = false  # Preserve sessions across restarts
```

### bytecode_caching_enabled

**Type:** String (Enum)  
**Default:** `"enabled"`  
**Environment Variable:** `GLOBAL_BYTECODE_CACHING_ENABLED`  
**Valid Values:** `"enabled"`, `"memory-only"`, `"disk-only"`, `"disabled"`

Controls compiled bytecode caching strategy for ASP scripts:
- `"enabled"` - Cache in both memory (tier 1) and disk (tier 2) for maximum performance
- `"memory-only"` - Cache only in RAM (tier 1); faster with higher memory usage
- `"disk-only"` - Cache only on disk (tier 2); slower but lower memory footprint
- `"disabled"` - No caching; recompile every request (development/debugging only)

**Performance Impact:**
```
enabled > memory-only > disk-only > disabled
```

**Example:**
```toml
bytecode_caching_enabled = "enabled"
```

### cache_max_size_mb

**Type:** Integer (megabytes)  
**Default:** `128`  
**Environment Variable:** `GLOBAL_CACHE_MAX_SIZE_MB`

Maximum size of the compiled bytecode cache in memory. When exceeded, least-recently-used scripts are evicted.

**Guidelines:**
- **Small Sites:** 64-128 MB
- **Medium Sites:** 128-256 MB
- **Large Sites:** 256-512 MB (requires proportional system RAM)

**Example:**
```toml
cache_max_size_mb = 128
```

### clean_cache_on_startup

**Type:** Boolean  
**Default:** `true`  
**Environment Variable:** `GLOBAL_CLEAN_CACHE_ON_STARTUP`

When enabled, clears the compiled script cache on server startup. Ensures new builds are executed but increases startup time.

**Guidelines:**
- `true` - Development (pick up recent changes)
- `false` - Production (faster startup)

**Example:**
```toml
clean_cache_on_startup = true
```

### vm_pool_size

**Type:** Integer  
**Default:** `10`  
**Environment Variable:** `GLOBAL_VM_POOL_SIZE`

Number of virtual machine instances in the execution pool. Each VM can execute one script simultaneously, so pool size determines concurrent execution capacity. A pool size of 10 VMs on a server with 512 MB of memory can handle approximately 2000 simultaneous requests for simple pages.

**Guidelines:**
- **Small Sites:** 5-10 VMs
- **Medium Sites:** 10-25 VMs
- **High-Traffic Sites:** 50-100+ VMs

**Tuning tip:** If the server is dropping or queuing requests, lower `vm_pool_size` and raise `golang_memory_limit_mb`. Request blocking is usually caused by Garbage Collector pressure rather than insufficient VM count. `golang_memory_limit_mb` has a stronger influence on throughput than `vm_pool_size`.

⚠️ Setting too high causes memory exhaustion; too low causes request queueing. You must use a value greater than 1.

**Example:**
```toml
vm_pool_size = 10
```

### golang_memory_limit_mb

**Type:** Integer (0 = unlimited)  
**Default:** `512`  
**Environment Variable:** `GLOBAL_GOLANG_MEMORY_LIMIT_MB`

Maximum memory the Go runtime is allowed to use (in megabytes). `0` means unlimited. This directive has a stronger influence on server throughput than `vm_pool_size` and directly affects how memory-intensive libraries such as G3ZSTD work.

**Guidelines:**
- **Development:** 256-512 MB
- **Production:** 512 MB or higher based on workload and number of concurrent VMs
- **Containerized:** Set to a value below the container memory limit

**Tuning tip:** If the server is missing or delaying requests, increase this value before raising `vm_pool_size`. Requests stalling are usually caused by Garbage Collector pressure triggered by a low memory ceiling rather than an insufficient VM pool.

⚠️ Note: The Go runtime may not strictly enforce this limit; actual memory usage can vary based on workload and GC behavior. Setting this too low may lead to performance degradation or out-of-memory errors.

**Example:**
```toml
golang_memory_limit_mb = 512
```

### session_flush_interval_seconds

**Type:** Integer (seconds)  
**Default:** `120`  
**Environment Variable:** `GLOBAL_SESSION_FLUSH_INTERVAL_SECONDS`

Interval for asynchronously flushing dirty in-memory sessions to disk. A value greater than `0` keeps session writes off the request hot path while still guaranteeing a safe flush on process shutdown. Higher values reduce disk I/O; lower values provide fresher data.

**Guidelines:**
- `0` - Synchronous writes on every change (safe but slower)
- `30-60` - Frequent writes (good for critical session data)
- `120` - Balanced default
- `300+` - Reduced I/O (suitable for read-heavy applications)

**Example:**
```toml
session_flush_interval_seconds = 120
```

### adodb_platform_architecture

**Type:** String (Enum)  
**Default:** `"auto"`  
**Environment Variable:** `GLOBAL_ADODB_PLATFORM_ARCHITECTURE`  
**Valid Values:** `"auto"`, `"amd64"`, `"386"`

Specifies platform architecture for ADODB Access database driver (Windows only).

- `"auto"` - Automatically detect platform architecture
- `"amd64"` - 64-bit (modern systems)
- `"386"` - 32-bit (legacy systems)

**Example:**
```toml
adodb_platform_architecture = "auto"
```

### execute_as_asp

**Type:** Array of Strings  
**Default:** `[".asp"]`  
**Environment Variable:** `GLOBAL_EXECUTE_AS_ASP` (comma-separated list)

File extensions to execute as ASP scripts. Other extensions are served as static files.

**Examples:**
```toml
# Default: only .asp files
execute_as_asp = [".asp"]

# Multiple formats
execute_as_asp = [".asp", ".aspx", ".cer", ".asa"]
```

### execute_as_vbscript

**Type:** Array of Strings  
**Default:** `[".vbs"]`  
**Environment Variable:** `GLOBAL_EXECUTE_AS_VBSCRIPT`

File extensions treated as pure VBScript code when `engine_mode` is set to `vbscript`. In this mode, ASP delimiters (`<% %>`) are not parsed, and the entire file is treated as source code.

### execute_as_javascript

**Type:** Array of Strings  
**Default:** `[".js"]`  
**Environment Variable:** `GLOBAL_EXECUTE_AS_JAVASCRIPT`

File extensions treated as pure JavaScript code when `engine_mode` is set to `javascript`. In this mode, ASP delimiters (`<% %>`) are not parsed, and the entire file is treated as source code.

### viper_watch_config

**Type:** Boolean  
**Default:** `false`  
**Environment Variable:** `GLOBAL_VIPER_WATCH_CONFIG`

When enabled, automatically reloads configuration file on changes without restarting the server.

⚠️ Note: Not fully implemented; manual restart currently required for most changes.

**Example:**
```toml
viper_watch_config = false
```

### viper_automatic_env

**Type:** Boolean  
**Default:** `true`  
**Environment Variable:** None (must be configured in `axonasp.toml`; read before automatic environment variable binding is enabled)

When enabled, automatically reads configuration values from environment variables. Allows environment variables to override `axonasp.toml` settings. Because `global.viper_automatic_env` is checked before environment variable binding is activated, setting `GLOBAL_VIPER_AUTOMATIC_ENV` or `VIPER_AUTOMATIC_ENV` in the environment has no effect.

**Example:**
```toml
viper_automatic_env = true
```

### temp_dir
**Type:** String (path)  
**Default:** `"./temp"`  
**Environment Variable:** `GLOBAL_TEMP_DIR`

This directory is used for storing temporary files created during the execution of ASP scripts, such as session data, cached compiled scripts, and other temporary resources. Make sure this directory is writable by the server process and has sufficient space to accommodate the temporary files generated by your applications.

**Example:**
```toml
temp_dir = "./temp"
```
---

## CLI Settings `[cli]`

Configuration for command-line interface (`axonasp-cli.exe`).

### enable_cli

**Type:** Boolean  
**Default:** `true`  
**Environment Variable:** `CLI_ENABLE_CLI`

When enabled, allows the TUI (Text User Interface) for interactive ASP script testing and execution. Required for CLI to function.

**Example:**
```toml
enable_cli = true
```

### enable_cli_run_from_command_line

**Type:** Boolean  
**Default:** `true`  
**Environment Variable:** `CLI_ENABLE_CLI_RUN_FROM_COMMAND_LINE`

When enabled, allows running ASP scripts directly from command line:
```powershell
.\axonasp-cli.exe -r script.asp
```

⚠️ Security: Only enable if you trust the scripts being executed.

**Example:**
```toml
enable_cli_run_from_command_line = true
```

### force_fresh_compile

**Type:** Boolean  
**Default:** `true`  
**Environment Variable:** `CLI_FORCE_FRESH_COMPILE`

When `true`, CLI execution always recompiles scripts and bypasses the bytecode cache entirely. When `false`, the CLI follows the global `bytecode_caching_enabled` setting.

**Guidelines:**
- `true` - Recommended for TUI and development workflows to ensure the latest script version is always executed
- `false` - Suitable for scheduled or automated `-r` runs where caching improves performance and scripts do not change frequently

### engine_mode

**Type:** String (Enum)  
**Default:** `"default"`  
**Environment Variable:** `CLI_ENGINE_MODE`  
**Valid Values:** `"default"`, `"vbscript"`, `"javascript"`

Sets the language mode for the CLI:
- `"default"` - Execute standard ASP (HTML + `<% %>` delimiters)
- `"vbscript"` - Execute pure VBScript (bypassing ASP delimiters for `.vbs` extensions)
- `"javascript"` - Execute pure JavaScript (bypassing ASP delimiters for `.js` extensions)

**Example:**
```toml
engine_mode = "default"
```

---

## Web Server Settings `[server]`

Configuration for HTTP web server (`axonasp-http.exe`).

### default_error_pages_directory

**Type:** String (path)  
**Default:** `"./www/error-pages"`  
**Environment Variable:** `SERVER_DEFAULT_ERROR_PAGES_DIRECTORY`

Directory containing custom error page templates (e.g., `404.asp`, `500.asp`).

**Files Searched:**
- `404.asp` / `404.html` - Not Found
- `500.asp` / `500.html` - Internal Server Error
- `403.asp` / `403.html` - Forbidden
- `400.asp` / `400.html` - Bad Request

**Example:**
```toml
default_error_pages_directory = "./www/error-pages"
```

### web_root

**Type:** String (path)  
**Default:** `"./www/"`  
**Environment Variable:** `SERVER_WEB_ROOT`

Root directory for serving files. All requests are resolved relative to this path.

**Security:** This setting cannot be overridden by `web.config` files.

**Example:**
```toml
web_root = "./www/"
```

### default_pages

**Type:** Array of Strings  
**Default:** `["index.asp", "default.asp", "index.html", ...]`  
**Environment Variable:** `SERVER_DEFAULT_PAGES` (comma-separated)

Default pages to serve when a directory is requested. Server tries each in order; serves the first found.

**Example:**
```toml
default_pages = [
  "index.asp",
  "default.asp",
  "index.html",
  "default.html",
]
```

### server_port

**Type:** Integer or String  
**Default:** `8801`  
**Environment Variable:** `SERVER_SERVER_PORT`

Port on which the HTTP server listens. Point your reverse proxy to this port.

⚠️ Recommended: Expose only through reverse proxy (nginx, Apache, IIS, Caddy)

**Example:**
```toml
server_port = 8801
```

### blocked_extensions

**Type:** Array of Strings  
**Default:** `[".asax", ".ascx", ".config", ".exe", ".dll", ...]`  
**Environment Variable:** `SERVER_BLOCKED_EXTENSIONS` (comma-separated)

File extensions that cannot be served by the web server (returns 404).

**Security Best Practices:**
- Block source code: `.cs`, `.vb`, `.asp`, `.asa`
- Block configuration: `.config`, `.toml`, `.env`
- Block binaries: `.exe`, `.dll`, `.so`
- Block database: `.mdb`, `.db`, `.sqlite`

**Example:**
```toml
blocked_extensions = [
  ".asa",
  ".cer",
  ".config",
  ".cs",
  ".exe",
  ".dll",
  ".env",
  ".mdb",
  ".toml",
  ".vb",
]
```

### blocked_files

**Type:** Array of Strings  
**Default:** `["MyInfo.xml"]`  
**Environment Variable:** `SERVER_BLOCKED_FILES` (comma-separated)

Specific files that cannot be served (returns 404).

**Example:**
```toml
blocked_files = [
  "MyInfo.xml",
  "web.config",
  "global.asa",
]
```

### blocked_dirs

**Type:** Array of Strings  
**Default:** `["./www/error-pages", "./www/axonasp-pages"]`  
**Environment Variable:** `SERVER_BLOCKED_DIRS` (comma-separated)

Directories that cannot be accessed directly (returns 404).

**Example:**
```toml
blocked_dirs = [
  "./www/error-pages",
  "./www/axonasp-pages",
  "./www/config",
]
```

### enable_webconfig

**Type:** Boolean  
**Default:** `true`  
**Environment Variable:** `SERVER_ENABLE_WEBCONFIG`

When enabled, reads `web.config` files from directories and applies settings. Allows per-application configuration (custom error pages, redirects, etc.).

**Example:**
```toml
enable_webconfig = true
```

### enable_directory_listing

**Type:** Boolean  
**Default:** `true`  
**Environment Variable:** `SERVER_ENABLE_DIRECTORY_LISTING`

When enabled, displays directory contents if no default page is found (security risk in production).

⚠️ Disable in production!

**Example:**
```toml
enable_directory_listing = false  # Production setting
```

### directory_listing_template

**Type:** String (path)  
**Default:** `"./www/axonasp-pages/directory-listing.html"`  
**Environment Variable:** `SERVER_DIRECTORY_LISTING_TEMPLATE`

HTML template used for directory listing UI. Can be customized to match site design.

**Example:**
```toml
directory_listing_template = "./www/axonasp-pages/directory-listing.html"
```

### engine_mode

**Type:** String (Enum)  
**Default:** `"default"`  
**Environment Variable:** `SERVER_ENGINE_MODE`  
**Valid Values:** `"default"`, `"vbscript"`, `"javascript"`

Sets the language mode for the HTTP server:
- `"default"` - Execute standard ASP (HTML + `<% %>` delimiters)
- `"vbscript"` - Execute pure VBScript (bypassing ASP delimiters for `.vbs` extensions)
- `"javascript"` - Execute pure JavaScript (bypassing ASP delimiters for `.js` extensions)

---

## FastCGI Server Settings `[fastcgi]`

Configuration for FastCGI application server (`axonasp-fastcgi.exe`).

### default_pages

**Type:** Array of Strings  
**Environment Variable:** `FASTCGI_DEFAULT_PAGES` (comma-separated)

Default pages for FastCGI mode (same purpose as HTTP server).

**Example:**
```toml
[fastcgi]
default_pages = [
  "index.asp",
  "default.asp",
  "index.html",
]
```

### server_port

**Type:** Integer or String  
**Default:** `9000`  
**Environment Variable:** `FASTCGI_SERVER_PORT`

Port or socket for FastCGI listener. Can be TCP port or Unix socket path.

**TCP Examples:**
```toml
server_port = 9000
```

**Unix Socket Examples (Linux/macOS):**
```toml
server_port = "unix:/tmp/axonasp.sock"
```

### engine_mode

**Type:** String (Enum)  
**Default:** `"default"`  
**Environment Variable:** `FASTCGI_ENGINE_MODE`  
**Valid Values:** `"default"`, `"vbscript"`, `"javascript"`

Sets the language mode for the FastCGI server:
- `"default"` - Execute standard ASP (HTML + `<% %>` delimiters)
- `"vbscript"` - Execute pure VBScript (bypassing ASP delimiters for `.vbs` extensions)
- `"javascript"` - Execute pure JavaScript (bypassing ASP delimiters for `.js` extensions)

---

## Database Configuration `[g3db]`

Configuration for G3DB library (multi-database support).

### MySQL Settings

```toml
[g3db]
mysql_host = "localhost"
mysql_port = 3306
mysql_user = "root"
mysql_pass = "password"
mysql_database = "test"
```

**Environment Variables:**
- `G3DB_MYSQL_HOST`
- `G3DB_MYSQL_PORT`
- `G3DB_MYSQL_USER`
- `G3DB_MYSQL_PASS`
- `G3DB_MYSQL_DATABASE`

### PostgreSQL Settings

```toml
[g3db]
postgres_host = "localhost"
postgres_port = 5432
postgres_user = "postgres"
postgres_pass = "password"
postgres_database = "test"
postgres_ssl_mode = "disable"
```

**Environment Variables:**
- `G3DB_POSTGRES_HOST`
- `G3DB_POSTGRES_PORT`
- `G3DB_POSTGRES_USER`
- `G3DB_POSTGRES_PASS`
- `G3DB_POSTGRES_DATABASE`
- `G3DB_POSTGRES_SSL_MODE`

*(Legacy `postgress_*` keys and `G3DB_POSTGRESS_*` environment variables are also accepted as fallbacks if the `postgres_*` keys are unset.)*

### MS SQL Server Settings

```toml
[g3db]
mssql_host = "localhost"
mssql_port = 1433
mssql_user = "sa"
mssql_pass = "password"
mssql_database = "test"
```

**Environment Variables:**
- `G3DB_MSSQL_HOST`
- `G3DB_MSSQL_PORT`
- `G3DB_MSSQL_USER`
- `G3DB_MSSQL_PASS`
- `G3DB_MSSQL_DATABASE`

### SQLite Settings

```toml
[g3db]
sqlite_path = "./database.db"
sqlite_busy_timeout = 5000
```

**Environment Variables:**
- `G3DB_SQLITE_PATH`
- `G3DB_SQLITE_BUSY_TIMEOUT` (milliseconds)

### Oracle Settings

```toml
[g3db]
# Option 1: Full DSN (takes precedence)
oracle_dsn = "oracle://user:password@localhost:1521/ORCLCDB"

# Option 2: Individual settings
oracle_host = "localhost"
oracle_port = 1521
oracle_user = "system"
oracle_pass = "password"
oracle_service = "ORCLCDB"
```

**Environment Variables:**
- `G3DB_ORACLE_DSN`
- `G3DB_ORACLE_HOST`
- `G3DB_ORACLE_PORT`
- `G3DB_ORACLE_USER`
- `G3DB_ORACLE_PASS`
- `G3DB_ORACLE_SERVICE`

---

## Mail Configuration `[g3mail]`

Configuration for G3Mail library (email sending).

### g3mail Settings

```toml
[g3mail]
smtp_host = "smtp.gmail.com"
smtp_port = 587
smtp_user = "your_email@gmail.com"
smtp_pass = "your_app_password"
smtp_from = "sender@example.com"
```

**Environment Variables (read directly by `G3MAIL` as fallback when properties are not set in script):**
- `SMTP_HOST`
- `SMTP_PORT`
- `SMTP_USER`
- `SMTP_PASS`
- `SMTP_FROM`

**Gmail Configuration Example:**
```toml
[g3mail]
smtp_host = "smtp.gmail.com"
smtp_port = 587  # Or 465 for TLS
smtp_user = "your-email@gmail.com"
smtp_pass = "your-app-specific-password"  # Not your account password!
smtp_from = "noreply@example.com"
```

⚠️ Security: Use environment variables or `.env` file for credentials, not hardcoded values.

---

## AxonASP Functions Settings `[axfunctions]`

Configuration for built-in Ax functions.

### enable_global_ax

**Type:** Boolean  
**Default:** `true`  
**Environment Variable:** `AXFUNCTIONS_ENABLE_GLOBAL_AX`

Makes Ax functions available globally without `Server.CreateObject()`.

**Default Method (with setting disabled):**
```vbscript
Set ax = Server.CreateObject("G3AXON.Functions")
result = ax.SomeFuction()
```

**When enabled:**
```vbscript
result = AxSomeFunction()  ' Direct call
```

⚠️ Breaks Classic ASP compatibility; enable only if needed.

**Example:**
```toml
enable_global_ax = true
```

### enable_axservershutdown_function

**Type:** Boolean  
**Default:** `false`  
**Environment Variable:** `AXFUNCTIONS_ENABLE_AXSERVERSHUTDOWN_FUNCTION`

When enabled, allows ASP scripts to call `AxShutdownAxonASPServer()` to shut down the server programmatically.

⚠️ **Security Warning:** Disable in production! Prevents unauthorized server shutdowns.

**Example:**
```toml
enable_axservershutdown_function = false
```

### ax_default_css_path

**Type:** String (path)  
**Default:** `"./www/axonasp-pages/css/axonasp.css"`  
**Environment Variable:** `AXFUNCTIONS_AX_DEFAULT_CSS_PATH`

CSS file used for built-in AxonASP pages (error pages, etc.).

**Example:**
```toml
ax_default_css_path = "./www/axonasp-pages/css/axonasp.css"
```

### ax_default_logo_path

**Type:** String (path)  
**Default:** `"./www/axonasp-pages/images/logo.png"`  
**Environment Variable:** `AXFUNCTIONS_AX_DEFAULT_LOGO_PATH`

Logo file for built-in AxonASP pages (served as inline base64 image).

**Example:**
```toml
ax_default_logo_path = "./www/axonasp-pages/images/logo.png"
```

---

## MCP Server Settings `[mcp]`

Configuration for Model Context Protocol server (`axonasp-mcp.exe`).

### mcp_mode

**Type:** String (Enum)  
**Default:** `"stdio"`  
**Environment Variable:** `MCP_MCP_MODE`  
**Valid Values:** `"stdio"`, `"sse"`

Communication mode for MCP:
- `"stdio"` - Standard input/output (local/CLI usage)
- `"sse"` - Server-Sent Events (remote web access)

**Example:**
```toml
mcp_mode = "stdio"
```

### mcp_sse_port

**Type:** Integer  
**Default:** `8000`  
**Environment Variable:** `MCP_MCP_SSE_PORT`

Port for SSE server when `mcp_mode` is set to `"sse"`. Ignored in `"stdio"` mode.

**Access Points:**
- SSE stream: `http://localhost:8000/sse`
- Command API: `http://localhost:8000/command`

**Example:**
```toml
mcp_sse_port = 8000
```

### mcp_docs

**Type:** String (path)  
**Default:** `"mcp/docs.md"`  
**Environment Variable:** `MCP_MCP_DOCS`

Path to markdown documentation file used by MCP for query responses and assistance.

**Example:**
```toml
mcp_docs = "mcp/docs.md"
```

---

## MSWC Component Settings `[mswc]`

Configuration for Microsoft Web Components (MSWC) like PageCounter.

### pagecounter_enabled

**Type:** Boolean  
**Default:** `false`  
**Environment Variable:** `MSWC_PAGECOUNTER_ENABLED`

When enabled, provides MSWC.PageCounter component for tracking page hit counts.

**Example:**
```toml
pagecounter_enabled = false
```

### pagecounter_file

**Type:** String (path)  
**Default:** `"./temp/hitcnt.gob"`  
**Environment Variable:** `MSWC_PAGECOUNTER_FILE`

Binary file (Go gob format) storing page hit counts. Created automatically if missing.

**Example:**
```toml
pagecounter_file = "./temp/hitcnt.gob"
```

### pagecounter_save_interval_seconds

**Type:** Integer (seconds)  
**Default:** `120`  
**Environment Variable:** `MSWC_PAGECOUNTER_SAVE_INTERVAL_SECONDS`

Interval for flushing hit counts to disk. Balances performance vs. data freshness.

**Guidelines:**
- Lower values → More frequent writes → Fresher data but higher I/O
- Higher values → Less frequent writes → Better performance but data loss on crash

**Example:**
```toml
pagecounter_save_interval_seconds = 120
```

---

## Service Wrapper Settings [service]

Configuration for the service wrapper binary (axonasp-service on Unix and axonasp-service.exe on Windows).

This mode is designed for easier setup, mainly on Windows and small installations. For advanced Unix production deployments, prefer native service manager configuration from the runtime service pages.

### service_name

**Type:** String  
**Default:** "AxonASPServer"  
**Environment Variable:** SERVICE_SERVICE_NAME

Internal service identifier used by the operating system service manager.

**Example:**
```toml
[service]
service_name = "AxonASPServer"
```

### service_display_name

**Type:** String  
**Default:** "G3pix AxonASP Server"  
**Environment Variable:** SERVICE_SERVICE_DISPLAY_NAME

Human-readable service name shown in service management tools.

**Example:**
```toml
[service]
service_display_name = "G3pix AxonASP Server"
```

### service_description

**Type:** String  
**Default:** "AxonASP Service running AxonASP Server. This is a wrapper used by axonasp-http or axonasp-fastcgi."  
**Environment Variable:** SERVICE_SERVICE_DESCRIPTION

Service description shown by the service manager.

**Example:**
```toml
[service]
service_description = "AxonASP Service running AxonASP Server. This is a wrapper used by axonasp-http or axonasp-fastcgi."
```

### service_executable_path

**Type:** String (path)  
**Default:** "./axonasp-http"  
**Environment Variable:** SERVICE_SERVICE_EXECUTABLE_PATH

Executable path started by the wrapper process.

Behavior:

- Relative paths are resolved from the wrapper executable directory.
- Windows automatically appends .exe when no file extension is provided.
- Can target axonasp-http or axonasp-fastcgi.

**Example:**
```toml
[service]
service_executable_path = "./axonasp-fastcgi"
```

### service_environment_variables

**Type:** Array of Strings (KEY=VALUE format)  
**Default:** []  
**Environment Variable:** SERVICE_SERVICE_ENVIRONMENT_VARIABLES

Environment variable entries applied to the child runtime process.

**Example:**
```toml
[service]
service_environment_variables = ["SERVER_SERVER_PORT=9901", "GLOBAL_DEFAULT_TIMEZONE=UTC"]
```

---

## Configuration Examples

### Development Setup

```toml
[global]
default_charset = "UTF-8"
default_mslcid = 1033
default_script_timeout = 300
enable_asp_debugging = true
enable_log_files = true
dump_preprocessed_source = false
clean_sessions_on_startup = false
bytecode_caching_enabled = "enabled"
cache_max_size_mb = 256
clean_cache_on_startup = true
vm_pool_size = 50
golang_memory_limit_mb = 256

[server]
server_port = 8801
enable_directory_listing = true
```

### Production Setup

```toml
[global]
default_charset = "UTF-8"
default_mslcid = 1033
default_script_timeout = 60
enable_asp_debugging = false
enable_log_files = true
dump_preprocessed_source = false
clean_sessions_on_startup = false
bytecode_caching_enabled = "enabled"
cache_max_size_mb = 512
clean_cache_on_startup = false
vm_pool_size = 200
golang_memory_limit_mb = 512

[server]
server_port = 8801
enable_directory_listing = false

[cli]
enable_cli = false
enable_cli_run_from_command_line = false
```

### Docker/Container Setup

Use environment variables instead of modifying the configuration file:

```yaml
# docker-compose.yml example
services:
  axonasp:
    image: axonasp:latest
    environment:
      GLOBAL_DEFAULT_SCRIPT_TIMEOUT: "120"
      SERVER_SERVER_PORT: "8801"
      GLOBAL_VM_POOL_SIZE: "100"
      GLOBAL_GOLANG_MEMORY_LIMIT_MB: "256"
      G3DB_MYSQL_HOST: "mysql"
      G3DB_MYSQL_USER: "appuser"
      G3DB_MYSQL_PASS: "securepassword"
    ports:
      - "8801:8801"
```

---

## Best Practices

1. **Development vs. Production:** Use different configurations or environment variables per environment
2. **Security:** Never hardcode credentials; use environment variables or `.env` files
3. **Performance:** Enable bytecode caching and set appropriate pool sizes
4. **Debugging:** Enable debugging/logging only during development
5. **Backups:** Keep backups of your configuration file before making changes
6. **Monitoring:** Monitor cache hit rates, error logs, and VM pool usage
7. **Scaling:** For high traffic, use reverse proxy with multiple AxonASP instances

For more information, see Project Structure and deployment examples.

- viper_watch_config enables live reload hooks in host runtimes that support it.
- Keep production values strict for blocking rules, memory limits, and directory listing policy.

## Code Example
```toml
[global]
default_script_timeout = 90
response_buffer_limit_mb = 8
bytecode_caching_enabled = "memory-only"
cache_max_size_mb = 256
vm_pool_size = 150
golang_memory_limit_mb = 512
enable_asp_debugging = false
enable_log_files = true

enable_webconfig = true
```