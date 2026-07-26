# Deploy AxonASP with Caddy Server Module

## Overview

The G3Pix AxonASP Caddy module provides an integrated HTTP handler that executes Classic ASP applications natively within the Caddy web server. Running AxonASP as a Caddy module is the easiest deployment model because it eliminates the need to configure, manage, and monitor separate FastCGI workers, upstream reverse proxy pools, or Unix domain sockets.

Key advantages of the Caddy module integration include:

- **Zero-Process Architecture:** The AxonASP VBScript and JavaScript runtime executes directly inside Caddy HTTP request handlers without external worker daemons.
- **Automatic HTTPS:** Benefit from Caddy's native automatic TLS certificate generation and renewal.
- **Minimal Configuration:** Enable ASP processing with a simple single-line directive in your site block.
- **Native Alignment:** Automatic integration with Caddy's native temporary storage, dynamic multi-tenant site isolation, customized directory index file resolution, and sensitive file cloaking.

## Installation and Compilation

You can obtain the custom Caddy server with the AxonASP module pre-packaged or compile it using `xcaddy`.

### Pre-Compiled Release Binaries

Pre-compiled Caddy executables containing the embedded AxonASP module are published on the official AxonASP repository release channel. Download the executable corresponding to your target platform (Windows, Linux, or macOS) from the repository releases page and place it in your system binary path (for example, `/usr/local/bin/caddy` or `C:\caddy\caddy.exe`).

### Compiling from Source with xcaddy

If you build from source or wish to bundle additional Caddy plugins, use `xcaddy` (the official Caddy builder tool).

First, install `xcaddy`:

```bash
go install github.com/caddyserver/xcaddy/cmd/xcaddy@latest
```

Next, navigate to the `caddy` module directory within the AxonASP repository and execute `xcaddy build`:

```bash
cd ./caddy
xcaddy build --with g3pix.com.br/axonasp/caddy=. --replace g3pix.com.br/axonasp=..
```

On Windows, you can also use the included PowerShell build script:

```powershell
Set-Location ./caddy
.\run_caddy.ps1
```

This compiles a custom `caddy` (or `caddy.exe`) binary containing the AxonASP module.

## Caddyfile Configuration

To enable AxonASP processing in your site block, add the `axonasp` directive inside a `route` block alongside Caddy's `file_server`.

### Basic Configuration Example

```caddyfile
example.com {
    root * /var/www/myaspapp
    route {
        axonasp
        file_server
    }
}
```

### Multi-Site Configuration with Directive Options

When hosting multiple distinct applications or setting specific paths, use subdirectives inside the `axonasp` block:

```caddyfile
site1.example.com {
    root * /var/www/site1
    route {
        axonasp {
            site_name site1_prod
            config_file /var/www/site1/config/axonasp.toml
            global_asa_path /var/www/site1/global.asa
        }
        file_server
    }
}

site2.example.com {
    root * /var/www/site2
    route {
        axonasp {
            site_name site2_prod
            global_asa_path /var/www/site2/global.asa
        }
        file_server
    }
}
```

### Directive Reference

| Subdirective | Type | Required | Description |
|---|---|---|---|
| `site_name` | String | Optional | Unique logical name for the site. Used to isolate temporary files, bytecode cache, and session data. |
| `config_file` | String | Optional | Absolute or relative path to a custom `axonasp.toml` configuration file for the site module. |
| `global_asa_path` | String | Optional | Absolute or relative path to the application `global.asa` file. |

## Runtime Features and Overrides

The AxonASP Caddy module includes native runtime overrides designed to harmonize with Caddy's architecture.

### Temporary Directory Alignment and Site Isolation

AxonASP automatically intercepts the `global.temp_dir` setting loaded from `axonasp.toml` and maps temporary storage to Caddy's native data and temporary directory hierarchy.

- **Native Storage Path:** Temporary files resolve inside Caddy's native temporary storage path (`caddy/temp/axonasp/<site_name>`).
- **Internal Subfolder Preservation:** The module creates and maintains internal subfolders `cache` (for compiled script bytecode) and `sessions` (for session state persistence).
- **Collision Prevention:** Dedicated subfolders are created per site or hostname (`site_name` if configured, or host header), allowing multiple Classic ASP applications to run on the same Caddy instance without session or cache collisions.

### Custom Index File Resolution

When a request arrives for a directory path (such as `/` or `/admin/`), the AxonASP module inspects the directory on disk and prioritizes ASP index files in this exact order:

1. `index.asp`
2. `default.asp`
3. `home.asp`
4. `main.asp`

If any of these four files exist in the target directory, AxonASP rewrites the request path and processes the target ASP script. If none of these four files exist on disk, AxonASP yields control back to Caddy's middleware pipeline, allowing `file_server` to serve standard default static files such as `index.html`.

### Cloaking Sensitive Files

To prevent unauthorized public web access to sensitive server files, the module enforces automatic request cloaking:

- **Targeted Files:** Direct HTTP requests for `global.asa` and `MyInfo.xml` (case-insensitively, anywhere in the URL path) are intercepted.
- **Client Response:** The server immediately returns an HTTP `404 Not Found` error to external web clients, concealing the existence of the files.
- **Internal Engine Access:** Internal AxonASP processes (such as `Application_OnStart` initialization, script inclusion, and server mapping) read and process these files directly from the filesystem without restriction.

## Configuration File Resolution

AxonASP configuration settings are loaded via Viper from `axonasp.toml`.

1. **Explicit Path:** If `config_file` is specified in the `axonasp` Caddyfile block, that configuration file is loaded directly.
2. **Automatic Resolution:** If no path is specified, the module searches in `./config/axonasp.toml`, the site root, or the executable directory.
3. **Environment Overrides:** Configuration keys can be overridden via environment variables when `global.viper_automatic_env` is set.

## Code Example

Below is a complete deployment setup demonstrating a multi-tenant Caddyfile configuration alongside a sample ASP page.

### Caddyfile

```caddyfile
:8080 {
    root * ./www
    route {
        axonasp {
            site_name portal_site
            global_asa_path ./www/global.asa
        }
        file_server
    }
}
```

### ASP Page (`./www/default.asp`)

```asp
<%@ Language="VBScript" %>
<!doctype html>
<html>
<head>
    <title>AxonASP Caddy Integration</title>
</head>
<body>
    <h1>Welcome to AxonASP on Caddy</h1>
    <p>Server Time: <%= Now() %></p>
    <p>Session ID: <%= Session.SessionID %></p>
    <p>Engine: <%= Request.ServerVariables("SERVER_SOFTWARE") %></p>
</body>
</html>
```
