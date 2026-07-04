# FastCGI Setup

## Overview

AxonASP provides a FastCGI application server (`axonasp-fastcgi.exe`) that integrates directly with Nginx, Apache or other servers using the FastCGI protocol. In this mode AxonASP acts as a backend process and the front-end web server handles all HTTP connections, static content, and TLS termination. There is no support for IIS FastCGI, you must use the http platform handler to proxy requests to the AxonASP default server.

AxonASP FastCGI supports **multi-host deployments** with different document roots. This allows a single FastCGI server process to serve content from multiple virtual hosts, each with its own document root directory. 

Attention: the FastCGI implementation does not support serving global.asa from multiple virtual hosts, it will only use the global.asa from the root directory definied by the key server.web_root or by the --config.global_asa flag. If you need to serve different global.asa files for different virtual hosts, you must run multiple FastCGI server processes with different configuration files.

There is a future plan to support multiple global.asa files for different virtual hosts by using a PHP-FPM like approach, but it is not fully implemented yet/tested, although the FPM implementation is already available in the source code and can be used to run multiple FastCGI processes with different global.asa files.

## Prerequisites

- `axonasp-fastcgi.exe` running and reachable on the configured port (default: 9000)
- Front-end web server with FastCGI support installed (nginx, Apache)

The FastCGI port is configured in `config/axonasp.toml`:

```toml
[fastcgi]
server_port = 9000
```

The same setting also accepts endpoint-style values:

- `9000` (port only)
- `127.0.0.1:9000` (host and port)
- `:9000` (bind localhost on a specific port)
- `unix:/tmp/axonasp.sock` (Unix socket, Linux/macOS)

The endpoint can be configured as a number or as a string in TOML:

```toml
[fastcgi]
server_port = "9000"
```

Unix socket example:

```toml
server_port = "unix:/tmp/axonasp.sock"
```

If the endpoint value is empty, AxonASP falls back to port `9000`.

## Startup Flags

FastCGI startup supports these relevant flags:

- `-c`, `--config.config_file`: Custom configuraton TOML file path.
- `--fastcgi.server_port`: Overrides `fastcgi.server_port` using the same endpoint format accepted in TOML.
- `--config.global_asa`: Optional directory that must contain the `global.asa` to be used. If the file is not found in this directory, AxonASP will terminate with an internal 500 error. If not set, AxonASP will fallback to `server.web_root` and then the current working directory of the executable.

Examples:

```powershell
.\axonasp-fastcgi.exe --fastcgi.server_port 9001
```

```powershell
.\axonasp-fastcgi.exe --fastcgi.server_port unix:/tmp/axonasp.sock
```

```powershell
.\axonasp-fastcgi.exe --config.global_asa /var/www/app
```

**Start AxonASP FastCGI:**

```powershell
.\axonasp-fastcgi.exe
```

**Start AxonASP FastCGI with a Custom Configuration File:**

You can launch a FastCGI server process with a custom configuration using the `-c` or `--config.config_file` flag. This allows you to start multiple distinct FastCGI server instances running on different ports or unix sockets with separate configurations (such as different session limits, web roots, or database connection strings):

```powershell
.\axonasp-fastcgi.exe -c .\config\app_fastcgi.toml
```

## global.asa Startup Resolution

At startup, G3Pix AxonASP resolves `global.asa` using this order:

1. If `--config.global_asa` is explicitly set, AxonASP checks `<directory>/global.asa` in that path only.
2. If `--config.global_asa` is not set, AxonASP checks `server.web_root`.
3. If not found in `server.web_root`, AxonASP checks the process current working directory.
4. If not found in either fallback location, AxonASP skips `global.asa` execution and continues startup.

Validation behavior for explicit flag use:

- If `--config.global_asa` is set and `global.asa` is missing (or invalid in that directory), AxonASP writes a startup log entry and terminates startup with an internal 500 state.

Startup logging behavior:

- When found, AxonASP logs the selected source directory for `global.asa`.
- When skipped by fallback, AxonASP logs that no file was found in `server.web_root` and current working directory.

## DOCUMENT_ROOT and SCRIPT_NAME Parameters

AxonASP FastCGI reads the following FastCGI CGI variables to support multi-host deployments:

- **`DOCUMENT_ROOT`**: The directory where the virtual host's files are located. When provided by the reverse proxy, files are resolved from this directory instead of the configured `RootDir`.
- **`SCRIPT_NAME`**: The virtual path to the requested ASP script (e.g., `/default.asp`). When provided, this takes priority over the URL path for path resolution.

If `DOCUMENT_ROOT` is not provided, the server falls back to the `RootDir` configured in the TOML file, ensuring backward compatibility with existing single-host setups.

## Nginx with FastCGI

### Single Virtual Host

```nginx
upstream axonasp_fcgi {
    server localhost:9000 max_fails=3 fail_timeout=30s;
}

server {
    listen 443 ssl http2;
    server_name myapp.example.com;

    ssl_certificate /etc/ssl/certs/myapp.crt;
    ssl_certificate_key /etc/ssl/private/myapp.key;
    ssl_protocols TLSv1.2 TLSv1.3;
    ssl_ciphers HIGH:!aNULL:!MD5;

    root /var/www/myapp;
    index default.asp index.asp;

    location ~ \.asp$ {
        fastcgi_pass axonasp_fcgi;
        # CRITICAL: Pass DOCUMENT_ROOT for multi-host support
        fastcgi_param DOCUMENT_ROOT $document_root;
        fastcgi_param SCRIPT_NAME $fastcgi_script_name;
        fastcgi_param SCRIPT_FILENAME $document_root$fastcgi_script_name;
        fastcgi_param REQUEST_METHOD $request_method;
        fastcgi_param QUERY_STRING $query_string;
        fastcgi_param SERVER_NAME $server_name;
        fastcgi_param SERVER_PORT $server_port;
        fastcgi_param REQUEST_URI $request_uri;
        fastcgi_param DOCUMENT_URI $document_uri;
        fastcgi_param HTTPS $https;
        include fastcgi_params;
    }

    location / {
        try_files $uri $uri/ =404;
    }
}
```

### Multiple Virtual Hosts (Single FastCGI Process)

The primary advantage of AxonASP FastCGI is supporting multiple virtual hosts with different document roots from a **single FastCGI server process**. This architecture is identical to PHP-FPM:

```nginx
upstream axonasp_fcgi {
    server localhost:9000 max_fails=3 fail_timeout=30s;
}

# Virtual Host #1
server {
    listen 443 ssl http2;
    server_name site1.example.com;
    root /var/www/site1;

    ssl_certificate /etc/ssl/certs/site1.crt;
    ssl_certificate_key /etc/ssl/private/site1.key;

    location ~ \.asp$ {
        fastcgi_pass axonasp_fcgi;
        fastcgi_param DOCUMENT_ROOT $document_root;
        fastcgi_param SCRIPT_NAME $fastcgi_script_name;
        fastcgi_param SCRIPT_FILENAME $document_root$fastcgi_script_name;
        include fastcgi_params;
    }
}

# Virtual Host #2
server {
    listen 443 ssl http2;
    server_name site2.example.com;
    root /var/www/site2;

    ssl_certificate /etc/ssl/certs/site2.crt;
    ssl_certificate_key /etc/ssl/private/site2.key;

    location ~ \.asp$ {
        fastcgi_pass axonasp_fcgi;
        fastcgi_param DOCUMENT_ROOT $document_root;
        fastcgi_param SCRIPT_NAME $fastcgi_script_name;
        fastcgi_param SCRIPT_FILENAME $document_root$fastcgi_script_name;
        include fastcgi_params;
    }
}
```

How it works:
1. Nginx receives request for `site1.example.com/index.asp`
2. Nginx sets `$document_root = /var/www/site1` and passes it to FastCGI
3. AxonASP FastCGI receives `DOCUMENT_ROOT=/var/www/site1` parameter
4. AxonASP loads `/var/www/site1/index.asp` (not from the configured RootDir)

This enables true multi-tenant ASP hosting from a single FastCGI process.

## Key FastCGI Parameters

| Parameter | Description | Required |
|-----------|-------------|----------|
| `DOCUMENT_ROOT` | Directory containing the virtual host's files | Yes (for multi-host) |
| `SCRIPT_NAME` | Virtual path to the ASP script | Yes (for proper path resolution) |
| `SCRIPT_FILENAME` | Absolute file path (informational, not used by AxonASP) | No |
| `REQUEST_METHOD` | HTTP method (GET, POST, etc.) | Yes |
| `QUERY_STRING` | URL query parameters | Yes |
| `SERVER_NAME` | Virtual host name | Yes |
| `SERVER_PORT` | HTTP/HTTPS port | Yes |
| `HTTPS` | "on" if HTTPS, "off" if HTTP | Yes |



## Apache with FastCGI

Apache with `mod_fcgid` or `mod_proxy_fcgi` automatically passes FastCGI CGI environment variables, including `DOCUMENT_ROOT` and `SCRIPT_NAME`, to the FastCGI application.

### mod_fcgid Configuration

```apache
<VirtualHost *:443>
    ServerName myapp.example.com
    DocumentRoot "/var/www/myapp"

    SSLEngine on
    SSLCertificateFile /etc/ssl/certs/myapp.crt
    SSLCertificateKeyFile /etc/ssl/private/myapp.key

    <IfModule mod_fcgid.c>
        FcgidConnectTimeout 30
        FcgidIdleTimeout 300
        FcgidMaxRequestLen 1073741824
    </IfModule>

    <FilesMatch "\.asp$">
        SetHandler fcgid-script
    </FilesMatch>

    FcgidWrapper "unix:/var/run/axonasp.sock" .asp
    # or FcgidWrapper "127.0.0.1:9000" .asp
</VirtualHost>
```

### mod_proxy_fcgi Configuration

```apache
<VirtualHost *:443>
    ServerName myapp.example.com
    DocumentRoot "/var/www/myapp"

    SSLEngine on
    SSLCertificateFile /etc/ssl/certs/myapp.crt
    SSLCertificateKeyFile /etc/ssl/private/myapp.key

    <FilesMatch "\.asp$">
        SetHandler "proxy:unix:/var/run/axonasp.sock|fcgi://localhost/"
        # or: SetHandler "proxy:fcgi://127.0.0.1:9000"
    </FilesMatch>
</VirtualHost>
```

Enable required modules:

```bash
a2enmod fcgid ssl
# or for proxy_fcgi:
a2enmod proxy proxy_fcgi ssl
systemctl restart apache2
```

Apache automatically sets `DOCUMENT_ROOT` to the `DocumentRoot` directive value and `SCRIPT_NAME` to the requested script path, enabling seamless multi-host support.

## IIS with FastCGI

AxonASP does not natively support IIS FastCGI, you must use the HttpPlatformHandler to proxy requests to the AxonASP default server. This is the way to integrate with IIS while maintaining full feature parity, as the FastCGI protocol is not natively supported on Windows because of lack of support for named pipes and other Windows IPC mechanisms in the FastCGI go implementation. Check the IIS Support documentation for more details.

## Unix Socket Mode

On Linux and macOS you can use a Unix domain socket instead of a TCP port for lower overhead:

```toml
[fastcgi]
server_port = "unix:/tmp/axonasp.sock"
```

Configure Nginx to connect via the socket:

```nginx
upstream axonasp_fcgi {
    server unix:/tmp/axonasp.sock;
}
```

## Remarks

- The FastCGI server supports the same ASP libraries and functions as the HTTP server. Feature parity is maintained between all runtime modes.
- `global.asa` startup lookup can be forced with `--config.global_asa`, or resolved by fallback (`server.web_root`, then current working directory).
- The FastCGI server does not serve static files directly. The front-end web server handles static content. 
- Increase `instanceMaxRequests` or `maxInstances` in IIS for high-traffic deployments.
- For Nginx, use `fastcgi_read_timeout` to match the AxonASP script timeout configured in `axonasp.toml`.
