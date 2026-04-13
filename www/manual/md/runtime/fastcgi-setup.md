# FastCGI Setup

## Overview

AxonASP provides a FastCGI application server (`axonasp-fastcgi.exe`) that integrates directly with Nginx, Apache, or IIS using the FastCGI protocol. In this mode AxonASP acts as a backend process and the front-end web server handles all HTTP connections, static content, and TLS termination.

AxonASP FastCGI fully supports **multi-host deployments** with different document roots, similar to PHP-FPM. This allows a single FastCGI server process to serve content from multiple virtual hosts, each with its own document root directory.

## Prerequisites

- `axonasp-fastcgi.exe` running and reachable on the configured port (default: 9000)
- Front-end web server with FastCGI support installed (nginx, Apache, IIS)

The FastCGI port is configured in `config/axonasp.toml`:

```toml
[fastcgi]
server_port = 9000
```

The port can also be a socket path on Linux and macOS:

```toml
server_port = "unix:/tmp/axonasp.sock"
```

**Start AxonASP FastCGI:**

```powershell
.\axonasp-fastcgi.exe
```

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

Configure IIS to use AxonASP as a FastCGI handler for `.asp` files:

```powershell
# Run as Administrator
Import-Module WebAdministration

# Add FastCGI application
Add-WebConfiguration -PSPath "IIS:\" `
    -Filter "system.webServer/fastCgi" `
    -Value @{
        fullPath = "C:\axonasp\axonasp-fastcgi.exe";
        arguments = "";
        instanceMaxRequests = 10000;
        maxInstances = 4;
        requestTimeout = 300;
    }

# Map .asp extension to the FastCGI handler
Add-WebConfigurationProperty `
    -PSPath "IIS:\Sites\Default Web Site" `
    -Filter "system.webServer/handlers" `
    -Name "." `
    -Value @{
        name = "AxonASP-FastCGI";
        path = "*.asp";
        verb = "*";
        modules = "FastCgiModule";
        scriptProcessor = "C:\axonasp\axonasp-fastcgi.exe";
        resourceType = "File";
        requireAccess = "Read";
    }

iisreset
```

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
- `global.asa` is loaded from the configured web root on startup, identical to the HTTP server, you should set the `server.web_root` key to the parent directory of `global.asa` for it to be loaded properly.
- The FastCGI server does not serve static files directly. The front-end web server handles static content. 
- Increase `instanceMaxRequests` or `maxInstances` in IIS for high-traffic deployments.
- For Nginx, use `fastcgi_read_timeout` to match the AxonASP script timeout configured in `axonasp.toml`.
