# FastCGI Setup

## Overview

AxonASP provides a FastCGI application server (`axonasp-fastcgi.exe`) that integrates directly with Nginx, Apache, or IIS using the FastCGI protocol. In this mode AxonASP acts as a backend process and the front-end web server handles all HTTP connections, static content, and TLS termination.

## Prerequisites

- `axonasp-fastcgi.exe` running and reachable on the configured port (default: 9000)
- Front-end web server with FastCGI support installed

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

## Nginx with FastCGI

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

## Apache with FastCGI

Requires `mod_fcgid` or `mod_proxy_fcgi`:

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

    FcgidWrapper "/path/to/axonasp-fastcgi.exe" .asp
</VirtualHost>
```

Enable required modules:

```bash
a2enmod fcgid ssl
systemctl restart apache2
```

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
- `global.asa` is loaded from the configured web root on startup, identical to the HTTP server.
- The FastCGI server does not serve static files directly. The front-end web server handles static content.
- Increase `instanceMaxRequests` or `maxInstances` in IIS for high-traffic deployments.
- For Nginx, use `fastcgi_read_timeout` to match the AxonASP script timeout configured in `axonasp.toml`.
