# Reverse Proxy Configuration

## Overview

The recommended way to deploy AxonASP in production is behind a reverse proxy. The proxy handles TLS termination, rate limiting, static asset caching, and load balancing, while AxonASP focuses on executing ASP scripts. AxonASP HTTP server listens on port 8801 by default.

## Prerequisites

- AxonASP HTTP server (`axonasp-http.exe`) running and reachable on localhost
- A reverse proxy (Nginx, Apache, Caddy, or IIS) installed on the same machine or network

**Start AxonASP:**

```powershell
.\axonasp-http.exe
```

## Why Not Expose AxonASP Directly

Do not expose `axonasp-http.exe` directly to public internet traffic:

- No centralized TLS/SSL termination
- No rate-limiting or DDoS protection
- No centralized request logging
- Increased attack surface

Use a reverse proxy as the public-facing entry point.

## Nginx (Proxy Mode)

```nginx
upstream axonasp_backend {
    server localhost:8801 max_fails=3 fail_timeout=30s;
}

server {
    listen 80;
    server_name myapp.example.com;
    return 301 https://$server_name$request_uri;
}

server {
    listen 443 ssl http2;
    server_name myapp.example.com;

    ssl_certificate /etc/ssl/certs/myapp.crt;
    ssl_certificate_key /etc/ssl/private/myapp.key;
    ssl_protocols TLSv1.2 TLSv1.3;
    ssl_ciphers HIGH:!aNULL:!MD5;

    location / {
        proxy_pass http://axonasp_backend;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
        proxy_buffering off;
        proxy_request_buffering off;
    }
}
```

## Apache (Proxy Mode)

```apache
<VirtualHost *:80>
    ServerName myapp.example.com
    RewriteEngine On
    RewriteCond %{HTTPS} off
    RewriteRule ^(.*)$ https://%{HTTP_HOST}%{REQUEST_URI} [L,R=301]
</VirtualHost>

<VirtualHost *:443>
    ServerName myapp.example.com

    SSLEngine on
    SSLCertificateFile /etc/ssl/certs/myapp.crt
    SSLCertificateKeyFile /etc/ssl/private/myapp.key
    SSLProtocol TLSv1.2 TLSv1.3

    ProxyRequests Off
    ProxyPreserveHost On

    <Location />
        ProxyPass http://localhost:8801/
        ProxyPassReverse http://localhost:8801/
        RequestHeader set X-Real-IP "%{REMOTE_ADDR}s"
        RequestHeader set X-Forwarded-For "%{HTTP:X-Forwarded-For}s, %{REMOTE_ADDR}s"
        RequestHeader set X-Forwarded-Proto "%{REQUEST_SCHEME}s"
    </Location>
</VirtualHost>
```

Enable required modules:

```bash
a2enmod proxy proxy_http ssl rewrite headers
systemctl restart apache2
```

## Caddy (Proxy Mode)

Caddy handles HTTPS certificates automatically through Let's Encrypt:

```caddyfile
myapp.example.com {
    reverse_proxy localhost:8801 {
        header_up X-Forwarded-Proto {scheme}
        header_up X-Forwarded-For {remote_host}
        header_up X-Real-IP {remote_host}
    }
}
```

Start Caddy:

```bash
caddy run
```

## IIS (Proxy Mode)

Configure IIS Application Request Routing to forward requests to AxonASP:

```xml
<!-- web.config on the IIS site -->
<configuration>
  <system.webServer>
    <rewrite>
      <rules>
        <rule name="Proxy to AxonASP" stopProcessing="true">
          <match url="(.*)" />
          <action type="Rewrite" url="http://localhost:8801/{R:1}" />
        </rule>
      </rules>
    </rewrite>
  </system.webServer>
</configuration>
```

Requires IIS Application Request Routing (ARR) and URL Rewrite modules.

## Load Balancing Multiple Instances

For high-traffic deployments, run multiple AxonASP instances on different ports and balance across them:

```nginx
upstream axonasp_cluster {
    server localhost:8801 weight=1;
    server localhost:8802 weight=1;
    server localhost:8803 weight=1;
}

server {
    listen 443 ssl http2;
    server_name myapp.example.com;

    location / {
        proxy_pass http://axonasp_cluster;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

Start each instance on a different port using environment variables:

```powershell
# Terminal 1
$env:SERVER_PORT = "8801"; .\axonasp-http.exe

# Terminal 2
$env:SERVER_PORT = "8802"; .\axonasp-http.exe

# Terminal 3
$env:SERVER_PORT = "8803"; .\axonasp-http.exe
```

## Remarks

- AxonASP does not handle TLS/SSL natively. Terminate SSL at the reverse proxy and forward plain HTTP to AxonASP.
- Set `proxy_buffering off` in Nginx when using `Response.Flush` or streaming responses in ASP scripts.
- The port can also be changed in `config/axonasp.toml` under `[server] server_port`.
- All proxy configurations above assume AxonASP runs on the same machine as the proxy on localhost. Adjust the upstream address for remote deployments.
