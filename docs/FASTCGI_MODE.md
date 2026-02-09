# AxonASP FastCGI Mode

## Overview

AxonASP can run in two modes:

1. **Standalone HTTP Server** - The default mode running on port 4050 (axonasp.exe)
2. **FastCGI Application Server** - For integration with web servers like nginx, Apache, IIS (axonaspcgi.exe)

FastCGI mode allows you to leverage existing web server infrastructure while using AxonASP as the ASP Classic application processor, similar to how PHP-FPM works with PHP.

## Building FastCGI Executable

```powershell
# Build the FastCGI application server
go build -o axonaspcgi.exe ./axonaspcgi
```

This creates `axonaspcgi.exe` which is separate from the main `axonasp.exe` standalone server.

## FastCGI Server Configuration

### Command Line Options

```powershell
axonaspcgi.exe [options]
```

Available options:

- `-listen` - FastCGI listen address (default: `127.0.0.1:9000`)
  - TCP socket: `127.0.0.1:9000` or `:9000`
  - Unix socket: `unix:/var/run/axonasp.sock`
- `-root` - Web root directory (default: `./www`)
- `-timezone` - Server timezone (default: `America/Sao_Paulo`)
- `-default` - Default page (default: `default.asp`)
- `-timeout` - Script execution timeout in seconds (default: `30`)
- `-debug` - Enable ASP debug mode (default: `false`)

### Environment Variables

Alternatively, you can configure using a `.env` file or environment variables:

```env
FCGI_LISTEN=127.0.0.1:9000
WEB_ROOT=./www
TIMEZONE=America/Sao_Paulo
DEFAULT_PAGE=default.asp
SCRIPT_TIMEOUT=30
DEBUG_ASP=false
```

### Starting the FastCGI Server

#### TCP Socket (Network)
```powershell
# Listen on localhost port 9000
.\axonaspcgi.exe -listen 127.0.0.1:9000 -root ./www

# Listen on all interfaces port 9000
.\axonaspcgi.exe -listen :9000 -root ./www
```

#### Unix Socket (Linux/macOS)
```bash
# Listen on Unix socket
./axonaspcgi -listen unix:/var/run/axonasp.sock -root /var/www/asp
```

## Web Server Integration

### Nginx Configuration

#### Basic FastCGI Configuration

```nginx
server {
    listen 80;
    server_name example.com;
    root /var/www/asp;
    index default.asp;

    location ~ \.(asp|aspx)$ {
        fastcgi_pass   127.0.0.1:9000;
        fastcgi_index  default.asp;
        
        fastcgi_param  SCRIPT_FILENAME  $document_root$fastcgi_script_name;
        fastcgi_param  QUERY_STRING     $query_string;
        fastcgi_param  REQUEST_METHOD   $request_method;
        fastcgi_param  CONTENT_TYPE     $content_type;
        fastcgi_param  CONTENT_LENGTH   $content_length;
        
        fastcgi_param  SCRIPT_NAME      $fastcgi_script_name;
        fastcgi_param  REQUEST_URI      $request_uri;
        fastcgi_param  DOCUMENT_URI     $document_uri;
        fastcgi_param  DOCUMENT_ROOT    $document_root;
        fastcgi_param  SERVER_PROTOCOL  $server_protocol;
        
        fastcgi_param  GATEWAY_INTERFACE  CGI/1.1;
        fastcgi_param  SERVER_SOFTWARE    nginx/$nginx_version;
        fastcgi_param  REMOTE_ADDR        $remote_addr;
        fastcgi_param  REMOTE_PORT        $remote_port;
        fastcgi_param  SERVER_ADDR        $server_addr;
        fastcgi_param  SERVER_PORT        $server_port;
        fastcgi_param  SERVER_NAME        $server_name;
        
        fastcgi_param  HTTPS              $https if_not_empty;
    }

    # Serve static files directly
    location ~* \.(jpg|jpeg|png|gif|ico|css|js|svg|woff|woff2|ttf|eot)$ {
        expires 30d;
        add_header Cache-Control "public, immutable";
    }
}
```

#### Unix Socket Configuration

```nginx
server {
    listen 80;
    server_name example.com;
    root /var/www/asp;
    
    location ~ \.(asp|aspx)$ {
        fastcgi_pass   unix:/var/run/axonasp.sock;
        include        fastcgi_params;
        fastcgi_param  SCRIPT_FILENAME  $document_root$fastcgi_script_name;
    }
}
```

#### Advanced Configuration with URL Rewriting

```nginx
server {
    listen 80;
    server_name example.com;
    root /var/www/asp;
    index default.asp;

    # Try to serve file directly, fallback to ASP routing
    location / {
        try_files $uri $uri/ @asp;
    }

    # ASP routing
    location @asp {
        rewrite ^/(.*)$ /router.asp?path=$1 last;
    }

    # Process ASP files
    location ~ \.(asp|aspx)$ {
        fastcgi_pass   127.0.0.1:9000;
        include        fastcgi_params;
        fastcgi_param  SCRIPT_FILENAME  $document_root$fastcgi_script_name;
        fastcgi_param  PATH_INFO        $fastcgi_path_info;
    }
}
```

### Apache Configuration

#### Using mod_proxy_fcgi (Apache 2.4+)

```apache
<VirtualHost *:80>
    ServerName example.com
    DocumentRoot /var/www/asp
    
    # Enable FastCGI proxy
    <FilesMatch "\.asp$">
        SetHandler "proxy:fcgi://127.0.0.1:9000"
    </FilesMatch>
    
    # Pass necessary environment variables
    SetEnvIf Request_URI "\.asp$" proxy-fcgi-pathinfo=true
</VirtualHost>
```

#### Using mod_fcgid

```apache
<VirtualHost *:80>
    ServerName example.com
    DocumentRoot /var/www/asp
    
    <Directory /var/www/asp>
        Options +ExecCGI
        AddHandler fcgid-script .asp
        FCGIWrapper "/usr/local/bin/axonaspcgi" .asp
    </Directory>
</VirtualHost>
```

### IIS Configuration (Windows)

1. Install FastCGI module for IIS if not already installed
2. Open IIS Manager
3. Select your site → Handler Mappings
4. Add Module Mapping:
   - Request path: `*.asp`
   - Module: `FastCgiModule`
   - Executable: `C:\path\to\axonaspcgi.exe`
   - Name: `ASP FastCGI`
5. Configure FastCGI settings in `applicationHost.config`:

```xml
<configuration>
    <system.webServer>
        <fastCgi>
            <application fullPath="C:\path\to\axonaspcgi.exe"
                         arguments="-listen 127.0.0.1:9000 -root C:\inetpub\wwwroot\asp"
                         maxInstances="4"
                         idleTimeout="300"
                         activityTimeout="30"
                         requestTimeout="90"
                         instanceMaxRequests="10000"
                         protocol="NamedPipe"
                         flushNamedPipe="false" />
        </fastCgi>
    </system.webServer>
</configuration>
```

## Features and Compatibility

### Supported Features

All AxonASP features work in FastCGI mode:

- ✅ Classic ASP Objects (Request, Response, Server, Session, Application)
- ✅ Server-side includes (SSI)
- ✅ Session management (via cookies)
- ✅ Application state
- ✅ Global.asa support (Application_OnStart, Session_OnStart, etc.)
- ✅ All G3 custom libraries (G3JSON, G3FILES, G3HTTP, etc.)
- ✅ Standard COM objects (ADODB, Scripting, MSXML2)
- ✅ File uploads
- ✅ Database connectivity
- ✅ VBScript execution

### Session Handling

Sessions work identically to standalone mode:
- Session data stored in `temp/session/` directory
- Session cookie name: `ASPSESSIONID`
- Default timeout: 20 minutes
- Automatic cleanup of expired sessions

### Application State

Application-level state is maintained within the FastCGI process:
- Shared across all requests to the same FastCGI process
- `Application_OnStart` executed when FastCGI server starts
- `Application_OnEnd` executed when FastCGI server stops

**Note**: When running multiple FastCGI worker processes, Application state is NOT shared between workers. Use external storage (database, Redis, etc.) for shared state in multi-process deployments.

## Performance Optimization

### Process Pool

Configure your web server to maintain a pool of FastCGI processes:

**Nginx:**
```nginx
fastcgi_pass   127.0.0.1:9000;
fastcgi_buffering on;
fastcgi_buffer_size 16k;
fastcgi_buffers 4 16k;
```

**Apache:**
```apache
FCGIWrapper "/usr/local/bin/axonaspcgi -listen :9000" virtual
MaxProcessCount 4
MinProcessCount 1
```

### Nginx Caching

Cache ASP output for frequently accessed pages:

```nginx
fastcgi_cache_path /var/cache/nginx/axonasp levels=1:2 keys_zone=ASP_CACHE:10m inactive=60m;
fastcgi_cache_key "$scheme$request_method$host$request_uri";

location ~ \.(asp|aspx)$ {
    fastcgi_pass 127.0.0.1:9000;
    include fastcgi_params;
    fastcgi_param SCRIPT_FILENAME $document_root$fastcgi_script_name;
    
    # Enable caching
    fastcgi_cache ASP_CACHE;
    fastcgi_cache_valid 200 60m;
    fastcgi_cache_bypass $arg_nocache;
    add_header X-Cache-Status $upstream_cache_status;
}
```

### Multiple Workers

Run multiple FastCGI processes for load balancing:

```powershell
# Windows - Start multiple instances on different ports
Start-Process axonaspcgi.exe -ArgumentList "-listen :9001 -root ./www"
Start-Process axonaspcgi.exe -ArgumentList "-listen :9002 -root ./www"
Start-Process axonaspcgi.exe -ArgumentList "-listen :9003 -root ./www"
```

```nginx
# Nginx upstream configuration
upstream axonasp_backend {
    server 127.0.0.1:9001;
    server 127.0.0.1:9002;
    server 127.0.0.1:9003;
}

server {
    location ~ \.(asp|aspx)$ {
        fastcgi_pass axonasp_backend;
        include fastcgi_params;
        fastcgi_param SCRIPT_FILENAME $document_root$fastcgi_script_name;
    }
}
```

## Comparison: Standalone vs FastCGI

| Feature | Standalone Mode | FastCGI Mode |
|---------|----------------|--------------|
| Web Server | Built-in | External (nginx, Apache, IIS) |
| Port | 4050 (default) | 9000 (default) or Unix socket |
| HTTPS | Requires proxy | Native support via web server |
| Static Files | Served by AxonASP | Served by web server (faster) |
| Load Balancing | Manual setup | Native via upstream |
| Reverse Proxy | Required for production | Built-in |
| Process Management | Manual | Web server handles |
| Restart | Manual | Graceful via web server |
| Virtual Hosts | Single site | Multiple sites via web server |

## Troubleshooting

### FastCGI Process Not Starting

```powershell
# Check if port is already in use
netstat -an | findstr :9000

# Test FastCGI server manually
.\axonaspcgi.exe -listen :9000 -root ./www -debug true
```

### 502 Bad Gateway (Nginx)

Check FastCGI process is running and accessible:

```bash
# Test connection to FastCGI port
telnet 127.0.0.1 9000

# Check nginx error log
tail -f /var/log/nginx/error.log
```

### Permission Denied (Unix Socket)

```bash
# Set correct permissions on socket
chmod 666 /var/run/axonasp.sock

# Or configure web server user to match socket owner
chown www-data:www-data /var/run/axonasp.sock
```

### Session Not Working

Ensure session directory is writable:

```powershell
# Windows
icacls temp\session /grant Everyone:(OI)(CI)F

# Linux
chmod 777 temp/session
```

### Application_OnStart Not Executing

Verify global.asa is in web root:

```powershell
# Should exist: ./www/global.asa
ls www/global.asa

# Check FastCGI startup logs
.\axonaspcgi.exe -listen :9000 -root ./www
# Look for: "global.asa loaded successfully"
# and "Application_OnStart executed successfully"
```

## Example Service Configuration

### Systemd Service (Linux)

Create `/etc/systemd/system/axonaspcgi.service`:

```ini
[Unit]
Description=AxonASP FastCGI Server
After=network.target

[Service]
Type=simple
User=www-data
Group=www-data
WorkingDirectory=/var/www/asp
ExecStart=/usr/local/bin/axonaspcgi -listen unix:/var/run/axonasp.sock -root /var/www/asp
Restart=always
RestartSec=10

[Install]
WantedBy=multi-user.target
```

Enable and start:

```bash
sudo systemctl daemon-reload
sudo systemctl enable axonaspcgi
sudo systemctl start axonaspcgi
sudo systemctl status axonaspcgi
```

### Windows Service (NSSM)

Using [NSSM](https://nssm.cc/) - Non-Sucking Service Manager:

```powershell
# Install NSSM
choco install nssm

# Create service
nssm install AxonASPFastCGI "C:\AxonASP\axonaspcgi.exe"
nssm set AxonASPFastCGI AppParameters "-listen :9000 -root C:\inetpub\wwwroot\asp"
nssm set AxonASPFastCGI AppDirectory "C:\AxonASP"
nssm set AxonASPFastCGI DisplayName "AxonASP FastCGI Server"
nssm set AxonASPFastCGI Description "ASP Classic application server via FastCGI"
nssm set AxonASPFastCGI Start SERVICE_AUTO_START

# Start service
nssm start AxonASPFastCGI

# Check status
nssm status AxonASPFastCGI
```

## Development and Testing

### Testing FastCGI Locally

1. Start FastCGI server:
```powershell
.\axonaspcgi.exe -listen :9000 -root ./www -debug true
```

2. Test with cgi-fcgi tool:
```bash
# Install fcgi tools
apt-get install libfcgi0ldbl

# Test request
SCRIPT_FILENAME=./www/default.asp \
QUERY_STRING="" \
REQUEST_METHOD=GET \
cgi-fcgi -bind -connect 127.0.0.1:9000
```

### Using with Docker

```dockerfile
FROM golang:1.21 AS builder
WORKDIR /app
COPY . .
RUN go build -o axonaspcgi ./axonaspcgi

FROM debian:bullseye-slim
RUN apt-get update && apt-get install -y ca-certificates && rm -rf /var/lib/apt/lists/*
COPY --from=builder /app/axonaspcgi /usr/local/bin/
COPY ./www /var/www/asp
EXPOSE 9000
CMD ["axonaspcgi", "-listen", ":9000", "-root", "/var/www/asp"]
```

Docker Compose with Nginx:

```yaml
version: '3.8'
services:
  axonaspcgi:
    build: .
    ports:
      - "9000:9000"
    volumes:
      - ./www:/var/www/asp
      - ./temp:/app/temp
    environment:
      - WEB_ROOT=/var/www/asp
      - SCRIPT_TIMEOUT=30
      
  nginx:
    image: nginx:alpine
    ports:
      - "80:80"
    volumes:
      - ./nginx.conf:/etc/nginx/nginx.conf
      - ./www:/var/www/asp
    depends_on:
      - axonaspcgi
```

## Conclusion

FastCGI mode provides a production-ready deployment option for AxonASP applications, allowing integration with enterprise web servers while maintaining full compatibility with ASP Classic functionality. Choose FastCGI mode when you need advanced web server features, better performance, or integration with existing infrastructure.
