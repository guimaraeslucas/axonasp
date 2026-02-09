# AxonASP FastCGI Quick Start Guide

This guide will help you quickly test and run AxonASP in FastCGI mode.

## Quick Test with Built-in Test Server

For testing purposes, we'll use the FastCGI server directly with a simple test:

1. **Build the FastCGI executable**:
```powershell
go build -o axonaspcgi.exe ./axonaspcgi
```

2. **Start the FastCGI server**:
```powershell
.\axonaspcgi.exe -listen :9000 -root ./www -debug true
```

3. **Test the server** (requires fcgi-fcgi or nginx setup)

## Testing with Nginx (Recommended)

### 1. Install Nginx

**Windows**: Download from https://nginx.org/en/download.html

**Linux**:
```bash
sudo apt-get install nginx
```

**macOS**:
```bash
brew install nginx
```

### 2. Create Nginx Configuration

Create a file `nginx-axonasp.conf`:

```nginx
server {
    listen 8080;
    server_name localhost;
    root /path/to/axonasp/www;
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
    }

    location ~* \.(jpg|jpeg|png|gif|ico|css|js)$ {
        expires 30d;
    }
}
```

### 3. Start Everything

**Terminal 1** - Start AxonASP FastCGI:
```powershell
.\axonaspcgi.exe -listen :9000 -root ./www
```

**Terminal 2** - Start Nginx:
```powershell
# Windows
nginx.exe -c nginx-axonasp.conf

# Linux/macOS
sudo nginx -c /path/to/nginx-axonasp.conf
```

### 4. Test It!

Open your browser and go to:
- http://localhost:8080/tests/test_fastcgi.asp

You should see the FastCGI test page with all tests passing!

## Testing with IIS (Windows)

### 1. Prerequisites

- Windows Server or Windows 10/11 with IIS installed
- IIS FastCGI module (usually pre-installed with IIS)

**Install IIS with FastCGI support:**
```powershell
# Run as Administrator
Enable-WindowsOptionalFeature -Online -FeatureName IIS-WebServerRole
Enable-WindowsOptionalFeature -Online -FeatureName IIS-CGI
```

Or via Server Manager:
- Open Server Manager → Add Roles and Features
- Select Web Server (IIS) → Application Development → CGI

### 2. Build and Start AxonASP FastCGI

```powershell
# Build the FastCGI executable
go build -o axonaspcgi.exe ./axonaspcgi

# Start FastCGI server on port 9000
.\axonaspcgi.exe -listen :9000 -root C:\inetpub\wwwroot\asp
```

Keep this terminal running in the background.

### 3. Configure IIS FastCGI Application

**Option A: Using IIS Manager (GUI)**

1. Open IIS Manager (`inetmgr.exe`)
2. Select your server in the left panel
3. Double-click **FastCGI Settings**
4. Click **Add Application...** in the right panel
5. Configure:
   - **Full Path**: `C:\path\to\axonaspcgi.exe`
   - **Arguments**: `-listen :9000 -root C:\inetpub\wwwroot\asp`
   - **Monitor changes to file**: (leave empty)
   - **Instance MaxRequests**: `10000`
   - **Activity Timeout**: `30`
   - **Request Timeout**: `90`
6. Click **OK**

**Option B: Using Command Line**

```powershell
# Add FastCGI application
%windir%\system32\inetsrv\appcmd.exe set config -section:system.webServer/fastCgi /+"[fullPath='C:\AxonASP\axonaspcgi.exe',arguments='-listen :9000 -root C:\inetpub\wwwroot\asp',maxInstances='4',idleTimeout='300',activityTimeout='30',requestTimeout='90',instanceMaxRequests='10000',protocol='NamedPipe',flushNamedPipe='false']" /commit:apphost
```

### 4. Add Handler Mapping

**Option A: Using IIS Manager (GUI)**

1. In IIS Manager, select your website (e.g., "Default Web Site")
2. Double-click **Handler Mappings**
3. Click **Add Module Mapping...** in the right panel
4. Configure:
   - **Request path**: `*.asp`
   - **Module**: `FastCgiModule`
   - **Executable**: `C:\path\to\axonaspcgi.exe|-listen :9000 -root C:\inetpub\wwwroot\asp`
   - **Name**: `ASP via AxonASP FastCGI`
5. Click **Request Restrictions...**
   - **Invoke handler only if request is mapped to**: File or folder
   - Click **OK**
6. Click **OK** and confirm the prompt

**Option B: Using web.config**

Create or edit `C:\inetpub\wwwroot\asp\web.config`:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<configuration>
    <system.webServer>
        <handlers>
            <add name="ASP-AxonASP" path="*.asp" verb="*" 
                 modules="FastCgiModule" 
                 scriptProcessor="C:\AxonASP\axonaspcgi.exe|-listen :9000 -root C:\inetpub\wwwroot\asp"
                 resourceType="Either" 
                 requireAccess="Script" />
        </handlers>
    </system.webServer>
</configuration>
```

**Option C: Using Command Line**

```powershell
# Add handler mapping
%windir%\system32\inetsrv\appcmd.exe set config "Default Web Site" -section:system.webServer/handlers /+"[name='ASP-AxonASP',path='*.asp',verb='*',modules='FastCgiModule',scriptProcessor='C:\AxonASP\axonaspcgi.exe|-listen :9000 -root C:\inetpub\wwwroot\asp',resourceType='Either']" /commit:apphost
```

### 5. Set Permissions

Ensure the IIS application pool identity has access to your files:

```powershell
# Grant permissions to IIS_IUSRS group
icacls "C:\inetpub\wwwroot\asp" /grant "IIS_IUSRS:(OI)(CI)F" /T
icacls "C:\AxonASP" /grant "IIS_IUSRS:(OI)(CI)RX" /T
```

### 6. Test IIS Configuration

1. Copy your ASP files to `C:\inetpub\wwwroot\asp\`
2. Make sure `axonaspcgi.exe` is running in a terminal
3. Open your browser and navigate to:
   - http://localhost/test_fastcgi.asp
   - http://localhost/default.asp

### 7. Running as Windows Service (Production)

For production, run axonaspcgi as a Windows Service using NSSM:

```powershell
# Install NSSM (Non-Sucking Service Manager)
choco install nssm

# Create Windows Service
nssm install AxonASPFastCGI "C:\AxonASP\axonaspcgi.exe"
nssm set AxonASPFastCGI AppParameters "-listen :9000 -root C:\inetpub\wwwroot\asp"
nssm set AxonASPFastCGI AppDirectory "C:\AxonASP"
nssm set AxonASPFastCGI DisplayName "AxonASP FastCGI Server"
nssm set AxonASPFastCGI Description "ASP Classic application server via FastCGI"
nssm set AxonASPFastCGI Start SERVICE_AUTO_START

# Start the service
nssm start AxonASPFastCGI

# Check status
nssm status AxonASPFastCGI
```

### Troubleshooting IIS

**"500 Internal Server Error"**
- Check IIS logs: `C:\inetpub\logs\LogFiles\W3SVC1\`
- Verify axonaspcgi.exe is running: `Get-Process axonaspcgi`
- Check FastCGI settings in IIS Manager

**"Cannot connect to FastCGI"**
- Verify port 9000 is open: `netstat -an | findstr :9000`
- Check Windows Firewall settings
- Ensure axonaspcgi.exe started successfully

**Permission denied errors**
- Run `icacls` commands above to grant proper permissions
- Check Application Pool identity has access to files

**ASP files download instead of executing**
- Verify handler mapping is correct
- Ensure `*.asp` is mapped to FastCgiModule
- Restart IIS: `iisreset`

## Testing Without Web Server

You can test the FastCGI protocol directly using curl and fcgi:

```bash
# Install cgi-fcgi (Linux)
sudo apt-get install libfcgi0ldbl

# Test request
SCRIPT_FILENAME=./www/tests/test_fastcgi.asp \
QUERY_STRING="" \
REQUEST_METHOD=GET \
cgi-fcgi -bind -connect 127.0.0.1:9000
```

## Troubleshooting

### "Connection refused" error
- Make sure axonaspcgi.exe is running on port 9000
- Check firewall settings
- Verify the port with: `netstat -an | findstr :9000` (Windows) or `lsof -i :9000` (Linux)

### "502 Bad Gateway" from Nginx
- Verify FastCGI server is running: `netstat -an | findstr :9000` (Windows) or `lsof -i :9000` (Linux)
- Check nginx error log for details
- Ensure nginx has permission to connect to the port

### "500 Internal Server Error" from IIS
- Check IIS logs in `C:\inetpub\logs\LogFiles\`
- Verify FastCGI application is configured correctly
- Ensure axonaspcgi.exe has proper permissions
- Check Event Viewer for application errors

### Session not working
- Ensure `temp/session` directory exists and is writable
- Check that cookies are enabled in your browser

### "File not found" errors
- Verify the `-root` parameter points to the correct directory
- Check that ASP files exist in the specified root directory

## Next Steps

Once you've verified FastCGI mode works:

1. Read the full [FastCGI Mode Documentation](FASTCGI_MODE.md)
2. Configure for production deployment
   - **Nginx/Apache**: Set up SSL certificates with Let's Encrypt
   - **IIS**: Configure Application Pool settings and SSL bindings
3. Set up load balancing with multiple workers
4. Implement caching strategies
5. Configure monitoring and logging
6. Set up automatic restarts (systemd on Linux, Windows Service on Windows)

## Configuration Comparison

| Feature | Nginx | Apache | IIS |
|---------|-------|--------|-----|
| **Setup Complexity** | Medium | Medium | Easy (GUI) |
| **Performance** | Excellent | Good | Good |
| **Static Files** | Excellent | Good | Good |
| **SSL/TLS** | Easy | Easy | Easy |
| **Load Balancing** | Built-in | With modules | ARR module |
| **Best For** | Linux production | Linux/Windows | Windows shops |
| **Configuration** | Config file | Config file | GUI or web.config |

## Need Help?

- GitHub Issues: https://github.com/guimaraeslucas/axonasp/issues
- Documentation: https://g3pix.com.br/axonasp
