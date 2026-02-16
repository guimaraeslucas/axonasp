<p align="center">
  <img src="errorpages/axonasp.svg" alt="G3Pix AxonASP Logo" width="400"/>
</p>

<h1 align="center">G3Pix AxonASP</h1>

<p align="center">
  <strong>High-Performance Classic ASP Runtime Engine</strong>
  <br>
  Run your new and legacy ASP Classic applications with modern speed and cross-platform compatibility
</p>

<p align="center">
  <img src="https://img.shields.io/badge/version-1.0-blue.svg" alt="Version 1.0"/>
  <img src="https://img.shields.io/badge/Go-1.25+-00ADD8.svg" alt="Go Version"/>
  <img src="https://img.shields.io/badge/platform-Windows%20%7C%20Linux%20%7C%20macOS-lightgrey.svg" alt="Platform"/>
  <img src="https://img.shields.io/badge/license-MPL-green.svg" alt="License"/>
</p>

---

## 🚀 Overview

**G3Pix AxonASP** is a cutting-edge, high-performance runtime engine for Classic ASP (Active Server Pages) applications. Built entirely in **GoLang**, it delivers exceptional speed, security, and cross-platform compatibility while maintaining near-perfect backward compatibility with legacy ASP codebase.

Say goodbye to Windows-only hosting and IIS dependencies. G3Pix AxonASP empowers you to run your Classic ASP applications on **Windows**, **Linux**, and **macOS** with **accelerated performance** and **minimal to zero code modifications**.

### ✨ Key Features

- **🔥 Blazing Fast**: Built with GoLang for superior performance compared to traditional IIS hosting
- **🌍 Cross-Platform**: Compile and deploy on Windows, Linux, or macOS - anywhere GoLang runs
- **🔄 Drop-In Replacement**: Run most Classic ASP applications without code modifications
- **🔒 Secure by Design**: Files served only from `www/` root directory, preventing unauthorized access
- **🛠️ Extended Functionality**: 60+ custom functions inspired by PHP for enhanced productivity
- **📦 Full COM Support**: ADODB (with Access database support), MSXML2, Scripting objects, WScript.Shell, ADOX
- **🗄️ Database Support**: SQLite, MySQL, PostgreSQL, MS SQL Server, Microsoft Access (Windows)
- **⚙️ Database Conversion Tool**: Built-in wizard to migrate Access databases to modern formats
- **⚙️ Dual Deployment**: Standalone server or FastCGI mode for nginx/Apache/IIS integration
- **🔌 IIS Compatibility**: web.config support for URL rewriting, redirects, and custom error pages
- **🚀 Production Ready**: Session management, Application state, Global.asa events, graceful shutdown

---

## 📋 Deployment Modes

AxonASP offers two deployment modes for different use cases:

| Mode | Binary | Best For | Details |
|------|--------|----------|---------|
| **Standalone** | `axonasp.exe` | Development, simple deployments, application specific proxy hosting | Built-in HTTP server on port 4050. Quick setup, no dependencies. |
| **FastCGI** | `axonaspcgi.exe` | Production, high-traffic sites, existing infrastructure | Integrates with nginx, Apache, IIS. See [FastCGI Guide](docs/FASTCGI_MODE.md) |

**Quick Start Standalone:**
```bash
go run main.go
# or build: go build -o axonasp.exe
./axonasp.exe
```

**Quick Start FastCGI:**
```bash
go build -o axonaspcgi.exe ./axonaspcgi
./axonaspcgi -listen :9000 -root ./www
# Then configure your web server to proxy .asp files to port 9000
```

**Install Linux (RPM distro - nginx + AxonASP proxy):**
```bash
wget https://raw.githubusercontent.com/guimaraeslucas/axonasp/main/linux_install.sh -O linux_install.sh && sed -i 's/\r$//' linux_install.sh && chmod +x linux_install.sh && bash linux_install.sh
# Then configure your web server to proxy .asp files to port 9000
```

---

## 📦 Installation

### Prerequisites

- **Go 1.25+** installed on your system
- Basic understanding of Classic ASP and VBScript

### Quick Start

1. **Clone the Repository**
   ```bash
   git clone https://github.com/guimaraeslucas/axonasp.git
   cd axonasp
   ```

2. **Install Dependencies**
   ```bash
   go mod download
   ```

3. **Configure Environment** (Optional)
   
   Create a `.env` file in the root directory (view .env.example to more options):
   ```env
   SERVER_PORT=4050
   WEB_ROOT=./www
   TIMEZONE=America/Sao_Paulo
   DEFAULT_PAGE=default.asp
   SCRIPT_TIMEOUT=30
   DEBUG_ASP=FALSE
   ERROR_404_MODE=IIS
   ```

4. **Run the Server**
   ```bash
   go run main.go
   ```

5. **Access Your Application**
   
   Open your browser and navigate to:
   ```
   http://localhost:4050
   ```

---

## 🏗️ Building for Production

### Build for Current Platform
```bash
go build 
or
./build.ps1 (windows - build all executables server/fastcgi)
```

### Cross-Platform Compilation

**Windows (64-bit)**
```bash
GOOS=windows GOARCH=amd64 go build 
```

**Linux (64-bit)**
```bash
GOOS=linux GOARCH=amd64 go build 
```

**macOS (Intel)**
```bash
GOOS=darwin GOARCH=amd64 go build 
```

**macOS (Apple Silicon)**
```bash
GOOS=darwin GOARCH=arm64 go build
```

After building, simply run the executable:
```bash
./axonasp          # Linux/macOS
axonasp.exe        # Windows
```

### FastCGI Mode

AxonASP can run as a FastCGI application server (similar to PHP-FPM), allowing integration with production web servers:

```bash
# Build FastCGI executable
go build -o axonaspcgi.exe ./axonaspcgi

# Run on TCP socket (default: 127.0.0.1:9000)
./axonaspcgi -listen :9000 -root ./www

# Run on Unix socket (Linux/macOS)
./axonaspcgi -listen unix:/var/run/axonasp.sock -root /var/www/asp
```

**FastCGI Mode Benefits:**
- ✅ Native web server integration (nginx, Apache, IIS)
- ✅ Better static file performance (served by web server)
- ✅ Production-grade SSL/TLS handling by server
- ✅ Advanced caching strategies

**When to use FastCGI:**
- Production deployments with high traffic
- Multiple applications behind one web server
- Need for advanced web server features (SSL, caching, compression)
- Integration with existing infrastructure

**Quick nginx Configuration:**
```nginx
server {
    listen 80;
    server_name example.com;
    root /var/www/asp;
    
    location ~ \.(asp|aspx)$ {
        fastcgi_pass   127.0.0.1:9000;
        include        fastcgi_params;
        fastcgi_param  SCRIPT_FILENAME  $document_root$fastcgi_script_name;
    }
}
```

📖 **Complete FastCGI Guide**: [docs/FASTCGI_MODE.md](docs/FASTCGI_MODE.md)  
🚀 **Quick Start**: [docs/FASTCGI_QUICKSTART.md](docs/FASTCGI_QUICKSTART.md)

---

## ⚙️ Configuration

G3Pix AxonASP uses a `.env` file for configuration. All settings are optional with sensible defaults.

### Available Configuration Options

| Variable | Default | Description |
|----------|---------|-------------|
| `SERVER_PORT` | `4050` | HTTP server port |
| `WEB_ROOT` | `./www` | Root directory for ASP files |
| `TIMEZONE` | `America/Sao_Paulo` | Server timezone |
| `DEFAULT_PAGE` | `index.asp,default.asp,...` | Default document hierarchy (comma-separated) |
| `SCRIPT_TIMEOUT` | `30` | Script execution timeout (seconds) |
| `DEBUG_ASP` | `FALSE` | Enable HTML stack traces and detailed logs |
| `CLEAN_SESSIONS` | `TRUE` | Auto-cleanup expired sessions on startup |
| `ASP_CACHE_TYPE` | `disk` | AST cache type: `memory` or `disk` |
| `ASP_CACHE_TTL_MINUTES` | `0` | TTL for AST cache (0 = forever) |
| `AXONASP_VM` | `FALSE` | Enable experimental bytecode VM |
| `VM_CACHE_TYPE` | `disk` | VM bytecode cache type: `memory` or `disk` |
| `VM_CACHE_TTL_MINUTES` | `0` | TTL for VM bytecode cache (0 = forever) |
| `MEMORY_LIMIT_MB` | `0` | Memory limit for the process (0 = no limit) |
| `ERROR_404_MODE` | `DEFAULT` | 404 handling mode: `DEFAULT` or `IIS` |
| `BLOCKED_EXTENSIONS` | (see .env) | Comma-separated list of blocked file extensions |
| `COM_PROVIDER` | `auto` | COM provider selection for Access (`auto` or `code`) |
| `SQL_TRACE` | `FALSE` | Enable verbose SQL tracing |
| `SMTP_HOST` | - | SMTP server hostname |
| `SMTP_PORT` | `587` | SMTP server port |
| `SMTP_USER` | - | SMTP authentication username |
| `SMTP_PASS` | - | SMTP authentication password |
| `SMTP_FROM` | - | Default sender email address |
| `MYSQL_HOST` | `localhost` | MySQL host for G3DB |
| `POSTGRES_HOST` | `localhost` | PostgreSQL host for G3DB |
| `MSSQL_HOST` | `localhost` | MS SQL Server host for G3DB |
| `SQLITE_PATH` | `./database.db` | SQLite database path for G3DB |

---

## 🌐 Reverse Proxy Configuration

G3Pix AxonASP supports **one application per server instance** due to how `global.asa` is loaded. To host multiple applications, run multiple instances on different ports.

### Nginx Configuration

```nginx
# Application 1 on port 4050
server {
    listen 80;
    server_name app1.example.com;

    location / {
        proxy_pass http://localhost:4050;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}

# Application 2 on port 4051
server {
    listen 80;
    server_name app2.example.com;

    location / {
        proxy_pass http://localhost:4051;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

### Apache Configuration

```apache
# Application 1 on port 4050
<VirtualHost *:80>
    ServerName app1.example.com
    
    ProxyPreserveHost On
    ProxyPass / http://localhost:4050/
    ProxyPassReverse / http://localhost:4050/
    
    <Proxy *>
        Order deny,allow
        Allow from all
    </Proxy>
</VirtualHost>

# Application 2 on port 4051
<VirtualHost *:80>
    ServerName app2.example.com
    
    ProxyPreserveHost On
    ProxyPass / http://localhost:4051/
    ProxyPassReverse / http://localhost:4051/
    
    <Proxy *>
        Order deny,allow
        Allow from all
    </Proxy>
</VirtualHost>
```

### IIS Configuration (URL Rewrite)

Install **Application Request Routing (ARR)** and **URL Rewrite** modules, then add to `web.config`:

```xml
<configuration>
    <system.webServer>
        <rewrite>
            <rules>
                <rule name="ReverseProxyToAxonASP" stopProcessing="true">
                    <match url="(.*)" />
                    <action type="Rewrite" url="http://localhost:4050/{R:1}" />
                </rule>
            </rules>
        </rewrite>
    </system.webServer>
</configuration>
```

---

## 🔀 IIS Compatibility Mode

AxonASP supports IIS-style URL rewriting and error handling through `web.config` files.

### Custom 404 Error Handling

**Three modes available** (configure via `ERROR_404_MODE` environment variable):

1. **`default`** - Static HTML error page
2. **`asp`** - Custom ASP error handler page
3. **`iis`** - IIS-style web.config httpErrors section

#### Mode 1: Default (Static HTML)
```env
ERROR_404_MODE=default
```
Serves `errorpages/404.html` for missing files.

#### Mode 2: Custom ASP Handler
```env
ERROR_404_MODE=asp
CUSTOM_404_PAGE=/errors/404.asp
```

Create `/www/errors/404.asp`:
```vbscript
<%
Response.Status = "404 Not Found"
originalUrl = Request.ServerVariables("URL")
%>
<h1>Page Not Found</h1>
<p>The requested URL <%= Server.HTMLEncode(originalUrl) %> was not found.</p>
```

#### Mode 3: IIS web.config
```env
ERROR_404_MODE=iis
```

Create `/www/web.config`:
```xml
<?xml version="1.0" encoding="UTF-8"?>
<configuration>
    <system.webServer>
        <httpErrors errorMode="Custom">
            <remove statusCode="404" />
            <error statusCode="404" path="/errors/404.asp" responseMode="ExecuteURL" />
        </httpErrors>
    </system.webServer>
</configuration>
```

### URL Rewriting with web.config

AxonASP reads `web.config` for URL rewrite rules:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<configuration>
    <system.webServer>
        <rewrite>
            <rules>
                <!-- Redirect non-www to www -->
                <rule name="Add WWW" stopProcessing="true">
                    <match url="(.*)" />
                    <conditions>
                        <add input="{HTTP_HOST}" pattern="^example\.com$" />
                    </conditions>
                    <action type="Redirect" url="http://www.example.com/{R:1}" redirectType="Permanent" />
                </rule>
                
                <!-- Rewrite clean URLs -->
                <rule name="Product Page" stopProcessing="true">
                    <match url="^product/([0-9]+)$" />
                    <action type="Rewrite" url="/product.asp?id={R:1}" />
                </rule>
                
                <!-- Redirect old paths -->
                <rule name="Old Blog" stopProcessing="true">
                    <match url="^old-blog/(.*)" />
                    <action type="Redirect" url="/blog/{R:1}" redirectType="Permanent" />
                </rule>
            </rules>
        </rewrite>
    </system.webServer>
</configuration>
```

**Supported Features:**
- ✅ URL rewriting (transparent to user)
- ✅ URL redirects (301/302)
- ✅ Pattern matching with regular expressions
- ✅ Capture groups `{R:1}`, `{R:2}`, etc.
- ✅ Server variable conditions `{HTTP_HOST}`, `{HTTPS}`, etc.
- ✅ Multiple rules with priority ordering
- ✅ `stopProcessing` attribute

### Running Multiple Instances

Start multiple AxonASP instances with different configurations:

```bash
# Terminal 1 - Application 1
SERVER_PORT=4050 WEB_ROOT=./app1/www ./axonasp

# Terminal 2 - Application 2
SERVER_PORT=4051 WEB_ROOT=./app2/www ./axonasp

# Terminal 3 - Application 3
SERVER_PORT=4052 WEB_ROOT=./app3/www ./axonasp
```

---

## 🎯 Features & Compatibility

### Supported Classic ASP Objects

- ✅ **Request Object**: QueryString, Form, ServerVariables, Cookies, ClientCertificate
- ✅ **Response Object**: Write, Redirect, Cookies, Buffer, ContentType
- ✅ **Server Object**: CreateObject, MapPath, URLEncode, HTMLEncode
- ✅ **Session Object**: Session state management with file-based persistence
- ✅ **Application Object**: Application-wide state management

### COM Object Support

#### ADODB (Database Access)

Full support for Active Data Objects (ADO) database operations:

- **`ADODB.Connection`** - Database connections and transactions
  - Supports: SQLite, MySQL, PostgreSQL, MS SQL Server
  - **Microsoft Access** (Windows only): `.mdb` and `.accdb` files
  - Connection pooling and transaction management
  - `Open()`, `Close()`, `Execute()`, `BeginTrans()`, `CommitTrans()`, `RollbackTrans()`

- **`ADODB.Recordset`** - Data retrieval and manipulation
  - Navigation: `MoveNext()`, `MovePrevious()`, `MoveFirst()`, `MoveLast()`
  - Editing: `AddNew()`, `Update()`, `Delete()`
  - Properties: `EOF`, `BOF`, `RecordCount`, `Fields` collection
  - Cursor types and lock strategies support

- **`ADODB.Stream`** - Binary and text stream handling
  - Read/write text and binary data
  - Charset encoding support (UTF-8, ASCII, etc.)
  - `ReadText()`, `WriteText()`, `Position`, `Size` properties

**📖 Complete ADODB documentation**: [docs/ADODB_IMPLEMENTATION.md](docs/ADODB_IMPLEMENTATION.md)

**🪟 Access Database Support (Windows only)**: [docs/ACCESS_DATABASE_SUPPORT.md](docs/ACCESS_DATABASE_SUPPORT.md)

##### Database Connection Strings

**SQLite** (Cross-platform):
```vbscript
conn.Open "Driver={SQLite3};Data Source=" & Server.MapPath("database.db")
```

**MySQL**:
```vbscript
conn.Open "Driver={MySQL};Server=localhost;Database=mydb;uid=user;pwd=pass"
```

**PostgreSQL**:
```vbscript
conn.Open "Driver={PostgreSQL};Server=localhost;Database=mydb;uid=postgres;pwd=pass"
```

**MS SQL Server**:
```vbscript
conn.Open "Provider=SQLOLEDB;Server=localhost;Database=mydb;uid=sa;pwd=pass"
```

**Microsoft Access (Windows only)**:
```vbscript
' Older .mdb format (Jet)
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\db\database.mdb"

' Newer .accdb format (ACE)
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\db\database.accdb"
```

**Connection Example**:
```vbscript
<%
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Driver={SQLite3};Data Source=" & Server.MapPath("app.db")

Set rs = conn.Execute("SELECT * FROM users WHERE active = 1")
Do While Not rs.EOF
    Response.Write rs.Fields("username") & "<br>"
    rs.MoveNext
Loop

rs.Close
conn.Close
%>
```

#### ADOX (Database Schema)

Database schema creation and management:

- **`ADOX.Catalog`** - Database structure operations
  - Create databases, tables, indexes
  - Manage users and groups
  - Schema modification

**Example: Creating Access Database**
```vbscript
Set catalog = Server.CreateObject("ADOX.Catalog")
catalog.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\db\mydb.mdb"

Set table = Server.CreateObject("ADOX.Table")
table.Name = "Customers"
table.Columns.Append "ID", 3  ' adInteger
table.Columns.Append "Name", 202, 255  ' adVarWChar
catalog.Tables.Append table
```

#### MSXML2 (XML & HTTP)

- **`MSXML2.ServerXMLHTTP`** - HTTP/HTTPS requests
  - Full HTTP method support (GET, POST, PUT, DELETE, PATCH)
  - Custom headers and authentication
  - `Open()`, `Send()`, `SetRequestHeader()`, `GetResponseHeader()`
  - `ResponseText`, `ResponseXML`, `Status`, `StatusText`

- **`MSXML2.DOMDocument`** - XML parsing and manipulation
  - `Load()`, `LoadXML()`, `Save()` - File and string operations
  - XPath queries: `SelectNodes()`, `SelectSingleNode()`
  - DOM manipulation: `CreateElement()`, `AppendChild()`, etc.
  - `DocumentElement`, `ChildNodes`, navigation properties

**📖 Full MSXML2 documentation**: [docs/MSXML2_IMPLEMENTATION.md](docs/MSXML2_IMPLEMENTATION.md)

#### Scripting Objects

- **`Scripting.Dictionary`** - Key-value associative arrays
  - `Add()`, `Remove()`, `Exists()`, `Keys()`, `Items()`
  - `CompareMode` - Binary or text comparison
  - Perfect for caching and data structures

- **`Scripting.FileSystemObject`** - File system operations
  - File operations: `FileExists()`, `CopyFile()`, `DeleteFile()`, `CreateTextFile()`
  - Folder operations: `FolderExists()`, `CreateFolder()`, `DeleteFolder()`
  - Path utilities: `GetBaseName()`, `GetExtensionName()`, `BuildPath()`
  - Text file I/O: `OpenTextFile()` with read/write/append modes
  - Special folders: `GetSpecialFolder(0)` - Windows, System, Temp

**📖 Complete Scripting Objects guide**: [docs/SCRIPTING_OBJECTS_IMPLEMENTATION.md](docs/SCRIPTING_OBJECTS_IMPLEMENTATION.md)

#### WScript.Shell

Windows Script Host Shell operations:

- **`WScript.Shell`** - System integration and automation
  - `Run()` - Execute external programs and commands
  - `Exec()` - Execute with output capture
  - `CreateShortcut()` - Create Windows shortcuts
  - `RegRead()`, `RegWrite()`, `RegDelete()` - Registry operations (Windows only)
  - `ExpandEnvironmentStrings()` - Expand environment variables
  - `SpecialFolders()` - Access special Windows folders

**Example: Running System Commands**
```vbscript
Set shell = Server.CreateObject("WScript.Shell")

' Run command and wait
shell.Run "notepad.exe", 1, true

' Execute and capture output
Set exec = shell.Exec("cmd /c dir")
Do While Not exec.StdOut.AtEndOfStream
    Response.Write exec.StdOut.ReadLine() & "<br>"
Loop

' Read environment variable
winDir = shell.ExpandEnvironmentStrings("%WINDIR%")
```

**📖 WScript.Shell documentation**: [docs/WSCRIPT_SHELL_IMPLEMENTATION.md](docs/WSCRIPT_SHELL_IMPLEMENTATION.md)

#### Database Conversion Tool

AxonASP includes a built-in wizard to help you migrate legacy Microsoft Access databases to modern formats:

- **Supported Targets**: SQLite, MySQL, PostgreSQL, MS SQL Server
- **Features**: Automatic schema mapping, asynchronous processing, transaction-based data migration
- **Access URL**: `http://localhost:4050/database-convert/`

**📖 Database Conversion Tool Guide**: [docs/DATABASE_CONVERSION_TOOL.md](docs/DATABASE_CONVERSION_TOOL.md)

### Custom G3 Libraries

G3Pix AxonASP extends Classic ASP with modern functionality through custom libraries:

#### G3JSON
```vbscript
Set json = Server.CreateObject("G3JSON")
Set obj = json.Parse("{""name"": ""John"", ""age"": 30}")
Response.Write obj.name  ' John
```

#### G3FILES
```vbscript
Set fs = Server.CreateObject("G3FILES")
content = fs.Read("data.txt")
fs.Write "output.txt", "Hello World"
fs.Append "log.txt", "New log entry"
```

#### G3HTTP
```vbscript
Set http = Server.CreateObject("G3HTTP")
Set result = http.Fetch("https://api.example.com/data", "GET")
Response.Write result.body
```

#### G3MAIL
```vbscript
Set mail = Server.CreateObject("G3MAIL")
mail.Send "to@example.com", "from@example.com", "Subject", "Body", False
```

#### G3CRYPTO
```vbscript
Set crypto = Server.CreateObject("G3CRYPTO")
uuid = crypto.UUID()
hash = crypto.HashPassword("mypassword")
isValid = crypto.VerifyPassword("mypassword", hash)
```

#### G3REGEXP
```vbscript
Set re = Server.CreateObject("G3REGEXP")
re.Pattern = "\\d+"
re.Global = True
Set matches = re.Execute("Order 123, batch 456")
Response.Write matches.Count
```

#### G3FILEUPLOADER
```vbscript
Set uploader = Server.CreateObject("G3FileUploader")
uploader.MaxFileSize = 10485760  ' 10MB
Set fileInfo = uploader.Process("fileField", Server.MapPath("/uploads/"), "newname.jpg")
Response.Write "Uploaded: " & fileInfo.SavedPath
```

#### G3ZIP
```vbscript
Set zip = Server.CreateObject("G3ZIP")
If zip.Create("temp/archive.zip") Then
    zip.AddFile "data.txt", "docs/data.txt"
    zip.AddText "hello.txt", "Hello from AxonASP!"
    zip.Close()
End If
```

#### G3FC [G3FC Archiver Tool](https://g3pix.com.br/g3fc/)
Modern high-performance encrypted container:
```vbscript
Set fc = Server.CreateObject("G3FC")
fc.Create "backup.g3fc", Array("www/data", "www/images"), "secret_password"
```

#### G3TEMPLATE
Go-style template rendering:
```vbscript
Set tpl = Server.CreateObject("G3TEMPLATE")
Set data = Server.CreateObject("Scripting.Dictionary")
data.Add "Name", "World"
Response.Write tpl.Render("templates/hello.tpl", data)
```

#### G3IMAGE
2D image drawing and rendering using `gg`:
```vbscript
Set img = Server.CreateObject("G3IMAGE")
img.NewContext 600, 240
img.SetHexColor "#111827"
img.Clear
img.SetHexColor "#10b981"
img.DrawCircle 120, 120, 70
img.Fill

bytes = img.RenderContent("png")
Response.ContentType = "image/png"
Response.BinaryWrite bytes

' Explicit release so objects can be garbage collected faster
img.Close
Set img = Nothing
```

#### G3PDF
PDF generation library (FPDF 1.86 translated to Go) with text, images, and HTML rendering:
```vbscript
Set pdf = Server.CreateObject("G3PDF")
pdf.AddPage
pdf.SetFont "helvetica", "B", 16
pdf.Cell 0, 10, "Hello from AxonASP PDF", 0, 1, "L", False, ""
pdf.WriteHTML "<p><b>HTML rendering</b> is supported.</p>"
pdf.Output "I", "sample.pdf", True
```

Supported aliases: `G3PDF`, `PDF`, `FPDF`.
For complete usage examples and API details, see `docs/PDF_LIB_IMPLEMENTATION.md`.

#### G3DB
Modern database library with full `database/sql` functionality:
```vbscript
' Open database connection
Set db = Server.CreateObject("G3DB")
db.Open("sqlite", ":memory:")

' Execute queries with prepared statements
Set rs = db.Query("SELECT * FROM users WHERE age > ?", 25)
Do While Not rs.EOF
    Response.Write rs("name") & " - " & rs("email") & "<br>"
    rs.MoveNext()
Loop
rs.Close()

' Transaction support
Set tx = db.Begin()
tx.Exec("INSERT INTO users (name, email) VALUES (?, ?)", "John", "john@example.com")
tx.Commit()

' Connection pool configuration
db.SetMaxOpenConns(25)
db.SetMaxIdleConns(10)

db.Close()
```

**Supported Databases**: MySQL, PostgreSQL, MS SQL Server, SQLite  
**Key Features**: Connection pooling, transactions, prepared statements, environment configuration

**📖 Complete library documentation**: See [docs/](docs/) folder for detailed guides on each library.

Quick links:
- [docs/G3JSON_IMPLEMENTATION.md](docs/G3JSON_IMPLEMENTATION.md)
- [docs/G3FILES_IMPLEMENTATION.md](docs/G3FILES_IMPLEMENTATION.md)
- [docs/G3HTTP_IMPLEMENTATION.md](docs/G3HTTP_IMPLEMENTATION.md)
- [docs/G3MAIL_IMPLEMENTATION.md](docs/G3MAIL_IMPLEMENTATION.md)
- [docs/G3CRYPTO_IMPLEMENTATION.md](docs/G3CRYPTO_IMPLEMENTATION.md)
- [docs/G3REGEXP_IMPLEMENTATION.md](docs/G3REGEXP_IMPLEMENTATION.md)
- [docs/G3FILEUPLOADER_IMPLEMENTATION.md](docs/G3FILEUPLOADER_IMPLEMENTATION.md)
- [docs/G3ZIP_IMPLEMENTATION.md](docs/G3ZIP_IMPLEMENTATION.md)
- [docs/G3FC_IMPLEMENTATION.md](docs/G3FC_IMPLEMENTATION.md)
- [docs/G3TEMPLATE_IMPLEMENTATION.md](docs/G3TEMPLATE_IMPLEMENTATION.md)
- [docs/G3DB_IMPLEMENTATION.md](docs/G3DB_IMPLEMENTATION.md)
- [docs/G3IMAGE_IMPLEMENTATION.md](docs/G3IMAGE_IMPLEMENTATION.md)
- [docs/PDF_LIB_IMPLEMENTATION.md](docs/PDF_LIB_IMPLEMENTATION.md)

---

## 🔧 Extended Custom Functions

G3Pix AxonASP includes **60+ custom functions** inspired by PHP, all prefixed with `Ax` for clarity. These functions enhance productivity and provide familiar functionality for developers coming from other platforms.

### Array Functions
- `AxArrayMerge(arr1, arr2, ...)` - Merge multiple arrays
- `AxArrayContains(value, array)` - Check if value exists in array
- `AxArrayMap(callback, array)` - Transform array elements
- `AxArrayFilter(callback, array)` - Filter array elements
- `AxCount(array)` - Return array length
- `AxExplode(delimiter, string)` - Split string into array
- `AxImplode(glue, array)` - Join array into string
- `AxArrayReverse(array)` - Reverse array order
- `AxRange(start, end, step)` - Create range array

### String Functions
- `AxStringReplace(search, replace, subject)` - Replace string
- `AxSprintf(format, args...)` - Format string (supports %s, %d, %f)
- `AxPad(string, length, pad_string, type)` - Pad string
- `AxRepeat(string, count)` - Repeat string
- `AxUcFirst(string)` - Uppercase first character
- `AxWordCount(string, format)` - Count words
- `AxNewLineToBr(string)` - Convert newlines to `<br>`
- `AxTrim(string, chars)` - Trim characters
- `AxStringGetCsv(string)` - Parse CSV string

### Math Functions
- `AxCeil(number)` - Round up
- `AxFloor(number)` - Round down
- `AxMax(values...)` - Return maximum value
- `AxMin(values...)` - Return minimum value
- `AxRand(min, max)` - Random integer
- `AxNumberFormat(number, decimals, dec_point, thousands_sep)` - Format number

### Type Checking
- `AxIsInt(value)` - Check if integer
- `AxIsFloat(value)` - Check if float
- `AxCTypeAlpha(string)` - Check if alphabetic
- `AxCTypeAlnum(string)` - Check if alphanumeric
- `AxEmpty(value)` - Check if empty
- `AxIsset(value)` - Check if set

### Date/Time Functions
- `AxTime()` - Get Unix timestamp
- `AxDate(format, timestamp)` - Format date/time

### Hashing & Encoding
- `AxMd5(string)` - MD5 hash
- `AxSha1(string)` - SHA1 hash
- `AxHash(algorithm, string)` - Hash with algorithm
- `AxBase64Encode(string)` - Base64 encode
- `AxBase64Decode(string)` - Base64 decode
- `AxUrlDecode(string)` - URL decode
- `AxHtmlSpecialChars(string)` - Encode HTML special characters
- `AxStripTags(string)` - Remove HTML tags
- `AxRgbToHex(r, g, b)` - Convert RGB to hex color

### Validation Functions
- `AxFilterValidateIp(ip)` - Validate IP address
- `AxFilterValidateEmail(email)` - Validate email address

### Request Functions
- `AxGetRequest()` - Get all request parameters
- `AxGetGet()` - Get query string parameters
- `AxGetPost()` - Get POST parameters

### Utility Functions
- `AxGenerateGuid()` - Generate GUID
- `AxBuildQueryString(dict)` - Build URL query string
- `AxVarDump(value)` - Debug output variable
- `Document.Write(string)` - HTML-safe output

**📖 Full documentation**: See [CUSTOM_FUNCTIONS.md](CUSTOM_FUNCTIONS.md)

---

## 📁 Project Structure

```
axonasp/
├── main.go                 # Entry point & HTTP server
├── .env                    # Configuration file
├── asp/                    # ASP parser & executor
│   ├── asp_parser.go       # VBScript parser
│   ├── asp_lexer.go        # Tokenizer
│   └── asp_executor.go     # Code executor
├── axonaspcgi/             # FastCGI mode
│   └── main.go             # FastCGI application server
├── server/                 # Core libraries
│   ├── request_object.go   # Request object
│   ├── response_object.go  # Response object
│   ├── session_object.go   # Session management
│   ├── server_object.go    # Server object
│   ├── application_object.go # Application state
│   ├── custom_functions.go # Custom Ax functions
│   ├── json_lib.go         # G3JSON library
│   ├── file_lib.go         # G3FILES library
│   ├── http_lib.go         # G3HTTP library
│   ├── mail_lib.go         # G3MAIL library
│   ├── crypto_lib.go       # G3CRYPTO library
│   ├── regexp_lib.go       # G3REGEXP library
│   ├── zip_lib.go          # G3ZIP library
│   ├── g3fc_lib.go         # G3FC library
│   ├── template_lib.go     # G3TEMPLATE library
│   ├── image_lib.go        # G3IMAGE library
│   ├── file_uploader_lib.go # G3FileUploader library
│   ├── g3db_lib.go         # G3DB library (modern database access)
│   ├── database_lib.go     # ADODB implementation
│   ├── adox_lib.go         # ADOX implementation
│   ├── msxml_lib.go        # MSXML2 implementation
│   ├── dictionary_lib.go   # Scripting.Dictionary
│   ├── wscript_shell_lib.go # WScript.Shell
│   ├── webconfig_parser.go # web.config reader
│   └── global_asa_manager.go # Global.asa handler
├── vbscript/               # VBScript compatibility layer
├── docs/                   # Documentation
│   ├── G3DB_IMPLEMENTATION.md
│   ├── ADODB_IMPLEMENTATION.md
│   ├── ACCESS_DATABASE_SUPPORT.md
│   ├── FASTCGI_MODE.md
│   ├── FASTCGI_QUICKSTART.md
│   ├── SCRIPTING_OBJECTS_IMPLEMENTATION.md
│   ├── WSCRIPT_SHELL_IMPLEMENTATION.md
│   ├── MSXML2_IMPLEMENTATION.md
│   ├── (OTHER LIBRARIES AND HELPERS)
│   └── CUSTOM_FUNCTIONS.md
├── www/                    # Web root (your ASP files here)
│   ├── default.asp         # Default document
│   ├── global.asa          # Application events
│   ├── web.config          # IIS-style configuration
│   └── tests/              # Test files
├── temp/session/           # Session storage
└── errorpages/             # Error templates (403, 404, 500)
```

---

## 🔐 Security

### Secure by Design

- **Sandboxed File Access**: All files must reside in the `www/` directory, preventing directory traversal attacks
- **Session Management**: Secure session handling with file-based persistence in `temp/session/`
- **Input Validation**: Built-in validation functions for IP addresses, emails, and more
- **HTML Encoding**: Automatic HTML encoding with `Document.Write()` to prevent XSS
- **Error Handling**: Graceful error pages without exposing sensitive information

### Best Practices

1. Always validate user input using validation functions
2. Use `Document.Write()` for user-generated content to prevent XSS
3. Store sensitive files outside the `www/` directory
4. Enable `DEBUG_ASP=FALSE` in production
5. Use HTTPS with reverse proxy in production environments
6. Regularly update dependencies with `go get -u`

---

### Enable ASP Debugging

In `.env`:
```env
DEBUG_ASP=TRUE
#SQL Tracing Information (very verbose)
SQL_TRACE=TRUE
```

This enables error description for debugging on console.

---

## 📊 Performance

G3Pix AxonASP delivers exceptional performance thanks to GoLang's efficiency:

- **Fast Startup**: Server starts in milliseconds
- **Low Memory Footprint**: Minimal resource consumption
- **Concurrent Request Handling**: Native Go concurrency for handling multiple requests
- **Optimized Parsing**: Efficient VBScript lexer and parser 

---

## 🌟 Why Choose G3Pix AxonASP?

| Feature | Traditional IIS | G3Pix AxonASP |
|---------|-----------------|---------------|
| **Platform** | Windows only | Windows, Linux, macOS |
| **Performance** | Standard | Accelerated (Go) |
| **Dependencies** | IIS, Windows Server | Single binary |
| **Deployment** | Complex | Simple binary or FastCGI |
| **Database Support** | Windows databases | SQLite, MySQL, PostgreSQL, SQL Server, Access |
| **Cost** | Windows licensing | Free & open source |
| **Modernization** | Limited | 60+ extended functions |
| **Container Ready** | Challenging | Docker-friendly |
| **Web Server Integration** | IIS only | nginx, Apache, IIS, FastCGI |
| **URL Rewriting** | IIS modules | Built-in web.config support |

---

## 🛠️ Global.asa Support

G3Pix AxonASP fully supports `global.asa` for application and session lifecycle events:

### Supported Events

- `Application_OnStart` - Fires when the server starts
- `Application_OnEnd` - Fires when the server shuts down
- `Session_OnStart` - Fires when a new session is created
- `Session_OnEnd` - Fires when a session expires


**⚠️ Important**: Each AxonASP instance supports **one application** per server due to global.asa loading. Run multiple instances for multiple applications.

---

## 🤝 Contributing

We welcome contributions! Please follow these guidelines:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

### Development Guidelines

- All code, comments, and documentation must be in **English**
- Follow Go best practices and conventions
- Add tests for new features in `www/tests/`
- Update documentation when adding features
- Keep commits atomic and descriptive

---

## 📄 License

This project is licensed under the MPL License - see the [LICENSE](LICENSE) file for details.

---

## 💬 Support & Community

- **Issues**: [GitHub Issues](https://github.com/guimaraeslucas/axonasp/issues)
- **Discussions**: [GitHub Discussions](https://github.com/guimaraeslucas/axonasp/discussions)
- **Website**: [https://g3pix.com.br/axonasp](https://g3pix.com.br/axonasp)

---

## 🗺️ Roadmap

- [X] ADODB database support (SQLite, MySQL, PostgreSQL, SQL Server)
- [X] Microsoft Access database support (Windows)
- [X] FastCGI mode for production deployments
- [X] IIS-style web.config URL rewriting
- [X] Custom 404 error handling (3 modes)
- [X] Scripting.FileSystemObject and Dictionary
- [X] WScript.Shell for system integration
- [X] ADOX for database schema management
- [X] Database Conversion Tool (Access to SQLite/MySQL/PG/MSSQL)
- [X] 60+ custom functions (Ax* functions)
- [X] Image creation
- [X] ZIP support
- [X] G3FC support
- [X] XML support
- [X] PDF support
- [ ] WebSocket support
- [ ] Built-in Redis session storage
- [ ] Docker official images
- [ ] OAuth2 authentication library
- [ ] REST API generator
- [ ] GraphQL support
- [ ] Implement VM/Compiler

---

## 🙏 Acknowledgments

Special thanks to:
- The Go community for an amazing language and ecosystem
- Classic ASP developers who keep legacy applications running
- Contributors and testers who help improve G3Pix AxonASP =)
- Pieter Cooreman (@PieterCooreman) for the help with real ASP code, tests and bug checks

---

<p align="center">
  <strong>Built with ❤️ by G3Pix</strong>
  <br>
  Making Classic ASP modern, fast, and cross-platform
</p>

<p align="center">
  <a href="https://github.com/guimaraeslucas/axonasp">⭐ Star us on GitHub</a>
  •
  <a href="https://github.com/guimaraeslucas/axonasp/issues">🐛 Report Bug</a>
  •
  <a href="https://github.com/guimaraeslucas/axonasp/issues">✨ Request Feature</a>
</p>
