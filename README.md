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
- **📦 COM Object Support**: Full support for ADODB, MSXML2, and custom G3 libraries
- **⚙️ Simple Configuration**: Environment-based configuration via `.env` file
- **🔌 Reverse Proxy Ready**: Seamlessly integrates with Nginx, Apache, or IIS as reverse proxy

---

## 📦 Installation

### Prerequisites

- **Go 1.21+** installed on your system
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
   
   Create a `.env` file in the root directory:
   ```env
   SERVER_PORT=4050
   WEB_ROOT=./www
   TIMEZONE=America/Sao_Paulo
   DEFAULT_PAGE=default.asp
   SCRIPT_TIMEOUT=30
   DEBUG_ASP=FALSE
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
go build -o axonasp
```

### Cross-Platform Compilation

**Windows (64-bit)**
```bash
GOOS=windows GOARCH=amd64 go build -o axonasp.exe
```

**Linux (64-bit)**
```bash
GOOS=linux GOARCH=amd64 go build -o axonasp
```

**macOS (Intel)**
```bash
GOOS=darwin GOARCH=amd64 go build -o axonasp
```

**macOS (Apple Silicon)**
```bash
GOOS=darwin GOARCH=arm64 go build -o axonasp
```

After building, simply run the executable:
```bash
./axonasp          # Linux/macOS
axonasp.exe        # Windows
```

---

## ⚙️ Configuration

G3Pix AxonASP uses a `.env` file for configuration. All settings are optional with sensible defaults.

### Available Configuration Options

| Variable | Default | Description |
|----------|---------|-------------|
| `SERVER_PORT` | `4050` | HTTP server port |
| `WEB_ROOT` | `./www` | Root directory for ASP files |
| `TIMEZONE` | `America/Sao_Paulo` | Server timezone |
| `DEFAULT_PAGE` | `default.asp` | Default document name |
| `SCRIPT_TIMEOUT` | `30` | Script execution timeout (seconds) |
| `DEBUG_ASP` | `FALSE` | Enable HTML stack traces in ASP files |
| `SMTP_HOST` | - | SMTP server hostname |
| `SMTP_PORT` | `587` | SMTP server port |
| `SMTP_USER` | - | SMTP authentication username |
| `SMTP_PASS` | - | SMTP authentication password |

### Example `.env` File

```env
SERVER_PORT=4050
WEB_ROOT=./www
TIMEZONE=America/New_York
DEFAULT_PAGE=index.asp
SCRIPT_TIMEOUT=60
DEBUG_ASP=TRUE

# SMTP Configuration
SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_USER=your-email@gmail.com
SMTP_PASS=your-app-password
```

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
- `ADODB.Connection` - Database connections (SQL Server, MySQL, PostgreSQL, SQLite)
- `ADODB.Recordset` - Data retrieval and manipulation
- `ADODB.Stream` - Binary and text stream handling

#### MSXML2 (XML & HTTP)
- `MSXML2.ServerXMLHTTP` - HTTP requests
- `MSXML2.DOMDocument` - XML parsing and manipulation

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

#### G3TEMPLATE
```vbscript
Set tpl = Server.CreateObject("G3TEMPLATE")
Set data = Server.CreateObject("Scripting.Dictionary")
data.Add "name", "John"
html = tpl.Render("template.html", data)
```

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
│   ├── database_lib.go     # ADODB implementation
│   ├── msxml_lib.go        # MSXML2 implementation
│   └── global_asa_manager.go # Global.asa handler
├── vbscript/               # VBScript compatibility layer
├── www/                    # Web root (your ASP files here)
│   ├── default.asp         # Default document
│   ├── global.asa          # Application events
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

## 🧪 Testing

Test files are available in the `www/tests/` directory. Access them via:

```
http://localhost:4050/tests/test_basics.asp
http://localhost:4050/tests/test_custom_functions.asp
http://localhost:4050/tests/test_database.asp
```

### Enable ASP Debugging

In `.env`:
```env
DEBUG_ASP=TRUE
```

This enables error description for debugging on console.

---

## 📊 Performance

G3Pix AxonASP delivers exceptional performance thanks to GoLang's efficiency:

- **Fast Startup**: Server starts in milliseconds
- **Low Memory Footprint**: Minimal resource consumption
- **Concurrent Request Handling**: Native Go concurrency for handling multiple requests
- **Optimized Parsing**: Efficient VBScript lexer and parser
- **Compiled Binary**: No interpreter overhead, runs as native code

---

## 🌟 Why Choose G3Pix AxonASP?

| Feature | Traditional IIS | G3Pix AxonASP |
|---------|-----------------|---------------|
| **Platform** | Windows only | Windows, Linux, macOS |
| **Performance** | Standard | Accelerated (Go) |
| **Dependencies** | IIS, Windows Server | Single binary |
| **Deployment** | Complex | Simple binary |
| **Cost** | Windows licensing | Free & open source |
| **Modernization** | Limited | Extended functions |
| **Container Ready** | Challenging | Docker-friendly |

---

## 🛠️ Global.asa Support

G3Pix AxonASP fully supports `global.asa` for application and session lifecycle events:

### Supported Events

- `Application_OnStart` - Fires when the server starts
- `Application_OnEnd` - Fires when the server shuts down
- `Session_OnStart` - Fires when a new session is created
- `Session_OnEnd` - Fires when a session expires

### Example global.asa

```vbscript
<script language="vbscript" runat="server">
Sub Application_OnStart
    Application("StartTime") = Now()
    Application("TotalVisitors") = 0
End Sub

Sub Session_OnStart
    Application("TotalVisitors") = Application("TotalVisitors") + 1
End Sub
</script>
```

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
- **Website**: [https://g3pix.com](https://g3pix.com.br/axonasp)

---

## 🗺️ Roadmap

- [ ] WebSocket support
- [ ] Built-in Redis session storage
- [ ] Performance monitoring dashboard
- [ ] Docker official images
- [ ] Additional database drivers
- [ ] REST API generator
- [ ] Hot-reload development mode

---

## 🙏 Acknowledgments

Special thanks to:
- The Go community for an amazing language and ecosystem
- Classic ASP developers who keep legacy applications running
- Contributors and testers who help improve G3Pix AxonASP =)

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
