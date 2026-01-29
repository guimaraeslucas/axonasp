Quick Instructions for Code Agents (G3Pix AxonASP)
Role: Expert GoLang Developer. Focus: Quality, precision, performance, security. Primary Constraint: ALL content (code, comments, documentation, output) must be in ENGLISH (US), regardless of the user's input language. Don't summarize, or explain the changes unless explicitly asked, just provide the code. Also, think and explain in english, even if asked in portuguese. Document only in English. Keep the license reader.

1. Architecture Overview
Main Server: main.go runs HTTP server on :4050, serving ./www.

Core Logic: asp/ contains asp_parser.go and asp_lexer.go (VBScript-Go integration).

Libraries: server/libs/ contains ASP implementations (FileSystem, JSON, HTTP, etc.).

Invoked via Server.CreateObject("LIB_NAME").

Must use the standard ASP execution context.

2. Development & Debugging
Environment: Windows Powershell.

Run: go run main.go.

Build: go build -o axonasp.exe -> ./axonasp.exe.

Testing: Access http://localhost:4050/tests/test_basics.asp or other test_*.asp files in www/tests.

ASP Debugging: Set <% debug_asp_code = "TRUE" %> in the ASP file for HTML stack traces.

Compilation Rule: ALWAYS compile Go code after editing to verify success. Do not compile for pure ASP edits.

3. Coding Standards & Conventions
Language: STRICT ENGLISH ONLY. Translate any non-English comments/UI/Code and any output or answer immediately.

VBScript Compatibility: Strict adherence to VBScript and ASP Classic standards.

Option Compare: Parser honors Option Compare Binary/Text at top of VBScript files; executor applies per-program compare mode (binary vs case-insensitive text) for all string comparisons (including Select Case, =/</> ops).

Variable Lookup: MUST ALWAYS be Case-insensitive (and stored as lowercase internally).

Session/App: Sessions stored in temp/session (Cookie: ASPSESSIONID); Application state in memory.

Includes: file = relative to current; virtual = relative to www/.

Documentation: Keep instructions in this file. Sync copilot-instructions.md and GEMINI.md on updates. Do not create new .md explanation files unless asked. The folder docs/ is reserved for user-facing documentation, follow the pattern from the files already created when adding new documentation.

New Libraries: Name as *_lib.go. Mimic VBScript nomenclature. Document only in English.

1. Configuration (.env)
File: .env in root (defaults in code).

Keys: SERVER_PORT (4050), WEB_ROOT (./www), TIMEZONE (America/Sao_Paulo), DEFAULT_PAGE (default.asp), SCRIPT_TIMEOUT (30), DEBUG_ASP (FALSE), SMTP settings.

5. API & Library Reference

### Custom G3 Libraries (Server.CreateObject)

**G3JSON** (json_lib.go)
- `Parse(jsonString)` - Parse JSON string to object/array
- `Stringify(object)` - Convert object to JSON string
- `NewObject()` - Create empty dictionary
- `NewArray()` - Create empty array
- `LoadFile(path)` - Load and parse JSON file
- Returns native Go maps/slices for VBScript subscript access

**G3FILES** (file_lib.go)
- `Read(path)` / `ReadText(path)` - Read entire file
- `Write(path, content)` / `WriteText(path, content)` - Create/overwrite file
- `Append(path, content)` / `AppendText(path, content)` - Append to file
- `Exists(path)` - Check file/folder existence
- `Size(path)` - Get file size in bytes
- `List(path)` - Get directory contents array
- `Delete(path)` / `Remove(path)` - Delete file
- `Copy(source, dest)` - Copy file atomically
- `Move(source, dest)` - Move file
- `Rename(path, newName)` - Rename file
- `MkDir(path)` / `MakeDir(path)` - Create directory hierarchy
- `DateCreated(path)`, `DateModified(path)` - Get timestamps
- Path security: All paths validated against root directory

**G3HTTP** (http_lib.go)
- `Fetch(url, [method], [body])` / `Request(...)` - Execute HTTP request
- Methods: GET, POST, PUT, DELETE, PATCH
- Automatic JSON response parsing to Dictionary objects
- Returns plain text for non-JSON responses
- 10-second timeout default
- Returns nil on error

**G3TEMPLATE** (template_lib.go)
- `Render(path, [data])` - Parse and render Go template
- Supports: {{.Variable}}, {{if}}, {{range}}, piping
- Data binding via objects/dictionaries
- Auto-escaping for security
- Perfect for email templates, reports, HTML generation

**G3MAIL** (mail_lib.go)
- `Send(host, port, user, pass, from, to, subject, body, [isHtml])` - Manual SMTP config
- `SendStandard(to, subject, body, [isHtml])` - Use .env SMTP settings
- Environment vars: SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS, SMTP_FROM
- HTML and plain text support
- Comma-separated recipients supported

**G3CRYPTO** (crypto_lib.go)
- `UUID()` - Generate RFC 4122 UUID v4 (36 chars with hyphens)
- `HashPassword(pass)` / `Hash(pass)` - bcrypt hashing (12 cost)
- `VerifyPassword(pass, hash)` / `Verify(pass, hash)` - Constant-time comparison
- Cryptographically secure random generation
- One-way hashing resistant to rainbow tables

**G3REGEXP** (regexp_lib.go)
- `Pattern` property - The regex pattern string
- `IgnoreCase` property - Case-insensitive (boolean)
- `Global` property - Match all (boolean)
- `Multiline` property - ^ and $ match line boundaries
- `Test(text)` - Boolean test if pattern matches
- `Execute(text)` - Returns Matches collection
- `Replace(text, replacement)` - Replace matched text
- Full Go regexp syntax with capture groups
- Returns RegExpMatches with Value, Index, Length for each match

**G3FileUploader** (file_uploader_lib.go)
- `Process(fieldName, targetDir, [newFileName])` - Upload single file
- `ProcessAll(targetDir)` - Upload all files in form
- `GetFileInfo(fieldName)` - Get file info without upload
- `GetAllFilesInfo()` - Get all files info
- Extension validation: `BlockExtension()`, `AllowExtension()`, `SetUseAllowedOnly()`
- `MaxFileSize` property for size limits (default 10MB)
- Returns rich metadata: Size, MIME type, paths, timestamps

### Standard COM Support

**MSXML2.ServerXMLHTTP** (msxml_lib.go)
- `Open(method, url, [async], [user], [pwd])` - Initialize request
- `SetRequestHeader(header, value)` - Add custom headers
- `Send([body])` - Execute request
- `GetResponseHeader(header)`, `GetAllResponseHeaders()` - Response headers
- `ResponseText`, `ResponseXML`, `ResponseBody` - Response access
- `Status`, `StatusText`, `ReadyState` - Response info
- Full HTTP method support with proper async/sync handling

**MSXML2.DOMDocument** (msxml_lib.go)
- `Load(path)` - Load XML file
- `LoadXML(xmlString)` - Parse XML string
- `Save(path)` - Save to file
- `DocumentElement` - Root element
- `CreateElement()`, `CreateAttribute()`, `CreateTextNode()`, `CreateCDATASection()`
- `SelectNodes(xpath)` - XPath query returning collection
- `SelectSingleNode(xpath)` - XPath returning single node
- Full DOM navigation: ChildNodes, ParentNode, FirstChild, LastChild, NextSibling
- Node manipulation: AppendChild, InsertBefore, RemoveChild, ReplaceChild
- `Xml`, `InnerXml`, `OuterXml` properties for serialization

**ADODB.Connection** (database_lib.go)
- `Open(connString, [user], [pwd])` - Open database
- `Close()` - Close connection
- `Execute(sql, [params])` - Execute query returning Recordset
- `BeginTrans()`, `CommitTrans()`, `RollbackTrans()` - Transactions
- `State` property - 0=closed, 1=open
- `Errors` collection for error details
- Supports: SQLite, MySQL, PostgreSQL, MS SQL Server
- Parameter binding prevents SQL injection

**ADODB.Recordset** (database_lib.go)
- Navigation: `MoveFirst()`, `MoveLast()`, `MoveNext()`, `MovePrevious()`, `Move()`
- `EOF`, `BOF` - End/Beginning of file
- `RecordCount` - Total records
- `AddNew()`, `Update([field], [value])`, `Delete()` - Record modification
- `Fields` collection - Column access
- `AbsolutePosition` - Current record position
- Supports cursor types and locking strategies

**ADODB.Stream** (database_lib.go)
- `Open([source], [mode], [options], [user], [pwd])` - Open stream
- `ReadText([chars])`, `WriteText(text)` - Text operations
- `Type` property - 1=Text, 2=Binary
- `Charset` property - Encoding (UTF-8, ASCII, etc.)
- `Position`, `Size` - Stream properties

**Scripting.Dictionary** (dictionary_lib.go)
- `Add(key, value)` - Add pair
- `Remove(key)`, `RemoveAll()` - Remove
- `Exists(key)` - Check existence
- `Item(key)` - Get/set via subscript or method
- `Keys()`, `Items()` - Get all keys/values as array
- `Count` property - Number of items
- `CompareMode` - 0=Binary, 1=TextCompare
- Full For Each support via enumeration

**Scripting.FileSystemObject** (file_lib.go)
- File: `FileExists()`, `GetFile()`, `CopyFile()`, `MoveFile()`, `DeleteFile()`, `CreateTextFile()`
- Folder: `FolderExists()`, `GetFolder()`, `CreateFolder()`, `MoveFolder()`, `DeleteFolder()`
- Text: `OpenTextFile(path, mode, [create], [format])` - Returns TextFile
- Path: `GetBaseName()`, `GetExtensionName()`, `GetParentFolderName()`, `GetDrive()`, `BuildPath()`
- Special: `GetSpecialFolder(type)` - 0=System, 1=Windows, 2=Temp
- FSOFile: Name, Path, Size, Type, DateCreated, DateModified, Copy(), Move(), Delete()
- FSOFolder: Name, Path, Size, Files collection, SubFolders collection, CreateFolder()
- FSOTextFile: ReadLine(), ReadAll(), Read(chars), WriteLine(), Write(), Close(), AtEndOfStream

### Database Connection Strings

**SQLite**:
```
Driver={SQLite3};Data Source=C:\path\database.db
```

**MySQL**:
```
Driver={MySQL};Server=localhost;Database=dbname;uid=user;pwd=pass
```

**PostgreSQL**:
```
Driver={PostgreSQL};Server=localhost;Database=dbname;uid=postgres;pwd=pass
```

**MS SQL Server**:
```
Provider=SQLOLEDB;Server=servername;Database=dbname;uid=sa;pwd=pass
```

### Documentation
Detailed implementation guides for each library are in docs/ folder:
- docs/G3JSON_IMPLEMENTATION.md
- docs/G3FILES_IMPLEMENTATION.md
- docs/G3HTTP_IMPLEMENTATION.md
- docs/G3TEMPLATE_IMPLEMENTATION.md
- docs/G3MAIL_IMPLEMENTATION.md
- docs/G3CRYPTO_IMPLEMENTATION.md
- docs/G3REGEXP_IMPLEMENTATION.md
- docs/MSXML2_IMPLEMENTATION.md
- docs/ADODB_IMPLEMENTATION.md
- docs/SCRIPTING_OBJECTS_IMPLEMENTATION.md
- docs/G3FILEUPLOADER_IMPLEMENTATION.md

6. Library Implementation Patterns

### Adding a New Library
1. Create `server/newlib_lib.go` following existing patterns
2. Implement Component interface: `GetProperty()`, `SetProperty()`, `CallMethod()`
3. Register in `server/executor_libraries.go` - Create wrapper type
4. Register in `server/executor.go` - Add to CreateObject() switch
5. Create documentation in `docs/NEWLIB_IMPLEMENTATION.md`
6. Add test file `www/tests/test_newlib.asp`
7. Update CUSTOM_FUNCTIONS.md if adding built-in functions

### Common Library Pattern
```go
type G3NEWLIB struct {
    ctx *ExecutionContext
}

func (lib *G3NEWLIB) GetProperty(name string) interface{} {
    // Return property value
}

func (lib *G3NEWLIB) SetProperty(name string, value interface{}) {}

func (lib *G3NEWLIB) CallMethod(name string, args ...interface{}) interface{} {
    switch strings.ToLower(name) {
    case "methodname":
        // Implement method
    }
    return nil
}
```

### Best Practices for Agents
- Always use lowercase for method/property lookups: `strings.ToLower(name)`
- Validate arguments before processing
- Return nil for errors (not panic)
- Use ExecutionContext for path resolution: `ctx.Server_MapPath()`
- Log errors via existing helpers, not stdout
- Preserve VBScript type semantics in return values
- Support both common names and abbreviations (e.g., Read/ReadText)
- Guard mutable state with mutexes if concurrent access possible
- Test with both uppercase and mixed-case method calls

7. Pull Request Guidelines
Update/Create test_*.asp in www/tests/ for every fix/feature.

Update Default.asp with changes.

Maintain G3Pix AxonASP branding.

Prioritize secure, testable, and small implementations.

8. Agent Quick-Start Checklist
- Keep responses and code comments strictly in ENGLISH (US).
- Prefer small, safe diffs; avoid touching server/deprecated/ except for reference.
- Respect VBScript semantics: case-insensitive identifiers, Option Compare rules, ByRef/ByVal behavior.
- Preserve ASP execution context when adding libraries or functions.
- When adding a library, name it *_lib.go and register via Server.CreateObject mapping.
- Sync updates between this file and GEMINI.md whenever instructions change.

9. Coding & Tooling Expectations
- Run gofmt on touched Go files; keep ASCII unless a file already needs non-ASCII.
- Compile after Go changes: go build -o go-asp.exe ./...
- Run tests when applicable: go test ./asp ./server ./VBScript-Go
- For ASP-only changes, do not rebuild; validate by hitting http://localhost:4050/tests/<test>.asp
- Favor explicit errors; avoid panics in request path; log via existing error handling helpers.
- Concurrency: session/app state is shared; guard mutable shared data when introducing goroutines.

10. ASP/VBScript Execution Notes
- Option Compare Binary/Text at file top sets comparison mode for that program; executor applies the chosen mode to all string comparisons.
- Includes: file path relative to current file; virtual path relative to www/ root.
- Session storage: temp/session with cookie ASPSESSIONID; Application lives in-memory.
- Variable lookup and storage are case-insensitive; store lowercase internally.
- Custom objects must match classic ASP expectations (e.g., ADODB-like APIs, MSXML2 object models).

11. Global.asa Support
- File Location: www/global.asa
- Supported Formats: Both `<% %>` and `<script runat="server">` blocks
- Events Supported:
  * Application_OnStart: Executed once at server startup
  * Application_OnEnd: Executed when server shuts down
  * Session_OnStart: Executed when a new session is created
  * Session_OnEnd: Executed when a session expires or is abandoned
- ASP Lexer Enhancement: Added support for `<script language="vbscript" runat="server">` blocks via regex matching
- Implementation: global_asa_manager.go handles loading, parsing, and executing global.asa events
- Testing: Use www/tests/test_global_asa.asp to verify global.asa functionality
