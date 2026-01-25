Quick Instructions for Code Agents (G3 AxonASP)
Role: Expert GoLang Developer. Focus: Quality, precision, performance, security. Primary Constraint: ALL content (code, comments, documentation, output) must be in ENGLISH (US), regardless of the user's input language. Don't summarize, or explain the changes unless explicitly asked, just provide the code. Also, think and explain in english, even if asked in portuguese. Document only in English.

1. Architecture Overview
Main Server: main.go runs HTTP server on :4050, serving ./www.

Core Logic: asp/ contains asp_parser.go and asp_lexer.go (VBScript-Go integration).

Libraries: server/libs/ contains ASP implementations (FileSystem, JSON, HTTP, etc.).

Invoked via Server.CreateObject("LIB_NAME").

Must use the standard ASP execution context.

Deprecated: server/deprecated/ is for reference only. DO NOT MODIFY OR USE IN PRODUCTION.

2. Development & Debugging
Environment: Windows Powershell.

Run: go run main.go.

Build: go build -o go-asp.exe -> ./go-asp.exe.

Testing: Access http://localhost:4050/tests/test_basics.asp or other test_*.asp files in www/.

ASP Debugging: Set <% debug_asp_code = "TRUE" %> in the ASP file for HTML stack traces.

Compilation Rule: ALWAYS compile Go code after editing to verify success. Do not compile for pure ASP edits.

3. Coding Standards & Conventions
Language: STRICT ENGLISH ONLY. Translate any non-English comments/UI/Code and any output or answer immediately.

VBScript Compatibility: Strict adherence to VBScript and ASP Classic standards.

Option Compare: Parser honors Option Compare Binary/Text at top of VBScript files; executor applies per-program compare mode (binary vs case-insensitive text) for all string comparisons (including Select Case, =/</> ops).

Variable Lookup: MUST ALWAYS be Case-insensitive (and stored as lowercase internally).

Session/App: Sessions stored in temp/session (Cookie: ASPSESSIONID); Application state in memory.

Includes: file = relative to current; virtual = relative to www/.

Documentation: Keep instructions in this file. Sync copilot-instructions.md and GEMINI.md on updates. Do not create new .md explanation files.

New Libraries: Name as *_lib.go. Mimic VBScript nomenclature. Document only in English.

4. Configuration (.env)
File: .env in root (defaults in code).

Keys: SERVER_PORT (4050), WEB_ROOT (./www), TIMEZONE (America/Sao_Paulo), DEFAULT_PAGE (default.asp), SCRIPT_TIMEOUT (30), DEBUG_ASP (FALSE), SMTP settings.

5. API & Library Reference
Custom G3 Libs (Server.CreateObject):
G3JSON: NewObject, Parse, Stringify, LoadFile.
G3FILES: Read, Write, Append, Exists, Size, List, Delete, MkDir.
G3HTTP: Fetch (method).
G3TEMPLATE: Render.
G3MAIL: Send, SendStandard.
G3CRYPTO: UUID, HashPassword, VerifyPassword.

Standard COM Support:
MSXML2: ServerXMLHTTP, DOMDocument (standard methods supported).
ADODB: Connection, Recordset, Stream.
Databases: SQLite (:memory:, file), MySQL, PostgreSQL, MS SQL Server.
Filtering: Supports in-memory filtering (=, <>, LIKE, etc.).

6. Pull Request Guidelines
Update/Create test_*.asp in www/ for every fix/feature.

Update Default.asp with changes.

Maintain G3 AxonASP branding.

Prioritize secure, testable, and small implementations.

7. Agent Quick-Start Checklist
- Keep responses and code comments strictly in ENGLISH (US).
- Prefer small, safe diffs; avoid touching server/deprecated/ except for reference.
- Respect VBScript semantics: case-insensitive identifiers, Option Compare rules, ByRef/ByVal behavior.
- Preserve ASP execution context when adding libraries or functions.
- When adding a library, name it *_lib.go and register via Server.CreateObject mapping.
- Sync updates between this file and GEMINI.md whenever instructions change.

8. Coding & Tooling Expectations
- Run gofmt on touched Go files; keep ASCII unless a file already needs non-ASCII.
- Compile after Go changes: go build -o go-asp.exe ./...
- Run tests when applicable: go test ./asp ./server ./VBScript-Go
- For ASP-only changes, do not rebuild; validate by hitting http://localhost:4050/<test>.asp
- Favor explicit errors; avoid panics in request path; log via existing error handling helpers.
- Concurrency: session/app state is shared; guard mutable shared data when introducing goroutines.

9. ASP/VBScript Execution Notes
- Option Compare Binary/Text at file top sets comparison mode for that program; executor applies the chosen mode to all string comparisons.
- Includes: file path relative to current file; virtual path relative to www/ root.
- Session storage: temp/session with cookie ASPSESSIONID; Application lives in-memory.
- Variable lookup and storage are case-insensitive; store lowercase internally.
- Custom objects must match classic ASP expectations (e.g., ADODB-like APIs, MSXML2 object models).

10. Global.asa Support
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
