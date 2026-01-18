Quick Instructions for Code Agents (G3 AxonASP)
Role: Expert GoLang Developer. Focus: Quality, precision, performance, security. Primary Constraint: ALL content (code, comments, documentation, output) must be in ENGLISH (US), regardless of the user's input language. Don't summarize, or explain the changes unless explicitly asked, just provide the code. Also, think and explain in english.

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

Testing: Access http://localhost:4050/test_basics.asp or other test_*.asp files in www/.

ASP Debugging: Set <% debug_asp_code = "TRUE" %> in the ASP file for HTML stack traces.

Compilation Rule: ALWAYS compile Go code after editing to verify success. Do not compile for pure ASP edits.

3. Coding Standards & Conventions
Language: STRICT ENGLISH ONLY. Translate any non-English comments/UI/Code and any output or answer immediately.

VBScript Compatibility: Strict adherence to VBScript and ASP Classic standards.

Variable Lookup: MUST ALWAYS be Case-insensitive (and stored as lowercase internally).

Session/App: Sessions stored in temp/session (Cookie: ASPSESSIONID); Application state in memory.

Includes: file = relative to current; virtual = relative to www/.

Documentation: Keep instructions in this file. Sync copilot-instructions.md and GEMINI.md on updates. Do not create new .md explanation files.

New Libraries: Name as *_lib.go. Mimic VBScript nomenclature. Document in English.

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