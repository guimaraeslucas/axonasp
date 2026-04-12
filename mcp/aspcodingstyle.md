STRICT CLASSIC ASP & VBSCRIPT CODING GUIDELINES

Overview: Classic ASP (Active Server Pages) powered by VBScript 5.8 has a highly specific syntax that predates modern programming paradigms. To ensure maximum compatibility, performance, and stability, all ASP code must strictly adhere to the following Microsoft IIS legacy standards. Follow STRICT guidelines to avoid syntax errors, runtime errors, and compatibility issues when running on the AxonASP platform.

1. ASP Directives and Delimiters
Modern parsers might be forgiving, but strict Classic ASP requires exact formatting for page directives.
* No Line Breaks in Directives: The processing directive `<%@ ... %>` must exist entirely on the very first line of the file. You cannot break it across multiple lines.
* No Extra Spacing: Keep directives compact. Do not add unnecessary spaces around the equal signs.
    * Correct: `<%@Language="VBSCRIPT" CodePage="65001"%>`
    * Incorrect: `<%@ Language = "VBSCRIPT" %>`
* Tag Closure:Never leave `<%` or `%>` tags unclosed or mismatched.

2. Control Structures (The "If-Then" Rule)
VBScript control structures are rigid. Modern shortcuts do not apply.
* Mandatory Block `If...Then...End If`: You must never use single-line `If...Then` statements (e.g., `If x = 1 Then y = 2`). While legacy VBScript technically allowed this, it is notoriously prone to parsing errors and scope bleeding in complex ASP pages. You MUST always use block syntax and explicitly close the statement with `End If`.
    * Correct:
        If x = 1 Then
            y = 2
        End If 
    * Incorrect: `If x = 1 Then y = 2`
* Loop Closures: `For` loops must end with `Next` (optionally `Next variableName`), `Do While` must end with `Loop`, and `While` must end with `Wend`.

3. Variable Declaration & Initialization
VBScript lacks modern variable declaration features.
* Mandatory `Option Explicit`: Every single VBScript page or included file should ideally enforce `Option Explicit` to prevent undeclared variables.
* No Inline Initialization: You cannot declare and assign a variable on the same line.
    * Correct:
        Dim myVar
        myVar = "Hello"
    * Incorrect: `Dim myVar = "Hello"`
* Everything is a Variant: There is no strong typing in VBScript. Do not attempt to type variables (e.g., `Dim x As Integer` will throw a syntax error).

4. Object Assignment (`Set` vs. `Let`)
This is the most common pitfall for modern developers writing VBScript.
* The `Set` Keyword: Whenever you are assigning an Object (like a Recordset, Connection, FSO, or Dictionary) to a variable, you must use the `Set` keyword.
    * Correct: `Set rs = Server.CreateObject("ADODB.Recordset")`
    * Incorrect: `rs = Server.CreateObject("ADODB.Recordset")`
* The `Nothing` Keyword: To destroy an object and free memory, you must explicitly set it to `Nothing`.
    * Correct: `Set rs = Nothing`

5. Method and Function Calling (Parentheses Rules)
VBScript's rules for parentheses are highly idiosyncratic compared to modern languages.
* Calling Subs or Methods (No Return Value): If you are calling a Sub or a method without expecting a return value, do not use parentheses, OR use the `Call` keyword with parentheses.
    * Correct: `Response.Write "Hello World"`
    * Correct: `Call Response.Write("Hello World")`
    * Incorrect: `Response.Write("Hello World")` (This passes the argument ByVal and can cause syntax errors with multiple arguments).
* Calling Functions (Expecting a Return Value): If you are assigning the result to a variable, you must use parentheses.
    * Correct: `myLen = Len("Hello")`

6. Major Quirks vs. Modern Languages
* No Short-Circuit Logic: In an `If A And B Then` statement, VBScript evaluates both `A` and `B`, even if `A` is false. This will cause crashes if `B` relies on `A` being true (e.g., `If Not rs.EOF And rs("Name") = "John" Then` will crash if the recordset is at EOF). You must nest your `If` statements.
* Error Handling: There is no `try...catch...finally` block. You must use inline error handling.
    * Enable: `On Error Resume Next`
    * Check: `If Err.Number <> 0 Then ...`
    * Disable/Reset: `On Error GoTo 0`
* String Concatenation: Always use the ampersand (`&`) to concatenate strings, never the plus sign (`+`), as `+` can inadvertently cause mathematical addition if both variants evaluate to numbers.
* Equality Operator: A single equals sign (`=`) acts as both the assignment operator AND the equality comparison operator. There is no `==` or `===`.
    * Assignment: `x = 5`
    * Comparison: `If x = 5 Then`
* Line Continuation: Statements cannot naturally break across lines. You must use the underscore character (`_`) preceded by a space to span a single logical line of code across multiple physical lines.

7. Server-Side Object Creation
When using `Server.CreateObject`, you must use the exact ProgID string as defined by the library or component. The server will return a native object handle that is compatible with VBScript's late binding. Always check the documentation for the correct ProgID and ensure that the corresponding library is properly registered on the server.

8. Use of the MCP Server and AxonASP Libraries
When creating objects with `Server.CreateObject`, prefer using the ProgIDs provided by the AxonASP manual for maximum compatibility and performance. These ProgIDs are designed to leverage the native implementations in AxonASP, which are optimized for the Classic ASP environment.
If you have access to the AxonASP MCP server, you should query it for information on code formatting, as well as available server functions and properties. If the MCP server is unavailable, refer to the manual located in ./www/manual/md or online at https://github.com/guimaraeslucas/axonasp/tree/main/www/manual/md/authoring/llm-classic-asp-coding.md.
The following objects are available natively. To ensure maximum performance and efficiency, always prioritize using these built-in objects over writing custom ASP scripts for the same functionality: G3MD, G3CRYPTO, G3Axon, G3JSON, G3DB, G3HTTP, G3MAIL, G3Image, G3FILES, G3Template, G3Zip, G3ZLIB, G3TAR, G3ZSTD, G3FC, WScript.Shell, ADOX.Catalog, MSWC.AdRotator, MSWC.BrowserType, MSWC.NextLink, MSWC.ContentRotator, MSWC.Counters, MSWC.PageCounter, MSWC.Tools, MSWC.MyInfo, MSWC.PermissionChecker, MSXML2.ServerXMLHTTP, MSXML2.XMLHTTP, Microsoft.XMLHTTP, MSXML2.DOMDocument, Microsoft.XMLDOM, G3PDF, G3FileUploader and upload compatibility aliases, Scripting.FileSystemObject, Scripting.Dictionary, ADODB.Stream, ADODB.Connection, ADODBOLE.Connection, ADODB.Recordset, ADODB.Command, VBScript.RegExp and RegExp.
```

## Code Example
```asp
<%@Language="VBSCRIPT"%>
<%
Option Explicit

Dim conn
Dim rs
Dim sql

Set conn = Server.CreateObject("ADODB.Connection")
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("./data.mdb")
conn.Open

sql = "SELECT name FROM users"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql, conn

If Not rs.EOF Then
    Response.Write rs("name")
End If

rs.Close
conn.Close
Set rs = Nothing
Set conn = Nothing
%>
```