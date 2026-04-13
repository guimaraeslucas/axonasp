# Server Object

## Overview

The `Server` object provides utility methods used throughout ASP scripts. Its most common uses are creating COM/native objects, translating virtual paths to physical file paths, and encoding strings for safe output or URL use. It is available in every ASP script with no setup.

---

## Methods

### Server.CreateObject

Instantiates a COM object or one of AxonASP's native library objects by its ProgID string. This is the main entry point for all built-in libraries such as G3JSON, G3DB, G3Mail, ADODB, Scripting.FileSystemObject, and so on.

**Syntax:**
```asp
Set obj = Server.CreateObject("ProgID")
```

**Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| ProgID | String | The programmatic identifier of the object to create |

**Return Value:** Object (use `Set` assignment)

**Example:**
```asp
<%
Set json = Server.CreateObject("G3JSON")
Dim data : data = json.Parse("{""name"":""Alice"",""age"":30}")
Response.Write data("name")    ' Alice
Set json = Nothing
%>
```

```asp
<%
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=...; Data Source=..."
' ...
conn.Close
Set conn = Nothing
%>
```

---

### Server.MapPath

Translates a virtual (URL-style) path into an absolute physical file system path on the server. This is needed whenever you must open a file using `FileSystemObject` or another file-access library.

**Syntax:**
```asp
Server.MapPath(path)
```

**Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| path | String | Virtual path. A leading `/` means the web root. A relative path is relative to the current script's directory. |

**Return Value:** String (absolute file system path)

**Example:**
```asp
<%
' Absolute virtual path
Dim fullPath
fullPath = Server.MapPath("/data/config.json")
Response.Write fullPath   ' e.g., C:\inetpub\wwwroot\data\config.json

' Relative path (relative to current script's directory)
Dim relPath
relPath = Server.MapPath("images/logo.png")
Response.Write relPath
%>
```

---

### Server.HTMLEncode

Escapes HTML special characters so that a string is safe to insert into an HTML page. Converts `<`, `>`, `&`, `"`, and `'` to their HTML entity equivalents.

**Syntax:**
```asp
Server.HTMLEncode(string)
```

**Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| string | String | The text to escape |

**Return Value:** String (HTML-safe text)

**Example:**
```asp
<%
Dim userInput : userInput = Request.Form("comment")
Response.Write "<p>" & Server.HTMLEncode(userInput) & "</p>"
' Prevents XSS: <script>alert(1)</script> becomes
' &lt;script&gt;alert(1)&lt;/script&gt;
%>
```

---

### Server.URLEncode

Encodes a string for safe use in a URL query string. Spaces become `+`, and special characters are percent-encoded.

**Syntax:**
```asp
Server.URLEncode(string)
```

**Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| string | String | The value to encode |

**Return Value:** String (URL-encoded)

**Example:**
```asp
<%
Dim searchTerm : searchTerm = "hello world & more"
Response.Redirect "/search.asp?q=" & Server.URLEncode(searchTerm)
' Redirects to /search.asp?q=hello+world+%26+more
%>
```

---

### Server.URLPathEncode

Encodes a string for use in a URL path segment. Slashes (`/`) are preserved as-is; only individual path components are encoded.

**Syntax:**
```asp
Server.URLPathEncode(string)
```

**Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| string | String | The path to encode |

**Return Value:** String (path-encoded)

**Example:**
```asp
<%
Dim folder : folder = "my documents/2026"
Dim encodedPath : encodedPath = Server.URLPathEncode(folder)
' Result: "my%20documents/2026"  (slash preserved)
%>
```

---

### Server.GetLastError

Returns the `ASPError` object populated with information about the most recent error in the current script. Useful inside custom error handlers.

**Syntax:**
```asp
Set err = Server.GetLastError()
```

**Return Value:** ASPError object

**Example:**
```asp
<%
On Error Resume Next
Dim x : x = 1 / 0   ' Division by zero
If Err.Number <> 0 Then
    Dim aspErr
    Set aspErr = Server.GetLastError()
    Response.Write "Error " & aspErr.Number & ": " & aspErr.Description
End If
%>
```

---

### Server.Execute

Executes another ASP file as if its content were inline in the current script. The called script shares the same `Request`, `Response`, `Session`, and `Application` objects. Execution returns to the calling page when the included script finishes.

**Syntax:**
```asp
Server.Execute(path)
```

**Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| path | String | Virtual path to the ASP file to execute |

**Return Value:** None

**Example:**
```asp
<%
' include a header component
Server.Execute "/includes/header.asp"
%>
<main>Page content here</main>
<%
Server.Execute "/includes/footer.asp"
%>
```

---

### Server.Transfer

Transfers control to another ASP file without issuing an HTTP redirect. The URL in the browser does not change. Unlike `Execute`, control does not return to the calling page.

**Syntax:**
```asp
Server.Transfer(path)
```

**Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| path | String | Virtual path to the ASP file |

**Return Value:** None (execution continues in the target file)

**Example:**
```asp
<%
If Session("UserID") = "" Then
    Server.Transfer "/errors/unauthorized.asp"
End If
%>
```

---

## Properties

### Server.ScriptTimeout

Gets or sets the maximum number of seconds that a script is allowed to run before the server terminates it with a timeout error. The default is set by `default_script_timeout` in `axonasp.toml` (default 60 seconds).

**Type:** Integer (read/write, seconds)

**Example:**
```asp
<%
' Allow a long-running export up to 5 minutes
Server.ScriptTimeout = 300

' ... export logic ...
%>
```

---

## Remarks

- `Server.MapPath` is sandboxed to the web root configured in `axonasp.toml`. Paths outside the web root resolve to the web root itself.
- `Server.HTMLEncode` should be applied to **all user-supplied content** before writing it to HTML output to prevent cross-site scripting (XSS).
- `Server.CreateObject` accepts the same ProgIDs as classic IIS ASP plus all AxonASP-specific libraries. A full list is in the Libraries section.
- `Server.Execute` and `Server.Transfer` can be used as alternatives to `#include` for dynamic component inclusion.
