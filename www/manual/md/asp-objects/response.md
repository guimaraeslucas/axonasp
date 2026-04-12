# Response Object

## Overview

The `Response` object controls everything that AxonASP sends back to the client browser. Through it, you write HTML output, set HTTP headers and cookies, control caching, issue redirects, and manage response buffering. It is available in every ASP script with no setup required.

---

## Methods

### Response.Write

Writes a string directly to the HTTP response. This is the primary method for generating output in an ASP page. The `<%= expression %>` shortcut is equivalent to `Response.Write expression`.

**Syntax:**
```asp
Response.Write(string)
<%= expression %>
```

**Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| string | String/Variant | The content to output |

**Return Value:** None

**Example:**
```asp
<%
Response.Write "<h1>Hello, World!</h1>"
Response.Write "Today is: " & Date()
%>
<p>Current time: <%= Time() %></p>
```

---

### Response.BinaryWrite

Writes raw binary data to the response. Used when sending file downloads, images, or other non-text content directly from a byte array.

**Syntax:**
```asp
Response.BinaryWrite(data)
```

**Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| data | Variant (Byte Array) | Binary content to write |

**Return Value:** None

**Example:**
```asp
<%
Response.ContentType = "image/png"
Response.BinaryWrite imageByteArray
%>
```

---

### Response.Redirect

Issues an HTTP 302 redirect to the specified URL and immediately stops the current script. Any buffered output is discarded.

**Syntax:**
```asp
Response.Redirect(url)
```

**Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| url | String | Absolute or relative URL to redirect to |

**Return Value:** None (does not return â€” execution stops)

**Example:**
```asp
<%
If Session("UserID") = "" Then
    Response.Redirect "/login.asp?reason=notloggedin"
End If
' Execution continues only when user is logged in
%>
```

---

### Response.End

Immediately stops script execution and sends the current buffer to the client. Any code after `Response.End` does not run.

**Syntax:**
```asp
Response.End()
```

**Return Value:** None (does not return)

**Example:**
```asp
<%
If Request.QueryString("debug") = "1" Then
    Response.Write "Debug mode active"
    Response.End
End If
' Normal page code continues here
%>
```

---

### Response.Flush

Forces any currently buffered output to be sent to the client immediately. Useful for long-running operations where you want to show progress.

**Syntax:**
```asp
Response.Flush()
```

**Return Value:** None

**Example:**
```asp
<%
Response.Write "Starting process...<br>"
Response.Flush

' Simulate work
Dim i
For i = 1 To 5
    Response.Write "Step " & i & " complete<br>"
    Response.Flush
Next

Response.Write "Done."
%>
```

---

### Response.Clear

Discards all content currently in the output buffer. Only works when buffering is enabled. Has no effect after `Response.Flush` has been called.

**Syntax:**
```asp
Response.Clear()
```

**Return Value:** None

**Example:**
```asp
<%
On Error Resume Next
' Attempt to build output
Response.Write BuildPage()
If Err.Number <> 0 Then
    Response.Clear
    Response.Write "<p>An error occurred. Please try again.</p>"
End If
%>
```

---

### Response.AddHeader

Adds or replaces an HTTP response header. Headers must be set before any output is flushed to the client.

**Syntax:**
```asp
Response.AddHeader(name, value)
```

**Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| name | String | HTTP header name |
| value | String | Header value |

**Return Value:** None

**Example:**
```asp
<%
Response.AddHeader "X-Content-Type-Options", "nosniff"
Response.AddHeader "X-Frame-Options", "DENY"
Response.AddHeader "Content-Disposition", "attachment; filename=""report.csv"""
%>
```

---

### Response.AppendToLog

Appends a custom message to the AxonASP server log file (`temp/server.log`). Useful for application-level audit trails without file system access.

**Syntax:**
```asp
Response.AppendToLog(message)
```

**Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| message | String | Message to write to the log |

**Return Value:** None

**Example:**
```asp
<%
Response.AppendToLog "User " & Session("UserID") & " accessed report at " & Now()
%>
```

---

## Properties

### Response.Buffer

Controls whether output is buffered before being sent to the client. When `True` (the default), all output is held in memory until the script ends or `Flush` is called. When `False`, each `Response.Write` is sent immediately.

**Type:** Boolean (read/write)  
**Default:** `True`

**Example:**
```asp
<%
Response.Buffer = False   ' Disable buffering â€” immediate output
Response.Write "This line goes to the browser right away."
%>
```

---

### Response.ContentType

Sets the MIME type of the response. The default is `text/html`. Change this when serving JSON, XML, CSV, images, or other content types.

**Type:** String (read/write)  
**Default:** `"text/html"`

**Example:**
```asp
<%
Response.ContentType = "application/json"
Response.Write "{""status"": ""ok""}"
%>
```

---

### Response.Charset

Sets the character encoding appended to the `Content-Type` header. The value must be a valid IANA charset name.

**Type:** String (read/write)  
**Default:** `"utf-8"`

**Example:**
```asp
<%
Response.Charset = "utf-8"
%>
```

---

### Response.Status

Sets the HTTP status line. Changing this before any output is flushed overrides the default `200 OK`.

**Type:** String (read/write)  
**Default:** `"200 OK"`

**Example:**
```asp
<%
Response.Status = "404 Not Found"
Response.Write "<h1>Page not found</h1>"
%>
```

```asp
<%
Response.Status = "403 Forbidden"
Response.End
%>
```

---

### Response.CacheControl

Sets the value of the `Cache-Control` HTTP header. Use `"Public"` to allow proxies to cache the response or `"Private"` (the default) to restrict caching to the client only.

**Type:** String (read/write)  
**Default:** `"Private"`

**Example:**
```asp
<%
Response.CacheControl = "Public"
Response.Expires = 60   ' cache for 60 minutes
%>
```

---

### Response.Expires

Sets the number of minutes from the current time before the response expires. A value of `0` or negative means the content is already expired (forces revalidation).

**Type:** Integer (read/write, minutes)

**Example:**
```asp
<%
Response.Expires = 0     ' Force browsers to revalidate
Response.CacheControl = "no-cache"
%>
```

---

### Response.ExpiresAbsolute

Sets an absolute date and time for the response to expire, rather than a relative offset.

**Type:** Date/String (read/write)

**Example:**
```asp
<%
Response.ExpiresAbsolute = "April 30, 2026 00:00:00"
%>
```

---

### Response.Cookies

Provides access to the collection of response cookies. Each entry can be a simple string value or a collection of sub-key/value pairs. Cookie attributes (domain, path, expiration, secure, HttpOnly) are set as properties of each cookie entry.

**Syntax:**
```asp
Response.Cookies("name") = "value"
Response.Cookies("name").Expires = DateAdd("d", 30, Now())
Response.Cookies("name").Domain  = "example.com"
Response.Cookies("name").Path    = "/"
Response.Cookies("name").Secure  = True
Response.Cookies("name").HTTPOnly = True

' Dictionary-style cookie (sub-keys)
Response.Cookies("prefs")("theme") = "dark"
Response.Cookies("prefs")("lang")  = "en"
Response.Cookies("prefs").Expires  = DateAdd("d", 365, Now())
```

**Example:**
```asp
<%
Response.Cookies("SessionToken") = "abc123xyz"
Response.Cookies("SessionToken").Path    = "/"
Response.Cookies("SessionToken").Secure  = True
Response.Cookies("SessionToken").HTTPOnly = True
Response.Cookies("SessionToken").Expires = DateAdd("d", 1, Now())
%>
```

---

### Response.IsClientConnected

Returns `True` if the client is still connected and receiving the response. Useful to abort long-running scripts when the browser has disconnected.

**Type:** Boolean (read-only)

**Example:**
```asp
<%
Dim i
For i = 1 To 1000000
    If Not Response.IsClientConnected Then
        Response.End
    End If
    ' ... process record i ...
Next
%>
```

---

### Response.CodePage

Sets the code page used to convert Unicode strings to bytes in the response. `65001` is UTF-8.

**Type:** Integer (read/write)  
**Default:** `65001` (UTF-8)

---

### Response.PICS

Sets the value of a `PICS-Label` HTTP header for content rating systems.

**Type:** String (read/write)

---

## Remarks

- **Buffering** is enabled by default. The buffer has a configurable maximum size (`response_buffer_limit_mb` in `axonasp.toml`). Exceeding it raises a runtime error.
- Calling `Response.End` or `Response.Redirect` causes a non-resumable signal that terminates script execution immediately â€” code after them in the same script does not run.
- Setting headers or cookies after `Response.Flush` has been called has no effect; headers must be sent before any body bytes.
- `Response.Write` accepts any VBScript Variant; the value is coerced to a string before output.
