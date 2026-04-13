# Request Object

## Overview

The `Request` object provides access to all information sent to the server by the current HTTP request. This includes query string parameters, HTML form data, cookies, server environment variables, and client certificate fields. You can read any of these through named collections, and the `Request` object itself supports a default shortcut that searches all collections in sequence.

The `Request` object is always available in every ASP script â€” no `Server.CreateObject` is needed.

---

## Collections

### Request.QueryString

Returns values from the URL query string. For example, the URL `/page.asp?color=blue&size=10` populates `QueryString` with `color` and `size`.

**Syntax:**
```asp
Request.QueryString("key")
Request.QueryString("key")(index)   ' 1-based index for multi-value keys
Request.QueryString("key").Count
```

**Example:**
```asp
<%
Dim color
color = Request.QueryString("color")
If color = "" Then
    color = "red"
End If
Response.Write "Selected color: " & color
%>
```

```asp
<%
' Enumerate all query string parameters
Dim k
For Each k In Request.QueryString
    Response.Write k & " = " & Request.QueryString(k) & "<br>"
Next
%>
```

---

### Request.Form

Returns values submitted via an HTML form using the `POST` method. The keys correspond to the `name` attributes of the form elements.

**Syntax:**
```asp
Request.Form("fieldname")
Request.Form("fieldname")(index)   ' 1-based, for multi-select
Request.Form("fieldname").Count
```

**Note:** Accessing `Request.Form` and calling `Request.BinaryRead` in the same script is not allowed. Use one or the other.

**Example:**
```asp
<%
Dim username, password
username = Request.Form("username")
password = Request.Form("password")

If Len(Trim(username)) = 0 Then
    Response.Write "Username is required."
    Response.End
End If

Response.Write "Welcome, " & Server.HTMLEncode(username)
%>
```

---

### Request.Cookies

Returns cookie values sent by the browser. Cookies with sub-keys (created using `Response.Cookies("name")("subkey")`) can be read with an extra subscript.

**Syntax:**
```asp
Request.Cookies("name")
Request.Cookies("name")("subkey")
Request.Cookies("name").HasKeys        ' True if cookie has sub-keys
```

**Example:**
```asp
<%
Dim sessionToken
sessionToken = Request.Cookies("AuthToken")
If sessionToken = "" Then
    Response.Redirect "/login.asp"
End If
%>
```

---

### Request.ServerVariables

Returns the value of HTTP headers and server environment variables. The variable names are standardized strings like `HTTP_HOST`, `REMOTE_ADDR`, `REQUEST_METHOD`, and so on.

**Syntax:**
```asp
Request.ServerVariables("variable_name")
```

**Common Variables:**

| Variable | Description |
|----------|-------------|
| `HTTP_HOST` | The hostname from the request |
| `HTTP_REFERER` | The referring page URL |
| `HTTP_USER_AGENT` | Browser/client identification string |
| `REMOTE_ADDR` | Client IP address |
| `REQUEST_METHOD` | `GET`, `POST`, `PUT`, etc. |
| `SCRIPT_NAME` | Path of the current script |
| `SERVER_NAME` | Server hostname |
| `SERVER_PORT` | Port number (e.g., `80`) |
| `SERVER_PROTOCOL` | `HTTP/1.1` |
| `HTTPS` | `on` if HTTPS, otherwise empty |
| `QUERY_STRING` | Raw query string |
| `CONTENT_TYPE` | Content type of POST body |
| `CONTENT_LENGTH` | Byte count of the POST body |

**Example:**
```asp
<%
Dim clientIP, method
clientIP = Request.ServerVariables("REMOTE_ADDR")
method   = Request.ServerVariables("REQUEST_METHOD")

Response.Write "Your IP: " & clientIP & "<br>"
Response.Write "Method:  " & method
%>
```

---

### Request.ClientCertificate

Returns fields from the client's SSL certificate when the connection uses certificate-based authentication.

**Syntax:**
```asp
Request.ClientCertificate("key")
```

**Common Keys:** `Subject`, `Issuer`, `ValidFrom`, `ValidUntil`, `SerialNumber`, `Certificate`

---

### Request (Default Collection)

When a key is passed directly to `Request(key)`, AxonASP searches all collections in this order:
1. `QueryString`
2. `Form`
3. `Cookies`
4. `ClientCertificate`
5. `ServerVariables`

This is a convenience shortcut. For clarity and performance, prefer specifying the collection explicitly.

**Example:**
```asp
<%
' Reads from whichever collection contains "id"
Dim id
id = Request("id")
%>
```

---

## Properties

### TotalBytes

Returns the total number of bytes in the request body. This is useful before calling `BinaryRead`.

**Syntax:**
```asp
Request.TotalBytes
```

**Type:** Long (read-only)

**Example:**
```asp
<%
Dim bodySize
bodySize = Request.TotalBytes
Response.Write "Body size: " & bodySize & " bytes"
%>
```

---

## Methods

### BinaryRead

Reads the raw bytes from the POST body. This is used when you need to process binary data â€” for example, a file upload or raw JSON body not processed through `Request.Form`.

After calling `BinaryRead`, accessing `Request.Form` is not permitted in the same request.

**Syntax:**
```asp
Request.BinaryRead(count)
```

**Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| count | Long | Number of bytes to read |

**Return Value:** Variant (Byte Array)

**Example:**
```asp
<%
Dim rawBytes, byteCount
byteCount = Request.TotalBytes
rawBytes  = Request.BinaryRead(byteCount)
' rawBytes is now a byte safearray containing the POST body
%>
```

---

## Remarks

- All collection lookups are **case-insensitive**.
- Multi-value form fields and query string parameters are accessible by 1-based numeric index.
- Iterating `For Each k In Request.QueryString` yields collection keys in insertion order.
- `Request.Form` is lazily loaded on first access; calling `BinaryRead` before it prevents form parsing.
- The default collection lookup order (`QueryString â†’ Form â†’ Cookies â†’ ClientCertificate â†’ ServerVariables`) applies only when you omit the collection name.
