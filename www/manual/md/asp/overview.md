# ASP Intrinsic Objects - Quick Reference

## Overview

Classic ASP provides six built-in objects that are always available inside any ASP script without any `Server.CreateObject` call. These objects form the foundation of every ASP application â€” they give your scripts access to the incoming HTTP request, the outgoing response, user sessions, application-wide state, server utilities, and error information.

AxonASP implements all six objects with full compatibility with the original IIS behavior.

---

## At a Glance

| Object | Purpose |
|--------|---------|
| **Request** | Reads data sent by the browser (query string, form fields, cookies, server variables) |
| **Response** | Writes output to the browser and controls the HTTP response (headers, cookies, redirects) |
| **Session** | Stores data for a single user across multiple page requests |
| **Application** | Stores data shared across all users and all requests for the application lifetime |
| **Server** | Utility methods â€” URL/HTML encoding, path translation, object creation, error info |
| **ASPError** | Provides structured error information when a script error occurs |

---

## Request

Represents data sent to the server by the client.

**Key collections:**

| Collection | What it Contains |
|-----------|----------------|
| `Request.QueryString("key")` | URL query string parameters (`?name=value`) |
| `Request.Form("key")` | HTML form fields submitted via POST |
| `Request.Cookies("name")` | Client cookies |
| `Request.ServerVariables("var")` | HTTP headers and server environment info |
| `Request.ClientCertificate("key")` | Client certificate fields (HTTPS) |
| `Request("key")` | Searches all collections in order |

**Key properties:** `Request.TotalBytes`  
**Key method:** `Request.BinaryRead(n)`

---

## Response

Controls what is sent back to the browser.

**Key methods:**

| Method | Purpose |
|--------|---------|
| `Response.Write(text)` | Output text to browser |
| `Response.Redirect(url)` | 302 redirect and stop |
| `Response.End()` | Immediately stop script execution |
| `Response.Flush()` | Send buffered content immediately |
| `Response.Clear()` | Clear the buffer |
| `Response.AddHeader(name, value)` | Add HTTP header |
| `Response.AppendToLog(message)` | Write to the server log |
| `Response.BinaryWrite(data)` | Send raw byte array |

**Key properties:** `Response.Buffer`, `Response.ContentType`, `Response.Status`, `Response.Charset`, `Response.CacheControl`, `Response.Expires`, `Response.ExpiresAbsolute`, `Response.Cookies`, `Response.IsClientConnected`

---

## Session

Stores user-specific data between requests using a cookie-based session ID.

**Usage:**
```asp
Session("UserName") = "Alice"
Dim name : name = Session("UserName")
Session.Abandon
```

**Key properties:** `Session.SessionID`, `Session.Timeout`, `Session.LCID`, `Session.CodePage`  
**Key methods:** `Session.Abandon()`  
**Key collections:** `Session.Contents`, `Session.StaticObjects`

---

## Application

Application-level data store shared among all users. Must use `Lock/Unlock` to prevent race conditions.

**Usage:**
```asp
Application.Lock
Application("HitCount") = Application("HitCount") + 1
Application.Unlock
```

**Key methods:** `Application.Lock()`, `Application.Unlock()`  
**Key collections:** `Application.Contents`, `Application.StaticObjects`

---

## Server

Provides utility methods used throughout scripts.

**Key methods:**

| Method | Purpose |
|--------|---------|
| `Server.CreateObject(progID)` | Instantiate a COM/native object |
| `Server.MapPath(path)` | Translate virtual path to physical path |
| `Server.HTMLEncode(str)` | Escape HTML special characters |
| `Server.URLEncode(str)` | Encode string for use in a URL query |
| `Server.URLPathEncode(str)` | Encode URL path segments |
| `Server.GetLastError()` | Return the current ASPError object |
| `Server.Execute(path)` | Execute another ASP file inline |
| `Server.Transfer(path)` | Transfer execution to another page |

**Key property:** `Server.ScriptTimeout`

---

## ASPError

Returned by `Server.GetLastError()` after an error occurs.

**Key properties:** `ASPCode`, `Number`, `Source`, `Description`, `File`, `Line`, `Column`, `Category`

---

## Detailed Pages

- Request Object
- Response Object
- Session Object
- Application Object
- Server Object
- ASPError Object
