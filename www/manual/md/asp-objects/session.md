# Session Object

## Overview

The `Session` object stores user-specific data that persists across multiple page requests within the same browsing session. AxonASP maintains session state by issuing a unique session cookie (`ASPSESSIONID`) to each visitor. Session data is stored in memory and flushed to disk asynchronously (interval configurable via `session_flush_interval_seconds` in `axonasp.toml`).

The `Session` object is always available and requires no initialization.

---

## Storing and Reading Session Values

Values are stored and retrieved using a dictionary-style syntax. Keys are case-insensitive strings; values can be any VBScript Variant â€” strings, numbers, arrays, or object references.

**Syntax:**
```asp
Session("key") = value
value = Session("key")
```

**Example â€” Login flow:**
```asp
<%
' login.asp â€” after verifying credentials
Session("UserID")   = 42
Session("UserName") = "alice"
Session("Role")     = "admin"
Response.Redirect "/dashboard.asp"
%>
```

```asp
<%
' dashboard.asp â€” check session
If Session("UserID") = "" Then
    Response.Redirect "/login.asp"
End If

Response.Write "Welcome, " & Session("UserName") & "!"
%>
```

---

## Methods

### Session.Abandon

Destroys the current session and all data stored in it. The session cookie is removed from the client. If the user makes another request after calling `Abandon`, a fresh session with a new ID is created.

**Syntax:**
```asp
Session.Abandon()
```

**Return Value:** None

**Example:**
```asp
<%
' logout.asp
Session.Abandon
Response.Redirect "/login.asp?msg=loggedout"
%>
```

---

## Properties

### Session.SessionID

Returns a unique identifier for the current session. In AxonASP, this is a string representation of the session cookie value. The numeric form is also available as an integer.

**Type:** String (read-only)

**Example:**
```asp
<%
Response.Write "Session ID: " & Session.SessionID
%>
```

---

### Session.Timeout

Gets or sets how long (in minutes) a session remains valid without activity. When the timeout expires, the session is abandoned and its data is discarded.

**Type:** Integer (read/write, minutes)  
**Default:** `20`

**Example:**
```asp
<%
' Extend session lifetime for premium users
If Session("Role") = "premium" Then
    Session.Timeout = 120
End If
%>
```

---

### Session.LCID

Gets or sets the locale identifier for the session. This affects how dates, times, and numbers are formatted when using VBScript formatting functions. A value of `1033` corresponds to English (United States).

**Type:** Integer (read/write)  
**Default:** Value from `default_mslcid` in `axonasp.toml`

**Example:**
```asp
<%
' Set locale to Brazilian Portuguese
Session.LCID = 1046
Response.Write FormatDateTime(Now(), 1)   ' Uses Portuguese date format
%>
```

---

### Session.CodePage

Gets or sets the code page used for string encoding within this session. `65001` is UTF-8.

**Type:** Integer (read/write)  
**Default:** `65001`

---

### Session.Contents

The `Contents` collection provides access to all session values set via `Session("key") = value`. You can enumerate it with `For Each` or use the `Remove` / `RemoveAll` methods to delete entries.

**Syntax:**
```asp
Session.Contents("key")
Session.Contents.Remove("key")
Session.Contents.RemoveAll()
```

**Example:**
```asp
<%
' List all session variables
Dim k
For Each k In Session.Contents
    Response.Write k & " = " & Session(k) & "<br>"
Next
%>
```

```asp
<%
' Remove a single value
Session.Contents.Remove("CartItems")
%>
```

---

### Session.StaticObjects

Contains objects declared with `<OBJECT>` tags in `global.asa` with `SCOPE="Session"`. These are instantiated the first time the session is accessed and are destroyed when the session ends.

**Type:** Collection (read-only enumeration)

---

## Remarks

- Session data is stored per user based on the `ASPSESSIONID` cookie. Cookies must be enabled in the browser for sessions to work.
- Storing **large objects** (arrays, nested dictionaries) in the session can increase memory and disk flush overhead. Store only essential identifiers and small values.
- `Session.Abandon` does not prevent a new session from being created if the user makes another request; it only clears the current session's data.
- In AxonASP, session values surviving a server restart depends on the `session_flush_interval_seconds` and `clean_sessions_on_startup` configuration settings.
- **Thread safety:** The session is protected internally; you do not need to `Lock`/`Unlock` for session access (unlike `Application`).
