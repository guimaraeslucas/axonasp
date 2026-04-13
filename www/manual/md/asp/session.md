# Manage Session State

## Overview

Use the `Session` object to persist user-specific values between HTTP requests in G3Pix AxonASP. The runtime identifies each visitor with an `ASPSESSIONID` cookie and maps that cookie to a server-side session store.

`Session` is built in. Do not instantiate it with `Server.CreateObject`.

## Prerequisites

- The client must accept cookies.
- Session storage must be enabled in your AxonASP runtime configuration.

## Store and Read Values

Use dictionary-style access to save and read values by key.

```asp
Session("UserID") = 42
Session("UserName") = "alice"

Dim currentUserName
currentUserName = Session("UserName")
```

### Return details for `Session("key")`

- Returns `Empty` when the key does not exist.
- Returns the exact value previously stored for that key when it exists.

## Methods

### Session.Abandon

Ends the current session and clears all values in that session.

#### Syntax

```asp
Session.Abandon
```

#### Return value

Returns no value.

#### Example

```asp
<%
Session.Abandon
Response.Redirect "/login.asp?msg=loggedout"
%>
```

## Properties

### Session.SessionID

Gets the current session identifier.

#### Type

String (read-only)

#### Return value

Returns a non-empty String that identifies the current session.

#### Example

```asp
<%
Response.Write "Session ID: " & Session.SessionID
%>
```

### Session.Timeout

Gets or sets the inactivity timeout in minutes.

#### Type

Integer (read/write)

#### Return value

Returns an Integer with the configured inactivity timeout.

#### Example

```asp
<%
If Session("Role") = "premium" Then
    Session.Timeout = 120
End If
%>
```

### Session.LCID

Gets or sets the locale identifier used by VBScript formatting functions in the current session.

#### Type

Integer (read/write)

#### Return value

Returns an Integer LCID value such as `1033`.

#### Example

```asp
<%
Session.LCID = 1033
Response.Write FormatDateTime(Now(), 1)
%>
```

### Session.CodePage

Gets or sets the session code page.

#### Type

Integer (read/write)

#### Return value

Returns an Integer code page value such as `65001`.

#### Example

```asp
<%
Session.CodePage = 65001
%>
```

### Session.Contents

Accesses all key/value pairs stored in the current session.

#### Syntax

```asp
Session.Contents("key")
Session.Contents.Remove "key"
Session.Contents.RemoveAll
```

#### Return value

- `Session.Contents("key")` returns `Empty` when the key does not exist.
- `Session.Contents("key")` returns the stored value when the key exists.
- `Session.Contents.Remove "key"` returns no value.
- `Session.Contents.RemoveAll` returns no value.

#### Example

```asp
<%
Dim k
For Each k In Session.Contents
    Response.Write k & " = " & CStr(Session.Contents(k)) & "<br>"
Next

Session.Contents.Remove "CartItems"
%>
```

### Session.StaticObjects

Gets the collection of session-scoped objects declared in `global.asa` with `SCOPE="Session"`.

#### Type

Collection (read-only)

#### Return value

Returns a collection object. It can be enumerated but not replaced.

## How It Works

- G3Pix AxonASP associates one `ASPSESSIONID` cookie with one server-side session record.
- Session keys are case-insensitive.
- When you call `Session.Abandon`, the current record is discarded. A later request can create a new session with a different ID.

## Best Practices

- Store compact values such as IDs and flags.
- Avoid large arrays and large object graphs in session state.
- Set `Session.Timeout` explicitly for workflows that require longer inactivity windows.
