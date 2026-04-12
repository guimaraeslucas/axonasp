# Application Object

## Overview

The `Application` object stores data that is shared across **all users and all requests** for the entire lifetime of the ASP application. Unlike `Session`, which is per-user, `Application` values are global â€” every visitor to your site reads from and writes to the same store.

Because multiple requests can run concurrently, you must call `Application.Lock` before writing and `Application.Unlock` immediately after to prevent data corruption.

---

## Storing and Reading Application Values

**Syntax:**
```asp
Application("key") = value
value = Application("key")
```

**Example â€” Page hit counter:**
```asp
<%
Application.Lock
Application("HitCount") = Application("HitCount") + 1
Application.Unlock

Response.Write "Page views since startup: " & Application("HitCount")
%>
```

---

## Methods

### Application.Lock

Prevents other scripts from modifying `Application` contents while you are performing a read-modify-write sequence. Only one request may hold the lock at a time; others will wait.

**Syntax:**
```asp
Application.Lock()
```

**Return Value:** None

**Important:** Always pair `Lock` with `Unlock`. Failing to unlock will block all other requests from ever writing to `Application` again.

---

### Application.Unlock

Releases the application lock previously acquired by `Lock`. This must be called after every `Lock`, even inside error handlers.

**Syntax:**
```asp
Application.Unlock()
```

**Return Value:** None

**Example â€” Safe counter update:**
```asp
<%
Application.Lock
On Error Resume Next
Application("Visitors") = Application("Visitors") + 1
If Err.Number <> 0 Then
    Application("Visitors") = 1
End If
Application.Unlock
%>
```

---

## Collections

### Application.Contents

Holds all values set via `Application("key") = value`. You can enumerate it with `For Each`, remove individual keys, or clear all entries.

**Syntax:**
```asp
Application.Contents("key")
Application.Contents.Remove("key")
Application.Contents.RemoveAll()
```

**Example â€” Enumerate application state:**
```asp
<%
Dim k
For Each k In Application.Contents
    Response.Write k & " = " & Application(k) & "<br>"
Next
%>
```

---

### Application.StaticObjects

Contains objects declared with `<OBJECT>` tags in `global.asa` that have `SCOPE="Application"`. These are instantiated once when the application starts and are destroyed only when the application shuts down.

**Type:** Collection (read-only enumeration)

---

## Remarks

- The `Lock` / `Unlock` pattern is **mandatory** any time you perform a read-modify-write cycle, such as incrementing a counter. Reading a value without modifying it generally does not require a lock, but be aware that the value may change between reads in a high-concurrency scenario.
- Do not store large objects (e.g., recordsets, file handles) in `Application`. Store only small, serializable values.
- Application state persists for the lifetime of the server process. If the server restarts, all `Application` values are lost unless your `global.asa` repopulates them in the `Application_OnStart` event.
- **Never** call `Response.End` or `Response.Redirect` while the application lock is held; this prevents the framework from calling `Unlock` and permanently blocks all writers.
