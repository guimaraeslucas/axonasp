# SetConnMaxLifetime Method

## Overview

The **SetConnMaxLifetime** method sets the maximum duration in seconds that a database connection may be reused in G3Pix AxonASP.

## Syntax

```asp
obj.SetConnMaxLifetime(seconds)
```

## Parameters and Arguments

- **seconds** (Integer, Required): The maximum lifetime for a connection in seconds. Use 0 for no limit.

## Return Values

Returns **Empty**.

## Remarks

- Connections older than the specified duration will be closed and removed from the pool, regardless of their idle state.
- This ensures that long-lived connections are periodically recycled, preventing potential memory leaks or stale connection issues on the database server.
- The default behavior is typically governed by the database driver, but this method allows for explicit control within G3Pix AxonASP.

## Code Example

```asp
<%
Dim db
Set db = Server.CreateObject("G3DB")

' Set connection lifetime to 1 hour (3600 seconds)
db.SetConnMaxLifetime 3600

If db.Open("mysql", "user:pass@tcp(localhost)/dbname") Then
    ' Operations...
    db.Close
End If

Set db = Nothing
%>
```
