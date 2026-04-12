# Fetch Method

## Overview

Fetches data from a remote source.

## Syntax

```asp
result = obj.Fetch(...)
```

## Parameters and Arguments

- url (String, Required): Absolute URL to request.
- method (String, Optional): HTTP verb (GET, POST, PUT, DELETE, etc.).
- headers (Variant, Optional): Header map/object.
- body (String, Optional): Request body payload.
- timeoutMs (Integer, Optional): Request timeout in milliseconds.
- Argument validation: invalid count or type raises runtime errors.

## Return Values

Returns a Variant result. Depending on the operation, this can be String, Boolean, Number, Array, Dictionary/object handle, or Empty.

## Remarks

- Method names are case-insensitive.
- Prefer explicit variable assignment and defensive checks before using returned values.
- For object values, use Set when assigning the return value.

## Code Example

```asp
<%
Option Explicit
Dim obj, result
Set obj = Server.CreateObject("G3HTTP")
result = obj.Fetch()
If IsObject(result) Then
    Response.Write "Object returned"
Else
    Response.Write CStr(result)
End If
Set obj = Nothing
%>
```



