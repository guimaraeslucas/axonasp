# Write Method

## Overview

Writes content to the active output target.

## Syntax

```asp
result = obj.Write(...)
```

## Parameters and Arguments

- path (String, Required): Target file path.
- content (String, Required): Text content to write.
- encoding (String, Optional): Text encoding, default UTF-8.
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
Set obj = Server.CreateObject("G3FILES")
result = obj.Write()
If IsObject(result) Then
    Response.Write "Object returned"
Else
    Response.Write CStr(result)
End If
Set obj = Nothing
%>
```



