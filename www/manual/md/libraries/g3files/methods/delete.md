# Delete Method

## Overview

Executes the Delete operation provided by the G3FILES library.

## Syntax

```asp
result = obj.Delete(...)
```

## Parameters and Arguments

- path (String, Required): File or folder path.
- recursive (Boolean, Optional): When true, deletes directories recursively.
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
result = obj.Delete()
If IsObject(result) Then
    Response.Write "Object returned"
Else
    Response.Write CStr(result)
End If
Set obj = Nothing
%>
```



