# RenderViaTemp Method

## Overview

Renders Via Temp from the current operation context.

## Syntax

```asp
result = obj.RenderViaTemp(...)
```

## Parameters and Arguments

- tempPath (String, Optional): Temporary file path or directory.
- contentType (String, Optional): Output mime type hint for response integration.
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
Set obj = Server.CreateObject("G3IMAGE")
result = obj.RenderViaTemp()
If IsObject(result) Then
    Response.Write "Object returned"
Else
    Response.Write CStr(result)
End If
Set obj = Nothing
%>
```



