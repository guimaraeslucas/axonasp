# DrawRectangle Method

## Overview

Draws Rectangle onto the current image canvas.

## Syntax

```asp
result = obj.DrawRectangle(...)
```

## Parameters and Arguments

- x (Number, Required): Left.
- y (Number, Required): Top.
- width (Number, Required): Rectangle width.
- height (Number, Required): Rectangle height.
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
result = obj.DrawRectangle()
If IsObject(result) Then
    Response.Write "Object returned"
Else
    Response.Write CStr(result)
End If
Set obj = Nothing
%>
```



