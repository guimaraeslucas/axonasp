# DrawStringAnchored Method

## Overview

Draws String Anchored onto the current image canvas.

## Syntax

```asp
result = obj.DrawStringAnchored(...)
```

## Parameters and Arguments

- text (String, Required): Text to draw.
- x (Number, Required): Anchor X.
- y (Number, Required): Anchor Y.
- anchor (String, Optional): Alignment anchor (for example, center).
- fontSize (Number, Optional): Font size.
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
result = obj.DrawStringAnchored()
If IsObject(result) Then
    Response.Write "Object returned"
Else
    Response.Write CStr(result)
End If
Set obj = Nothing
%>
```



