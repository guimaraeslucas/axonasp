# SetColor Method

## Overview

Sets Color for the G3IMAGE library.

## Syntax

```asp
result = obj.SetColor(...)
```

## Parameters and Arguments

- red (Integer, Required): 0-255.
- green (Integer, Required): 0-255.
- blue (Integer, Required): 0-255.
- alpha (Integer, Optional): 0-255 or library-specific range.
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
result = obj.SetColor()
If IsObject(result) Then
    Response.Write "Object returned"
Else
    Response.Write CStr(result)
End If
Set obj = Nothing
%>
```



