# MeasureString Method

## Overview

Measures String and returns dimension metrics.

## Syntax

```asp
result = obj.MeasureString(...)
```

## Parameters and Arguments

- text (String, Required): Text to measure.
- fontSize (Number, Optional): Font size used for measurement.
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
result = obj.MeasureString()
If IsObject(result) Then
    Response.Write "Object returned"
Else
    Response.Write CStr(result)
End If
Set obj = Nothing
%>
```



