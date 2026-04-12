# SetFont Method

## Overview

Sets Font for the G3PDF library.

## Syntax

```asp
result = obj.SetFont(...)
`````

## Parameters and Arguments

- family (String, Required): Font family name.
- style (String, Optional): Style flags (B, I, U).
- size (Number, Required): Font size.
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
Set obj = Server.CreateObject("G3PDF")
result = obj.SetFont()
If IsObject(result) Then
    Response.Write "Object returned"
Else
    Response.Write CStr(result)
End If
Set obj = Nothing
%>
`````





