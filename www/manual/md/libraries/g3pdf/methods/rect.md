# Rect Method

## Overview

Draws a rectangle in the current PDF page.

## Syntax

```asp
result = obj.Rect(...)
`````

## Parameters and Arguments

- x (Number, Required): Left.
- y (Number, Required): Top.
- width (Number, Required): Width.
- height (Number, Required): Height.
- style (String, Optional): Draw style (D/F/DF).
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
result = obj.Rect()
If IsObject(result) Then
    Response.Write "Object returned"
Else
    Response.Write CStr(result)
End If
Set obj = Nothing
%>
`````





