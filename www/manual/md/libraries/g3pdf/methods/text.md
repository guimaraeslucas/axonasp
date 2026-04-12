# Text Method

## Overview

Writes text at absolute coordinates in the current PDF page.

## Syntax

```asp
result = obj.Text(...)
`````

## Parameters and Arguments

- x (Number, Required): X coordinate.
- y (Number, Required): Y coordinate.
- text (String, Required): Text to draw.
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
result = obj.Text()
If IsObject(result) Then
    Response.Write "Object returned"
Else
    Response.Write CStr(result)
End If
Set obj = Nothing
%>
`````





