# Ln Method

## Overview

Moves the PDF cursor to the next line position.

## Syntax

```asp
result = obj.Ln(...)
`````

## Parameters and Arguments

- height (Number, Optional): Vertical move after line break.
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
result = obj.Ln()
If IsObject(result) Then
    Response.Write "Object returned"
Else
    Response.Write CStr(result)
End If
Set obj = Nothing
%>
`````





