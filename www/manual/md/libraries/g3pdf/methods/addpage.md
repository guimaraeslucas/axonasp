# AddPage Method

## Overview

Adds Page to the current operation context.

## Syntax

```asp
result = obj.AddPage(...)
`````

## Parameters and Arguments

- orientation (String, Optional): Page orientation (P/L).
- size (String, Optional): Page size (A4, Letter, etc.).
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
result = obj.AddPage()
If IsObject(result) Then
    Response.Write "Object returned"
Else
    Response.Write CStr(result)
End If
Set obj = Nothing
%>
`````





