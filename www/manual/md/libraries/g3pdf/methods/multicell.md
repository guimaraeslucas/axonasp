# MultiCell Method

## Overview

Writes wrapped multi-line text blocks in the current PDF page.

## Syntax

```asp
result = obj.MultiCell(...)
`````

## Parameters and Arguments

- width (Number, Required): Cell width.
- height (Number, Required): Line height.
- text (String, Required): Multi-line text.
- border (Variant, Optional): Border mode.
- align (String, Optional): Alignment.
- fill (Boolean, Optional): Fill background.
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
result = obj.MultiCell()
If IsObject(result) Then
    Response.Write "Object returned"
Else
    Response.Write CStr(result)
End If
Set obj = Nothing
%>
`````





