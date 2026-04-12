# Cell Method

## Overview

Writes a positioned cell into the current PDF page.

## Syntax

```asp
result = obj.Cell(...)
`````

## Parameters and Arguments

- width (Number, Required): Cell width.
- height (Number, Required): Cell height.
- text (String, Optional): Cell text.
- border (Variant, Optional): Border mode.
- ln (Integer, Optional): Line break behavior.
- align (String, Optional): Horizontal alignment.
- fill (Boolean, Optional): Fill background when true.
- link (String, Optional): Optional link target.
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
result = obj.Cell()
If IsObject(result) Then
    Response.Write "Object returned"
Else
    Response.Write CStr(result)
End If
Set obj = Nothing
%>
`````





