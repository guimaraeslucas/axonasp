# Image Method

## Overview

Places an image into the current PDF page.

## Syntax

```asp
result = obj.Image(...)
`````

## Parameters and Arguments

- imagePath (String, Required): Image file path.
- x (Number, Required): X coordinate.
- y (Number, Required): Y coordinate.
- width (Number, Optional): Render width.
- height (Number, Optional): Render height.
- imageType (String, Optional): Type hint (PNG/JPG/etc.).
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
result = obj.Image()
If IsObject(result) Then
    Response.Write "Object returned"
Else
    Response.Write CStr(result)
End If
Set obj = Nothing
%>
`````





