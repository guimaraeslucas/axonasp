# Process Method

## Overview

Processes current inputs using current configuration.

## Syntax

```asp
result = obj.Process(...)
`````

## Parameters and Arguments

- markdownText (String, Required): Markdown input content.
- options (Variant, Optional): Rendering options map.
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
Set obj = Server.CreateObject("G3MD")
result = obj.Process()
If IsObject(result) Then
    Response.Write "Object returned"
Else
    Response.Write CStr(result)
End If
Set obj = Nothing
%>
`````





