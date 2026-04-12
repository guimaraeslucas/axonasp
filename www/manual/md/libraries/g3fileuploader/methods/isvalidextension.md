# IsValidExtension Method

## Overview

Checks whether Valid Extension is valid or enabled.

## Syntax

```asp
result = obj.IsValidExtension(...)
`````

## Parameters and Arguments

- fileNameOrExtension (String, Required): Candidate file name or extension to validate.
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
Set obj = Server.CreateObject("G3FILEUPLOADER")
result = obj.IsValidExtension()
If IsObject(result) Then
    Response.Write "Object returned"
Else
    Response.Write CStr(result)
End If
Set obj = Nothing
%>
`````





