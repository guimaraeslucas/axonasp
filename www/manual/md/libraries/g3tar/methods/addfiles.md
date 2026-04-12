# AddFiles Method

## Overview

Adds Files to the current operation context.

## Syntax

```asp
result = obj.AddFiles(...)
`````

## Parameters and Arguments

- sourcePaths (Variant, Required): Array/list of source file paths.
- archiveRoot (String, Optional): Root folder inside archive.
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
Set obj = Server.CreateObject("G3TAR")
result = obj.AddFiles()
If IsObject(result) Then
    Response.Write "Object returned"
Else
    Response.Write CStr(result)
End If
Set obj = Nothing
%>
`````





