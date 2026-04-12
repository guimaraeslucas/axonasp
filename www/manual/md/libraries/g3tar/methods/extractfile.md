# ExtractFile Method

## Overview

Extracts File from an archive into output paths.

## Syntax

```asp
result = obj.ExtractFile(...)
`````

## Parameters and Arguments

- archiveName (String, Required): Entry path/name inside archive.
- outputPath (String, Required): Destination file path.
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
result = obj.ExtractFile()
If IsObject(result) Then
    Response.Write "Object returned"
Else
    Response.Write CStr(result)
End If
Set obj = Nothing
%>
`````





