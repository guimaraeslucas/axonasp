# AddFile Method

## Overview

Adds File to the current operation context.

## Syntax

```asp
result = obj.AddFile(...)
```

## Parameters and Arguments

- filePath (String, Required): Source file path.
- archiveName (String, Optional): Name/path inside archive.
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
Set obj = Server.CreateObject("G3ZIP")
result = obj.AddFile()
If IsObject(result) Then
    Response.Write "Object returned"
Else
    Response.Write CStr(result)
End If
Set obj = Nothing
%>
```



