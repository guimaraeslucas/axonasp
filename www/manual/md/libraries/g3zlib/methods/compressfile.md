# CompressFile Method

## Overview

Compresses File into a compact representation.

## Syntax

```asp
result = obj.CompressFile(...)
```

## Parameters and Arguments

- inputPath (String, Required): Source file path.
- outputPath (String, Required): Compressed output path.
- level (Integer, Optional): Compression level.
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
Set obj = Server.CreateObject("G3ZLIB")
result = obj.CompressFile()
If IsObject(result) Then
    Response.Write "Object returned"
Else
    Response.Write CStr(result)
End If
Set obj = Nothing
%>
```



