# ConvertFileEncoding Method

## Overview

Converts File Encoding between supported formats.

## Syntax

```asp
result = obj.ConvertFileEncoding(...)
```

## Parameters and Arguments

- filePath (String, Required): Input file path.
- sourceEncoding (String, Required): Source encoding label.
- targetEncoding (String, Required): Target encoding label.
- outputPath (String, Optional): Optional output path; overwrite behavior depends on implementation.
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
Set obj = Server.CreateObject("G3FILES")
result = obj.ConvertFileEncoding()
If IsObject(result) Then
    Response.Write "Object returned"
Else
    Response.Write CStr(result)
End If
Set obj = Nothing
%>
```



