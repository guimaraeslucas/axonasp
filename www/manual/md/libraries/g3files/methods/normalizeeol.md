# NormalizeEOL Method

## Overview

Normalizes EOL to a consistent format.

## Syntax

```asp
result = obj.NormalizeEOL(...)
```

## Parameters and Arguments

- text (String, Required): Input text.
- eolStyle (String, Optional): Target line ending (CRLF, LF).
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
result = obj.NormalizeEOL()
If IsObject(result) Then
    Response.Write "Object returned"
Else
    Response.Write CStr(result)
End If
Set obj = Nothing
%>
```



