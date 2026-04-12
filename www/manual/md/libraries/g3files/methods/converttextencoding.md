# ConvertTextEncoding Method

## Overview

Converts Text Encoding between supported formats.

## Syntax

```asp
result = obj.ConvertTextEncoding(...)
```

## Parameters and Arguments

- text (String, Required): Input text.
- sourceEncoding (String, Required): Source encoding label.
- targetEncoding (String, Required): Target encoding label.
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
result = obj.ConvertTextEncoding()
If IsObject(result) Then
    Response.Write "Object returned"
Else
    Response.Write CStr(result)
End If
Set obj = Nothing
%>
```



