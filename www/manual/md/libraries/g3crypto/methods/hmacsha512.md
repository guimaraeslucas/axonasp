# HmacSHA512 Method

## Overview

Computes a cryptographic result using the HmacSHA512 operation.

## Syntax

```asp
result = obj.HmacSHA512(...)
```

## Parameters and Arguments

- key (String, Required): Secret key.
- message (String, Required): Message to sign.
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
Set obj = Server.CreateObject("G3Crypto")
result = obj.HmacSHA512()
If IsObject(result) Then
    Response.Write "Object returned"
Else
    Response.Write CStr(result)
End If
Set obj = Nothing
%>
```



