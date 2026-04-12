# OpenFromEnv Method

## Overview

Opens From Env for subsequent operations.

## Syntax

```asp
result = obj.OpenFromEnv(...)
`````

## Parameters and Arguments

- envKey (String, Required): Environment variable containing DSN/connection string.
- driverName (String, Optional): Driver/provider name when not embedded in env value.
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
Set obj = Server.CreateObject("G3DB")
result = obj.OpenFromEnv()
If IsObject(result) Then
    Response.Write "Object returned"
Else
    Response.Write CStr(result)
End If
Set obj = Nothing
%>
`````





