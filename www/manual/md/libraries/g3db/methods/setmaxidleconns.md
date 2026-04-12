# SetMaxIdleConns Method

## Overview

Sets Max Idle Conns for the G3DB library.

## Syntax

```asp
result = obj.SetMaxIdleConns(...)
`````

## Parameters and Arguments

- maxIdle (Integer, Required): Maximum number of idle connections.
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
result = obj.SetMaxIdleConns()
If IsObject(result) Then
    Response.Write "Object returned"
Else
    Response.Write CStr(result)
End If
Set obj = Nothing
%>
`````





