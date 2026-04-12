# Open Method

## Overview

The Open method is exposed by the ADODB.Stream object in AxonASP.

## Syntax

```asp
result = obj.Open(...)
```
## Parameters and Arguments

- Parameters (Variant, Optional): Accepted arguments depend on runtime dispatch for this object.
- Argument validation: Invalid argument count or types raise runtime errors.

## Return Values

Returns a Variant result. Depending on operation, this can be String, Boolean, Number, Array, object handle, or Empty.

## Remarks

- Method names are case-insensitive.
- Use Set for object return values.

## Code Example

```asp
<%
Option Explicit
Dim obj, result
Set obj = Server.CreateObject("ADODB.Stream")
result = obj.Open()
If IsObject(result) Then
    Response.Write "Object returned"
Else
    Response.Write CStr(result)
End If
Set obj = Nothing
%>
```