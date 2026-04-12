# Recordset.Requery Method

## Overview

The Recordset.Requery method is exposed by the ADODB.Connection object in AxonASP.

## Syntax

```asp
result = obj.Recordset.Requery(...)
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
Set obj = Server.CreateObject("ADODB.Connection")
result = obj.Recordset.Requery()
If IsObject(result) Then
    Response.Write "Object returned"
Else
    Response.Write CStr(result)
End If
Set obj = Nothing
%>
```