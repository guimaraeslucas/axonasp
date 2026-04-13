# Axgetconfigkeys

## Overview

Returns an array with all available AxonASP configuration keys.

## Syntax

```asp
result = obj.Axgetconfigkeys(...)
```

## Parameters and Arguments

- Parameters (Variant, Optional): This method accepts arguments according to runtime dispatch behavior.
- Validation: argument count and type checks are handled at runtime by AxonASP.

## Return Values

- Returns an Array or a String.


## Remarks

- Method names are case-insensitive.
- For object return values, use Set when assigning the return value.

## Code Example

```asp
<%
Option Explicit
Dim obj, result
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")
result = obj.Axgetconfigkeys()
If IsObject(result) Then
    Response.Write "Object returned"
Else
    Response.Write CStr(result)
End If
Set obj = Nothing
%>
```


