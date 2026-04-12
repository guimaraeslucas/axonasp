# Axgetenv Method

## Overview

Reads an environment variable by name and returns its value.

## Syntax

```asp
result = obj.Axgetenv(...)
```

## Parameters and Arguments

- Parameters (Variant, Optional): This method accepts arguments according to runtime dispatch behavior.
- Validation: argument count and type checks are handled at runtime by AxonASP.

## Return Values

- Returns a Variant compatible with Classic ASP/VBScript.
- Depending on operation, the result can be String, Boolean, Number, Array, or Empty.

## Remarks

- Method names are case-insensitive.
- For object return values, use Set when assigning the return value.

## Code Example

```asp
<%
Option Explicit
Dim obj, result
Set obj = Server.CreateObject("G3AXON.Functions")
result = obj.Axgetenv()
If IsObject(result) Then
    Response.Write "Object returned"
Else
    Response.Write CStr(result)
End If
Set obj = Nothing
%>
```


