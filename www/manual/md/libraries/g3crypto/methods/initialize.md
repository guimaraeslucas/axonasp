# Initialize Method

## Overview

Executes the Initialize operation provided by the G3CRYPTO library.

## Syntax

```asp
result = obj.Initialize(...)
```

## Parameters and Arguments

- algorithm (String, Optional): Preferred algorithm context or preset.
- options (String, Optional): Additional provider-specific initialization options.
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
result = obj.Initialize()
If IsObject(result) Then
    Response.Write "Object returned"
Else
    Response.Write CStr(result)
End If
Set obj = Nothing
%>
```



