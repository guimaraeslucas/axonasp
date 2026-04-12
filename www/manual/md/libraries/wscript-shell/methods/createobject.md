# CreateObject Method

## Overview

The CreateObject method is exposed by the WScript.Shell library object. Use it to execute this library operation from Classic ASP/VBScript with AxonASP runtime behavior.

## Syntax

```asp
result = obj.CreateObject(...)
`````

## Parameters and Arguments

- Parameters (Variant, Optional): This method accepts arguments according to the runtime dispatch of the WScript.Shell object.
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
Set obj = Server.CreateObject("WScript.Shell")
result = obj.CreateObject()
If IsObject(result) Then
    Response.Write "Object returned"
Else
    Response.Write CStr(result)
End If
Set obj = Nothing
%>
`````

