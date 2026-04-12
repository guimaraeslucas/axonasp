# PageCounter.Reset Method

## Overview
Calls the Reset member on the PageCounter compatibility object.

## Syntax
```asp
Dim obj
Set obj = Server.CreateObject("MSWC.PageCounter")
obj.Reset [path]
```

## Parameters and Arguments
- Parameters are validated by runtime dispatch for MSWC.PageCounter.
- Invalid argument count or incompatible values can raise runtime errors.

## Return Values
Returns a Variant-compatible value according to operation semantics.

## Remarks
- Resets current page or provided path counter.
- Use Set when assigning object return values.
- Member names are case-insensitive.

## Code Example
```asp
<%
Dim obj, result
Set obj = Server.CreateObject("MSWC.PageCounter")
result = Empty
On Error Resume Next
result = obj.Reset()
On Error GoTo 0
Response.Write CStr(result)
Set obj = Nothing
%>
```